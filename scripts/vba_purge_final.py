"""
VBA Purge + Patch — final implementation.

Proven technique from NVISO security research:
- Strip p-code (PerformanceCache) from all module streams → Excel uses source code only
- Set MODULEOFFSET=0 in dir stream → tells Excel source starts at byte 0
- Zero _VBA_PROJECT stream body → removes p-code bytecache
- Patch modConfig + modHTTP source with new User-Agent

Mini-FAT streams (ThisWorkbook, Sheet1, dir stream, _VBA_PROJECT):
  - Read/write via root entry data area (64-byte mini-sectors)
Regular FAT streams (all large module streams):
  - Read/write via normal 512-byte sectors
"""
import zipfile, io, struct, olefile, math
from oletools.olevba import decompress_stream as ol_decomp

XLAM_BAK3 = '/home/user/workspace/sec-edgar-xbrl-excel-addin/dist/SEC_XBRL_Addin.xlam.bak3'
XLAM_OUT  = '/home/user/workspace/sec-edgar-xbrl-excel-addin/dist/SEC_XBRL_Addin.xlam'

OLD_UA  = b"SEC XBRL Excel Add-in sec-addin@github.io"
NEW_UA  = b"SEC-XBRL-Addin sec-xbrl-addin@outlook.com"
OLD_CMT = b'"SEC XBRL Excel Add-in sec-addin@github.io"'
NEW_CMT = b'"SEC-XBRL-Addin sec-xbrl-addin@outlook.com"'
assert len(OLD_UA)==len(NEW_UA)==41 and len(OLD_CMT)==len(NEW_CMT)==43

SECTOR_SIZE = 512
MINI_SECTOR_SIZE = 64

MODULE_TEXT_OFFSETS = {
    'ThisWorkbook': 1779, 'Sheet1': 835, 'modMain': 67618, 'modConfig': 8979,
    'modExcelWriter': 22054, 'modJSONParser': 21061, 'modHTTP': 17778,
    'modTickerLookup': 8434, 'modClassifier': 19142, 'modProgress': 6835,
    'modRibbon': 3919, 'JsonConverter': 46937,
}
MODULE_NAMES_ORDER = list(MODULE_TEXT_OFFSETS.keys())
# Discovered MODULEOFFSET record positions in decompressed dir stream:
MODULEOFFSET_DIRPOS = [590, 702, 820, 950, 1110, 1264, 1382, 1548, 1702, 1844, 1974, 2128]


def compress_vba(data: bytes) -> bytes:
    data = bytes(data)
    result = bytearray(b'\x01')
    dc, total = 0, len(data)
    while dc < total:
        dcs = dc; chunk_buf = bytearray()
        while dc < total and (dc - dcs) < 4096:
            fbi = len(chunk_buf); chunk_buf.append(0); flag = 0
            for bit in range(8):
                if dc >= total or (dc - dcs) >= 4096: break
                diff = dc - dcs
                if diff == 0:
                    chunk_buf.append(data[dc]); dc += 1; continue
                bc = max(4, int(math.ceil(math.log(diff, 2))))
                lm = 0xFFFF >> bc; om = (~lm) & 0xFFFF; ml = lm+3; mo = 1<<bc
                ws = max(0, dc-mo); limit = min(ml, total-dc); bl, bo = 0, 0
                for j in range(dc-1, ws-1, -1):
                    ml2 = 0
                    while ml2 < limit and data[j+ml2] == data[dc+ml2]: ml2 += 1
                    if ml2 > bl: bl = ml2; bo = dc-j
                    if bl >= limit: break
                if bl >= 3:
                    chunk_buf.extend(struct.pack('<H', ((bo-1)<<(16-bc)|(bl-3))&0xFFFF))
                    flag |= (1<<bit); dc += bl
                else:
                    chunk_buf.append(data[dc]); dc += 1
            chunk_buf[fbi] = flag
        if len(chunk_buf)+2 < 4098:
            result.extend(struct.pack('<H', 0xB000|(len(chunk_buf)-1)&0x0FFF))
            result.extend(chunk_buf)
        else:
            result.extend(struct.pack('<H', 0x3FFF))
            result.extend(data[dcs:dcs+4096].ljust(4096, b'\x00'))
    return bytes(result)


class OLEPatcher:
    """Handles reading/writing OLE streams in both regular and mini FAT."""
    
    def __init__(self, vba_bin_bytes: bytes):
        self.data = bytearray(vba_bin_bytes)
        self.ole = olefile.OleFileIO(io.BytesIO(vba_bin_bytes))
        self.fat = self.ole.fat
        # Force mini-FAT to load by opening any mini-stream first
        # (ole.minifat is None until a mini-stream is accessed)
        try:
            self.ole.openstream(['VBA', 'ThisWorkbook']).read()
        except Exception:
            try:
                self.ole.openstream(['VBA', 'dir']).read()
            except Exception:
                pass
        self.minifat = self.ole.minifat
        self._build_root_data()
    
    def _build_root_data(self):
        """Read root entry data (mini-stream container) and map it."""
        root = self.ole.root
        self.root_chain = self._get_regular_chain(root.isectStart)
        self.root_data = bytearray()
        for sec in self.root_chain:
            self.root_data.extend(self._read_regular_sector(sec))
    
    def _read_regular_sector(self, sec_num):
        off = SECTOR_SIZE + sec_num * SECTOR_SIZE
        return bytes(self.data[off:off+SECTOR_SIZE])
    
    def _get_regular_chain(self, start_sec):
        chain = [start_sec]
        cur = start_sec
        while True:
            nxt = self.fat[cur]
            if nxt >= 0xFFFFFFF0: break
            chain.append(nxt); cur = nxt
        return chain
    
    def _get_mini_chain(self, start_ms):
        chain = [start_ms]
        cur = start_ms
        while True:
            nxt = self.minifat[cur]
            if nxt >= 0xFFFFFFF0: break
            chain.append(nxt); cur = nxt
        return chain
    
    def is_mini(self, sid):
        entry = self.ole.direntries[sid]
        return entry.size < 4096 and len(self.minifat) > 0
    
    def get_chain_capacity(self, sid):
        entry = self.ole.direntries[sid]
        if self.is_mini(sid):
            chain = self._get_mini_chain(entry.isectStart)
            return len(chain) * MINI_SECTOR_SIZE
        else:
            chain = self._get_regular_chain(entry.isectStart)
            return len(chain) * SECTOR_SIZE
    
    def write_stream(self, sid, new_data: bytes):
        """Write new_data to stream sid. Data must fit in existing chain."""
        entry = self.ole.direntries[sid]
        
        if self.is_mini(sid):
            # Write to mini-FAT sectors within root data
            chain = self._get_mini_chain(entry.isectStart)
            cap = len(chain) * MINI_SECTOR_SIZE
            if len(new_data) > cap:
                raise ValueError(f"Data {len(new_data)} > mini cap {cap}")
            padded = bytearray(new_data) + b'\x00' * (cap - len(new_data))
            for i, ms in enumerate(chain):
                off = ms * MINI_SECTOR_SIZE
                self.root_data[off:off+MINI_SECTOR_SIZE] = padded[i*MINI_SECTOR_SIZE:(i+1)*MINI_SECTOR_SIZE]
            self._flush_root_data()
        else:
            # Write to regular FAT sectors
            chain = self._get_regular_chain(entry.isectStart)
            cap = len(chain) * SECTOR_SIZE
            if len(new_data) > cap:
                raise ValueError(f"Data {len(new_data)} > regular cap {cap}")
            padded = bytearray(new_data) + b'\x00' * (cap - len(new_data))
            for i, sec in enumerate(chain):
                off = SECTOR_SIZE + sec * SECTOR_SIZE
                self.data[off:off+SECTOR_SIZE] = padded[i*SECTOR_SIZE:(i+1)*SECTOR_SIZE]
    
    def _flush_root_data(self):
        """Write root data back to regular FAT sectors."""
        padded = bytearray(self.root_data)
        cap = len(self.root_chain) * SECTOR_SIZE
        padded.extend(b'\x00' * (cap - len(padded)))
        for i, sec in enumerate(self.root_chain):
            off = SECTOR_SIZE + sec * SECTOR_SIZE
            self.data[off:off+SECTOR_SIZE] = padded[i*SECTOR_SIZE:(i+1)*SECTOR_SIZE]
    
    def update_dir_entry_size(self, sid, new_size: int):
        """Update the size field in the directory entry for sid."""
        dirsect_start = struct.unpack_from('<I', bytes(self.data), 48)[0]
        dir_chain = [dirsect_start]
        cur = dirsect_start
        while True:
            nxt = self.fat[cur]
            if nxt >= 0xFFFFFFF0: break
            dir_chain.append(nxt); cur = nxt
        epp = SECTOR_SIZE // 128
        ts = dir_chain[sid // epp]
        eoff = SECTOR_SIZE + ts * SECTOR_SIZE + (sid % epp) * 128
        size_off = eoff + 120
        struct.pack_into('<I', self.data, size_off, new_size)
    
    def get_bytes(self):
        return bytes(self.data)


# ─────────────────────────────────────────────────────────────
print("=" * 60)
print("Loading bak3...")
with zipfile.ZipFile(XLAM_BAK3, 'r') as z:
    all_files = {n: z.read(n) for n in z.namelist()}
    namelist = z.namelist()

patcher = OLEPatcher(all_files['xl/vbaProject.bin'])
ole = patcher.ole


# ─────────────────────────────────────────────────────────────
# Step 1: Patch modConfig source
# ─────────────────────────────────────────────────────────────
print("\nStep 1: Patch modConfig source (OLD_UA → NEW_UA)...")
config_sid = ole._find(['VBA', 'modConfig'])
config_stream = ole.openstream(['VBA', 'modConfig']).read()
config_decomp = ol_decomp(config_stream[MODULE_TEXT_OFFSETS['modConfig']:])
ua_pos = config_decomp.find(OLD_UA)
print(f"  OLD_UA at {ua_pos}")
config_patched = bytearray(config_decomp); config_patched[ua_pos:ua_pos+41] = NEW_UA
config_new_comp = compress_vba(bytes(config_patched))
assert ol_decomp(config_new_comp) == bytes(config_patched), "modConfig roundtrip FAILED"
assert NEW_UA in ol_decomp(config_new_comp)
print(f"  ✓ {len(config_new_comp)} bytes")

print("Patch modHTTP source (comment)...")
http_sid = ole._find(['VBA', 'modHTTP'])
http_stream = ole.openstream(['VBA', 'modHTTP']).read()
http_decomp = ol_decomp(http_stream[MODULE_TEXT_OFFSETS['modHTTP']:])
cmt_pos = http_decomp.find(OLD_CMT)
print(f"  OLD_CMT at {cmt_pos}")
http_patched = bytearray(http_decomp); http_patched[cmt_pos:cmt_pos+43] = NEW_CMT
http_new_comp = compress_vba(bytes(http_patched))
assert ol_decomp(http_new_comp) == bytes(http_patched), "modHTTP roundtrip FAILED"
print(f"  ✓ {len(http_new_comp)} bytes")


# ─────────────────────────────────────────────────────────────
# Step 2: VBA Purge — strip p-code from large (regular-FAT) module streams
# ─────────────────────────────────────────────────────────────
print("\nStep 2: VBA Purge — strip p-code from regular-FAT module streams...")
for name, text_off in MODULE_TEXT_OFFSETS.items():
    sid = ole._find(['VBA', name])
    if sid is None:
        print(f"  {name:20s}: NOT FOUND, skip")
        continue
    if patcher.is_mini(sid):
        print(f"  {name:20s}: mini-FAT, skip (tiny stub module)")
        continue
    
    stream = ole.openstream(['VBA', name]).read()
    
    if name == 'modConfig':
        new_src = config_new_comp
    elif name == 'modHTTP':
        new_src = http_new_comp
    else:
        new_src = stream[text_off:]
    
    cap = patcher.get_chain_capacity(sid)
    if len(new_src) > cap:
        print(f"  {name:20s}: OVERFLOW {len(new_src)} > {cap}!")
        import sys; sys.exit(1)
    
    patcher.write_stream(sid, new_src)
    # CRITICAL: Do NOT update dir entry size for regular-FAT streams.
    # If we shrink the size below 4096, olefile reclassifies the stream as
    # mini-FAT and tries to read from wrong location. Keep original size
    # so the stream stays in regular-FAT territory. Excel uses MODULEOFFSET=0
    # and reads the source starting at byte 0; it doesn't rely on size.
    # patcher.update_dir_entry_size(sid, len(new_src))  ← intentionally omitted
    print(f"  {name:20s}: {len(stream)} → {len(new_src)} bytes written (size field kept at {len(stream)}, cap={cap})")


# ─────────────────────────────────────────────────────────────
# Step 3: Patch dir stream — all MODULEOFFSET → 0
# ─────────────────────────────────────────────────────────────
print("\nStep 3: Patch dir stream — set MODULEOFFSET records to 0...")
dir_sid = ole._find(['VBA', 'dir'])
dir_comp = ole.openstream(['VBA', 'dir']).read()
dir_decomp = bytearray(ol_decomp(dir_comp))

for i, dp in enumerate(MODULEOFFSET_DIRPOS):
    old_v = struct.unpack_from('<I', dir_decomp, dp+6)[0]
    struct.pack_into('<I', dir_decomp, dp+6, 0)
    print(f"  {MODULE_NAMES_ORDER[i]:20s}: {old_v} → 0")

dir_decomp = bytes(dir_decomp)
new_dir_comp = compress_vba(dir_decomp)
assert ol_decomp(new_dir_comp) == dir_decomp, "dir roundtrip FAILED"
print(f"  Dir: {len(dir_comp)} → {len(new_dir_comp)} bytes")

patcher.write_stream(dir_sid, new_dir_comp)
patcher.update_dir_entry_size(dir_sid, len(new_dir_comp))
print(f"  ✓ Dir stream written (mini-FAT={patcher.is_mini(dir_sid)})")


# ─────────────────────────────────────────────────────────────
# Step 4: Zero _VBA_PROJECT stream
# ─────────────────────────────────────────────────────────────
print("\nStep 4: Zero _VBA_PROJECT stream...")
try:
    vp_paths = [p for p in [['_VBA_PROJECT'], ['VBA', '_VBA_PROJECT']] 
                if any(str(p) in str(ole.listdir()) for _ in [1])]
    # Try to find it
    vp_stream = None
    vp_sid = None
    for path in [['_VBA_PROJECT'], ['VBA', '_VBA_PROJECT']]:
        try:
            vp_sid = ole._find(path)
            vp_stream = ole.openstream(path).read()
            break
        except:
            pass
    if vp_stream is not None:
        header_7 = vp_stream[:7]
        print(f"  _VBA_PROJECT: {len(vp_stream)} bytes, header={header_7.hex()}")
        patcher.write_stream(vp_sid, header_7)
        patcher.update_dir_entry_size(vp_sid, len(header_7))
        print(f"  ✓ Reduced to 7 bytes")
    else:
        print(f"  _VBA_PROJECT not found as standalone stream (may be in VBA/ subdir)")
        # Try common paths
        for path in ole.listdir():
            if '_VBA_PROJECT' in path[-1]:
                print(f"  Found at: {path}")
except Exception as e:
    print(f"  Error: {e}")


# ─────────────────────────────────────────────────────────────
# Step 5: Repack ZIP
# ─────────────────────────────────────────────────────────────
print("\nStep 5: Repacking XLAM...")
out_buf = io.BytesIO()
with zipfile.ZipFile(XLAM_BAK3, 'r') as zin:
    with zipfile.ZipFile(out_buf, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            d = patcher.get_bytes() if item.filename == 'xl/vbaProject.bin' else zin.read(item.filename)
            zout.writestr(item, d)
xlam_bytes = out_buf.getvalue()
with open(XLAM_OUT, 'wb') as f:
    f.write(xlam_bytes)
print(f"  Written {XLAM_OUT} ({len(xlam_bytes)} bytes)")


# ─────────────────────────────────────────────────────────────
# Final verification
# ─────────────────────────────────────────────────────────────
print("\n" + "=" * 60)
print("Final verification...")
with zipfile.ZipFile(XLAM_OUT, 'r') as z:
    vba_c = z.read('xl/vbaProject.bin')
ole_c = olefile.OleFileIO(io.BytesIO(vba_c))

# Check MODULEOFFSET values
dir_c_comp = ole_c.openstream(['VBA', 'dir']).read()
dir_c_decomp = bytearray(ol_decomp(dir_c_comp))
offset_vals = [struct.unpack_from('<I', dir_c_decomp, dp+6)[0] for dp in MODULEOFFSET_DIRPOS]
print(f"  MODULEOFFSET values: {offset_vals}")
print(f"  All zero: {'✓' if all(v==0 for v in offset_vals) else '✗'}")

# Check modConfig — after purge, stream IS the compressed source (offset=0)
sc = ole_c.openstream(['VBA', 'modConfig']).read()
dc = ol_decomp(sc)  # whole stream is now just CompressedSourceCode
print(f"  modConfig stream: {len(sc)} bytes (purged — whole stream is source)")
print(f"  modConfig NEW_UA: {'✓' if NEW_UA in dc else '✗'}")
print(f"  modConfig OLD_UA: {'✗ (good)' if OLD_UA not in dc else '✓ BAD'}")

# Check modHTTP — same
sh = ole_c.openstream(['VBA', 'modHTTP']).read()
dh = ol_decomp(sh)
print(f"  modHTTP NEW_CMT:  {'✓' if NEW_CMT in dh else '✗'}")
print(f"  modHTTP OLD_CMT:  {'✗ (good)' if OLD_CMT not in dh else '✓ BAD'}")
print(f"  modHTTP NEW_UA:   {'✓' if NEW_UA in dh else '✗'}")
print(f"  modHTTP OLD_UA:   {'✗ (good)' if OLD_UA not in dh else '✓ BAD'}")

# oletools macro extraction
from oletools.olevba import VBA_Parser
vba_parser = VBA_Parser('check.xlam', data=open(XLAM_OUT,'rb').read())
if vba_parser.detect_vba_macros():
    mods = list(vba_parser.extract_macros())
    print(f"\n  oletools: {len(mods)} modules")
    all_clean = True
    for (_, _, name, code) in mods:
        old = 'sec-addin@github.io' in code
        if old: all_clean = False
        new_u = 'sec-xbrl-addin@outlook.com' in code
        n = len(code.splitlines())
        print(f"    {name:20s}: {n:4d} lines | NEW={new_u} OLD={'BAD' if old else 'ok'}")
    print(f"  All clean: {'✓' if all_clean else '✗'}")

print("\n✓ DONE")
