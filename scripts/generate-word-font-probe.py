#!/usr/bin/env python3
"""Generate the first-party, single-glyph Word line-metric fallback probe.

The triangle and two-em advance are authored here. No external font bytes are
read. Paste stdout into packages/render/src/word-font-probe.ts after changes.
The font uses only bounded static TrueType tables and contains U+6C49.
"""
import struct, base64
p = lambda fmt, *v: struct.pack('>' + fmt, *v)
t = {}
t['head'] = p('IIIIHHQQhhhhHHhhh', 65536, 65536, 0, 1594834165, 3, 1000, 0, 0, 0, 0, 1900, 700, 0, 8, 2, 1, 0)
t['hhea'] = p('IhhhHhhhhhhhhhhhH', 65536, 800, -200, 0, 2000, 0, 100, 1900, 1, 0, 0, 0, 0, 0, 0, 0, 2)
t['maxp'] = p('IH' + 'H' * 13, 65536, 2, 3, 1, 0, 0, 2, 0, 0, 0, 0, 0, 0, 0, 0)
t['hmtx'] = p('HhHh', 1000, 0, 2000, 0)
g = p('hhhhhHH', 1, 0, 0, 1900, 700, 2, 0) + bytes([49, 33, 1]) + p('hhh', 1900, -1900, 700)
g += b'\x00' * (-len(g) % 4)
t['glyf'] = g
t['loca'] = p('III', 0, 0, len(g))
sub = p('HHHHHHH', 4, 32, 0, 4, 4, 1, 0) + p('HHHHHHHHH', 27721, 65535, 0, 27721, 65535, 1 - 27721 & 65535, 1, 0, 0)
t['cmap'] = p('HHHHI', 0, 1, 3, 1, 12) + sub
names = []
strings = b''
for (i, s) in [(1, 'Doffi Metric Probe'), (2, 'Regular'), (4, 'Doffi Metric Probe'), (6, 'DoffiMetricProbe')]:
    b = s.encode('utf-16-be')
    names.append(p('HHHHHH', 3, 1, 1033, i, len(b), len(strings)))
    strings += b
t['name'] = p('HHH', 0, len(names), 6 + 12 * len(names)) + b''.join(names) + strings
t['post'] = p('IIhhIIIII', 196608, 0, 0, 0, 0, 0, 0, 0, 0)
os = bytearray(78)
os[0:10] = p('HhHHH', 0, 1000, 400, 5, 0)
os[42:46] = p('I', 1 << 27)
os[58:62] = b'DOFF'
os[62:78] = p('HHHhhhHH', 64, 27721, 27721, 800, -200, 0, 800, 200)
t['OS/2'] = bytes(os)

def checksum(b):
    b += b'\x00' * (-len(b) % 4)
    return sum(struct.unpack('>' + 'I' * (len(b) // 4), b)) & 4294967295
n = len(t)
o = 12 + 16 * n
records = []
data = b''
offsets = {}
for (name, b) in sorted(t.items()):
    offsets[name] = o
    records.append(name.encode() + p('III', checksum(b), o, len(b)))
    b += b'\x00' * (-len(b) % 4)
    data += b
    o += len(b)
font = bytearray(p('IHHHH', 65536, n, 128, 3, n * 16 - 128) + b''.join(records) + data)
struct.pack_into('>I', font, offsets['head'] + 8, 2981146554 - checksum(font) & 4294967295)
print(base64.b64encode(font).decode())
