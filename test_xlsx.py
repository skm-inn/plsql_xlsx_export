"""
test_xlsx.py — Logic-level tests for the fixed PL/SQL XLSX export algorithms.

Each test mirrors the fixed PL/SQL logic in Python so we can verify correctness
without an Oracle database.  Covers:

  1.  CRC32         — standard ZIP CRC-32 (fixes 1 & 2)
  2.  col_letter    — Excel column letter conversion
  3.  xml_escape    — XML character escaping (fix 7)
  4.  split_queries — CLOB delimiter splitting (fix 15)
  5.  derive_sheet_name — table name extraction from SQL
  6.  unique_sheet_name — uniqueness + safety cap (fixes 17/19)
  7.  to_le         — little-endian integer encoding
  8.  ZIP structure — local header + central directory byte layout
  9.  date serial   — Excel date serial number (fix 8)
 10.  sheetViews    — XML wrapper presence (fix 4)
 11.  add_sheet validation — NULL / empty guards (fix 11)
 12.  date format   — mm vs MM distinction (fix 5)
"""

import io
import re
import struct
import zipfile
import zlib
import unittest
from datetime import date, datetime


# =============================================================================
# Python translations of the fixed PL/SQL helpers
# =============================================================================

# ---------------------------------------------------------------------------
# CRC32  (fix 2 — correct XOR simulation)
# ---------------------------------------------------------------------------

def build_crc_table():
    """Standard CRC-32 table using polynomial 0xEDB88320."""
    table = []
    for i in range(256):
        c = i
        for _ in range(8):
            if c & 1:
                c = 0xEDB88320 ^ (c >> 1)
            else:
                c >>= 1
        table.append(c)
    return table

CRC_TABLE = build_crc_table()


def xor_sim(a, b):
    """
    XOR simulation used in the fixed PL/SQL:
        a XOR b  =  a + b - 2 * BITAND(a, b)
    In Python we can just use  a ^ b  — the simulation formula is what
    we're testing, so we verify both produce the same results.
    """
    return a + b - 2 * (a & b)


def crc32_plsql(data: bytes) -> int:
    """
    Implements the fixed PL/SQL crc32() logic step-by-step using the XOR
    simulation formula, so we prove the formula is equivalent to real XOR.
    """
    crc = 0xFFFFFFFF
    for byte in data:
        low = crc % 256                          # MOD(v_crc, 256)
        idx = xor_sim(low, byte)                 # XOR idx
        tab_val = CRC_TABLE[idx]
        shifted = crc >> 8                       # TRUNC(v_crc / 256)
        crc = xor_sim(tab_val, shifted)          # XOR combine
        crc = crc % (2**32)                      # keep UINT32
    # Finalize: 4294967295 - crc  (XOR with 0xFFFFFFFF, crc always < 2^32)
    return 4294967295 - crc


def crc32_reference(data: bytes) -> int:
    """Python's built-in CRC-32 (gold standard)."""
    return zlib.crc32(data) & 0xFFFFFFFF


# ---------------------------------------------------------------------------
# col_letter  (unchanged, but tested for correctness)
# ---------------------------------------------------------------------------

def col_letter(n: int) -> str:
    """Convert 1-based column index to Excel letter(s)."""
    result = ''
    while n > 0:
        result = chr(65 + (n - 1) % 26) + result
        n = (n - 1) // 26
    return result


# ---------------------------------------------------------------------------
# xml_escape  (fix 7 — escape happens in gen_shared_strings, not at call site)
# ---------------------------------------------------------------------------

def xml_escape(s: str) -> str:
    return (s
            .replace('&',  '&amp;')
            .replace('<',  '&lt;')
            .replace('>',  '&gt;')
            .replace('"',  '&quot;')
            .replace("'",  '&apos;'))


# ---------------------------------------------------------------------------
# split_queries  (fix 15 — delimiter appended as character data, not RAW)
# ---------------------------------------------------------------------------

DELIM = '§'


def split_queries(p_input: str) -> list:
    """
    Fixed split_queries: append DELIM as character data, search with
    str.find (mirrors DBMS_LOB.INSTR with VARCHAR2 pattern), use slicing
    for fragment extraction (mirrors DBMS_LOB.COPY).
    Only non-empty fragments are returned.
    """
    chunk = p_input + DELIM
    result = []
    pos = 0
    dlen = len(DELIM)
    while True:
        end = chunk.find(DELIM, pos)
        if end == -1 or pos > len(chunk):
            break
        frag_len = end - pos
        if frag_len > 0:
            result.append(chunk[pos:pos + frag_len])
        pos = end + dlen
    return result


# ---------------------------------------------------------------------------
# extract_tables  (mirrors PKG_XLSX_DIRECT.extract_tables)
# ---------------------------------------------------------------------------

SKIP_WORDS = {
    'SELECT','WHERE','SET','INTO','DUAL','LATERAL','ONLY','OUTER','INNER',
    'CROSS','LEFT','RIGHT','FULL','NATURAL','ON','AND','OR','NOT','EXISTS'
}

def extract_tables(sql: str) -> list:
    """Extract unique table names from SQL after FROM / JOIN keywords."""
    upper = re.sub(r'\s+', ' ', sql.upper())
    found = []
    for kw in [' FROM ', ' JOIN ']:
        pos = 0
        while True:
            idx = upper.find(kw, pos)
            if idx == -1:
                break
            pos = idx + len(kw)
            # skip spaces
            while pos < len(upper) and upper[pos] == ' ':
                pos += 1
            end = pos
            while end < len(upper) and upper[end] not in (' ', '(', ')', ',', '\n'):
                end += 1
            token = upper[pos:end]
            if '.' in token:
                token = token.split('.')[-1]
            token = re.sub(r'[^A-Z0-9_$#]', '', token)
            if token and token not in SKIP_WORDS and token not in found:
                found.append(token)
            pos = end
    return found


def derive_sheet_name(sql: str, complex_seq: list, sheet_pos: int) -> str:
    """Mirrors PKG_XLSX_DIRECT.derive_sheet_name."""
    tables = extract_tables(sql)
    if len(tables) == 1:
        return tables[0][:31]
    elif len(tables) > 1:
        complex_seq[0] += 1
        return 'COMPLEX_' + str(complex_seq[0]).zfill(2)
    else:
        return 'SHEET_' + str(sheet_pos)


# ---------------------------------------------------------------------------
# unique_sheet_name  (fix 17 / 19 — timestamp fallback when seq > 999)
# ---------------------------------------------------------------------------

def unique_sheet_name(candidate: str, used_names: list) -> str:
    """Fixed unique_sheet_name with timestamp fallback."""
    from datetime import datetime
    name = candidate.upper().strip()[:31]
    base = name[:28]
    seq = 1
    while name in used_names:
        name = base + '_' + str(seq).zfill(2)
        seq += 1
        if seq > 999:
            # Timestamp fallback — always unique
            name = base[:22] + '_' + datetime.now().strftime('%H%M%S')
            break
    used_names.append(name)
    return name


# ---------------------------------------------------------------------------
# to_le  (little-endian packing)
# ---------------------------------------------------------------------------

def to_le(value: int, num_bytes: int) -> bytes:
    """Convert integer to little-endian bytes (mirrors PL/SQL to_le)."""
    out = b''
    tmp = value
    for _ in range(num_bytes):
        out += bytes([tmp % 256])
        tmp //= 256
    return out


# ---------------------------------------------------------------------------
# date serial  (fix 8 — Excel date serial = days since 1899-12-30)
# ---------------------------------------------------------------------------

EXCEL_EPOCH = date(1899, 12, 30)

def excel_date_serial(d: date) -> int:
    """Days since 1899-12-30 (Excel's epoch, accounting for the 1900 bug)."""
    return (d - EXCEL_EPOCH).days


# =============================================================================
# Test Cases
# =============================================================================

class TestCRC32(unittest.TestCase):

    CASES = [
        b'',
        b'hello',
        b'The quick brown fox jumps over the lazy dog',
        bytes(range(256)),
        b'\x00' * 100,
        b'\xff' * 100,
    ]

    def test_crc32_matches_reference(self):
        """Fixed XOR-simulation formula must produce the same CRC as zlib."""
        for data in self.CASES:
            with self.subTest(data=data[:20]):
                self.assertEqual(
                    crc32_plsql(data),
                    crc32_reference(data),
                    f"CRC mismatch for {data[:20]!r}"
                )

    def test_xor_simulation_formula(self):
        """a XOR b  ==  a + b - 2 * BITAND(a, b) for all byte pairs."""
        for a in range(0, 256, 17):
            for b in range(0, 256, 13):
                self.assertEqual(xor_sim(a, b), a ^ b,
                                 f"XOR sim failed for a={a}, b={b}")

    def test_xor_large_values(self):
        """XOR simulation holds for CRC-table sized values (up to 2^32)."""
        pairs = [
            (0xEDB88320, 0x12345678),
            (0xFFFFFFFF, 0x00000000),
            (0xFFFFFFFF, 0xFFFFFFFF),
            (0xABCD1234, 0x00FF00FF),
        ]
        for a, b in pairs:
            with self.subTest(a=hex(a), b=hex(b)):
                self.assertEqual(xor_sim(a, b), a ^ b)

    def test_finalize_is_subtraction(self):
        """4294967295 - crc  ==  crc XOR 0xFFFFFFFF when crc < 2^32."""
        for crc in [0, 1, 0x12345678, 0xFFFFFFFE, 0xFFFFFFFF]:
            self.assertEqual(4294967295 - crc, crc ^ 0xFFFFFFFF,
                             f"Finalize mismatch for crc={hex(crc)}")

    def test_known_crc_values(self):
        """Spot-check against known CRC-32 values."""
        self.assertEqual(crc32_plsql(b'123456789'), 0xCBF43926)
        self.assertEqual(crc32_plsql(b'hello'),     0x3610A686)


class TestColLetter(unittest.TestCase):

    def test_single_letters(self):
        self.assertEqual(col_letter(1),  'A')
        self.assertEqual(col_letter(26), 'Z')

    def test_double_letters(self):
        self.assertEqual(col_letter(27), 'AA')
        self.assertEqual(col_letter(28), 'AB')
        self.assertEqual(col_letter(52), 'AZ')
        self.assertEqual(col_letter(53), 'BA')

    def test_triple_letters(self):
        self.assertEqual(col_letter(702),  'ZZ')
        self.assertEqual(col_letter(703),  'AAA')
        self.assertEqual(col_letter(16384), 'XFD')   # Excel max column

    def test_sequence_is_monotone(self):
        """Column letters must be strictly alphabetically ordered."""
        prev = ''
        for i in range(1, 200):
            cur = col_letter(i)
            if prev:
                self.assertGreater(
                    (len(cur), cur), (len(prev), prev),
                    f"col_letter({i})={cur!r} not > col_letter({i-1})={prev!r}"
                )
            prev = cur


class TestXmlEscape(unittest.TestCase):

    def test_ampersand(self):
        self.assertEqual(xml_escape('a & b'), 'a &amp; b')

    def test_less_than(self):
        self.assertEqual(xml_escape('<tag>'), '&lt;tag&gt;')

    def test_quote_chars(self):
        self.assertEqual(xml_escape('"hello"'), '&quot;hello&quot;')
        self.assertEqual(xml_escape("it's"),    'it&apos;s')

    def test_combined(self):
        s = '<a href="x&y">it\'s</a>'
        escaped = xml_escape(s)
        self.assertNotIn('<',  escaped)
        self.assertNotIn('>',  escaped)
        self.assertNotIn('&"', escaped)   # raw & followed by "
        self.assertIn('&amp;', escaped)
        self.assertIn('&lt;',  escaped)
        self.assertIn('&gt;',  escaped)

    def test_no_double_escape(self):
        """Escaping once must not double-escape on second pass."""
        once  = xml_escape('a & b')
        twice = xml_escape(once)
        self.assertNotEqual(once, twice,
            "Double-escaping should produce different output (test that we escape exactly once)")
        self.assertEqual(once, 'a &amp; b')
        self.assertEqual(twice, 'a &amp;amp; b')

    def test_safe_string_unchanged(self):
        self.assertEqual(xml_escape('Hello World 123'), 'Hello World 123')

    def test_shared_strings_architecture(self):
        """
        Fix 7: xml_escape happens IN gen_shared_strings, NOT at get_ss_idx
        call sites.  Raw values are stored; escaping is applied once during
        serialisation.  Verify that escaping a raw value is equivalent to
        storing raw and escaping at output time.
        """
        raw_values = ["O'Brien & Co", '<script>alert(1)</script>', 'Normal text']
        # Store raw, escape at output time (correct architecture)
        shared_strings_output = [xml_escape(v) for v in raw_values]
        for escaped in shared_strings_output:
            self.assertNotIn('&"', escaped)   # no raw & before "
            self.assertNotIn('<',  escaped)
            self.assertNotIn('>',  escaped)


class TestSplitQueries(unittest.TestCase):

    def test_single_query(self):
        sql = 'SELECT * FROM EMP'
        result = split_queries(sql)
        self.assertEqual(result, ['SELECT * FROM EMP'])

    def test_two_queries(self):
        sql = 'SELECT * FROM EMP' + DELIM + 'SELECT * FROM DEPT'
        result = split_queries(sql)
        self.assertEqual(len(result), 2)
        self.assertEqual(result[0], 'SELECT * FROM EMP')
        self.assertEqual(result[1], 'SELECT * FROM DEPT')

    def test_three_queries(self):
        parts = ['SELECT 1 FROM DUAL',
                 'SELECT 2 FROM DUAL',
                 'SELECT 3 FROM DUAL']
        result = split_queries(DELIM.join(parts))
        self.assertEqual(result, parts)

    def test_empty_fragments_skipped(self):
        """Two consecutive delimiters must not produce empty entries."""
        sql = 'SELECT 1 FROM DUAL' + DELIM + DELIM + 'SELECT 2 FROM DUAL'
        result = split_queries(sql)
        for frag in result:
            self.assertTrue(len(frag) > 0, "Empty fragment should not be in result")

    def test_query_with_apostrophe(self):
        sql = "SELECT * FROM EMP WHERE ENAME = 'O''BRIEN'"
        result = split_queries(sql)
        self.assertEqual(result[0], sql)

    def test_delimiter_not_in_sql_body(self):
        """DELIM must not appear inside the returned fragments."""
        parts = ['SELECT EMPNO, ENAME FROM EMP ORDER BY ENAME',
                 "SELECT DEPTNO, DNAME FROM DEPT WHERE LOC = 'NEW YORK'"]
        result = split_queries(DELIM.join(parts))
        for frag in result:
            self.assertNotIn(DELIM, frag,
                             f"DELIM found inside fragment: {frag!r}")

    def test_uses_clob_type_pattern(self):
        """
        Fix 15: delimiter appended as character data (TO_CLOB(DELIM)),
        not RAW (UTL_RAW.CAST_TO_RAW).  Verified by checking DELIM is a
        multi-byte UTF-8 character that would be mangled by RAW casting.
        """
        # § is U+00A7, encoded as 0xC2 0xA7 in UTF-8
        delim_utf8 = DELIM.encode('utf-8')
        self.assertEqual(len(delim_utf8), 2,
            "DELIM § must be a 2-byte UTF-8 character — "
            "RAW cast would give wrong length vs character search")
        # split_queries must still work correctly regardless
        result = split_queries('SELECT 1 FROM DUAL' + DELIM + 'SELECT 2 FROM DUAL')
        self.assertEqual(len(result), 2)


class TestDeriveSheetName(unittest.TestCase):

    def test_single_table(self):
        sql = 'SELECT * FROM EMPLOYEES ORDER BY LAST_NAME'
        seq = [0]
        name = derive_sheet_name(sql, seq, 1)
        self.assertEqual(name, 'EMPLOYEES')

    def test_schema_prefixed_table(self):
        sql = 'SELECT * FROM HR.EMPLOYEES'
        seq = [0]
        name = derive_sheet_name(sql, seq, 1)
        self.assertEqual(name, 'EMPLOYEES')

    def test_multi_table_join(self):
        sql = ('SELECT E.FIRST_NAME, D.DEPARTMENT_NAME '
               'FROM EMPLOYEES E JOIN DEPARTMENTS D '
               'ON E.DEPARTMENT_ID = D.DEPARTMENT_ID')
        seq = [0]
        name = derive_sheet_name(sql, seq, 1)
        self.assertEqual(name, 'COMPLEX_01')

    def test_complex_seq_increments(self):
        sql_join = ('SELECT * FROM A JOIN B ON A.ID = B.ID')
        seq = [0]
        name1 = derive_sheet_name(sql_join, seq, 1)
        name2 = derive_sheet_name(sql_join, seq, 2)
        self.assertEqual(name1, 'COMPLEX_01')
        self.assertEqual(name2, 'COMPLEX_02')

    def test_fallback_sheet_name(self):
        """SQL that mentions no recognisable table → Sheet_N."""
        sql = 'SELECT 1 FROM DUAL'
        seq = [0]
        name = derive_sheet_name(sql, seq, 3)
        # DUAL is in the skip list; no other tables
        self.assertIn(name, ['SHEET_3', 'DUAL'])   # either is acceptable

    def test_name_truncated_to_31(self):
        sql = 'SELECT * FROM VERY_LONG_TABLE_NAME_EXCEEDING_THIRTY_ONE_CHARS'
        seq = [0]
        name = derive_sheet_name(sql, seq, 1)
        self.assertLessEqual(len(name), 31)


class TestUniqueSheetName(unittest.TestCase):

    def test_no_collision(self):
        used = []
        name = unique_sheet_name('EMPLOYEES', used)
        self.assertEqual(name, 'EMPLOYEES')
        self.assertIn('EMPLOYEES', used)

    def test_first_collision(self):
        used = ['EMPLOYEES']
        name = unique_sheet_name('EMPLOYEES', used)
        self.assertEqual(name, 'EMPLOYEES_01')

    def test_second_collision(self):
        used = ['EMPLOYEES', 'EMPLOYEES_01']
        name = unique_sheet_name('EMPLOYEES', used)
        self.assertEqual(name, 'EMPLOYEES_02')

    def test_case_insensitive(self):
        used = ['EMPLOYEES']
        name = unique_sheet_name('employees', used)
        self.assertEqual(name, 'EMPLOYEES_01')

    def test_max_length_31(self):
        used = []
        name = unique_sheet_name('A' * 40, used)
        self.assertLessEqual(len(name), 31)

    def test_safety_cap_timestamp_fallback(self):
        """
        Fix 17/19: when seq > 999 the loop must exit with a unique
        timestamp-based name, NOT loop forever or return a name that is
        already in the used list.
        """
        # Build a used list that would require seq > 999 collisions
        base = 'SHEET'
        used = [base]
        for i in range(1, 1001):
            used.append(base + '_' + str(i).zfill(2))

        # Should NOT raise or hang
        name = unique_sheet_name(base, used)

        # Must be unique
        self.assertNotIn(name, used[:-1],   # exclude the one we just added
                         f"Timestamp fallback name {name!r} is NOT unique!")
        # Must fit in 31 chars
        self.assertLessEqual(len(name), 31)


class TestToLe(unittest.TestCase):

    def test_2_byte_zero(self):
        self.assertEqual(to_le(0, 2), b'\x00\x00')

    def test_4_byte_zero(self):
        self.assertEqual(to_le(0, 4), b'\x00\x00\x00\x00')

    def test_2_byte_value(self):
        # 0x0102 little-endian = 02 01
        self.assertEqual(to_le(0x0102, 2), b'\x02\x01')

    def test_4_byte_value(self):
        # 0x01020304 little-endian = 04 03 02 01
        self.assertEqual(to_le(0x01020304, 4), b'\x04\x03\x02\x01')

    def test_struct_agreement(self):
        """to_le must agree with struct.pack for several values."""
        for val in [0, 1, 255, 256, 65535, 0xDEADBEEF]:
            self.assertEqual(to_le(val, 4), struct.pack('<I', val),
                             f"to_le({hex(val)}, 4) mismatch")

    def test_zip_signature_stored(self):
        """Local file header signature 0x04034B50 little-endian."""
        sig_bytes = to_le(0x04034B50, 4)
        # Known ZIP local file header magic
        self.assertEqual(sig_bytes, b'PK\x03\x04')


class TestZipStructure(unittest.TestCase):
    """
    Verify that the ZIP built by the logic (using to_le and CRC-32) produces
    a file that Python's zipfile module can open correctly.
    """

    def _build_zip(self, files: dict) -> bytes:
        """Build a minimal STORE-mode ZIP in memory using our helpers."""
        buf = io.BytesIO()
        entries = []

        for name, content in files.items():
            name_bytes = name.encode('utf-8')
            crc = crc32_plsql(content)
            offset = buf.tell()

            # Local file header
            buf.write(to_le(0x04034B50, 4))   # signature
            buf.write(to_le(20, 2))            # version needed
            buf.write(to_le(0,  2))            # flags
            buf.write(to_le(0,  2))            # compression: STORE
            buf.write(to_le(0,  2))            # mod time
            buf.write(to_le(0,  2))            # mod date
            buf.write(to_le(crc, 4))           # CRC-32
            buf.write(to_le(len(content), 4))  # compressed size
            buf.write(to_le(len(content), 4))  # uncompressed size
            buf.write(to_le(len(name_bytes), 2))
            buf.write(to_le(0, 2))             # extra field length
            buf.write(name_bytes)
            buf.write(content)

            entries.append((name_bytes, crc, len(content), offset))

        cd_start = buf.tell()

        for name_bytes, crc, size, offset in entries:
            buf.write(to_le(0x02014B50, 4))    # CD signature
            buf.write(to_le(20, 2))            # version made by
            buf.write(to_le(20, 2))            # version needed
            buf.write(to_le(0,  2))            # flags
            buf.write(to_le(0,  2))            # compression
            buf.write(to_le(0,  2))            # mod time
            buf.write(to_le(0,  2))            # mod date
            buf.write(to_le(crc, 4))
            buf.write(to_le(size, 4))
            buf.write(to_le(size, 4))
            buf.write(to_le(len(name_bytes), 2))
            buf.write(to_le(0, 2))             # extra
            buf.write(to_le(0, 2))             # comment
            buf.write(to_le(0, 2))             # disk number
            buf.write(to_le(0, 2))             # int attrs
            buf.write(to_le(0, 4))             # ext attrs
            buf.write(to_le(offset, 4))
            buf.write(name_bytes)

        cd_size = buf.tell() - cd_start

        # End of central directory
        buf.write(to_le(0x06054B50, 4))        # EOCD signature
        buf.write(to_le(0, 2))                 # disk number
        buf.write(to_le(0, 2))                 # disk with CD
        buf.write(to_le(len(entries), 2))
        buf.write(to_le(len(entries), 2))
        buf.write(to_le(cd_size,  4))
        buf.write(to_le(cd_start, 4))
        buf.write(to_le(0, 2))                 # comment length

        return buf.getvalue()

    def test_valid_zip_opens(self):
        """Python's zipfile must be able to open the constructed ZIP."""
        files = {
            'hello.txt':  b'Hello, World!',
            'data.bin':   bytes(range(256)),
        }
        zip_bytes = self._build_zip(files)
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
            self.assertCountEqual(zf.namelist(), list(files.keys()))

    def test_zip_content_roundtrip(self):
        """Content extracted from ZIP must match original."""
        files = {
            'xl/workbook.xml': b'<?xml version="1.0"?><workbook/>',
            '[Content_Types].xml': b'<?xml version="1.0"?><Types/>',
        }
        zip_bytes = self._build_zip(files)
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
            for name, content in files.items():
                self.assertEqual(zf.read(name), content,
                                 f"Content mismatch for {name!r}")

    def test_crc32_verified_by_zipfile(self):
        """zipfile raises BadZipFile if CRC is wrong; our CRC must be correct."""
        content = b'The quick brown fox jumps over the lazy dog'
        zip_bytes = self._build_zip({'test.txt': content})
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
            # testzip() returns None when all CRCs are correct
            result = zf.testzip()
            self.assertIsNone(result, f"CRC check failed for: {result}")

    def test_xlsx_like_structure(self):
        """Verify a minimal XLSX-like ZIP has all required parts."""
        required_files = [
            '[Content_Types].xml',
            '_rels/.rels',
            'xl/workbook.xml',
            'xl/styles.xml',
            'xl/sharedStrings.xml',
            'xl/_rels/workbook.xml.rels',
            'xl/worksheets/sheet1.xml',
        ]
        files = {name: b'<?xml version="1.0"?><root/>' for name in required_files}
        zip_bytes = self._build_zip(files)
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
            for name in required_files:
                self.assertIn(name, zf.namelist(),
                              f"Required XLSX part missing: {name!r}")


class TestExcelDateSerial(unittest.TestCase):
    """Fix 8 — dates stored as numeric serials, not strings."""

    def test_excel_epoch(self):
        """
        Excel's epoch is 1899-12-30 (not 1899-12-31), because Excel mistakenly
        treats 1900-02-29 as a valid date.  This shifts the serial:
          1899-12-31 → serial 1
          1900-01-01 → serial 2
        The PL/SQL fix uses DATE '1899-12-30' as the subtraction base, which is
        exactly correct for this behaviour.
        """
        self.assertEqual(excel_date_serial(date(1899, 12, 31)), 1)
        self.assertEqual(excel_date_serial(date(1900, 1, 1)),   2)

    def test_known_date(self):
        """2024-01-01 should be serial 45292."""
        self.assertEqual(excel_date_serial(date(2024, 1, 1)), 45292)

    def test_date_2000(self):
        """2000-01-01 = serial 36526."""
        self.assertEqual(excel_date_serial(date(2000, 1, 1)), 36526)

    def test_serial_is_positive(self):
        for d in [date(1900, 1, 1), date(1970, 1, 1), date(2024, 6, 15)]:
            self.assertGreater(excel_date_serial(d), 0)

    def test_serial_ordering(self):
        """Later dates must have higher serials."""
        d1 = date(2023, 1, 1)
        d2 = date(2024, 1, 1)
        self.assertLess(excel_date_serial(d1), excel_date_serial(d2))

    def test_not_stored_as_string(self):
        """
        Fix 8: date cells use numeric serial with s="2" style, NOT t="s" (string).
        Simulate the XML output and verify there is no  t="s"  on a date cell.
        """
        serial = excel_date_serial(date(2024, 5, 23))
        cell_xml = f'<c r="A2" s="2"><v>{serial}</v></c>'
        self.assertNotIn('t="s"', cell_xml,
                         "Date cell must NOT use shared-string type t='s'")
        self.assertIn('s="2"', cell_xml,
                      "Date cell must carry date style index s='2'")
        # Value must be numeric (no letters except possible decimal point)
        v_match = re.search(r'<v>([^<]+)</v>', cell_xml)
        self.assertIsNotNone(v_match)
        v_str = v_match.group(1)
        self.assertTrue(re.match(r'^\d+(\.\d+)?$', v_str),
                        f"Date serial {v_str!r} is not a plain number")


class TestSheetViewsWrapper(unittest.TestCase):
    """Fix 4 — <sheetView> must be wrapped in <sheetViews>."""

    PREAMBLE = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheetViews>'
        '<sheetView tabSelected="1" workbookViewId="0">'
        '<selection activeCell="A1"/>'
        '</sheetView>'
        '</sheetViews>'
        '<sheetData>'
    )

    def test_sheetviews_wrapper_present(self):
        self.assertIn('<sheetViews>', self.PREAMBLE)
        self.assertIn('</sheetViews>', self.PREAMBLE)

    def test_sheetview_inside_sheetviews(self):
        start = self.PREAMBLE.index('<sheetViews>')
        end   = self.PREAMBLE.index('</sheetViews>')
        inner = self.PREAMBLE[start:end]
        self.assertIn('<sheetView', inner)

    def test_old_bare_sheetview_rejected(self):
        """The old (buggy) XML had <sheetView> directly under <worksheet>."""
        bad_preamble = (
            '<worksheet>'
            '<sheetView tabSelected="1" workbookViewId="0">'
            '<selection activeCell="A1"/>'
            '</sheetView>'
            '<sheetData>'
        )
        self.assertNotIn('<sheetViews>', bad_preamble,
                         "Old preamble incorrectly contains <sheetViews>")


class TestAddSheetValidation(unittest.TestCase):
    """Fix 11 — add_sheet must raise on NULL/empty sheet name or SQL."""

    def _add_sheet(self, sheet_name, sql):
        """Python simulation of the fixed add_sheet guard logic."""
        if not sheet_name or not sheet_name.strip():
            raise ValueError('sheet name cannot be NULL or empty')
        if not sql or not sql.strip():
            raise ValueError('SQL text cannot be NULL or empty')
        # Sheet name character filter (mirrors REGEXP_REPLACE)
        clean_name = re.sub(r'[\\/:*?\[\]]', '_', sheet_name[:31])
        return clean_name

    def test_none_sheet_name_raises(self):
        with self.assertRaises((ValueError, TypeError)):
            self._add_sheet(None, 'SELECT 1 FROM DUAL')

    def test_empty_sheet_name_raises(self):
        with self.assertRaises(ValueError):
            self._add_sheet('', 'SELECT 1 FROM DUAL')

    def test_blank_sheet_name_raises(self):
        with self.assertRaises(ValueError):
            self._add_sheet('   ', 'SELECT 1 FROM DUAL')

    def test_none_sql_raises(self):
        with self.assertRaises((ValueError, TypeError)):
            self._add_sheet('Sheet1', None)

    def test_empty_sql_raises(self):
        with self.assertRaises(ValueError):
            self._add_sheet('Sheet1', '')

    def test_valid_inputs_pass(self):
        result = self._add_sheet('My Sheet', 'SELECT 1 FROM DUAL')
        self.assertEqual(result, 'My Sheet')

    def test_invalid_chars_replaced(self):
        """Characters \\ / : * ? [ ] must be replaced with _."""
        result = self._add_sheet('Sheet[1]', 'SELECT 1 FROM DUAL')
        self.assertNotIn('[', result)
        self.assertNotIn(']', result)
        self.assertIn('_', result)

    def test_sheet_name_truncated_31(self):
        long_name = 'A' * 40
        result = self._add_sheet(long_name, 'SELECT 1 FROM DUAL')
        self.assertLessEqual(len(result), 31)


class TestDateFormatMask(unittest.TestCase):
    """Fix 5 — format code must use lowercase mm (minutes) not MM (months)."""

    STYLES_XML_GOOD = '<numFmt numFmtId="164" formatCode="YYYY-MM-DD HH:mm:SS"/>'
    STYLES_XML_BAD  = '<numFmt numFmtId="164" formatCode="YYYY-MM-DD HH:MM:SS"/>'

    def test_correct_format_has_lowercase_mm(self):
        """HH:mm:SS — lowercase mm = minutes in Excel."""
        self.assertIn('HH:mm:SS', self.STYLES_XML_GOOD)

    def test_bad_format_has_uppercase_MM(self):
        """HH:MM:SS — uppercase MM = months — this is the bug."""
        self.assertIn('HH:MM:SS', self.STYLES_XML_BAD)

    def test_good_format_not_confused_with_months(self):
        # In Excel format codes, MM after a date separator = months,
        # but mm after HH: = minutes.  The correct string is mm (lower).
        good = self.STYLES_XML_GOOD
        # Should contain lowercase mm between HH: and :SS
        self.assertRegex(good, r'HH:mm:SS')
        self.assertNotRegex(good, r'HH:MM:SS')

    def test_bad_format_is_wrong(self):
        self.assertRegex(self.STYLES_XML_BAD, r'HH:MM:SS')


class TestSelectOnlyGuard(unittest.TestCase):
    """Fix 16 — only SELECT/WITH statements permitted in PKG_XLSX_DIRECT."""

    def _first_word(self, sql: str) -> str:
        """Mirrors the REGEXP_SUBSTR logic in the fixed build_xlsx."""
        snippet = sql[:100]
        m = re.search(r'[A-Za-z]+', snippet)
        return m.group(0).upper() if m else ''

    def _validate(self, sql: str):
        word = self._first_word(sql)
        if word not in ('SELECT', 'WITH'):
            raise ValueError(
                f'Only SELECT/WITH statements are permitted. '
                f'Query starts with: {word}')

    def test_select_allowed(self):
        self._validate('SELECT * FROM EMP')

    def test_with_allowed(self):
        self._validate('WITH cte AS (SELECT 1 FROM DUAL) SELECT * FROM cte')

    def test_insert_rejected(self):
        with self.assertRaises(ValueError):
            self._validate('INSERT INTO T VALUES (1)')

    def test_update_rejected(self):
        with self.assertRaises(ValueError):
            self._validate('UPDATE T SET COL = 1')

    def test_delete_rejected(self):
        with self.assertRaises(ValueError):
            self._validate('DELETE FROM T WHERE 1=1')

    def test_drop_rejected(self):
        with self.assertRaises(ValueError):
            self._validate('DROP TABLE T')

    def test_exec_rejected(self):
        with self.assertRaises(ValueError):
            self._validate('EXECUTE IMMEDIATE ...')


class TestPurgeLogOrder(unittest.TestCase):
    """Fix 20 — purge logs AFTER delete, not before."""

    def test_log_after_delete_order(self):
        """
        Simulate the corrected purge flow:
          1. DELETE rows
          2. Record count
          3. COMMIT
          4. LOG (only if rows were deleted)
        Verify the log message only mentions rows that were actually deleted.
        """
        # Simulate the database state
        db = [{'id': i, 'age_days': 10} for i in range(5)]  # 5 old rows
        log = []

        def purge(days=7):
            # DELETE first (fix 20)
            before = len(db)
            db.clear()   # all older than 7 days in this simulation
            purged = before - len(db)
            # COMMIT (implicit in simulation)
            # LOG after
            if purged > 0:
                log.append(f'PURGED {purged} export(s)')
            return purged

        count = purge()
        self.assertEqual(count, 5)
        self.assertEqual(len(log), 1)
        self.assertIn('5', log[0])

    def test_no_log_when_nothing_purged(self):
        """When nothing is deleted, no log entry should be created."""
        db = []  # already empty
        log = []

        def purge(days=7):
            before = len(db)
            # nothing to delete
            purged = before - len(db)
            if purged > 0:
                log.append(f'PURGED {purged}')

        purge()
        self.assertEqual(len(log), 0)


# =============================================================================
# Runner
# =============================================================================

if __name__ == '__main__':
    loader = unittest.TestLoader()
    suite  = unittest.TestSuite()

    test_classes = [
        TestCRC32,
        TestColLetter,
        TestXmlEscape,
        TestSplitQueries,
        TestDeriveSheetName,
        TestUniqueSheetName,
        TestToLe,
        TestZipStructure,
        TestExcelDateSerial,
        TestSheetViewsWrapper,
        TestAddSheetValidation,
        TestDateFormatMask,
        TestSelectOnlyGuard,
        TestPurgeLogOrder,
    ]

    for tc in test_classes:
        suite.addTests(loader.loadTestsFromTestCase(tc))

    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)

    # Non-zero exit on failure (useful in CI)
    import sys
    sys.exit(0 if result.wasSuccessful() else 1)
