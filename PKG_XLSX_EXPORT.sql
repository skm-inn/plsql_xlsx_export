-- =============================================================================
-- PKG_XLSX_EXPORT - Pure PL/SQL XLSX Generator (Core Engine)
-- Compatible: Oracle 12c, 19c
-- No APEX | No UTL_FILE | No Shell | No External Tools | No Table dependency
--
-- Responsibilities (ONLY these — nothing else):
--   1. Accept sheet registrations (init / add_sheet)
--   2. Execute each SQL query via DBMS_SQL
--   3. Build a valid XLSX file as a BLOB in memory (ZIP + XML)
--   4. Return that BLOB to the caller via build_blob
--
-- This package has ZERO table references.
-- It compiles and runs in ANY environment including read-only schemas.
--
-- Persistence (storing the BLOB) is the responsibility of the caller:
--   - PKG_XLSX_WRAPPER  : stores into XLSX_EXPORT_RESULTS (full-privilege env)
--   - PKG_XLSX_DIRECT   : returns BLOB directly to SQL caller (read-only env)
-- =============================================================================

-- -----------------------------------------------------------------------------
-- STEP 1: Package Specification
-- -----------------------------------------------------------------------------
CREATE OR REPLACE PACKAGE PKG_XLSX_EXPORT AUTHID CURRENT_USER AS

  /**
   * Reset the sheet registry. Call before starting a new workbook.
   * p_workbook_name: Label passed through to the BLOB caller for naming.
   */
  PROCEDURE init(
    p_workbook_name IN VARCHAR2 DEFAULT 'Export_' || TO_CHAR(SYSDATE,'YYYYMMDD_HH24MI')
  );

  /**
   * Register a SQL query as a worksheet.
   * p_sheet_name : Tab name in Excel (max 31 chars)
   * p_sql        : Any SELECT statement.
   */
  PROCEDURE add_sheet(
    p_sheet_name IN VARCHAR2,
    p_sql        IN CLOB
  );

  /**
   * Execute all registered queries, build the XLSX file, and return it
   * as a BLOB. The caller decides what to do with it — store, stream,
   * or return directly to SQL. No table access here.
   */
  FUNCTION build_blob RETURN BLOB;

  /**
   * Expose the current workbook name so callers (PKG_XLSX_WRAPPER etc.)
   * can read it back for storage/naming without needing a separate parameter.
   */
  FUNCTION get_workbook_name RETURN VARCHAR2;

END PKG_XLSX_EXPORT;
/

-- -----------------------------------------------------------------------------
-- STEP 2: Package Body
-- -----------------------------------------------------------------------------
CREATE OR REPLACE PACKAGE BODY PKG_XLSX_EXPORT AS

  -- ============================================================
  -- SECTION A: Internal Types & State
  -- ============================================================

  TYPE t_sheet IS RECORD (
    sheet_name  VARCHAR2(31),
    sql_text    CLOB
  );

  TYPE t_sheet_tab IS TABLE OF t_sheet INDEX BY PLS_INTEGER;

  g_sheets        t_sheet_tab;
  g_workbook_name VARCHAR2(200);
  g_sheet_count   PLS_INTEGER := 0;

  -- ============================================================
  -- SECTION B: ZIP Builder (embedded AS_ZIP logic, adapted)
  --            Builds a valid ZIP archive entirely in PL/SQL
  --            using RAW/BLOB manipulation.
  -- ============================================================

  -- ZIP structures live here during build
  TYPE t_zip_entry IS RECORD (
    file_name    VARCHAR2(1000),
    content      BLOB,
    crc32        NUMBER,
    comp_size    NUMBER,
    uncomp_size  NUMBER,
    offset       NUMBER
  );
  TYPE t_zip_tab IS TABLE OF t_zip_entry INDEX BY PLS_INTEGER;

  g_zip_entries t_zip_tab;
  g_zip_blob    BLOB;

  -- ---- CRC32 Table (standard polynomial 0xEDB88320) ----
  TYPE t_crc_tab IS TABLE OF NUMBER INDEX BY PLS_INTEGER;
  g_crc_tab t_crc_tab;

  PROCEDURE init_crc_table IS
    v_c  NUMBER;
    v_k  PLS_INTEGER;
  BEGIN
    FOR i IN 0..255 LOOP
      v_c := i;
      FOR j IN 0..7 LOOP
        IF MOD(v_c, 2) = 1 THEN
          v_c := 3988292384 - TRUNC(v_c / 2); -- 0xEDB88320
        ELSE
          v_c := TRUNC(v_c / 2);
        END IF;
      END LOOP;
      g_crc_tab(i) := v_c;
    END LOOP;
  END;

  FUNCTION crc32(p_data IN BLOB) RETURN NUMBER IS
    v_crc    NUMBER := 4294967295;
    v_len    NUMBER;
    v_chunk  RAW(32767);
    v_pos    NUMBER := 1;
    v_byte   NUMBER;
    v_idx    NUMBER;
    v_low    NUMBER;
  BEGIN
    IF g_crc_tab.COUNT = 0 THEN init_crc_table; END IF;
    v_len := DBMS_LOB.GETLENGTH(p_data);
    WHILE v_pos <= v_len LOOP
      v_chunk := DBMS_LOB.SUBSTR(p_data, LEAST(32767, v_len - v_pos + 1), v_pos);
      FOR i IN 1..UTL_RAW.LENGTH(v_chunk) LOOP
        v_byte := TO_NUMBER(RAWTOHEX(UTL_RAW.SUBSTR(v_chunk, i, 1)), 'XX');
        v_low  := MOD(v_crc, 256);
        -- XOR: a XOR b = a + b - 2 * BITAND(a, b)
        v_idx  := v_low + v_byte - 2 * BITAND(v_low, v_byte);
        v_crc  := g_crc_tab(v_idx) + TRUNC(v_crc / 256)
                - 2 * BITAND(g_crc_tab(v_idx), TRUNC(v_crc / 256));
        v_crc  := MOD(v_crc, 4294967296);
      END LOOP;
      v_pos := v_pos + 32767;
    END LOOP;
    -- Finalize: XOR with 0xFFFFFFFF = 4294967295 - crc (since crc < 2^32)
    RETURN 4294967295 - v_crc;
  END;

  -- ---- Little-endian helpers ----

  /** Convert integer to N-byte little-endian RAW */
  FUNCTION to_le(p_val IN NUMBER, p_bytes IN PLS_INTEGER) RETURN RAW IS
    v_hex VARCHAR2(16);
    v_len PLS_INTEGER := p_bytes * 2;
    v_out VARCHAR2(16) := '';
    v_tmp NUMBER := p_val;
  BEGIN
    FOR i IN 1..p_bytes LOOP
      v_out := v_out || LPAD(TO_CHAR(MOD(v_tmp, 256), 'XX'), 2, '0');
      v_tmp := TRUNC(v_tmp / 256);
    END LOOP;
    RETURN HEXTORAW(v_out);
  END;

  PROCEDURE append_raw(p_blob IN OUT NOCOPY BLOB, p_raw IN RAW) IS
  BEGIN
    IF p_raw IS NOT NULL AND UTL_RAW.LENGTH(p_raw) > 0 THEN
      DBMS_LOB.WRITEAPPEND(p_blob, UTL_RAW.LENGTH(p_raw), p_raw);
    END IF;
  END;

  PROCEDURE append_blob(p_dest IN OUT NOCOPY BLOB, p_src IN BLOB) IS
    v_len  NUMBER;
    v_pos  NUMBER := 1;
    v_amt  NUMBER := 32767;
    v_buf  RAW(32767);
  BEGIN
    v_len := DBMS_LOB.GETLENGTH(p_src);
    WHILE v_pos <= v_len LOOP
      DBMS_LOB.READ(p_src, LEAST(v_amt, v_len - v_pos + 1), v_pos, v_buf);
      DBMS_LOB.WRITEAPPEND(p_dest, UTL_RAW.LENGTH(v_buf), v_buf);
      v_pos := v_pos + v_amt;
    END LOOP;
  END;

  /** Convert VARCHAR2/CLOB text to BLOB (UTF-8) */
  FUNCTION str_to_blob(p_str IN CLOB) RETURN BLOB IS
    v_blob BLOB;
    v_dest_off NUMBER := 1;
    v_src_off  NUMBER := 1;
    v_lang     NUMBER := DBMS_LOB.DEFAULT_LANG_CTX;
    v_warn     NUMBER;
  BEGIN
    DBMS_LOB.CREATETEMPORARY(v_blob, TRUE);
    DBMS_LOB.CONVERTTOBLOB(v_blob, p_str, DBMS_LOB.LOBMAXSIZE,
                           v_dest_off, v_src_off,
                           NLS_CHARSET_ID('AL32UTF8'), v_lang, v_warn);
    RETURN v_blob;
  END;

  /** Add a file (as BLOB) into the in-memory ZIP structure */
  PROCEDURE zip_add_file(
    p_name    IN VARCHAR2,
    p_content IN BLOB
  ) IS
    v_entry t_zip_entry;
    v_idx   PLS_INTEGER;
  BEGIN
    v_idx := g_zip_entries.COUNT + 1;
    v_entry.file_name   := p_name;
    v_entry.content     := p_content;
    v_entry.uncomp_size := NVL(DBMS_LOB.GETLENGTH(p_content), 0);
    v_entry.comp_size   := v_entry.uncomp_size; -- STORE mode (no compression)
    v_entry.crc32       := crc32(p_content);
    v_entry.offset      := NVL(DBMS_LOB.GETLENGTH(g_zip_blob), 0);
    g_zip_entries(v_idx) := v_entry;

    -- Write Local File Header
    append_raw(g_zip_blob, HEXTORAW('504B0304'));           -- signature
    append_raw(g_zip_blob, to_le(20, 2));                  -- version needed: 2.0
    append_raw(g_zip_blob, to_le(0,  2));                  -- flags
    append_raw(g_zip_blob, to_le(0,  2));                  -- compression: STORED
    append_raw(g_zip_blob, to_le(0,  2));                  -- mod time
    append_raw(g_zip_blob, to_le(0,  2));                  -- mod date
    append_raw(g_zip_blob, to_le(v_entry.crc32,       4)); -- crc32
    append_raw(g_zip_blob, to_le(v_entry.comp_size,   4)); -- compressed size
    append_raw(g_zip_blob, to_le(v_entry.uncomp_size, 4)); -- uncompressed size
    append_raw(g_zip_blob, to_le(LENGTH(p_name),      2)); -- filename length
    append_raw(g_zip_blob, to_le(0, 2));                   -- extra field length
    append_raw(g_zip_blob, UTL_RAW.CAST_TO_RAW(p_name));  -- filename
    -- File data
    append_blob(g_zip_blob, p_content);
  END;

  /** Finalise ZIP: write Central Directory + End of Central Directory */
  PROCEDURE zip_finish IS
    v_cd_start NUMBER;
    v_cd_size  NUMBER;
    v_e        t_zip_entry;
    v_fname_raw RAW(1000);
  BEGIN
    v_cd_start := NVL(DBMS_LOB.GETLENGTH(g_zip_blob), 0);

    FOR i IN 1..g_zip_entries.COUNT LOOP
      v_e := g_zip_entries(i);
      v_fname_raw := UTL_RAW.CAST_TO_RAW(v_e.file_name);
      -- Central Directory Entry
      append_raw(g_zip_blob, HEXTORAW('504B0102'));            -- signature
      append_raw(g_zip_blob, to_le(20, 2));                   -- version made by
      append_raw(g_zip_blob, to_le(20, 2));                   -- version needed
      append_raw(g_zip_blob, to_le(0,  2));                   -- flags
      append_raw(g_zip_blob, to_le(0,  2));                   -- compression STORED
      append_raw(g_zip_blob, to_le(0,  2));                   -- mod time
      append_raw(g_zip_blob, to_le(0,  2));                   -- mod date
      append_raw(g_zip_blob, to_le(v_e.crc32,       4));      -- crc32
      append_raw(g_zip_blob, to_le(v_e.comp_size,   4));      -- comp size
      append_raw(g_zip_blob, to_le(v_e.uncomp_size, 4));      -- uncomp size
      append_raw(g_zip_blob, to_le(UTL_RAW.LENGTH(v_fname_raw), 2)); -- name len
      append_raw(g_zip_blob, to_le(0, 2));                    -- extra len
      append_raw(g_zip_blob, to_le(0, 2));                    -- comment len
      append_raw(g_zip_blob, to_le(0, 2));                    -- disk number
      append_raw(g_zip_blob, to_le(0, 2));                    -- int attributes
      append_raw(g_zip_blob, to_le(0, 4));                    -- ext attributes
      append_raw(g_zip_blob, to_le(v_e.offset, 4));           -- local header offset
      append_raw(g_zip_blob, v_fname_raw);                    -- filename
    END LOOP;

    v_cd_size := NVL(DBMS_LOB.GETLENGTH(g_zip_blob), 0) - v_cd_start;

    -- End of Central Directory Record
    append_raw(g_zip_blob, HEXTORAW('504B0506'));              -- signature
    append_raw(g_zip_blob, to_le(0, 2));                      -- disk number
    append_raw(g_zip_blob, to_le(0, 2));                      -- disk with CD
    append_raw(g_zip_blob, to_le(g_zip_entries.COUNT, 2));    -- entries on disk
    append_raw(g_zip_blob, to_le(g_zip_entries.COUNT, 2));    -- total entries
    append_raw(g_zip_blob, to_le(v_cd_size,   4));            -- CD size
    append_raw(g_zip_blob, to_le(v_cd_start,  4));            -- CD offset
    append_raw(g_zip_blob, to_le(0, 2));                      -- comment length
  END;

  -- ============================================================
  -- SECTION C: XLSX XML Generators
  -- ============================================================

  FUNCTION xml_escape(p_str IN VARCHAR2) RETURN VARCHAR2 IS
  BEGIN
    RETURN REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(
           p_str,
           '&',  '&amp;'),
           '<',  '&lt;'),
           '>',  '&gt;'),
           '"',  '&quot;'),
           '''', '&apos;');
  END;

  /** [Content_Types].xml */
  FUNCTION gen_content_types(p_sheet_count IN PLS_INTEGER) RETURN CLOB IS
    v_xml CLOB;
  BEGIN
    v_xml := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          || '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          || '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          || '<Default Extension="xml"  ContentType="application/xml"/>'
          || '<Override PartName="/xl/workbook.xml"'
          || ' ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
          || '<Override PartName="/xl/styles.xml"'
          || ' ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
          || '<Override PartName="/xl/sharedStrings.xml"'
          || ' ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>';
    FOR i IN 1..p_sheet_count LOOP
      v_xml := v_xml
            || '<Override PartName="/xl/worksheets/sheet' || i || '.xml"'
            || ' ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
    END LOOP;
    v_xml := v_xml || '</Types>';
    RETURN v_xml;
  END;

  /** _rels/.rels */
  FUNCTION gen_rels RETURN CLOB IS
  BEGIN
    RETURN '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        || '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        || '<Relationship Id="rId1"'
        || ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"'
        || ' Target="xl/workbook.xml"/>'
        || '</Relationships>';
  END;

  /** xl/_rels/workbook.xml.rels */
  FUNCTION gen_workbook_rels(p_sheet_count IN PLS_INTEGER) RETURN CLOB IS
    v_xml CLOB;
  BEGIN
    v_xml := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          || '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
    FOR i IN 1..p_sheet_count LOOP
      v_xml := v_xml
            || '<Relationship Id="rId' || i || '"'
            || ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"'
            || ' Target="worksheets/sheet' || i || '.xml"/>';
    END LOOP;
    v_xml := v_xml
          || '<Relationship Id="rId' || (p_sheet_count + 1) || '"'
          || ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"'
          || ' Target="sharedStrings.xml"/>'
          || '<Relationship Id="rId' || (p_sheet_count + 2) || '"'
          || ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"'
          || ' Target="styles.xml"/>'
          || '</Relationships>';
    RETURN v_xml;
  END;

  /** xl/workbook.xml */
  FUNCTION gen_workbook RETURN CLOB IS
    v_xml CLOB;
  BEGIN
    v_xml := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          || '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
          || ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
          || '<sheets>';
    FOR i IN 1..g_sheet_count LOOP
      v_xml := v_xml
            || '<sheet name="' || xml_escape(SUBSTR(g_sheets(i).sheet_name, 1, 31))
            || '" sheetId="' || i
            || '" r:id="rId' || i || '"/>';
    END LOOP;
    v_xml := v_xml || '</sheets></workbook>';
    RETURN v_xml;
  END;

  /** xl/styles.xml — header (bold) + date + number + default styles */
  FUNCTION gen_styles RETURN CLOB IS
  BEGIN
    RETURN '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        || '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        -- Number formats
        || '<numFmts count="1">'
        || '<numFmt numFmtId="164" formatCode="YYYY-MM-DD HH:mm:SS"/>'
        || '</numFmts>'
        -- Fonts: 0=normal, 1=bold header
        || '<fonts count="2">'
        || '<font><sz val="11"/><name val="Calibri"/></font>'
        || '<font><b/><sz val="11"/><name val="Calibri"/></font>'
        || '</fonts>'
        -- Fills: 0=none, 1=grey125 (required), 2=header blue
        || '<fills count="3">'
        || '<fill><patternFill patternType="none"/></fill>'
        || '<fill><patternFill patternType="gray125"/></fill>'
        || '<fill><patternFill patternType="solid"><fgColor rgb="FF4472C4"/></patternFill></fill>'
        || '</fills>'
        -- Borders: 0=none
        || '<borders count="1">'
        || '<border><left/><right/><top/><bottom/><diagonal/></border>'
        || '</borders>'
        -- Cell style xfs
        || '<cellStyleXfs count="1">'
        || '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>'
        || '</cellStyleXfs>'
        -- Cell xfs:
        -- 0 = default (string/number)
        -- 1 = header (bold + blue fill + white font)
        -- 2 = date format
        || '<cellXfs count="3">'
        || '<xf numFmtId="0"   fontId="0" fillId="0" borderId="0" xfId="0"/>'
        || '<xf numFmtId="0"   fontId="1" fillId="2" borderId="0" xfId="0">'
        || '<alignment horizontal="center"/></xf>'
        || '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        || '</cellXfs>'
        || '</styleSheet>';
  END;

  -- ============================================================
  -- SECTION D: Sheet Data Generator (core engine)
  --            Uses DBMS_SQL for truly dynamic any-column query
  -- ============================================================

  /**
   * Returns TYPE code bucket:
   *   1 = VARCHAR2/CHAR/NVARCHAR  -> string cell (t="s")
   *   2 = NUMBER                  -> number cell (t="n")
   *   3 = DATE/TIMESTAMP          -> date cell   (t="n", style 2)
   *   4 = CLOB/LONG               -> string cell
   */
  FUNCTION col_type_bucket(p_dbms_type IN PLS_INTEGER) RETURN PLS_INTEGER IS
  BEGIN
    CASE p_dbms_type
      WHEN  1 THEN RETURN 1;  -- VARCHAR2
      WHEN  2 THEN RETURN 2;  -- NUMBER
      WHEN  8 THEN RETURN 4;  -- LONG
      WHEN  9 THEN RETURN 1;  -- VARCHAR2 (alt code)
      WHEN 12 THEN RETURN 3;  -- DATE
      WHEN 23 THEN RETURN 1;  -- RAW (treat as string)
      WHEN 96 THEN RETURN 1;  -- CHAR
      WHEN 100 THEN RETURN 2; -- BINARY_FLOAT
      WHEN 101 THEN RETURN 2; -- BINARY_DOUBLE
      WHEN 112 THEN RETURN 4; -- CLOB
      WHEN 178 THEN RETURN 3; -- TIME
      WHEN 180 THEN RETURN 3; -- TIMESTAMP
      WHEN 181 THEN RETURN 3; -- TIMESTAMP WITH TZ
      WHEN 182 THEN RETURN 3; -- INTERVAL YM
      WHEN 183 THEN RETURN 3; -- INTERVAL DS
      WHEN 231 THEN RETURN 3; -- TIMESTAMP LOCAL TZ
      ELSE RETURN 1;           -- Default: treat as string
    END CASE;
  END;

  /**
   * Convert column number to Excel letter (A, B, ... Z, AA, AB, ...)
   */
  FUNCTION col_letter(p_col IN PLS_INTEGER) RETURN VARCHAR2 IS
    v_result VARCHAR2(10) := '';
    v_col    PLS_INTEGER  := p_col;
  BEGIN
    WHILE v_col > 0 LOOP
      v_result := CHR(65 + MOD(v_col - 1, 26)) || v_result;
      v_col    := TRUNC((v_col - 1) / 26);
    END LOOP;
    RETURN v_result;
  END;

  /**
   * Execute a SQL query and generate the xl/worksheets/sheetN.xml content.
   * Also populates the shared strings list passed in.
   *
   * p_sql         : SELECT statement
   * p_ss_list     : IN/OUT shared strings collection
   * p_ss_index    : IN/OUT current shared string index
   * Returns       : CLOB of sheet XML
   */
  FUNCTION gen_sheet_xml(
    p_sql      IN CLOB,
    p_ss_list  IN OUT NOCOPY DBMS_SQL.VARCHAR2A,
    p_ss_index IN OUT NOCOPY PLS_INTEGER
  ) RETURN CLOB IS

    v_cursor      INTEGER;
    v_ret         INTEGER;
    v_col_cnt     INTEGER;
    v_desc_tab    DBMS_SQL.DESC_TAB;
    v_type        PLS_INTEGER;
    v_xml         CLOB;
    v_row_clob    CLOB;
    v_cell_ref    VARCHAR2(10);
    v_row_num     PLS_INTEGER := 1;

    -- Column value buffers
    TYPE t_varchar_tab IS TABLE OF VARCHAR2(32767) INDEX BY PLS_INTEGER;
    TYPE t_number_tab  IS TABLE OF NUMBER          INDEX BY PLS_INTEGER;
    TYPE t_date_tab    IS TABLE OF DATE            INDEX BY PLS_INTEGER;
    TYPE t_clob_tab    IS TABLE OF CLOB            INDEX BY PLS_INTEGER;

    v_varchar_val  t_varchar_tab;
    v_number_val   t_number_tab;
    v_date_val     t_date_tab;
    v_clob_val     t_clob_tab;

    TYPE t_bucket_tab IS TABLE OF PLS_INTEGER INDEX BY PLS_INTEGER;
    v_buckets t_bucket_tab;

    -- Shared string helper
    FUNCTION get_ss_idx(p_str IN VARCHAR2) RETURN PLS_INTEGER IS
      v_idx PLS_INTEGER := p_ss_index;
    BEGIN
      p_ss_list(v_idx) := p_str;
      p_ss_index := p_ss_index + 1;
      RETURN v_idx;
    END;

  BEGIN
    DBMS_LOB.CREATETEMPORARY(v_xml, TRUE);
    DBMS_LOB.CREATETEMPORARY(v_row_clob, TRUE);

    v_cursor := DBMS_SQL.OPEN_CURSOR;
    BEGIN
      DBMS_SQL.PARSE(v_cursor, p_sql, DBMS_SQL.NATIVE);
      DBMS_SQL.DESCRIBE_COLUMNS(v_cursor, v_col_cnt, v_desc_tab);

      -- Define columns based on type
      FOR i IN 1..v_col_cnt LOOP
        v_type := v_desc_tab(i).col_type;
        v_buckets(i) := col_type_bucket(v_type);
        CASE v_buckets(i)
          WHEN 1 THEN DBMS_SQL.DEFINE_COLUMN(v_cursor, i, v_varchar_val(1), 32767);
          WHEN 2 THEN DBMS_SQL.DEFINE_COLUMN(v_cursor, i, v_number_val(1));
          WHEN 3 THEN DBMS_SQL.DEFINE_COLUMN(v_cursor, i, v_date_val(1));
          WHEN 4 THEN DBMS_SQL.DEFINE_COLUMN(v_cursor, i, v_clob_val(1));
          ELSE        DBMS_SQL.DEFINE_COLUMN(v_cursor, i, v_varchar_val(1), 32767);
        END CASE;
      END LOOP;

      v_ret := DBMS_SQL.EXECUTE(v_cursor);

      -- XML preamble
      DBMS_LOB.APPEND(v_xml,
        TO_CLOB('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             || '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
             || ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
             || '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="A1"/></sheetView></sheetViews>'
             || '<sheetData>'));

      -- Header row (style 1 = bold blue)
      DBMS_LOB.TRIM(v_row_clob, 0);
      DBMS_LOB.APPEND(v_row_clob, TO_CLOB('<row r="1">'));
      FOR i IN 1..v_col_cnt LOOP
        v_cell_ref := col_letter(i) || '1';
        DBMS_LOB.APPEND(v_row_clob,
          TO_CLOB('<c r="' || v_cell_ref || '" t="s" s="1">'
               || '<v>' || get_ss_idx(SUBSTR(v_desc_tab(i).col_name, 1, 32767)) || '</v>'
               || '</c>'));
      END LOOP;
      DBMS_LOB.APPEND(v_row_clob, TO_CLOB('</row>'));
      DBMS_LOB.APPEND(v_xml, v_row_clob);

      v_row_num := 2;

      -- Data rows
      WHILE DBMS_SQL.FETCH_ROWS(v_cursor) > 0 LOOP
        DBMS_LOB.TRIM(v_row_clob, 0);
        DBMS_LOB.APPEND(v_row_clob, TO_CLOB('<row r="' || v_row_num || '">'));

        FOR i IN 1..v_col_cnt LOOP
          v_cell_ref := col_letter(i) || v_row_num;
          CASE v_buckets(i)
            WHEN 1 THEN
              DBMS_SQL.COLUMN_VALUE(v_cursor, i, v_varchar_val(i));
              IF v_varchar_val(i) IS NOT NULL THEN
                DBMS_LOB.APPEND(v_row_clob,
                  TO_CLOB('<c r="' || v_cell_ref || '" t="s">'
                       || '<v>' || get_ss_idx(SUBSTR(v_varchar_val(i), 1, 32767)) || '</v>'
                       || '</c>'));
              END IF;
            WHEN 2 THEN
              DBMS_SQL.COLUMN_VALUE(v_cursor, i, v_number_val(i));
              IF v_number_val(i) IS NOT NULL THEN
                DBMS_LOB.APPEND(v_row_clob,
                  TO_CLOB('<c r="' || v_cell_ref || '">'
                       || '<v>' || TO_CHAR(v_number_val(i)) || '</v>'
                       || '</c>'));
              END IF;
            WHEN 3 THEN
              DBMS_SQL.COLUMN_VALUE(v_cursor, i, v_date_val(i));
              IF v_date_val(i) IS NOT NULL THEN
                -- Excel date serial: days since 1899-12-30 (accounting for Excel's 1900 leap year bug)
                DECLARE
                  v_serial NUMBER;
                BEGIN
                  v_serial := (v_date_val(i) - DATE '1899-12-30');
                  DBMS_LOB.APPEND(v_row_clob,
                    TO_CLOB('<c r="' || v_cell_ref || '" s="2">'
                         || '<v>' || TO_CHAR(v_serial, 'FM99999999990.9999999999') || '</v>'
                         || '</c>'));
                END;
              END IF;
            WHEN 4 THEN
              DBMS_SQL.COLUMN_VALUE(v_cursor, i, v_clob_val(i));
              IF v_clob_val(i) IS NOT NULL AND DBMS_LOB.GETLENGTH(v_clob_val(i)) > 0 THEN
                DBMS_LOB.APPEND(v_row_clob,
                  TO_CLOB('<c r="' || v_cell_ref || '" t="s">'
                       || '<v>' || get_ss_idx(SUBSTR(v_clob_val(i), 1, 32767)) || '</v>'
                       || '</c>'));
              END IF;
            ELSE NULL;
          END CASE;
        END LOOP;

        DBMS_LOB.APPEND(v_row_clob, TO_CLOB('</row>'));
        DBMS_LOB.APPEND(v_xml, v_row_clob);
        v_row_num := v_row_num + 1;
      END LOOP;

      DBMS_LOB.APPEND(v_xml, TO_CLOB('</sheetData></worksheet>'));

    EXCEPTION
      WHEN OTHERS THEN
        IF DBMS_SQL.IS_OPEN(v_cursor) THEN DBMS_SQL.CLOSE_CURSOR(v_cursor); END IF;
        IF v_row_clob IS NOT NULL THEN DBMS_LOB.FREETEMPORARY(v_row_clob); END IF;
        RAISE;
    END;

    DBMS_SQL.CLOSE_CURSOR(v_cursor);
    DBMS_LOB.FREETEMPORARY(v_row_clob);
    RETURN v_xml;
  END gen_sheet_xml;

  /** Build xl/sharedStrings.xml from the collected string table */
  FUNCTION gen_shared_strings(
    p_ss_list  IN DBMS_SQL.VARCHAR2A,
    p_ss_index IN PLS_INTEGER
  ) RETURN CLOB IS
    v_xml CLOB;
  BEGIN
    DBMS_LOB.CREATETEMPORARY(v_xml, TRUE);
    DBMS_LOB.APPEND(v_xml,
      TO_CLOB('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           || '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
           || ' count="' || (p_ss_index) || '"'
           || ' uniqueCount="' || (p_ss_index) || '">'));
    FOR i IN 0..(p_ss_index - 1) LOOP
      DBMS_LOB.APPEND(v_xml, TO_CLOB('<si><t xml:space="preserve">'
                                  || xml_escape(p_ss_list(i)) || '</t></si>'));
    END LOOP;
    DBMS_LOB.APPEND(v_xml, TO_CLOB('</sst>'));
    RETURN v_xml;
  END;

  -- ============================================================
  -- SECTION E: Public API Implementation
  -- ============================================================

  PROCEDURE init(
    p_workbook_name IN VARCHAR2 DEFAULT 'Export_' || TO_CHAR(SYSDATE,'YYYYMMDD_HH24MI')
  ) IS
  BEGIN
    g_sheets.DELETE;
    g_sheet_count   := 0;
    g_workbook_name := NVL(p_workbook_name, 'Export');
    -- Reset ZIP state
    g_zip_entries.DELETE;
    IF g_zip_blob IS NOT NULL THEN
      DBMS_LOB.FREETEMPORARY(g_zip_blob);
    END IF;
    DBMS_LOB.CREATETEMPORARY(g_zip_blob, TRUE);
  END;

  PROCEDURE add_sheet(
    p_sheet_name IN VARCHAR2,
    p_sql        IN CLOB
  ) IS
  BEGIN
    IF TRIM(p_sheet_name) IS NULL THEN
      RAISE_APPLICATION_ERROR(-20002, 'PKG_XLSX_EXPORT.add_sheet: sheet name cannot be NULL or empty.');
    END IF;
    IF p_sql IS NULL OR DBMS_LOB.GETLENGTH(p_sql) = 0 THEN
      RAISE_APPLICATION_ERROR(-20003, 'PKG_XLSX_EXPORT.add_sheet: SQL text cannot be NULL or empty.');
    END IF;
    g_sheet_count := g_sheet_count + 1;
    g_sheets(g_sheet_count).sheet_name := REGEXP_REPLACE(SUBSTR(p_sheet_name, 1, 31), '[\\/:*?\[\]]', '_');
    g_sheets(g_sheet_count).sql_text   := p_sql;
  END;

  FUNCTION get_workbook_name RETURN VARCHAR2 IS
  BEGIN
    RETURN g_workbook_name;
  END;

  FUNCTION build_blob RETURN BLOB IS
    v_ss_list   DBMS_SQL.VARCHAR2A;
    v_ss_index  PLS_INTEGER := 0;
    v_sheet_xml CLOB;
    v_blob      BLOB;
  BEGIN
    IF g_sheet_count = 0 THEN
      RAISE_APPLICATION_ERROR(-20001,
        'PKG_XLSX_EXPORT: No sheets registered. Call add_sheet first.');
    END IF;

    -- Always reset ZIP state to prevent stale entries from previous calls
    g_zip_entries.DELETE;
    IF g_zip_blob IS NOT NULL THEN
      DBMS_LOB.FREETEMPORARY(g_zip_blob);
    END IF;
    DBMS_LOB.CREATETEMPORARY(g_zip_blob, TRUE);

    TYPE t_clob_tab IS TABLE OF CLOB INDEX BY PLS_INTEGER;
    v_sheet_xmls t_clob_tab;

    FOR i IN 1..g_sheet_count LOOP
      v_sheet_xmls(i) := gen_sheet_xml(
                           g_sheets(i).sql_text, v_ss_list, v_ss_index);
    END LOOP;

    zip_add_file('[Content_Types].xml',        str_to_blob(gen_content_types(g_sheet_count)));
    zip_add_file('_rels/.rels',                str_to_blob(gen_rels));
    zip_add_file('xl/workbook.xml',            str_to_blob(gen_workbook));
    zip_add_file('xl/styles.xml',              str_to_blob(gen_styles));
    zip_add_file('xl/sharedStrings.xml',       str_to_blob(gen_shared_strings(v_ss_list, v_ss_index)));
    zip_add_file('xl/_rels/workbook.xml.rels', str_to_blob(gen_workbook_rels(g_sheet_count)));

    FOR i IN 1..g_sheet_count LOOP
      zip_add_file('xl/worksheets/sheet' || i || '.xml',
                   str_to_blob(v_sheet_xmls(i)));
    END LOOP;

    zip_finish;

    DBMS_LOB.CREATETEMPORARY(v_blob, TRUE);
    DBMS_LOB.COPY(v_blob, g_zip_blob, DBMS_LOB.GETLENGTH(g_zip_blob));
    RETURN v_blob;
  END build_blob;

END PKG_XLSX_EXPORT;
/


-- =============================================================================
-- USAGE EXAMPLES
-- =============================================================================

/*

-- ── build_blob — use in any environment, no table needed ─────────────────────
DECLARE
  v_blob BLOB;
BEGIN
  PKG_XLSX_EXPORT.init('My_Report');

  PKG_XLSX_EXPORT.add_sheet(
    p_sheet_name => 'Employees',
    p_sql        => 'SELECT EMPLOYEE_ID, FIRST_NAME, LAST_NAME, SALARY
                     FROM EMPLOYEES ORDER BY LAST_NAME'
  );

  PKG_XLSX_EXPORT.add_sheet(
    p_sheet_name => 'Departments',
    p_sql        => 'SELECT DEPARTMENT_ID, DEPARTMENT_NAME FROM DEPARTMENTS'
  );

  -- Returns BLOB — caller stores, streams, or downloads it
  v_blob := PKG_XLSX_EXPORT.build_blob;
  DBMS_OUTPUT.PUT_LINE('Blob size: ' || DBMS_LOB.GETLENGTH(v_blob) || ' bytes');
END;
/

-- NOTE: For end-to-end usage including storage and download, use either:
--   PKG_XLSX_WRAPPER  (full-privilege environments — stores to table)
--   PKG_XLSX_DIRECT   (read-only environments  — returns from SELECT)

*/
