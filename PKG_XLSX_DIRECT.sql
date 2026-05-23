-- =============================================================================
-- PKG_XLSX_DIRECT
-- SQL-callable XLSX generator — returns BLOB directly from a SELECT statement.
-- No tables required. No CREATE TABLE privilege needed.
-- DBA only needs to grant: CREATE PACKAGE + EXECUTE on DBMS_SQL + DBMS_LOB.
--
-- Depends on : PKG_XLSX_EXPORT (must be deployed first by DBA)
-- Compatible : Oracle 12c, 19c
--
-- USAGE — invoke directly from SQL:
--
--   SELECT PKG_XLSX_DIRECT.generate_xlsx(
--              p_workbook_name => 'Deposit_Issue',
--              p_queries       => 'SELECT * FROM ICTM_ACC WHERE ROWNUM<=10'
--                              || PKG_XLSX_DIRECT.DELIM
--                              || q'[SELECT * FROM STTM_CUST_ACCOUNT
--                                    WHERE CUST_NAME LIKE 'A%']'
--          )
--   FROM DUAL;
--
-- Then save the BLOB cell as Deposit_Issue.xlsx from SQL Developer
-- or PL/SQL Developer.
--
-- WORKSHEET NAMING — same auto-derive rules as PKG_XLSX_WRAPPER:
--   Single table in SQL   → table name
--   Multiple tables/JOINs → Complex_01, Complex_02, ...
--   Unparseable SQL       → Sheet_1, Sheet_2, ...
--   Duplicate name        → suffixed _01, _02, ...
--
-- QUERY SEPARATOR — default is  §  (section sign, Unicode 00A7).
-- Change PKG_XLSX_DIRECT.DELIM if that character appears in your SQL.
-- =============================================================================


-- =============================================================================
-- STEP 1: Package Specification
-- =============================================================================
CREATE OR REPLACE PACKAGE PKG_XLSX_DIRECT AUTHID CURRENT_USER AS

  -- ---------------------------------------------------------------------------
  -- Query delimiter used to split multiple SQL statements in p_queries.
  -- Default: § (section sign).  Override if your SQL contains this character.
  -- Usage:    q1 || PKG_XLSX_DIRECT.DELIM || q2 || PKG_XLSX_DIRECT.DELIM || q3
  -- ---------------------------------------------------------------------------
  DELIM CONSTANT VARCHAR2(3) := '§';

  -- ---------------------------------------------------------------------------
  -- get_delim: SQL-callable accessor for the DELIM constant.
  -- Use this from SQL context — package constants are not reachable from SQL.
  -- Example: 'SELECT 1 FROM DUAL' || PKG_XLSX_DIRECT.get_delim() || 'SELECT 2 FROM DUAL'
  -- ---------------------------------------------------------------------------
  FUNCTION get_delim RETURN VARCHAR2;

  -- ---------------------------------------------------------------------------
  -- Overload 1 — sheet names auto-derived from SQL table names (simplest)
  --
  -- p_workbook_name : Workbook/file name (no .xlsx extension needed).
  --                   If NULL → SCHEMA_YYYYMMDD_HH24MISS
  -- p_queries       : One or more SELECT statements separated by DELIM.
  --
  -- Returns         : BLOB — a valid .xlsx file ready to download.
  --
  -- Example (single query):
  --   SELECT PKG_XLSX_DIRECT.generate_xlsx(
  --              'My_Report',
  --              'SELECT * FROM EMP WHERE DEPTNO = 10'
  --          ) FROM DUAL;
  --
  -- Example (multiple queries):
  --   SELECT PKG_XLSX_DIRECT.generate_xlsx(
  --              'Deposit_Issue',
  --              'SELECT * FROM ICTM_ACC WHERE ROWNUM<=10'
  --              || PKG_XLSX_DIRECT.DELIM
  --              || q'[SELECT * FROM STTM_CUST_ACCOUNT WHERE CUST_NAME LIKE 'A%']'
  --          ) FROM DUAL;
  -- ---------------------------------------------------------------------------
  FUNCTION generate_xlsx(
    p_workbook_name IN VARCHAR2 DEFAULT NULL,
    p_queries       IN CLOB
  ) RETURN BLOB;

  -- ---------------------------------------------------------------------------
  -- Overload 2 — sheet names explicitly provided
  --
  -- p_sheet_names   : Sheet names separated by DELIM, matching order of p_queries.
  --                   If fewer names than queries, remaining sheets are auto-named.
  --
  -- Example:
  --   SELECT PKG_XLSX_DIRECT.generate_xlsx(
  --              'Deposit_Issue',
  --              'Accounts'      || PKG_XLSX_DIRECT.DELIM || 'Customers',
  --              'SELECT * FROM ICTM_ACC WHERE ROWNUM<=10'
  --              || PKG_XLSX_DIRECT.DELIM
  --              || q'[SELECT * FROM STTM_CUST_ACCOUNT WHERE CUST_NAME LIKE 'A%']'
  --          ) FROM DUAL;
  -- ---------------------------------------------------------------------------
  FUNCTION generate_xlsx(
    p_workbook_name IN VARCHAR2 DEFAULT NULL,
    p_sheet_names   IN VARCHAR2,
    p_queries       IN CLOB
  ) RETURN BLOB;

  PROCEDURE enable_debug;   -- turn on DBMS_OUTPUT debug tracing
  PROCEDURE disable_debug;  -- turn off debug tracing (default)

END PKG_XLSX_DIRECT;
/


-- =============================================================================
-- STEP 2: Package Body
-- =============================================================================
CREATE OR REPLACE PACKAGE BODY PKG_XLSX_DIRECT AS

  -- ===========================================================================
  -- SECTION A: Internal types
  -- ===========================================================================
  TYPE t_str_tab  IS TABLE OF VARCHAR2(32767) INDEX BY PLS_INTEGER;
  TYPE t_clob_tab IS TABLE OF CLOB            INDEX BY PLS_INTEGER;

  -- ============================================================
  -- Debug flag — enable at runtime: PKG_XLSX_DIRECT.enable_debug;
  -- ============================================================
  g_debug BOOLEAN := FALSE;

  PROCEDURE dbg(p_msg IN VARCHAR2) IS
  BEGIN
    IF g_debug THEN
      DBMS_OUTPUT.PUT_LINE('[' || TO_CHAR(SYSTIMESTAMP,'HH24:MI:SS.FF3') || '] ' || p_msg);
    END IF;
  END dbg;

  PROCEDURE enable_debug  IS BEGIN g_debug := TRUE;  END;
  PROCEDURE disable_debug IS BEGIN g_debug := FALSE; END;

  -- ===========================================================================
  -- SECTION B: String splitter
  --   Splits a delimited CLOB into a t_str_tab / t_clob_tab collection.
  -- ===========================================================================

  /** Split a VARCHAR2 list (sheet names) by DELIM into t_str_tab */
  FUNCTION split_names(p_input IN VARCHAR2) RETURN t_str_tab IS
    v_result t_str_tab;
    v_pos    PLS_INTEGER := 1;
    v_end    PLS_INTEGER;
    v_idx    PLS_INTEGER := 0;
    v_str    VARCHAR2(32767) := p_input || DELIM;
    v_dlen   PLS_INTEGER     := LENGTH(DELIM);
  BEGIN
    LOOP
      v_end := INSTR(v_str, DELIM, v_pos);
      EXIT WHEN v_end = 0;
      v_idx := v_idx + 1;
      v_result(v_idx) := TRIM(SUBSTR(v_str, v_pos, v_end - v_pos));
      v_pos := v_end + v_dlen;
      EXIT WHEN v_pos > LENGTH(v_str);
    END LOOP;
    dbg('split_names: ' || v_idx || ' name(s) found');
    RETURN v_result;
  END split_names;

  /** Split a CLOB containing delimited SQL statements into t_clob_tab */
  FUNCTION split_queries(p_input IN CLOB) RETURN t_clob_tab IS
    v_result t_clob_tab;
    v_pos    NUMBER      := 1;
    v_end    NUMBER;
    v_idx    PLS_INTEGER := 0;
    v_total  NUMBER;
    v_dlen   PLS_INTEGER := LENGTH(DELIM);
    v_chunk  CLOB;
    v_frag_len NUMBER;
  BEGIN
    -- Append delimiter as CLOB character data (not RAW)
    DBMS_LOB.CREATETEMPORARY(v_chunk, TRUE);
    DBMS_LOB.APPEND(v_chunk, p_input);
    DBMS_LOB.APPEND(v_chunk, TO_CLOB(DELIM));
    v_total := DBMS_LOB.GETLENGTH(v_chunk);

    LOOP
      -- Search with VARCHAR2 pattern (correct overload for CLOB)
      v_end := DBMS_LOB.INSTR(v_chunk, DELIM, v_pos);
      EXIT WHEN v_end = 0 OR v_pos > v_total;

      v_frag_len := v_end - v_pos;
      IF v_frag_len > 0 THEN
        v_idx := v_idx + 1;
        DBMS_LOB.CREATETEMPORARY(v_result(v_idx), TRUE);
        -- Use DBMS_LOB.COPY for safe extraction of any-length fragment
        DBMS_LOB.COPY(v_result(v_idx), v_chunk, v_frag_len, 1, v_pos);
      END IF;

      v_pos := v_end + v_dlen;
    END LOOP;

    DBMS_LOB.FREETEMPORARY(v_chunk);
    dbg('split_queries: ' || v_idx || ' query/queries found');
    RETURN v_result;
  EXCEPTION
    WHEN OTHERS THEN
      BEGIN DBMS_LOB.FREETEMPORARY(v_chunk); EXCEPTION WHEN OTHERS THEN NULL; END;
      RAISE;
  END split_queries;

  -- ===========================================================================
  -- SECTION C: Sheet name derivation (mirrors PKG_XLSX_WRAPPER logic)
  -- ===========================================================================

  /** Extract unique table names from a SQL string, pipe-delimited */
  FUNCTION extract_tables(p_sql IN CLOB) RETURN VARCHAR2 IS
    v_upper  VARCHAR2(32767);
    v_result VARCHAR2(4000) := '';
    v_pos    PLS_INTEGER;
    v_end    PLS_INTEGER;
    v_token  VARCHAR2(200);

    c_skip CONSTANT VARCHAR2(500) :=
      '|SELECT|WHERE|SET|INTO|DUAL|LATERAL|ONLY|OUTER|INNER|CROSS|'
   || 'LEFT|RIGHT|FULL|NATURAL|ON|AND|OR|NOT|EXISTS|';

    FUNCTION known(p_t IN VARCHAR2) RETURN BOOLEAN IS
    BEGIN
      RETURN INSTR('|' || v_result || '|', '|' || p_t || '|') > 0;
    END;

    PROCEDURE try_add(p_raw IN VARCHAR2) IS
      v_t VARCHAR2(200) := UPPER(TRIM(p_raw));
    BEGIN
      IF INSTR(v_t, '.') > 0 THEN
        v_t := SUBSTR(v_t, INSTR(v_t, '.') + 1);
      END IF;
      v_t := REGEXP_REPLACE(v_t, '[^A-Z0-9_$#]', '');
      IF v_t IS NULL
         OR INSTR(c_skip, '|'||v_t||'|') > 0
         OR known(v_t)
      THEN RETURN; END IF;
      v_result := v_result
               || CASE WHEN v_result IS NOT NULL THEN '|' END
               || v_t;
    END;

  BEGIN
    v_upper := UPPER(REGEXP_REPLACE(SUBSTR(p_sql,1,32767), '\s+', ' '));
    FOR kw IN (SELECT COLUMN_VALUE AS KW
               FROM   TABLE(SYS.ODCIVARCHAR2LIST(' FROM ',' JOIN '))) LOOP
      v_pos := 1;
      LOOP
        v_pos := INSTR(v_upper, kw.KW, v_pos);
        EXIT WHEN v_pos = 0;
        v_pos := v_pos + LENGTH(kw.KW);
        WHILE v_pos <= LENGTH(v_upper)
          AND SUBSTR(v_upper, v_pos, 1) = ' ' LOOP
          v_pos := v_pos + 1;
        END LOOP;
        v_end := v_pos;
        WHILE v_end <= LENGTH(v_upper)
          AND SUBSTR(v_upper, v_end, 1)
              NOT IN (' ','(',')',',',CHR(10)) LOOP
          v_end := v_end + 1;
        END LOOP;
        v_token := SUBSTR(v_upper, v_pos, v_end - v_pos);
        IF v_token IS NOT NULL THEN try_add(v_token); END IF;
        v_pos := v_end;
      END LOOP;
    END LOOP;
    RETURN v_result;
  END extract_tables;

  /** Derive sheet name from SQL when caller did not supply one */
  FUNCTION derive_sheet_name(
    p_sql         IN     CLOB,
    p_complex_seq IN OUT NOCOPY PLS_INTEGER,
    p_sheet_pos   IN     PLS_INTEGER
  ) RETURN VARCHAR2 IS
    v_tables VARCHAR2(4000);
    v_count  PLS_INTEGER := 0;
    v_first  VARCHAR2(200);
    v_pos    PLS_INTEGER := 1;
    v_end    PLS_INTEGER;
  BEGIN
    v_tables := extract_tables(p_sql);
    IF v_tables IS NOT NULL THEN
      v_end   := INSTR(v_tables||'|','|',v_pos);
      v_first := SUBSTR(v_tables, v_pos, v_end - v_pos);
      v_count := 1;
      v_pos   := v_end + 1;
      LOOP
        v_end := INSTR(v_tables||'|','|',v_pos);
        EXIT WHEN v_end = 0 OR v_pos > LENGTH(v_tables);
        v_count := v_count + 1;
        v_pos   := v_end + 1;
      END LOOP;
    END IF;
    IF v_count = 1 THEN
      dbg('derive_sheet_name: 1 table → "' || SUBSTR(v_first,1,31) || '"');
      RETURN SUBSTR(v_first, 1, 31);
    ELSIF v_count > 1 THEN
      p_complex_seq := p_complex_seq + 1;
      dbg('derive_sheet_name: ' || v_count || ' tables → "COMPLEX_' || LPAD(p_complex_seq,2,'0') || '"');
      RETURN 'COMPLEX_' || LPAD(p_complex_seq, 2, '0');
    ELSE
      dbg('derive_sheet_name: no tables → "SHEET_' || p_sheet_pos || '"');
      RETURN 'SHEET_' || p_sheet_pos;
    END IF;
  END derive_sheet_name;

  /** Ensure sheet name is unique within the workbook session */
  FUNCTION unique_sheet_name(
    p_candidate  IN     VARCHAR2,
    p_used_names IN OUT NOCOPY VARCHAR2
  ) RETURN VARCHAR2 IS
    v_name VARCHAR2(31) := UPPER(SUBSTR(TRIM(p_candidate),1,31));
    v_base VARCHAR2(28) := SUBSTR(v_name,1,28);
    v_seq  PLS_INTEGER  := 1;
  BEGIN
    LOOP
      EXIT WHEN INSTR('|'||p_used_names||'|','|'||v_name||'|') = 0;
      dbg('unique_sheet_name: "' || p_candidate || '" duplicate → trying "' || v_base||'_'||LPAD(v_seq,2,'0') || '"');
      v_name := v_base||'_'||LPAD(v_seq,2,'0');
      v_seq  := v_seq + 1;
      IF v_seq > 999 THEN
        v_name := SUBSTR(v_base, 1, 22) || '_' || TO_CHAR(SYSDATE, 'HH24MISS');
        EXIT;
      END IF;
    END LOOP;
    p_used_names := p_used_names
                 || CASE WHEN p_used_names IS NOT NULL THEN '|' END
                 || v_name;
    dbg('unique_sheet_name: resolved → "' || v_name || '"');
    RETURN v_name;
  END unique_sheet_name;

  /** Default workbook name when none supplied */
  FUNCTION default_workbook_name RETURN VARCHAR2 IS
  BEGIN
    RETURN SYS_CONTEXT('USERENV','CURRENT_SCHEMA')
        || '_' || TO_CHAR(SYSDATE,'YYYYMMDD_HH24MISS');
  END;

  -- ===========================================================================
  -- SECTION D: Core builder — shared by both overloads
  -- ===========================================================================
  FUNCTION build_xlsx(
    p_workbook_name IN VARCHAR2,
    p_sheet_names   IN t_str_tab,   -- may be empty; auto-derived if missing
    p_sqls          IN t_clob_tab
  ) RETURN BLOB IS
    v_used_sheets VARCHAR2(4000) := '';
    v_complex_seq PLS_INTEGER    := 0;
    v_sheet_count PLS_INTEGER    := p_sqls.COUNT;
    v_sheet_name  VARCHAR2(31);
    v_wbname      VARCHAR2(200);
    v_blob        BLOB;
  BEGIN
    IF v_sheet_count = 0 THEN
      RAISE_APPLICATION_ERROR(-20010,
        'PKG_XLSX_DIRECT: No SQL queries provided.');
    END IF;

    -- Resolve workbook name
    v_wbname := NVL(NULLIF(TRIM(p_workbook_name),''),
                    default_workbook_name);
    dbg('build_xlsx: workbook="' || v_wbname || '" — ' || v_sheet_count || ' query/queries');

    -- Validate each query is a SELECT or WITH statement (prevent DML injection)
    FOR i IN 1..v_sheet_count LOOP
      DECLARE
        v_first_word VARCHAR2(10);
      BEGIN
        v_first_word := UPPER(TRIM(REGEXP_SUBSTR(
                         DBMS_LOB.SUBSTR(p_sqls(i), 100, 1),
                         '[A-Z]+')));
        IF v_first_word NOT IN ('SELECT', 'WITH') THEN
          RAISE_APPLICATION_ERROR(-20012,
            'PKG_XLSX_DIRECT: Only SELECT/WITH statements are permitted. Query ' || i ||
            ' starts with: ' || v_first_word);
        END IF;
        dbg('sql_guard[' || i || ']: starts with "' || v_first_word || '" — OK');
      END;
    END LOOP;

    -- Initialise core engine
    PKG_XLSX_EXPORT.init(v_wbname);

    -- Register each sheet
    FOR i IN 1..v_sheet_count LOOP

      -- Resolve sheet name: explicit > auto-derived
      IF p_sheet_names.EXISTS(i)
         AND TRIM(p_sheet_names(i)) IS NOT NULL
      THEN
        v_sheet_name := TRIM(p_sheet_names(i));
      ELSE
        v_sheet_name := derive_sheet_name(p_sqls(i), v_complex_seq, i);
      END IF;

      -- Enforce uniqueness within this workbook
      v_sheet_name := unique_sheet_name(v_sheet_name, v_used_sheets);

      dbg('build_xlsx: sheet ' || i || '/' || v_sheet_count || ' → "' || v_sheet_name || '"');
      PKG_XLSX_EXPORT.add_sheet(v_sheet_name, p_sqls(i));
    END LOOP;

    -- Build and return BLOB directly — no table write
    v_blob := PKG_XLSX_EXPORT.build_blob;
    dbg('build_xlsx: BLOB size=' || NVL(DBMS_LOB.GETLENGTH(v_blob),0) || ' bytes — done');
    RETURN v_blob;

  EXCEPTION
    WHEN OTHERS THEN
      RAISE_APPLICATION_ERROR(-20011,
        'PKG_XLSX_DIRECT.build_xlsx failed: ' || SQLERRM
        || CHR(10) || DBMS_UTILITY.FORMAT_ERROR_BACKTRACE);
  END build_xlsx;

  -- ===========================================================================
  -- SECTION E: get_delim — SQL-callable accessor for the DELIM constant
  -- ===========================================================================
  FUNCTION get_delim RETURN VARCHAR2 IS
  BEGIN
    RETURN DELIM;
  END get_delim;

  -- ===========================================================================
  -- SECTION F: Public overload 1 — auto sheet names
  -- ===========================================================================
  FUNCTION generate_xlsx(
    p_workbook_name IN VARCHAR2 DEFAULT NULL,
    p_queries       IN CLOB
  ) RETURN BLOB IS
    v_sqls       t_clob_tab;
    v_empty_names t_str_tab;   -- empty: all sheet names auto-derived
  BEGIN
    dbg('generate_xlsx(auto-names): p_workbook="' || NVL(p_workbook_name,'(auto)') || '"');
    v_sqls := split_queries(p_queries);
    RETURN build_xlsx(p_workbook_name, v_empty_names, v_sqls);
  END generate_xlsx;

  -- ===========================================================================
  -- SECTION G: Public overload 2 — explicit sheet names
  -- ===========================================================================
  FUNCTION generate_xlsx(
    p_workbook_name IN VARCHAR2 DEFAULT NULL,
    p_sheet_names   IN VARCHAR2,
    p_queries       IN CLOB
  ) RETURN BLOB IS
    v_sqls  t_clob_tab;
    v_names t_str_tab;
  BEGIN
    dbg('generate_xlsx(explicit-names): p_workbook="' || NVL(p_workbook_name,'(auto)') || '"');
    v_sqls  := split_queries(p_queries);
    v_names := split_names(p_sheet_names);
    RETURN build_xlsx(p_workbook_name, v_names, v_sqls);
  END generate_xlsx;

END PKG_XLSX_DIRECT;
/


-- =============================================================================
-- USAGE EXAMPLES
-- =============================================================================

/*

-- ════════════════════════════════════════════════════════════════════
-- EXAMPLE 1: Single query, workbook name supplied, sheet auto-named
--   Workbook → DEPOSIT_ISSUE
--   Sheet 1  → ICTM_ACC  (auto-derived from table name)
-- ════════════════════════════════════════════════════════════════════
SELECT PKG_XLSX_DIRECT.generate_xlsx(
           'Deposit_Issue',
           'SELECT * FROM ICTM_ACC WHERE ROWNUM <= 10'
       )
FROM DUAL;


-- ════════════════════════════════════════════════════════════════════
-- EXAMPLE 2: Multiple queries, sheet names auto-derived
--   Workbook → DEPOSIT_ISSUE
--   Sheet 1  → ICTM_ACC
--   Sheet 2  → STTM_CUST_ACCOUNT  (LIKE with character filter)
-- ════════════════════════════════════════════════════════════════════
SELECT PKG_XLSX_DIRECT.generate_xlsx(
           'Deposit_Issue',
           'SELECT * FROM ICTM_ACC WHERE ROWNUM <= 10'
           || PKG_XLSX_DIRECT.DELIM
           || q'[SELECT * FROM STTM_CUST_ACCOUNT WHERE CUST_NAME LIKE 'A%']'
       )
FROM DUAL;


-- ════════════════════════════════════════════════════════════════════
-- EXAMPLE 3: Multiple queries with explicit sheet names (overload 2)
--   Workbook → DEPOSIT_ISSUE
--   Sheet 1  → Accounts
--   Sheet 2  → Customers
-- ════════════════════════════════════════════════════════════════════
SELECT PKG_XLSX_DIRECT.generate_xlsx(
           'Deposit_Issue',
           -- sheet names (pipe-delimited, same order as queries)
           'Accounts' || PKG_XLSX_DIRECT.DELIM || 'Customers',
           -- queries
           'SELECT * FROM ICTM_ACC WHERE ROWNUM <= 10'
           || PKG_XLSX_DIRECT.DELIM
           || q'[SELECT * FROM STTM_CUST_ACCOUNT WHERE CUST_NAME LIKE 'A%']'
       )
FROM DUAL;


-- ════════════════════════════════════════════════════════════════════
-- EXAMPLE 4: No workbook name — auto-named SCHEMA_YYYYMMDD_HH24MISS
-- ════════════════════════════════════════════════════════════════════
SELECT PKG_XLSX_DIRECT.generate_xlsx(
           p_queries => 'SELECT EMPNO, ENAME, SAL FROM EMP ORDER BY ENAME'
       )
FROM DUAL;


-- ════════════════════════════════════════════════════════════════════
-- EXAMPLE 5: WHERE clause with a plain character filter
--   Regular string: single quotes doubled  '' A% ''
-- ════════════════════════════════════════════════════════════════════
SELECT PKG_XLSX_DIRECT.generate_xlsx(
           'EMP_Report',
           'SELECT EMPNO, ENAME, JOB, SAL FROM EMP WHERE ENAME LIKE ''A%'''
       )
FROM DUAL;

-- Same using q'[...]' — no escaping needed:
SELECT PKG_XLSX_DIRECT.generate_xlsx(
           'EMP_Report',
           q'[SELECT EMPNO, ENAME, JOB, SAL FROM EMP WHERE ENAME LIKE 'A%']'
       )
FROM DUAL;


-- ════════════════════════════════════════════════════════════════════
-- EXAMPLE 6: WHERE clause where value contains an apostrophe (O'BRIEN)
-- ════════════════════════════════════════════════════════════════════

-- q'[...]' recommended — write value naturally, only SQL '' needed:
SELECT PKG_XLSX_DIRECT.generate_xlsx(
           'EMP_Report',
           q'[SELECT EMPNO, ENAME, JOB FROM EMP WHERE ENAME = 'O''BRIEN']'
       )
FROM DUAL;

-- Regular string — every quote doubled:
SELECT PKG_XLSX_DIRECT.generate_xlsx(
           'EMP_Report',
           'SELECT EMPNO, ENAME, JOB FROM EMP WHERE ENAME = ''O''''BRIEN'''
       )
FROM DUAL;


-- ════════════════════════════════════════════════════════════════════
-- EXAMPLE 7: Value comes from a SQL bind / variable (PL/SQL block)
-- ════════════════════════════════════════════════════════════════════
DECLARE
  v_name VARCHAR2(100) := 'O''BRIEN';
  v_blob BLOB;
BEGIN
  v_blob := PKG_XLSX_DIRECT.generate_xlsx(
                'EMP_Report',
                'SELECT EMPNO, ENAME, JOB FROM EMP '
             || 'WHERE  ENAME = ' || CHR(39) || v_name || CHR(39)
            );
  -- use v_blob as needed (e.g. insert into another table, send via UTL_MAIL)
  DBMS_OUTPUT.PUT_LINE('Size: ' || DBMS_LOB.GETLENGTH(v_blob) || ' bytes');
END;
/


-- ════════════════════════════════════════════════════════════════════
-- HOW TO DOWNLOAD from SQL Developer / PL/SQL Developer
-- ════════════════════════════════════════════════════════════════════
--
-- 1. Run any of the SELECT ... FROM DUAL examples above
-- 2. The result grid shows one row with a BLOB column
-- 3. SQL Developer    : click the BLOB cell → pencil/download icon
--                       → Save As → rename to  YourWorkbook.xlsx
--    PL/SQL Developer : right-click BLOB cell → Save to File
--                       → rename to  YourWorkbook.xlsx
-- 4. Open in Excel — all sheets, headers, and data will be present.


-- ════════════════════════════════════════════════════════════════════
-- CHANGING THE DELIMITER
-- ════════════════════════════════════════════════════════════════════
-- Default delimiter is § (section sign).
-- If your SQL text contains § use a different delimiter by
-- replacing DELIM in the package body Section A constant, e.g:
--   DELIM CONSTANT VARCHAR2(3) := '|~|';
-- Then recompile the package.

*/
