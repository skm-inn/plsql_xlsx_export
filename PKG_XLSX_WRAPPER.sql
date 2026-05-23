-- =============================================================================
-- PKG_XLSX_WRAPPER
-- Smart wrapper over PKG_XLSX_EXPORT
--
-- This package owns ALL table-related objects:
--   XLSX_EXPORT_RESULTS  — stores generated BLOBs with UID tracking
--   XLSX_EXPORT_LOG      — audit trail
--
-- PKG_XLSX_EXPORT (core engine) has NO table references whatsoever.
-- This clean separation means PKG_XLSX_EXPORT compiles and runs in
-- read-only schemas too. Only deploy THIS package when CREATE TABLE
-- privilege is available.
--
-- NAMING RULES:
--   WORKBOOK NAME:
--     - Use p_workbook_name if supplied
--     - If not supplied: <CURRENT_SCHEMA>_YYYYMMDD_HH24MISS
--     - If name already exists in table: suffix _01, _02, ...
--
--   WORKSHEET NAME (each sheet/tab):
--     - Use p_sheetN_name if supplied
--     - If not supplied: extract primary table name from the SQL
--         * Single table in FROM  → that table name
--         * Multiple tables/JOINs → Complex_01, Complex_02, ...
--         * No parseable table    → Sheet_N (fallback)
--     - If sheet name duplicated within same workbook: suffix _01, _02, ...
--
-- OTHER FEATURES:
--   - Unique UID (SYS_GUID) per export for tracking and download
--   - Success notification via DBMS_OUTPUT + ready-to-run retrieval SQL
--   - Manual delete by UID
--   - Auto-purge exports older than 7 days (runs on every generate call)
--   - Full audit trail in XLSX_EXPORT_LOG
--
-- Depends on : PKG_XLSX_EXPORT (deploy first — no table needed for that)
-- Compatible : Oracle 12c, 19c
-- =============================================================================


-- =============================================================================
-- STEP 1: Create XLSX_EXPORT_RESULTS table (safe / idempotent)
--         THIS is where the table lives — not in PKG_XLSX_EXPORT.
-- =============================================================================
BEGIN
  EXECUTE IMMEDIATE q'[
    CREATE TABLE XLSX_EXPORT_RESULTS (
      EXPORT_ID   NUMBER        GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
      EXPORT_UID  VARCHAR2(32),
      EXPORT_NAME VARCHAR2(200) NOT NULL,
      CREATED_ON  TIMESTAMP     DEFAULT SYSTIMESTAMP,
      XLSX_BLOB   BLOB,
      STATUS      VARCHAR2(20)  DEFAULT 'PENDING',
      ERROR_MSG   VARCHAR2(4000)
    )
  ]';
EXCEPTION
  WHEN OTHERS THEN IF SQLCODE = -955 THEN NULL; ELSE RAISE; END IF;
END;
/

BEGIN
  EXECUTE IMMEDIATE
    'CREATE UNIQUE INDEX UX_XLSX_EXPORT_UID ON XLSX_EXPORT_RESULTS(EXPORT_UID)';
EXCEPTION
  WHEN OTHERS THEN IF SQLCODE = -955 THEN NULL; ELSE RAISE; END IF;
END;
/


-- =============================================================================
-- STEP 2: Audit log table (safe / idempotent)
-- =============================================================================
BEGIN
  EXECUTE IMMEDIATE q'[
    CREATE TABLE XLSX_EXPORT_LOG (
      LOG_ID     NUMBER GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
      EXPORT_UID VARCHAR2(32),
      ACTION     VARCHAR2(50),
      MESSAGE    VARCHAR2(4000),
      LOG_TIME   TIMESTAMP DEFAULT SYSTIMESTAMP
    )
  ]';
EXCEPTION
  WHEN OTHERS THEN IF SQLCODE = -955 THEN NULL; ELSE RAISE; END IF;
END;
/


-- =============================================================================
-- STEP 3: Package Specification
-- =============================================================================
CREATE OR REPLACE PACKAGE PKG_XLSX_WRAPPER AS

  /**
   * Main entry point. Generates a multi-sheet XLSX workbook.
   *
   * p_workbook_name : Optional.
   *                   Supplied  → used as workbook name (uniqueness enforced).
   *                   Omitted   → auto: CURRENT_SCHEMA_YYYYMMDD_HH24MISS
   *
   * p_sheetN_name   : Optional per sheet.
   *                   Supplied  → used as worksheet tab name.
   *                   Omitted   → auto-derived from the SQL's primary table:
   *                               single table  → table name
   *                               multi-table   → Complex_01, Complex_02, ...
   *                               unparseable   → Sheet_N
   *
   * Returns: EXPORT_UID (32-char hex) — use this to download or delete.
   */
  FUNCTION generate_xlsx(
    p_workbook_name IN VARCHAR2 DEFAULT NULL,
    -- Sheet 1 (required)
    p_sheet1_name   IN VARCHAR2 DEFAULT NULL,
    p_sheet1_sql    IN CLOB,
    -- Sheets 2-10 (optional)
    p_sheet2_name   IN VARCHAR2 DEFAULT NULL,  p_sheet2_sql  IN CLOB DEFAULT NULL,
    p_sheet3_name   IN VARCHAR2 DEFAULT NULL,  p_sheet3_sql  IN CLOB DEFAULT NULL,
    p_sheet4_name   IN VARCHAR2 DEFAULT NULL,  p_sheet4_sql  IN CLOB DEFAULT NULL,
    p_sheet5_name   IN VARCHAR2 DEFAULT NULL,  p_sheet5_sql  IN CLOB DEFAULT NULL,
    p_sheet6_name   IN VARCHAR2 DEFAULT NULL,  p_sheet6_sql  IN CLOB DEFAULT NULL,
    p_sheet7_name   IN VARCHAR2 DEFAULT NULL,  p_sheet7_sql  IN CLOB DEFAULT NULL,
    p_sheet8_name   IN VARCHAR2 DEFAULT NULL,  p_sheet8_sql  IN CLOB DEFAULT NULL,
    p_sheet9_name   IN VARCHAR2 DEFAULT NULL,  p_sheet9_sql  IN CLOB DEFAULT NULL,
    p_sheet10_name  IN VARCHAR2 DEFAULT NULL,  p_sheet10_sql IN CLOB DEFAULT NULL
  ) RETURN VARCHAR2;

  /** Delete a specific export by UID. */
  PROCEDURE delete_export(p_export_uid IN VARCHAR2);

  /** Purge exports older than p_days days (default 7). */
  PROCEDURE purge_old_exports(p_days IN NUMBER DEFAULT 7);

  /** Print all exports to DBMS_OUTPUT. */
  PROCEDURE list_exports;

  /** Re-print the download SQL for a given UID. */
  PROCEDURE print_retrieval_sql(p_export_uid IN VARCHAR2);

END PKG_XLSX_WRAPPER;
/


-- =============================================================================
-- STEP 4: Package Body
-- =============================================================================
CREATE OR REPLACE PACKAGE BODY PKG_XLSX_WRAPPER AS

  -- ===========================================================================
  -- SECTION A: Utility helpers
  -- ===========================================================================

  PROCEDURE print_line(p_len IN PLS_INTEGER DEFAULT 68,
                       p_ch  IN VARCHAR2    DEFAULT '-') IS
  BEGIN
    DBMS_OUTPUT.PUT_LINE(RPAD(p_ch, p_len, p_ch));
  END;

  -- Autonomous so it commits independently of the main transaction
  PROCEDURE log_action(p_uid IN VARCHAR2,
                       p_act IN VARCHAR2,
                       p_msg IN VARCHAR2) IS
    PRAGMA AUTONOMOUS_TRANSACTION;
  BEGIN
    INSERT INTO XLSX_EXPORT_LOG(EXPORT_UID, ACTION, MESSAGE)
    VALUES (p_uid, p_act, SUBSTR(p_msg, 1, 4000));
    COMMIT;
  END;

  -- ===========================================================================
  -- SECTION B: Workbook name resolution
  -- ===========================================================================

  /**
   * Returns a workbook name that does not already exist in XLSX_EXPORT_RESULTS.
   * If p_base already exists, tries p_base_01, p_base_02, ...
   */
  FUNCTION unique_workbook_name(p_base IN VARCHAR2) RETURN VARCHAR2 IS
    v_name    VARCHAR2(200) := UPPER(TRIM(p_base));
    v_exists  PLS_INTEGER;
    v_seq     PLS_INTEGER := 1;
  BEGIN
    LOOP
      SELECT COUNT(*) INTO v_exists
      FROM   XLSX_EXPORT_RESULTS
      WHERE  UPPER(EXPORT_NAME) = v_name;
      EXIT WHEN v_exists = 0;
      v_name := UPPER(TRIM(p_base)) || '_' || LPAD(v_seq, 2, '0');
      v_seq  := v_seq + 1;
      EXIT WHEN v_seq > 999;        -- safety cap
    END LOOP;
    RETURN v_name;
  END;

  /**
   * Derive default workbook name when none supplied:
   *   <CURRENT_SCHEMA>_YYYYMMDD_HH24MISS
   */
  FUNCTION default_workbook_name RETURN VARCHAR2 IS
  BEGIN
    RETURN SYS_CONTEXT('USERENV', 'CURRENT_SCHEMA')
        || '_'
        || TO_CHAR(SYSDATE, 'YYYYMMDD_HH24MISS');
  END;

  -- ===========================================================================
  -- SECTION C: Worksheet name resolution (the key new logic)
  -- ===========================================================================

  /**
   * Extract all table/view names referenced in a SQL string.
   * Looks at tokens immediately after FROM and JOIN keywords.
   * Returns a pipe-delimited list e.g.  EMPLOYEES|DEPARTMENTS
   * Schema prefixes (SCHEMA.TABLE) are stripped to just TABLE.
   */
  FUNCTION extract_tables_from_sql(p_sql IN CLOB) RETURN VARCHAR2 IS

    v_upper   VARCHAR2(32767);
    v_result  VARCHAR2(4000) := '';
    v_pos     PLS_INTEGER;
    v_end     PLS_INTEGER;
    v_token   VARCHAR2(200);

    -- Keywords that immediately follow FROM/JOIN but are NOT table names
    c_skip CONSTANT VARCHAR2(500) :=
      '|SELECT|WHERE|SET|INTO|DUAL|LATERAL|ONLY|OUTER|INNER|CROSS|'
   || 'LEFT|RIGHT|FULL|NATURAL|ON|AND|OR|NOT|EXISTS|';

    FUNCTION already_in(p_tbl IN VARCHAR2) RETURN BOOLEAN IS
    BEGIN
      RETURN INSTR('|' || v_result || '|', '|' || p_tbl || '|') > 0;
    END;

    PROCEDURE try_add(p_raw IN VARCHAR2) IS
      v_tbl VARCHAR2(200);
    BEGIN
      -- Strip schema prefix
      v_tbl := UPPER(TRIM(p_raw));
      IF INSTR(v_tbl, '.') > 0 THEN
        v_tbl := SUBSTR(v_tbl, INSTR(v_tbl, '.') + 1);
      END IF;
      -- Strip alias hint characters
      v_tbl := REGEXP_REPLACE(v_tbl, '[^A-Z0-9_$#]', '');
      IF v_tbl IS NULL
         OR INSTR(c_skip, '|' || v_tbl || '|') > 0
         OR already_in(v_tbl)
      THEN RETURN;
      END IF;
      v_result := v_result
               || CASE WHEN v_result IS NOT NULL THEN '|' END
               || v_tbl;
    END;

  BEGIN
    -- Normalise whitespace
    v_upper := UPPER(REGEXP_REPLACE(SUBSTR(p_sql, 1, 32767), '\s+', ' '));

    FOR kw IN (SELECT COLUMN_VALUE AS KW
               FROM   TABLE(SYS.ODCIVARCHAR2LIST(' FROM ',' JOIN '))) LOOP
      v_pos := 1;
      LOOP
        v_pos := INSTR(v_upper, kw.KW, v_pos);
        EXIT WHEN v_pos = 0;
        v_pos := v_pos + LENGTH(kw.KW);

        -- Skip leading spaces
        WHILE v_pos <= LENGTH(v_upper)
          AND SUBSTR(v_upper, v_pos, 1) = ' ' LOOP
          v_pos := v_pos + 1;
        END LOOP;

        -- Read next token (stop at space / ( / , )
        v_end := v_pos;
        WHILE v_end <= LENGTH(v_upper)
          AND SUBSTR(v_upper, v_end, 1) NOT IN (' ','(',')',',',CHR(10)) LOOP
          v_end := v_end + 1;
        END LOOP;

        v_token := SUBSTR(v_upper, v_pos, v_end - v_pos);
        IF v_token IS NOT NULL THEN try_add(v_token); END IF;
        v_pos := v_end;
      END LOOP;
    END LOOP;

    RETURN v_result;
  END extract_tables_from_sql;

  /**
   * Derive a worksheet name from a SQL string when caller did not supply one.
   *
   *   1 unique table  → that table name  (max 31 chars)
   *   2+ tables       → Complex_NN  where NN auto-increments across
   *                     all sheets in the current workbook session
   *   0 tables found  → Sheet_N  (positional fallback)
   *
   * p_complex_seq : IN/OUT counter so successive complex sheets get
   *                 different numbers within the same workbook.
   * p_sheet_pos   : 1-based position of this sheet (for Sheet_N fallback).
   */
  FUNCTION derive_sheet_name(
    p_sql         IN     CLOB,
    p_complex_seq IN OUT NOCOPY PLS_INTEGER,
    p_sheet_pos   IN     PLS_INTEGER
  ) RETURN VARCHAR2 IS
    v_tables  VARCHAR2(4000);
    v_count   PLS_INTEGER := 0;
    v_first   VARCHAR2(200);
    v_pos     PLS_INTEGER := 1;
    v_end     PLS_INTEGER;
  BEGIN
    v_tables := extract_tables_from_sql(p_sql);

    -- Count pipe-delimited tokens
    IF v_tables IS NOT NULL THEN
      v_end := INSTR(v_tables || '|', '|', v_pos);
      v_first := SUBSTR(v_tables, v_pos, v_end - v_pos);
      v_count := 1;
      v_pos   := v_end + 1;
      LOOP
        v_end := INSTR(v_tables || '|', '|', v_pos);
        EXIT WHEN v_end = 0 OR v_pos > LENGTH(v_tables);
        v_count := v_count + 1;
        v_pos   := v_end + 1;
      END LOOP;
    END IF;

    IF v_count = 1 THEN
      RETURN SUBSTR(v_first, 1, 31);
    ELSIF v_count > 1 THEN
      p_complex_seq := p_complex_seq + 1;
      RETURN 'COMPLEX_' || LPAD(p_complex_seq, 2, '0');
    ELSE
      RETURN 'SHEET_' || p_sheet_pos;
    END IF;
  END derive_sheet_name;

  /**
   * Ensure a worksheet name is unique within the current workbook session.
   * Uses a simple in-memory pipe-delimited list of already-used names.
   * If duplicate found → append _01, _02, ...
   */
  FUNCTION unique_sheet_name(
    p_candidate  IN     VARCHAR2,
    p_used_names IN OUT NOCOPY VARCHAR2   -- pipe-delimited running list
  ) RETURN VARCHAR2 IS
    v_name  VARCHAR2(31) := UPPER(SUBSTR(TRIM(p_candidate), 1, 31));
    v_base  VARCHAR2(28);
    v_seq   PLS_INTEGER := 1;
  BEGIN
    -- Trim base to leave room for _NN suffix (31 - 3 = 28)
    v_base := SUBSTR(v_name, 1, 28);

    LOOP
      EXIT WHEN INSTR('|' || p_used_names || '|',
                      '|' || v_name || '|') = 0;
      v_name := v_base || '_' || LPAD(v_seq, 2, '0');
      v_seq  := v_seq + 1;
      IF v_seq > 999 THEN
        v_name := SUBSTR(v_base, 1, 22) || '_' || TO_CHAR(SYSDATE, 'HH24MISS');
        EXIT;
      END IF;
    END LOOP;

    p_used_names := p_used_names
                 || CASE WHEN p_used_names IS NOT NULL THEN '|' END
                 || v_name;
    RETURN v_name;
  END unique_sheet_name;


  -- ===========================================================================
  -- SECTION D: Public utilities
  -- ===========================================================================

  PROCEDURE print_retrieval_sql(p_export_uid IN VARCHAR2) IS
    v_name  VARCHAR2(200);
    v_id    NUMBER;
    v_bytes NUMBER;
  BEGIN
    BEGIN
      SELECT EXPORT_NAME, EXPORT_ID,
             NVL(DBMS_LOB.GETLENGTH(XLSX_BLOB), 0)
      INTO   v_name, v_id, v_bytes
      FROM   XLSX_EXPORT_RESULTS
      WHERE  EXPORT_UID = p_export_uid;
    EXCEPTION
      WHEN NO_DATA_FOUND THEN
        DBMS_OUTPUT.PUT_LINE('No export found for UID: ' || p_export_uid);
        RETURN;
    END;

    DBMS_OUTPUT.PUT_LINE('');
    print_line(68, '=');
    DBMS_OUTPUT.PUT_LINE('  EXPORT COMPLETE');
    print_line(68, '-');
    DBMS_OUTPUT.PUT_LINE('  Workbook  : ' || v_name || '.xlsx');
    DBMS_OUTPUT.PUT_LINE('  Export ID : ' || v_id);
    DBMS_OUTPUT.PUT_LINE('  UID       : ' || p_export_uid);
    DBMS_OUTPUT.PUT_LINE('  Size      : ' || ROUND(v_bytes / 1024, 1) || ' KB');
    print_line(68, '-');
    DBMS_OUTPUT.PUT_LINE('  DOWNLOAD — run in SQL Developer / PL/SQL Developer:');
    DBMS_OUTPUT.PUT_LINE('');
    DBMS_OUTPUT.PUT_LINE('    SELECT XLSX_BLOB');
    DBMS_OUTPUT.PUT_LINE('    FROM   XLSX_EXPORT_RESULTS');
    DBMS_OUTPUT.PUT_LINE('    WHERE  EXPORT_UID = ''' || p_export_uid || ''';');
    DBMS_OUTPUT.PUT_LINE('');
    DBMS_OUTPUT.PUT_LINE('  Save the BLOB cell as: ' || v_name || '.xlsx');
    print_line(68, '-');
    DBMS_OUTPUT.PUT_LINE('  DELETE when done:');
    DBMS_OUTPUT.PUT_LINE('');
    DBMS_OUTPUT.PUT_LINE('    EXEC PKG_XLSX_WRAPPER.delete_export('''
                      || p_export_uid || ''');');
    print_line(68, '=');
    DBMS_OUTPUT.PUT_LINE('');
  END print_retrieval_sql;

  -- ---------------------------------------------------------------------------

  PROCEDURE delete_export(p_export_uid IN VARCHAR2) IS
    v_name  VARCHAR2(200);
    v_cnt   PLS_INTEGER;
  BEGIN
    SELECT COUNT(*), MAX(EXPORT_NAME)
    INTO   v_cnt, v_name
    FROM   XLSX_EXPORT_RESULTS
    WHERE  EXPORT_UID = p_export_uid;

    IF v_cnt = 0 THEN
      DBMS_OUTPUT.PUT_LINE('[DELETE] UID not found: ' || p_export_uid);
      RETURN;
    END IF;

    DELETE FROM XLSX_EXPORT_RESULTS WHERE EXPORT_UID = p_export_uid;
    COMMIT;
    log_action(p_export_uid, 'DELETED', 'Manually deleted: ' || v_name);
    DBMS_OUTPUT.PUT_LINE('[DELETE] Removed: ' || v_name
                      || '  (UID: ' || p_export_uid || ')');
  END delete_export;

  -- ---------------------------------------------------------------------------

  PROCEDURE purge_old_exports(p_days IN NUMBER DEFAULT 7) IS
    v_purged PLS_INTEGER := 0;
  BEGIN
    DELETE FROM XLSX_EXPORT_RESULTS
    WHERE  CREATED_ON < SYSTIMESTAMP - NUMTODSINTERVAL(p_days, 'DAY');
    v_purged := SQL%ROWCOUNT;
    COMMIT;
    IF v_purged > 0 THEN
      log_action(NULL, 'PURGED',
                 v_purged || ' export(s) auto-purged (>' || p_days || ' days)');
    END IF;
    DELETE FROM XLSX_EXPORT_LOG
    WHERE LOG_TIME < SYSTIMESTAMP - NUMTODSINTERVAL(p_days * 2, 'DAY');
    COMMIT;
    DBMS_OUTPUT.PUT_LINE('[PURGE] ' || v_purged
      || ' export(s) older than ' || p_days || ' day(s) removed.');
  END purge_old_exports;

  -- ---------------------------------------------------------------------------

  PROCEDURE list_exports IS
    v_cnt PLS_INTEGER := 0;
  BEGIN
    print_line(68, '=');
    DBMS_OUTPUT.PUT_LINE('  XLSX EXPORT REPOSITORY');
    print_line(68, '-');
    DBMS_OUTPUT.PUT_LINE(
      RPAD('ID',  5) || '  ' || RPAD('WORKBOOK NAME', 32) || '  ' ||
      RPAD('STATUS', 9) || '  ' || RPAD('KB', 8) || '  CREATED');
    print_line(68, '-');

    FOR r IN (
      SELECT EXPORT_ID,
             EXPORT_NAME,
             EXPORT_UID,
             STATUS,
             ROUND(NVL(DBMS_LOB.GETLENGTH(XLSX_BLOB),0) / 1024, 1) AS KB,
             TO_CHAR(CREATED_ON, 'YYYY-MM-DD HH24:MI:SS') AS DT
      FROM   XLSX_EXPORT_RESULTS
      ORDER  BY CREATED_ON DESC
    ) LOOP
      DBMS_OUTPUT.PUT_LINE(
        RPAD(r.EXPORT_ID, 5) || '  ' ||
        RPAD(SUBSTR(r.EXPORT_NAME, 1, 32), 32) || '  ' ||
        RPAD(r.STATUS, 9) || '  ' ||
        RPAD(r.KB, 8) || '  ' || r.DT
      );
      DBMS_OUTPUT.PUT_LINE('         UID: ' || r.EXPORT_UID);
      v_cnt := v_cnt + 1;
    END LOOP;

    IF v_cnt = 0 THEN
      DBMS_OUTPUT.PUT_LINE('  (no exports found)');
    END IF;
    print_line(68, '-');
    DBMS_OUTPUT.PUT_LINE('  Total: ' || v_cnt || ' export(s)');
    print_line(68, '=');
  END list_exports;


  -- ===========================================================================
  -- SECTION E: generate_xlsx — main entry point
  -- ===========================================================================
  FUNCTION generate_xlsx(
    p_workbook_name IN VARCHAR2 DEFAULT NULL,
    p_sheet1_name   IN VARCHAR2 DEFAULT NULL,  p_sheet1_sql  IN CLOB,
    p_sheet2_name   IN VARCHAR2 DEFAULT NULL,  p_sheet2_sql  IN CLOB DEFAULT NULL,
    p_sheet3_name   IN VARCHAR2 DEFAULT NULL,  p_sheet3_sql  IN CLOB DEFAULT NULL,
    p_sheet4_name   IN VARCHAR2 DEFAULT NULL,  p_sheet4_sql  IN CLOB DEFAULT NULL,
    p_sheet5_name   IN VARCHAR2 DEFAULT NULL,  p_sheet5_sql  IN CLOB DEFAULT NULL,
    p_sheet6_name   IN VARCHAR2 DEFAULT NULL,  p_sheet6_sql  IN CLOB DEFAULT NULL,
    p_sheet7_name   IN VARCHAR2 DEFAULT NULL,  p_sheet7_sql  IN CLOB DEFAULT NULL,
    p_sheet8_name   IN VARCHAR2 DEFAULT NULL,  p_sheet8_sql  IN CLOB DEFAULT NULL,
    p_sheet9_name   IN VARCHAR2 DEFAULT NULL,  p_sheet9_sql  IN CLOB DEFAULT NULL,
    p_sheet10_name  IN VARCHAR2 DEFAULT NULL,  p_sheet10_sql IN CLOB DEFAULT NULL
  ) RETURN VARCHAR2 IS

    TYPE t_str_tab  IS TABLE OF VARCHAR2(31)  INDEX BY PLS_INTEGER;
    TYPE t_clob_tab IS TABLE OF CLOB          INDEX BY PLS_INTEGER;

    v_raw_names   t_str_tab;
    v_sqls        t_clob_tab;
    v_final_names t_str_tab;
    v_cnt         PLS_INTEGER := 0;

    v_uid          VARCHAR2(32);
    v_export_id    NUMBER;
    v_xlsx_blob    BLOB;
    v_workbook     VARCHAR2(200);
    v_used_sheets  VARCHAR2(4000) := '';  -- tracks used sheet names within workbook
    v_complex_seq  PLS_INTEGER    := 0;  -- counter for Complex_NN sheet names
    v_derived      VARCHAR2(31);

  BEGIN

    -- ── 0. Auto-purge silently ─────────────────────────────────────────────
    BEGIN purge_old_exports(7); EXCEPTION WHEN OTHERS THEN NULL; END;

    -- ── 1. UID ────────────────────────────────────────────────────────────
    v_uid := RAWTOHEX(SYS_GUID());

    -- ── 2. Collect sheets into arrays ─────────────────────────────────────
    -- Sheet 1 is mandatory (p_sheet1_sql has no DEFAULT NULL)
    v_cnt := 1;
    v_raw_names(1) := p_sheet1_name;
    v_sqls(1)      := p_sheet1_sql;

    -- Sheets 2-10 included only when SQL is provided
    IF p_sheet2_sql  IS NOT NULL THEN v_cnt:=v_cnt+1; v_raw_names(v_cnt):=p_sheet2_name;  v_sqls(v_cnt):=p_sheet2_sql;  END IF;
    IF p_sheet3_sql  IS NOT NULL THEN v_cnt:=v_cnt+1; v_raw_names(v_cnt):=p_sheet3_name;  v_sqls(v_cnt):=p_sheet3_sql;  END IF;
    IF p_sheet4_sql  IS NOT NULL THEN v_cnt:=v_cnt+1; v_raw_names(v_cnt):=p_sheet4_name;  v_sqls(v_cnt):=p_sheet4_sql;  END IF;
    IF p_sheet5_sql  IS NOT NULL THEN v_cnt:=v_cnt+1; v_raw_names(v_cnt):=p_sheet5_name;  v_sqls(v_cnt):=p_sheet5_sql;  END IF;
    IF p_sheet6_sql  IS NOT NULL THEN v_cnt:=v_cnt+1; v_raw_names(v_cnt):=p_sheet6_name;  v_sqls(v_cnt):=p_sheet6_sql;  END IF;
    IF p_sheet7_sql  IS NOT NULL THEN v_cnt:=v_cnt+1; v_raw_names(v_cnt):=p_sheet7_name;  v_sqls(v_cnt):=p_sheet7_sql;  END IF;
    IF p_sheet8_sql  IS NOT NULL THEN v_cnt:=v_cnt+1; v_raw_names(v_cnt):=p_sheet8_name;  v_sqls(v_cnt):=p_sheet8_sql;  END IF;
    IF p_sheet9_sql  IS NOT NULL THEN v_cnt:=v_cnt+1; v_raw_names(v_cnt):=p_sheet9_name;  v_sqls(v_cnt):=p_sheet9_sql;  END IF;
    IF p_sheet10_sql IS NOT NULL THEN v_cnt:=v_cnt+1; v_raw_names(v_cnt):=p_sheet10_name; v_sqls(v_cnt):=p_sheet10_sql; END IF;

    -- ── 3. Resolve WORKBOOK name ───────────────────────────────────────────
    v_workbook := unique_workbook_name(
                    CASE WHEN p_workbook_name IS NOT NULL
                         THEN p_workbook_name
                         ELSE default_workbook_name
                    END
                  );

    -- ── 4. Resolve WORKSHEET names ─────────────────────────────────────────
    FOR i IN 1..v_cnt LOOP
      IF v_raw_names(i) IS NOT NULL THEN
        -- Caller provided a name — just enforce uniqueness within workbook
        v_derived := v_raw_names(i);
      ELSE
        -- Auto-derive from SQL
        v_derived := derive_sheet_name(v_sqls(i), v_complex_seq, i);
      END IF;
      -- Ensure no duplicate tab names within this workbook
      v_final_names(i) := unique_sheet_name(v_derived, v_used_sheets);
    END LOOP;

    -- ── 5. Log + announce ─────────────────────────────────────────────────
    log_action(v_uid, 'INITIATED',
               'Workbook: ' || v_workbook || ' | Sheets: ' || v_cnt);

    DBMS_OUTPUT.PUT_LINE('');
    print_line(68, '=');
    DBMS_OUTPUT.PUT_LINE('  Generating: ' || v_workbook || '.xlsx');
    DBMS_OUTPUT.PUT_LINE('  UID   : ' || v_uid);
    DBMS_OUTPUT.PUT_LINE('  Sheets: ' || v_cnt);
    print_line(68, '-');
    FOR i IN 1..v_cnt LOOP
      DBMS_OUTPUT.PUT_LINE('  Sheet ' || i || ' : ' || v_final_names(i));
    END LOOP;
    print_line(68, '=');

    -- ── 6. Build XLSX via base package (returns BLOB — no table touch) ──────
    PKG_XLSX_EXPORT.init(v_workbook);
    FOR i IN 1..v_cnt LOOP
      PKG_XLSX_EXPORT.add_sheet(v_final_names(i), v_sqls(i));
    END LOOP;

    -- build_blob: pure engine call — zero table dependency
    v_xlsx_blob := PKG_XLSX_EXPORT.build_blob;

    -- ── 7. Wrapper owns the INSERT — table lives here, not in core engine ────
    INSERT INTO XLSX_EXPORT_RESULTS
           (EXPORT_UID, EXPORT_NAME, XLSX_BLOB, STATUS)
    VALUES (v_uid, v_workbook, v_xlsx_blob, 'COMPLETE')
    RETURNING EXPORT_ID INTO v_export_id;
    COMMIT;

    log_action(v_uid, 'COMPLETED',
               'Export ID: ' || v_export_id || ' | ' || v_workbook);

    -- ── 8. Success notification ───────────────────────────────────────────
    print_retrieval_sql(v_uid);

    RETURN v_uid;

  EXCEPTION
    WHEN OTHERS THEN
      log_action(v_uid, 'FAILED', SQLERRM);
      DBMS_OUTPUT.PUT_LINE('[ERROR] Export failed. Check XLSX_EXPORT_LOG for UID: ' || v_uid);
      RAISE;
  END generate_xlsx;

END PKG_XLSX_WRAPPER;
/


-- =============================================================================
-- USAGE EXAMPLES
-- =============================================================================

/*

-- ════════════════════════════════════════════════════════════════════
-- EXAMPLE 1: No workbook name, no sheet names
--   Workbook → MYSCHEMA_20240522_143055
--   Sheet 1  → USER_OBJECTS        (single table)
--   Sheet 2  → COMPLEX_01          (USER_SEGMENTS + USER_EXTENTS)
--   Sheet 3  → USER_SOURCE         (single table)
-- ════════════════════════════════════════════════════════════════════
DECLARE
  v_uid VARCHAR2(32);
BEGIN
  v_uid := PKG_XLSX_WRAPPER.generate_xlsx(
    p_sheet1_sql  => 'SELECT OBJECT_NAME, OBJECT_TYPE, STATUS, CREATED
                      FROM USER_OBJECTS ORDER BY OBJECT_TYPE',
    p_sheet2_sql  => 'SELECT S.SEGMENT_NAME, S.BYTES, E.EXTENT_ID
                      FROM USER_SEGMENTS S
                      JOIN USER_EXTENTS E ON E.SEGMENT_NAME = S.SEGMENT_NAME
                      WHERE ROWNUM <= 100',
    p_sheet3_sql  => 'SELECT NAME, TYPE, LINE, TEXT
                      FROM USER_SOURCE WHERE ROWNUM <= 200'
  );
  DBMS_OUTPUT.PUT_LINE('UID: ' || v_uid);
END;
/


-- ════════════════════════════════════════════════════════════════════
-- EXAMPLE 2: Workbook name supplied, sheet names auto-derived
--   Workbook → HR_MONTHLY_REPORT  (or HR_MONTHLY_REPORT_01 if exists)
--   Sheet 1  → EMPLOYEES
--   Sheet 2  → COMPLEX_01   (EMPLOYEES + DEPARTMENTS joined)
-- ════════════════════════════════════════════════════════════════════
DECLARE
  v_uid VARCHAR2(32);
BEGIN
  v_uid := PKG_XLSX_WRAPPER.generate_xlsx(
    p_workbook_name => 'HR_MONTHLY_REPORT',
    p_sheet1_sql    => 'SELECT EMPLOYEE_ID, FIRST_NAME, LAST_NAME, SALARY
                        FROM EMPLOYEES ORDER BY LAST_NAME',
    p_sheet2_sql    => q'[SELECT E.FIRST_NAME||' '||E.LAST_NAME AS NAME,
                                 D.DEPARTMENT_NAME, E.SALARY
                          FROM   EMPLOYEES E
                          JOIN   DEPARTMENTS D
                            ON   D.DEPARTMENT_ID = E.DEPARTMENT_ID
                          ORDER  BY E.SALARY DESC]'
  );
END;
/


-- ════════════════════════════════════════════════════════════════════
-- EXAMPLE 3: All names explicitly provided
--   Workbook → AUDIT_PACK
--   Sheet 1  → Access Log
--   Sheet 2  → Error Summary
-- ════════════════════════════════════════════════════════════════════
DECLARE
  v_uid VARCHAR2(32);
BEGIN
  v_uid := PKG_XLSX_WRAPPER.generate_xlsx(
    p_workbook_name => 'AUDIT_PACK',
    p_sheet1_name   => 'Access Log',
    p_sheet1_sql    => 'SELECT * FROM ACCESS_LOG   WHERE ROWNUM <= 500',
    p_sheet2_name   => 'Error Summary',
    p_sheet2_sql    => 'SELECT * FROM ERROR_SUMMARY WHERE ROWNUM <= 500'
  );
END;
/


-- ════════════════════════════════════════════════════════════════════
-- EXAMPLE 4: WHERE clause with a character parameter (LIKE filter)
--
-- Scenario : Export employees whose name starts with 'A'.
--
-- Rule     : A string literal inside a SQL passed as a PL/SQL VARCHAR2
--            needs its single quotes doubled up ( ' → '' ).
--
--   Plain SQL you would type:
--       WHERE ENAME LIKE 'A%'
--
--   Inside a PL/SQL quoted string, each ' becomes '':
--       WHERE ENAME LIKE ''A%''
--
-- Workbook → FILTERED_EMPLOYEES
-- Sheet 1  → EMP_NAME_A      (manual sheet name supplied)
-- ════════════════════════════════════════════════════════════════════
DECLARE
  v_uid VARCHAR2(32);
BEGIN
  v_uid := PKG_XLSX_WRAPPER.generate_xlsx(
    p_workbook_name => 'FILTERED_EMPLOYEES',
    p_sheet1_name   => 'EMP_NAME_A',
    p_sheet1_sql    =>
        'SELECT EMPNO, ENAME, JOB, SAL, HIREDATE ' ||
        'FROM   EMP '                               ||
        'WHERE  ENAME LIKE ''A%'' '                 ||
        'ORDER  BY ENAME'
  );
  DBMS_OUTPUT.PUT_LINE('UID: ' || v_uid);
END;
/

-- ── Alternative: use the q'[...]' quoting mechanism ──────────────────
-- q'[...]' lets you write single quotes naturally inside the brackets
-- without doubling them.  Choose whichever style you find more readable.
DECLARE
  v_uid VARCHAR2(32);
BEGIN
  v_uid := PKG_XLSX_WRAPPER.generate_xlsx(
    p_workbook_name => 'FILTERED_EMPLOYEES',
    p_sheet1_name   => 'EMP_NAME_A',
    p_sheet1_sql    => q'[
        SELECT EMPNO, ENAME, JOB, SAL, HIREDATE
        FROM   EMP
        WHERE  ENAME LIKE 'A%'
        ORDER  BY ENAME
    ]'
  );
  DBMS_OUTPUT.PUT_LINE('UID: ' || v_uid);
END;
/


-- ════════════════════════════════════════════════════════════════════
-- EXAMPLE 5: WHERE clause where the value itself contains a single quote
--
-- Scenario : Filter by a name that contains an apostrophe, e.g. O'BRIEN.
--
-- The value  O'BRIEN  has one embedded quote.
-- Each embedded quote must be escaped differently depending on technique:
--
--   Technique A — doubled quotes inside a regular string:
--       Plain value : O'BRIEN
--       In SQL text : 'O''BRIEN'        ← outer quotes + inner quote doubled
--       In PL/SQL   : '''O''''BRIEN'''  ← every quote in the whole string doubled
--       Confusing!  Use Technique B instead for values with embedded quotes.
--
--   Technique B — q'[...]' alternative quoting (recommended):
--       Write the value exactly as it appears; no escaping needed inside q'[...]'.
--       Only restriction: the bracket pair [ ] must not appear in your SQL text
--       (use q'{ }' or q'< >' if your SQL contains square brackets).
--
-- Workbook → APOSTROPHE_DEMO
-- Sheet 1  → EMP_OBRIEN
-- ════════════════════════════════════════════════════════════════════

-- ── Technique A: doubled-quote escaping (regular string) ─────────────
-- Breakdown of  '''O''''BRIEN'''  :
--   ''       → opening  '  of the SQL literal
--   O        → the letter O
--   ''''     → the embedded apostrophe  '  (two pairs: close + reopen string)
--   BRIEN    → rest of the name
--   ''       → closing  '  of the SQL literal
DECLARE
  v_uid VARCHAR2(32);
BEGIN
  v_uid := PKG_XLSX_WRAPPER.generate_xlsx(
    p_workbook_name => 'APOSTROPHE_DEMO',
    p_sheet1_name   => 'EMP_OBRIEN',
    p_sheet1_sql    =>
        'SELECT EMPNO, ENAME, JOB, SAL '  ||
        'FROM   EMP '                     ||
        'WHERE  ENAME = ''O''''BRIEN'' '  ||   -- represents: ENAME = 'O''BRIEN'
        'ORDER  BY EMPNO'                      -- which in SQL means: O'BRIEN
  );
  DBMS_OUTPUT.PUT_LINE('UID: ' || v_uid);
END;
/

-- ── Technique B: q'[...]' alternative quoting (recommended) ──────────
-- Write O'BRIEN naturally; no escaping needed.
DECLARE
  v_uid VARCHAR2(32);
BEGIN
  v_uid := PKG_XLSX_WRAPPER.generate_xlsx(
    p_workbook_name => 'APOSTROPHE_DEMO',
    p_sheet1_name   => 'EMP_OBRIEN',
    p_sheet1_sql    => q'[
        SELECT EMPNO, ENAME, JOB, SAL
        FROM   EMP
        WHERE  ENAME = 'O''BRIEN'
        ORDER  BY EMPNO
    ]'
  );
  -- Inside q'[...]' the outer delimiters are [ and ].
  -- The SQL string itself still needs  ''  for the embedded apostrophe
  -- because that escaping is part of SQL syntax, not PL/SQL syntax.
  DBMS_OUTPUT.PUT_LINE('UID: ' || v_uid);
END;
/

-- ── Technique C: build SQL dynamically using a variable ──────────────
-- Cleanest for parameterised calls — no escaping confusion at all.
DECLARE
  v_uid  VARCHAR2(32);
  v_name VARCHAR2(100) := 'O''BRIEN';   -- store the actual value; PL/SQL handles ''
  v_sql  CLOB;
BEGIN
  -- Build the SQL string; CHR(39) is a single quote character
  v_sql := 'SELECT EMPNO, ENAME, JOB, SAL '
        || 'FROM   EMP '
        || 'WHERE  ENAME = ' || CHR(39) || v_name || CHR(39) || ' '
        || 'ORDER  BY EMPNO';

  v_uid := PKG_XLSX_WRAPPER.generate_xlsx(
    p_workbook_name => 'APOSTROPHE_DEMO',
    p_sheet1_name   => 'EMP_OBRIEN',
    p_sheet1_sql    => v_sql
  );
  DBMS_OUTPUT.PUT_LINE('UID: ' || v_uid);
END;
/


-- ════════════════════════════════════════════════════════════════════
-- QUICK REFERENCE: Single-quote escaping cheat sheet
-- ════════════════════════════════════════════════════════════════════
--
--  What you want in SQL         | Regular string     | q'[...]'
--  ─────────────────────────────┼────────────────────┼──────────────
--  WHERE COL = 'SMITH'          | ''SMITH''          | 'SMITH'
--  WHERE COL LIKE 'A%'          | ''A%''             | 'A%'
--  WHERE COL = 'O''BRIEN'       | ''O''''BRIEN''     | 'O''BRIEN'
--  WHERE COL = 'ST JOHN''S'     | ''ST JOHN''''S''   | 'ST JOHN''S'
--  Concatenation: 'MR '||ENAME  | ''MR ''||ENAME     | 'MR '||ENAME
--
--  Key rule: in a regular PL/SQL string every single quote is doubled.
--            In q'[...]' only the SQL-level embedded quotes are doubled;
--            the PL/SQL wrapper quotes disappear.
--
-- ════════════════════════════════════════════════════════════════════


-- ════════════════════════════════════════════════════════════════════
-- MANAGE & DOWNLOAD
-- ════════════════════════════════════════════════════════════════════

-- List all exports
EXEC PKG_XLSX_WRAPPER.list_exports;

-- Re-print download SQL for a UID
EXEC PKG_XLSX_WRAPPER.print_retrieval_sql('<uid>');

-- Download (run in SQL Developer / PL/SQL Developer, save BLOB as .xlsx)
SELECT XLSX_BLOB FROM XLSX_EXPORT_RESULTS WHERE EXPORT_UID = '<uid>';

-- Delete one export
EXEC PKG_XLSX_WRAPPER.delete_export('<uid>');

-- Manual purge (older than 7 days)
EXEC PKG_XLSX_WRAPPER.purge_old_exports;

-- View audit trail
SELECT ACTION, MESSAGE, LOG_TIME
FROM   XLSX_EXPORT_LOG
ORDER  BY LOG_TIME DESC;

*/
