# plsql_xlsx_export — Pure PL/SQL XLSX Generator for Oracle

Generate multi-sheet Excel workbooks entirely inside Oracle — no APEX, no UTL_FILE, no shell scripts, no external tools, no client installs required.

The toolkit validates a `.xlsx` file entirely in memory as a BLOB, stores it in a database table, and makes it downloadable directly from SQL Developer or PL/SQL Developer.

Compatible: **Oracle 12c · Oracle 19c**

---

## Packages at a Glance

| Package | Role |
|---|---|
| `PKG_XLSX_EXPORT` | Core engine — ZIP builder, XLSX XML generator, BLOB output. Zero table references. Compiles in any schema. |
| `PKG_XLSX_WRAPPER` | Smart wrapper — auto-naming, UID tracking, success notifications, auto-purge. Requires `CREATE TABLE`. |
| `PKG_XLSX_DIRECT` | SQL-callable function — returns BLOB directly from a `SELECT`. No table required. |

---

## How It Works

```
Your SQL Session
      │
      ▼
PKG_XLSX_WRAPPER.generate_xlsx(workbook_name, sheet_sql_1, sheet_sql_2, ...)
  ├─ Resolves workbook name  (supplied  OR  auto: SCHEMA_YYYYMMDD_HH24MISS)
  ├─ Resolves worksheet names (supplied  OR  auto-derived from SQL table names)
  └─ Calls PKG_XLSX_EXPORT internally
        ├─ DBMS_SQL     → executes each query, captures columns + rows
        ├─ XML builder  → generates XLSX parts (workbook, sheets, styles, sharedStrings)
        └─ ZIP builder  → assembles raw ZIP bytes as BLOB (pure PL/SQL, no UTL_COMPRESS)
  ├─ Stamps a unique UID on the saved record
  └─ Prints success notification + ready-to-run download SQL
```

---

## Prerequisites

| Privilege | Required For | Notes |
|---|---|---|
| `CREATE PACKAGE` | All schemas | Required on your schema only |
| `CREATE TABLE` | `PKG_XLSX_WRAPPER` only | For output and audit tables |
| `EXECUTE ON DBMS_SQL` | All | Usually granted by default |
| `EXECUTE ON DBMS_LOB` | All | Usually granted by default |
| SQL Developer or PL/SQL Developer | Download | For saving the BLOB as `.xlsx` |

**Verify your privileges:**
```sql
SELECT PRIVILEGE FROM USER_SYS_PRIVS
WHERE  PRIVILEGE IN ('CREATE PACKAGE','CREATE TABLE','CREATE PROCEDURE')
ORDER  BY 1;
```

---

## Installation

Deploy in order — each file depends on the previous one.

| Environment | PKG_XLSX_EXPORT | PKG_XLSX_WRAPPER | PKG_XLSX_DIRECT |
|---|:---:|:---:|:---:|
| Full privileges | ✅ Required | ✅ Deploy | ⬜ Optional |
| Read-only / no CREATE TABLE | ✅ Required | ⬜ Skip | ✅ Deploy |

### Step 1 — Deploy the core engine (ALL environments)

`PKG_XLSX_EXPORT` has zero table references. It compiles cleanly in any schema regardless of privileges — read-only or not.

```sql
@PKG_XLSX_EXPORT.sql
```

**What it creates:**

| Object | Type | Purpose |
|---|---|---|
| `PKG_XLSX_EXPORT` | Package spec | Public API: `init`, `add_sheet`, `build_blob`, `get_workbook_name` |
| `PKG_XLSX_EXPORT` | Package body | ZIP builder, XLSX XML builder, BLOB assembler |

**Verify:**
```sql
SELECT OBJECT_NAME, OBJECT_TYPE, STATUS
FROM   USER_OBJECTS
WHERE  OBJECT_NAME = 'PKG_XLSX_EXPORT';
-- STATUS must be VALID
```

If invalid:
```sql
SELECT LINE, TEXT FROM USER_ERRORS
WHERE  NAME = 'PKG_XLSX_EXPORT'
ORDER  BY LINE;
```

---

### Step 2a — Full-privilege environment: deploy the wrapper

`PKG_XLSX_WRAPPER` creates `XLSX_EXPORT_RESULTS` and `XLSX_EXPORT_LOG`, then calls `PKG_XLSX_EXPORT.build_blob` and does its own `INSERT`. Requires `CREATE TABLE`.

```sql
@PKG_XLSX_WRAPPER.sql
```

**What it creates:**

| Object | Type | Purpose |
|---|---|---|
| `XLSX_EXPORT_RESULTS` | Table | Stores generated XLSX BLOBs with UID tracking |
| `UX_XLSX_EXPORT_UID` | Unique index | Enforces UID uniqueness |
| `XLSX_EXPORT_LOG` | Table | Audit trail (INITIATED / COMPLETED / FAILED / DELETED / PURGED) |
| `PKG_XLSX_WRAPPER` | Package | Auto-naming, UID, notifications, purge |

**Verify:**
```sql
SELECT OBJECT_NAME, OBJECT_TYPE, STATUS
FROM   USER_OBJECTS
WHERE  OBJECT_NAME = 'PKG_XLSX_WRAPPER';
```

---

### Step 2b — Read-only environment: deploy the direct function

`PKG_XLSX_DIRECT` requires only `CREATE PACKAGE`. It calls `PKG_XLSX_EXPORT.build_blob` and returns the BLOB directly to the SQL caller. No tables needed.

```sql
@PKG_XLSX_DIRECT.sql
```

**What it creates:**

| Object | Type | Purpose |
|---|---|---|
| `PKG_XLSX_DIRECT` | Package | SQL-callable BLOB function, no table access |

**Verify:**
```sql
SELECT OBJECT_NAME, OBJECT_TYPE, STATUS
FROM   USER_OBJECTS
WHERE  OBJECT_NAME = 'PKG_XLSX_DIRECT';
```

---

## Naming Behaviour

### Workbook name (the `.xlsx` filename)

| Scenario | Result |
|---|---|
| `p_workbook_name` supplied | That name is used |
| `p_workbook_name` not supplied | `SCHEMA_YYYYMMDD_HH24MISS` |
| Name already exists in table | Auto-suffixed: `MY_REPORT_01`, `MY_REPORT_02`, … |

### Worksheet name (each tab inside the workbook)

| Scenario | Result |
|---|---|
| `p_sheetN_name` supplied | That name is used |
| Not supplied — single table in SQL | Table name (e.g. `EMPLOYEES`) |
| Not supplied — multiple tables / JOINs | `COMPLEX_01`, `COMPLEX_02`, … |
| Not supplied — unparseable SQL | `SHEET_1`, `SHEET_2`, … |
| Duplicate name within same workbook | Auto-suffixed: `EMPLOYEES_01` |

---

## PKG_XLSX_WRAPPER — Usage Examples

### Example 1 — No names supplied (fully automatic)

```sql
DECLARE
  v_uid VARCHAR2(32);
BEGIN
  v_uid := PKG_XLSX_WRAPPER.generate_xlsx(
    p_sheet1_sql => 'SELECT OBJECT_NAME, OBJECT_TYPE, STATUS, CREATED
                     FROM USER_OBJECTS ORDER BY OBJECT_TYPE',
    p_sheet2_sql => 'SELECT S.SEGMENT_NAME, S.BYTES, E.EXTENT_ID
                     FROM USER_SEGMENTS S
                     JOIN  USER_EXTENTS E ON E.SEGMENT_NAME = S.SEGMENT_NAME
                     WHERE ROWNUM <= 100',
    p_sheet3_sql => 'SELECT NAME, TYPE, LINE, TEXT
                     FROM USER_SOURCE WHERE ROWNUM <= 200'
  );
  DBMS_OUTPUT.PUT_LINE('UID: ' || v_uid);
END;
/
-- Workbook → MYSCHEMA_20240522_143055
-- Sheet 1  → USER_OBJECTS
-- Sheet 2  → COMPLEX_01   (two tables joined)
-- Sheet 3  → USER_SOURCE
```

### Example 2 — Workbook name supplied, sheet names auto-derived

```sql
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
-- Workbook → HR_MONTHLY_REPORT  (or HR_MONTHLY_REPORT_01 if name exists)
-- Sheet 1  → EMPLOYEES
-- Sheet 2  → COMPLEX_01
```

### Example 3 — All names explicitly provided

```sql
DECLARE
  v_uid VARCHAR2(32);
BEGIN
  v_uid := PKG_XLSX_WRAPPER.generate_xlsx(
    p_workbook_name => 'AUDIT_PACK',
    p_sheet1_name   => 'Access Log',
    p_sheet1_sql    => 'SELECT * FROM ACCESS_LOG    WHERE ROWNUM <= 500',
    p_sheet2_name   => 'Error Summary',
    p_sheet2_sql    => 'SELECT * FROM ERROR_SUMMARY WHERE ROWNUM <= 500',
    p_sheet3_name   => 'User Activity',
    p_sheet3_sql    => 'SELECT * FROM USER_ACTIVITY WHERE ROWNUM <= 500'
  );
END;
/
```

### Example 4 — WHERE clause with a character filter (`LIKE 'A%'`)

Single quotes inside a PL/SQL string must be doubled (`''`).

```sql
-- Plain SQL: WHERE ENAME LIKE 'A%'
-- In PL/SQL string, each ' becomes '': WHERE ENAME LIKE ''A%''

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
END;
/
```

**Alternative — `q'[...]'` quoting (recommended for readability):**

```sql
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
END;
/
-- Inside q'[...]', the bracket pair [ ] acts as the string delimiter.
-- Write single quotes naturally without doubling them.
```

### Example 5 — WHERE value contains an apostrophe (`O'BRIEN`)

```sql
-- Technique A: doubled-quote escaping
DECLARE
  v_uid VARCHAR2(32);
BEGIN
  v_uid := PKG_XLSX_WRAPPER.generate_xlsx(
    p_workbook_name => 'APOSTROPHE_DEMO',
    p_sheet1_name   => 'EMP_OBRIEN',
    p_sheet1_sql    =>
        'SELECT EMPNO, ENAME, JOB, SAL '  ||
        'FROM   EMP '                     ||
        'WHERE  ENAME = ''O''''BRIEN'' '  ||   -- SQL: ENAME = 'O''BRIEN'
        'ORDER  BY EMPNO'
  );
END;
/

-- Technique B: q'[...]' alternative quoting (recommended)
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
END;
/

-- Technique C: CHR(39) — cleanest for parameterised values
DECLARE
  v_uid  VARCHAR2(32);
  v_name VARCHAR2(100) := 'O''BRIEN';
  v_sql  CLOB;
BEGIN
  v_sql := 'SELECT EMPNO, ENAME, JOB, SAL '
        || 'FROM   EMP '
        || 'WHERE  ENAME = ' || CHR(39) || v_name || CHR(39) || ' '
        || 'ORDER  BY EMPNO';
  v_uid := PKG_XLSX_WRAPPER.generate_xlsx(
    p_workbook_name => 'APOSTROPHE_DEMO',
    p_sheet1_name   => 'EMP_OBRIEN',
    p_sheet1_sql    => v_sql
  );
END;
/
```

### Single-quote escaping cheat sheet

| SQL you want | Regular PL/SQL string | `q'[...]'` |
|---|---|---|
| `WHERE COL = 'SMITH'` | `''SMITH''` | `'SMITH'` |
| `WHERE COL LIKE 'A%'` | `''A%''` | `'A%'` |
| `WHERE COL = 'O''BRIEN'` | `''O''''BRIEN''` | `'O''BRIEN'` |
| `WHERE COL = 'ST JOHN''S'` | `''ST JOHN''''S''` | `'ST JOHN''S'` |
| `'MR ' \|\| ENAME` | `''MR '' \|\| ENAME` | `'MR ' \|\| ENAME` |

> **Key rule:** In a regular PL/SQL string every single quote is doubled. In `q'[...]'` only the SQL-level embedded apostrophes are doubled — the PL/SQL wrapper quotes disappear entirely.

---

## Downloading the XLSX

When `generate_xlsx` completes, `DBMS_OUTPUT` automatically prints:

```
====================================================================
  EXPORT COMPLETE
--------------------------------------------------------------------
  Workbook  : HR_MONTHLY_REPORT.xlsx
  Export ID : 5
  UID       : A3F29C1B...
  Size      : 48.3 KB
--------------------------------------------------------------------
  DOWNLOAD — run in SQL Developer / PL/SQL Developer:

    SELECT XLSX_BLOB
    FROM   XLSX_EXPORT_RESULTS
    WHERE  EXPORT_UID = 'A3F29C1B...';

  Save the BLOB cell as: HR_MONTHLY_REPORT.xlsx
--------------------------------------------------------------------
  DELETE when done:

    EXEC PKG_XLSX_WRAPPER.delete_export('A3F29C1B...');
====================================================================
```

**In SQL Developer:** run the `SELECT`, click the BLOB cell → pencil/download icon → Save As → rename to `.xlsx`

**In PL/SQL Developer:** run the `SELECT`, right-click the BLOB cell → Save to File → rename to `.xlsx`

---

## Managing Exports (PKG_XLSX_WRAPPER)

```sql
-- List all exports
EXEC PKG_XLSX_WRAPPER.list_exports;

-- Re-print download SQL for a UID
EXEC PKG_XLSX_WRAPPER.print_retrieval_sql('A3F29C1B...');

-- Download
SELECT XLSX_BLOB FROM XLSX_EXPORT_RESULTS WHERE EXPORT_UID = 'A3F29C1B...';

-- Delete one export
EXEC PKG_XLSX_WRAPPER.delete_export('A3F29C1B...');

-- Purge exports older than 7 days (also runs automatically on every generate call)
EXEC PKG_XLSX_WRAPPER.purge_old_exports;

-- Custom retention period (e.g. 3 days)
EXEC PKG_XLSX_WRAPPER.purge_old_exports(3);

-- View audit trail
SELECT ACTION, MESSAGE, LOG_TIME
FROM   XLSX_EXPORT_LOG
ORDER  BY LOG_TIME DESC;
```

---

## PKG_XLSX_DIRECT — Usage Examples

Use this package when you cannot create tables, or when you need the BLOB returned directly to a SQL caller (e.g. piped through `UTL_MAIL`, stored programmatically, or downloaded straight from SQL Developer without a table).

All queries are passed as a single delimited string separated by `~` (tilde — `PKG_XLSX_DIRECT.get_delim()`), making it callable directly from a `SELECT` statement.

### How the delimiter works

```sql
-- Single query — no delimiter needed
'SELECT * FROM EMP'

-- Multiple queries — separate with PKG_XLSX_DIRECT.get_delim()
'SELECT * FROM EMP'
|| PKG_XLSX_DIRECT.get_delim()
|| 'SELECT * FROM DEPT'
|| PKG_XLSX_DIRECT.get_delim()
|| 'SELECT * FROM SALGRADE'
```

> **Note:** Always use `PKG_XLSX_DIRECT.get_delim()` (the function) in SQL context, not `PKG_XLSX_DIRECT.DELIM` (the constant). Package constants are not reachable from SQL — only functions are SQL-callable (PLS-00221).

### Function signature

```sql
FUNCTION generate_xlsx(
  p_workbook_name IN VARCHAR2 DEFAULT NULL,
  p_queries       IN CLOB,
  p_sheet_names   IN VARCHAR2 DEFAULT NULL,  -- NULL = auto-derive from SQL
  p_debug         IN VARCHAR2 DEFAULT 'N'    -- 'Y' enables DBMS_OUTPUT tracing
) RETURN BLOB;
```

> Always use **named notation** (`p_param => value`) to avoid overload ambiguity (PLS-307), especially when supplying `p_debug` or `p_sheet_names`.

### Example 1 — Auto sheet names

```sql
SELECT PKG_XLSX_DIRECT.generate_xlsx(
           p_workbook_name => 'Deposit_Issue',
           p_queries       => 'SELECT * FROM ICTM_ACC WHERE ROWNUM <= 10'
                           || PKG_XLSX_DIRECT.get_delim()
                           || q'[SELECT * FROM STTM_CUST_ACCOUNT WHERE CUST_NAME LIKE 'A%']'
       )
FROM DUAL;
-- Sheet 1 → ICTM_ACC
-- Sheet 2 → STTM_CUST_ACCOUNT
```

### Example 2 — Explicit sheet names

```sql
SELECT PKG_XLSX_DIRECT.generate_xlsx(
           p_workbook_name => 'Deposit_Issue',
           p_queries       => 'SELECT * FROM ICTM_ACC WHERE ROWNUM <= 10'
                           || PKG_XLSX_DIRECT.get_delim()
                           || q'[SELECT * FROM STTM_CUST_ACCOUNT WHERE CUST_NAME LIKE 'A%']',
           p_sheet_names   => 'Accounts' || PKG_XLSX_DIRECT.get_delim() || 'Customers'
       )
FROM DUAL;
-- Sheet 1 → Accounts
-- Sheet 2 → Customers
```

### Example 3 — No workbook name (auto-named)

```sql
SELECT PKG_XLSX_DIRECT.generate_xlsx(
           p_queries => 'SELECT EMPNO, ENAME, SAL FROM EMP ORDER BY ENAME'
       )
FROM DUAL;
-- Workbook → MYSCHEMA_20240522_143055
-- Sheet 1  → EMP
```

### Example 4 — With debug tracing

```sql
SELECT PKG_XLSX_DIRECT.generate_xlsx(
           p_workbook_name => 'Deposit_Issue',
           p_queries       => 'SELECT * FROM ICTM_ACC WHERE ROWNUM <= 10'
                           || PKG_XLSX_DIRECT.get_delim()
                           || q'[SELECT * FROM STTM_CUST_ACCOUNT WHERE ACCOUNT_CLASS LIKE 'A%']',
           p_debug         => 'Y'
       )
FROM DUAL;
-- DBMS_OUTPUT will show timestamped trace lines: [HH:MI:SS.FFF] message
```

### Example 5 — WHERE clause with character filter

```sql
-- Using q'[...]' — no escaping needed:
SELECT PKG_XLSX_DIRECT.generate_xlsx(
           p_workbook_name => 'EMP_Report',
           p_queries       => q'[SELECT EMPNO, ENAME, JOB, SAL FROM EMP WHERE ENAME LIKE 'A%']'
       )
FROM DUAL;

-- Using regular string — every quote doubled:
SELECT PKG_XLSX_DIRECT.generate_xlsx(
           p_workbook_name => 'EMP_Report',
           p_queries       => 'SELECT EMPNO, ENAME, JOB, SAL FROM EMP WHERE ENAME LIKE ''A%'''
       )
FROM DUAL;
```

### Example 6 — WHERE value contains apostrophe

```sql
-- q'[...]' — only SQL-level apostrophe needs doubling:
SELECT PKG_XLSX_DIRECT.generate_xlsx(
           p_workbook_name => 'EMP_Report',
           p_queries       => q'[SELECT EMPNO, ENAME, JOB FROM EMP WHERE ENAME = 'O''BRIEN']'
       )
FROM DUAL;
```

### How to download from SQL Developer / PL/SQL Developer

1. Run a `SELECT PKG_XLSX_DIRECT.generate_xlsx(...) FROM DUAL`
2. The result grid shows one row with a BLOB column
3. **SQL Developer:** click the BLOB cell → pencil/download icon → Save As → rename to `Workbook.xlsx`
4. **PL/SQL Developer:** right-click the BLOB cell → Save to File → rename to `Workbook.xlsx`
5. Open in Excel

### Changing the delimiter

The default delimiter is `~` (tilde, ASCII 126). This character:
- Is a single byte in **every** Oracle character set (US7ASCII → AL32UTF8)
- Is reliably found by `DBMS_LOB.INSTR` on all database character sets
- **Caution:** `~` can appear in SQL (REGEXP patterns, custom `LIKE ESCAPE` clauses). If your SQL contains `~`, change the constant.

> **Why not `§`?** The section sign (U+00A7) is 2 bytes in AL32UTF8. `DBMS_LOB.INSTR` with a multi-byte VARCHAR2 pattern fails silently on single-byte NLS_CHARACTERSET databases (e.g. WE8ISO8859P1), returning 0 and causing the query split to produce 0 results.

To change the delimiter, update the `DELIM` constant in the package spec and recompile:

```sql
-- In PKG_XLSX_DIRECT spec, change:
DELIM CONSTANT VARCHAR2(3) := '~';
-- to any character that cannot appear in your SQL, e.g.:
DELIM CONSTANT VARCHAR2(3) := CHR(30);  -- ASCII Record Separator (safest)
```

---

## Using PKG_XLSX_EXPORT Directly

Use `PKG_XLSX_EXPORT` directly only if you need programmatic BLOB access beyond what the wrapper and direct packages offer (e.g. more than 10 sheets, custom streaming).

```sql
DECLARE
  v_blob BLOB;
BEGIN
  PKG_XLSX_EXPORT.init('My_Export');

  PKG_XLSX_EXPORT.add_sheet('Data',
    'SELECT * FROM MY_TABLE WHERE ROWNUM <= 1000');

  PKG_XLSX_EXPORT.add_sheet('Summary',
    'SELECT DEPT, COUNT(*) CNT, SUM(SAL) TOTAL FROM EMP GROUP BY DEPT');

  v_blob := PKG_XLSX_EXPORT.build_blob;
  DBMS_OUTPUT.PUT_LINE('Size: ' || DBMS_LOB.GETLENGTH(v_blob) || ' bytes');
  -- use v_blob as needed (insert, mail, stream, etc.)
END;
/
```

---

## Debug Tracing

All three packages support per-call debug tracing. Pass `p_debug => 'Y'` to `generate_xlsx` (in DIRECT or WRAPPER), or call `enable_debug` / `disable_debug` directly for session-level control.

```sql
-- Session-level (PKG_XLSX_EXPORT / PKG_XLSX_DIRECT / PKG_XLSX_WRAPPER)
EXEC PKG_XLSX_EXPORT.enable_debug;
EXEC PKG_XLSX_DIRECT.enable_debug;
EXEC PKG_XLSX_WRAPPER.enable_debug;

-- Per-call (DIRECT and WRAPPER only)
-- Pass p_debug => 'Y' — debug auto-disabled at function exit
```

Debug output format: `[HH:MI:SS.FFF] message`

---

## API Reference

### PKG_XLSX_DIRECT

| Function / Procedure | Parameters | Description |
|---|---|---|
| `generate_xlsx` | `p_workbook_name DEFAULT NULL`<br>`p_queries CLOB`<br>`p_sheet_names DEFAULT NULL`<br>`p_debug DEFAULT 'N'` | Builds workbook, returns BLOB. Sheet names auto-derived if `p_sheet_names` is NULL. |
| `get_delim()` | — | Returns the `~` delimiter constant. Use this in SQL context instead of `PKG_XLSX_DIRECT.DELIM`. |
| `enable_debug` | — | Turns on DBMS_OUTPUT tracing for this session |
| `disable_debug` | — | Turns off tracing (default) |

### PKG_XLSX_WRAPPER

| Function / Procedure | Parameters | Description |
|---|---|---|
| `generate_xlsx` | `p_workbook_name DEFAULT NULL`<br>`p_sheet1_name DEFAULT NULL`, `p_sheet1_sql CLOB`<br>… up to sheet 10 …<br>`p_debug DEFAULT 'N'` | Builds workbook, saves BLOB to table, returns UID |
| `delete_export` | `p_export_uid` | Deletes one export by UID |
| `purge_old_exports` | `p_days DEFAULT 7` | Deletes exports older than N days |
| `list_exports` | — | Prints all exports to DBMS_OUTPUT |
| `print_retrieval_sql` | `p_export_uid` | Re-prints the download SQL for a UID |
| `enable_debug` | — | Turns on DBMS_OUTPUT tracing |
| `disable_debug` | — | Turns off tracing (default) |

### PKG_XLSX_EXPORT (Core Engine)

| Function / Procedure | Description |
|---|---|
| `init(p_workbook_name)` | Resets all state. Call before starting a new workbook. |
| `add_sheet(p_sheet_name, p_sql)` | Registers a SQL query as a worksheet. |
| `build_blob` | Executes all queries, builds the XLSX, returns raw BLOB. No table access. |
| `get_workbook_name` | Returns the current workbook name (set by `init`). |
| `enable_debug` / `disable_debug` | Toggle DBMS_OUTPUT debug tracing. |

---

## Data Type Mapping

| Oracle Type | Excel Output | Notes |
|---|---|---|
| `VARCHAR2`, `CHAR`, `NVARCHAR2` | String | Via shared string table |
| `NUMBER`, `BINARY_FLOAT`, `BINARY_DOUBLE` | Numeric | Native Excel number cell |
| `DATE`, `TIMESTAMP` (all variants) | Date/Time | Formatted `YYYY-MM-DD HH:MI:SS` |
| `INTERVAL YEAR TO MONTH` | String | Converted via Oracle implicit TO_CHAR |
| `INTERVAL DAY TO SECOND` | String | Converted via Oracle implicit TO_CHAR |
| `CLOB`, `LONG` | String | Truncated at 32,767 chars (Excel cell limit) |
| All other types | String | Safe fallback |

---

## XLSX Output Characteristics

| Feature | Detail |
|---|---|
| **Header row** | Bold, white text, Green Accent 6 background (`#70AD47`), centered, auto-generated from column names |
| **Sheet background** | Explicit solid white (`#FFFFFF`) on all data and date cells |
| **Encoding** | UTF-8 throughout |
| **Compression** | STORED mode (no DEFLATE) — Excel opens natively, no corruption risk |
| **Shared strings** | All text centralised in `xl/sharedStrings.xml` per OOXML spec |
| **Multi-sheet** | Up to 10 sheets per call via wrapper; unlimited via base package |

---

## Limitations

| Limitation | Detail |
|---|---|
| No cell formulas | Data-only output |
| No column auto-width | All columns use Excel default width |
| CLOB/LONG truncated at 32,767 chars | Matches Excel's practical cell limit |
| No bind variables in SQL | Pass literal values or use `CHR(39)` / `q'[...]'` approach in `add_sheet` |
| No DEFLATE compression | STORED mode ZIP — larger files for text-heavy data |
| Max sheet name 31 chars | Excel hard limit, enforced by package |
| Max 10 sheets via wrapper | Use `PKG_XLSX_EXPORT` directly for more |

---

## Compatibility

| Environment | Status |
|---|---|
| Oracle 12c | ✅ Tested |
| Oracle 19c | ✅ Tested |
| SQL Developer (BLOB download) | ✅ Tested |
| PL/SQL Developer (BLOB download) | ✅ Tested |
| Microsoft Excel 2016+ | ✅ Tested |
| LibreOffice Calc | ✅ Tested |

---

## License

MIT — free to use, modify, and distribute.
