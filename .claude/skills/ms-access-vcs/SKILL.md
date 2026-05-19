---
name: ms-access-vcs
description: Working with Microsoft Access databases exported via VCS (Version Control System for Access). Use when editing Access form definitions (.bas), report definitions, query files, VBA modules, linked table definitions, or VCS configuration files. Covers structural rules, common gotchas, and safe editing practices for the binary text format used by Access's LoadFromText/SaveAsText.
---

# Microsoft Access VCS Project Guide

Rules and gotchas for programmatically editing Access objects exported by VCS 4.x. Violations cause silent corruption or `Error 2128` on import via `LoadComponentFromText`.

## Reference Files

Read the relevant reference file before working on a specific object type:

| When working on... | Read this reference |
|---|---|
| Form `.bas` layout or form controls | `references/forms.md` |
| Report `.bas` layout, sections, or sorting/grouping | `references/reports.md` |
| Images in forms or reports (add, replace, convert from OLE) | `references/images.md` |
| Conditional formatting (forms or reports) | `references/conditional-formatting.md` |
| Queries (`.bas` + `.sql`) | `references/queries.md` |
| Tables, relationships, or ODBC connections | `references/tables-and-relationships.md` |
| VBA modules (`.bas`) or code-behind (`.cls`) | `references/vba.md` |
| Project config files (vcs-options, vcs-index, etc.) | `references/project-config.md` |
| Macros, shared images, or import/export specs | `references/other-objects.md` |

## Project Overview

A VCS-exported Access project lives in `<DatabaseName>.accdb.src/`. The directory name **must** match the `.accdb` filename with `.src` appended.

### File Pairing Rules

| Object Type | Files | Notes |
|-------------|-------|-------|
| Form | `.bas` + optional `.cls` + optional `.json` | `.bas` = layout, `.cls` = VBA code-behind (omitted if no code), `.json` = print settings |
| Report | `.bas` + optional `.cls` + optional `.json` | Same binary text format as forms, `.cls` = VBA code-behind (uncommon but valid), `.json` = print settings |
| Query | `.bas` + `.sql` | `.bas` = metadata/designer state, `.sql` = SQL text. Always paired |
| Module | `.bas` | Standard VBA module |
| Relationship | `.json` | Foreign key definition, join type, referential integrity |
| Linked Table | `.json` | Connection string, source table, attributes |
| Local Table | `.xml` | XSD schema defining columns, data types, indexes. For tables stored inside the .accdb |
| Macro | `.bas` | Access macro actions (not VBA). Simpler binary text format |
| Shared Image | `.json` + image file | `.json` = metadata (name, extension, hash), paired with actual image file (`.png`, `.jpg`, etc.) |
| Import/Export Spec | `.json` | Column definitions, delimiters, data types |
| Theme | `.thmx` | Office theme (binary OOXML archive — do not edit) |

When deleting an object, delete ALL its associated files plus its entry in `vcs-index.json`.

### Encoding

All JSON configuration files and `.sql` files use **UTF-8 with BOM** (byte order mark `EF BB BF`). The BOM must be preserved when editing. Binary text `.bas` files (forms, reports, macros) also use UTF-8 with BOM. `.bas`, `.cls`, and tbldef/relation `.json` files additionally use **CRLF** line endings (enforced via `.gitattributes`).

**Creating a new file from scratch:** standard write tools (including Claude's `Write` tool, `echo`, `cat <<EOF`) silently produce UTF-8 *without* BOM and use **LF** endings — this will corrupt the import. To create a new VCS file, emit the bytes explicitly:

```bash
printf '\xEF\xBB\xBF{\r\n  "Info": { ... }\r\n}\r\n' > path/to/new-file.json
```

Verify after writing: `od -c -N 5 path/to/new-file.json` should show `357 273 277` (the BOM) as the first three bytes, and the file should contain `\r` before each `\n`. *Editing* an existing file with `Edit`/`Read` preserves the original encoding — the BOM-stripping issue only applies to brand-new files created via `Write`.

## Safe Deletion Checklist

Before deleting any Access object:

1. **Search for references** in all `.cls`, `.bas`, `.sql` files. Objects may be referenced by:
   - VBA code (`DoCmd.OpenForm`, `DoCmd.OpenReport`, `DoCmd.OutputTo`)
   - Query SQL (`FROM tableName`, subreport `SourceObject`)
   - Form RecordSource properties
   - DLookup expressions
2. **Check for name collisions.** Objects with similar names may serve completely different purposes (e.g., an active report and a dead integration table that share a prefix).
3. **Delete all paired files** (.bas + .sql, .bas + .cls + .json, etc.)
4. **Remove from config files:** vcs-index.json, db-connection.json, hidden-attributes.json
5. **Renumber TabIndex** values in any form/report section where a control was removed
6. **Test import** in Access after changes

## Common Pitfalls

- **Commented-out code referencing kept objects:** Leave commented VBA lines alone if they reference objects you're keeping.
- **db-connection.json:** Must contain only one connection entry matching the canonical ODBC connection string used by all linked tables and pass-through queries. Never introduce plaintext passwords or user-specific credentials.
