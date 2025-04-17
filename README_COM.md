# MCP Access_COM (COM-based MS Access Reader)

This MCP app allows you to interact with Microsoft Access database modules, saved queries, forms, and MSys tables using the COM interface (via pywin32). It complements the Access-mdb MCP by providing access to VBA modules, forms, and system tables not accessible via ODBC/ADO.

## Workflow & Requirements
- **User must open the MDB/ACCDB file in Access UI first** (to resolve linked tables and bypass security prompts)
- MCP attaches to the running Access instance (does not launch a new one)
- Requires: Python 3.7+, pywin32, Microsoft Access installed

## Features / MCP Tools
- **connect_access_db**: Attach to running Access instance with open DB
- **disconnect_access_db**: Detach from Access instance
- **list_modules / get_module_code**: List VBA modules, get code
- **list_queries / get_query_sql**: List saved queries, get SQL
- **list_forms / get_form_properties**: List forms, get properties/controls
- **list_msys_tables / get_msys_table_data**: List MSys/system tables, read data

## Usage
1. Open your Access database in Access UI (ensure all links/macros/security are handled)
2. Start the MCP server:
   ```bash
   python server.py
   ```
3. Use the API endpoints to access modules, queries, forms, and MSys tables

## Concurrency & Safety
- This MCP attaches to an already-open Access session. Do not run multiple COM automations on the same DB at once.
- Access-mdb MCP (ODBC/ADO) can be used in parallel for read-only data access.

## Roadmap
- [x] Attach to running Access (not launch new)
- [x] List/read modules, queries, forms, MSys tables
- [ ] Export features, macro execution, advanced sync with Access-COM MCP

## Standalone MCP Server (2025-04-16)

- `access_com.py` is now the only entry point for MCP COM tools.
- All COM endpoints (connect, disconnect, list/get modules, queries, forms, MSys tables, etc.) are registered as MCP tools in this file.
- `server.py` is no longer needed for MCP operation and can be archived or removed if not required.

### How to Run

1. Open your Access database in the Access UI (e.g., `ific3033.mdb`).
2. Start the MCP server:
   ```bash
   python access_com.py
   ```
3. Use Windsurf, Claude Desktop, or any MCP client to call the COM tools (e.g., connect, list_modules, get_module_code, etc.).

### Notes
- All progress and troubleshooting are tracked in `TASKS.md`.
- For ODBC/ADO endpoints, see legacy code or create new MCP tools as needed.
- For further development, update this file and `TASKS.md` regularly.

---
This README is for the COM-based Access MCP only. For ODBC/ADO-based access, see README.md.
