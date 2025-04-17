"""
IMPORTANT:
After creating or modifying objects (queries, tables, etc) via COM automation, always disconnect the COM session before reopening or refreshing in the Access UI. This prevents file locking (.ldb) issues and ensures all changes are visible to both automation and Access UI. If the COM session is left open, the Access UI may not see new or changed objects until the session is closed.
"""

import win32com.client
import pythoncom
from mcp.server.fastmcp import FastMCP

mcp = FastMCP("Access-COM MCP")

class AccessCOMManager:
    def __init__(self):
        self.access_app = None
        self.db = None

    def connect(self):
        pythoncom.CoInitialize()
        self.access_app = win32com.client.GetActiveObject("Access.Application")
        self.db = self.access_app.CurrentDb()

    def disconnect(self):
        self.access_app = None
        self.db = None

    def list_modules(self):
        components = self.access_app.VBE.VBProjects(1).VBComponents
        return [c.Name for c in components]

    def get_module_code(self, module_name):
        components = self.access_app.VBE.VBProjects(1).VBComponents
        for c in components:
            if c.Name == module_name:
                return c.CodeModule.Lines(1, c.CodeModule.CountOfLines)
        return None

    def list_queries(self):
        return [q.Name for q in self.db.QueryDefs]

    def get_query_sql(self, query_name):
        for q in self.db.QueryDefs:
            if q.Name == query_name:
                return q.SQL
        return None

    def create_query(self, name, sql):
        # Delete if exists (Access does not allow duplicate QueryDef names)
        for q in self.db.QueryDefs:
            if q.Name == name:
                self.db.QueryDefs.Delete(name)
                break
        return self.db.CreateQueryDef(name, sql)

    def list_forms(self):
        return [f.Name for f in self.access_app.CurrentProject.AllForms]

    def get_form_properties(self, form_name):
        for f in self.access_app.CurrentProject.AllForms:
            if f.Name == form_name:
                return {"Name": f.Name, "IsLoaded": f.IsLoaded}
        return None

    def list_msys_tables(self):
        return [t.Name for t in self.db.TableDefs if t.Name.startswith('MSys')]

    def get_msys_table_data(self, table_name, limit=20):
        rs = self.db.OpenRecordset(table_name)
        results = []
        count = 0
        while not rs.EOF and count < limit:
            row = {}
            for i in range(rs.Fields.Count):
                row[rs.Fields.Item(i).Name] = rs.Fields.Item(i).Value
            results.append(row)
            rs.MoveNext()
            count += 1
        return results

# GLOBAL singleton instance: do NOT re-instantiate in any function!
com_manager = AccessCOMManager()

@mcp.tool()
def connect():
    try:
        return com_manager.connect()
    except Exception as e:
        return f"Error: {e}"

@mcp.tool()
def disconnect():
    try:
        return com_manager.disconnect()
    except Exception as e:
        return f"Error: {e}"

@mcp.tool()
def list_modules(full: bool = False):
    """
    List VBA modules. Returns first 5 by default. Set full=True to get all.
    """
    try:
        modules = com_manager.list_modules()
        if not full:
            return modules[:5]
        return modules
    except Exception as e:
        return f"Error: {e}"

@mcp.tool()
def get_module_code(module_name: str):
    try:
        return com_manager.get_module_code(module_name)
    except Exception as e:
        return f"Error: {e}"

@mcp.tool()
def list_queries(full: bool = False):
    """
    List queries. Returns first 5 by default. Set full=True to get all.
    """
    try:
        queries = com_manager.list_queries()
        if not full:
            return queries[:5]
        return queries
    except Exception as e:
        return f"Error: {e}"

@mcp.tool()
def get_query_sql(query_name: str):
    try:
        return com_manager.get_query_sql(query_name)
    except Exception as e:
        return f"Error: {e}"

@mcp.tool()
def list_querydefs_full():
    """
    List all QueryDefs with their Name, SQL, and Attributes for debugging.
    """
    try:
        results = []
        for q in com_manager.db.QueryDefs:
            try:
                results.append({
                    'Name': q.Name,
                    'SQL': getattr(q, 'SQL', None),
                    'Attributes': getattr(q, 'Attributes', None),
                    'Type': getattr(q, 'Type', None)
                })
            except Exception as e:
                results.append({'Name': getattr(q, 'Name', '??'), 'Error': str(e)})
        return results
    except Exception as e:
        return f"Error: {e}"

@mcp.tool()
def create_query(name: str, sql: str):
    """
    Create or replace a saved query (QueryDef) in the current Access database.
    """
    try:
        result = com_manager.create_query(name, sql)
        return f"Query '{name}' created."
    except Exception as e:
        return f"Error: {e}"

@mcp.tool()
def list_forms(full: bool = False):
    """
    List forms. Returns first 5 by default. Set full=True to get all.
    """
    try:
        forms = com_manager.list_forms()
        if not full:
            return forms[:5]
        return forms
    except Exception as e:
        return f"Error: {e}"

@mcp.tool()
def get_form_properties(form_name: str):
    try:
        return com_manager.get_form_properties(form_name)
    except Exception as e:
        return f"Error: {e}"

@mcp.tool()
def list_msys_tables(full: bool = False):
    """
    List MSys tables. Returns first 5 by default. Set full=True to get all.
    """
    try:
        tables = com_manager.list_msys_tables()
        if not full:
            return tables[:5]
        return tables
    except Exception as e:
        return f"Error: {e}"

@mcp.tool()
def get_msys_table_data(table_name: str, limit: int = 5, full: bool = False):
    """
    Get data from an MSys table. Returns 5 rows by default. Set full=True for all available rows up to 100.
    """
    try:
        if full:
            return com_manager.get_msys_table_data(table_name, 100)
        return com_manager.get_msys_table_data(table_name, limit)
    except Exception as e:
        return f"Error: {e}"

import csv
import os

# --- New MCP Tools: Linked Tables & Macros ---

@mcp.tool()
def list_linked_tables(full: bool = False):
    """
    List all linked tables with metadata (Name, ForeignName, Database, Connect, Type).
    Returns first 5 by default. Set full=True to get all.
    Uses COM if possible, otherwise falls back to MSysObjects.csv.
    """
    try:
        # Try using COM (live Access instance)
        msys = com_manager.db.OpenRecordset('MSysObjects')
        results = []
        while not msys.EOF:
            # Type=6 is linked table
            try:
                typ = msys.Fields('Type').Value
                connect = msys.Fields('Connect').Value if 'Connect' in msys.Fields else None
                db = msys.Fields('Database').Value if 'Database' in msys.Fields else None
                name = msys.Fields('Name').Value
                foreign = msys.Fields('ForeignName').Value if 'ForeignName' in msys.Fields else None
            except Exception:
                msys.MoveNext()
                continue
            if typ == 6 and (connect or db):
                results.append({
                    'Name': name,
                    'ForeignName': foreign,
                    'Database': db,
                    'Connect': connect,
                    'Type': typ
                })
            msys.MoveNext()
        if not full:
            return results[:5]
        return results
    except Exception:
        # Fallback: parse MSysObjects.csv
        csv_path = os.path.join(os.path.dirname(__file__), 'MSysObjects.csv')
        if not os.path.exists(csv_path):
            return 'MSysObjects.csv not found and COM not available.'
        with open(csv_path, newline='', encoding='utf-8') as f:
            reader = csv.DictReader(f, delimiter=';')
            results = []
            for row in reader:
                try:
                    typ = int(row.get('Type', '0'))
                except Exception:
                    continue
                connect = row.get('Connect')
                db = row.get('Database')
                name = row.get('Name')
                foreign = row.get('ForeignName')
                if typ == 6 and (connect or db):
                    results.append({
                        'Name': name,
                        'ForeignName': foreign,
                        'Database': db,
                        'Connect': connect,
                        'Type': typ
                    })
            if not full:
                return results[:5]
            return results

@mcp.tool()
def list_macros(full: bool = False):
    """
    List all macros with metadata (Name, Type). Returns first 5 by default. Set full=True to get all. Uses COM if possible, otherwise falls back to MSysObjects.csv.
    """
    try:
        msys = com_manager.db.OpenRecordset('MSysObjects')
        results = []
        while not msys.EOF:
            try:
                typ = msys.Fields('Type').Value
                name = msys.Fields('Name').Value
            except Exception:
                msys.MoveNext()
                continue
            if typ == -32766:
                results.append({'Name': name, 'Type': typ})
            msys.MoveNext()
        if not full:
            return results[:5]
        return results
    except Exception:
        # Fallback: parse MSysObjects.csv
        csv_path = os.path.join(os.path.dirname(__file__), 'MSysObjects.csv')
        if not os.path.exists(csv_path):
            return 'MSysObjects.csv not found and COM not available.'
        with open(csv_path, newline='', encoding='utf-8') as f:
            reader = csv.DictReader(f, delimiter=';')
            results = []
            for row in reader:
                try:
                    typ = int(row.get('Type', '0'))
                except Exception:
                    continue
                name = row.get('Name')
                if typ == -32766:
                    results.append({'Name': name, 'Type': typ})
            if not full:
                return results[:5]
            return results

# --- End new MCP tools ---

if __name__ == "__main__":
    mcp.run()
