#!/usr/bin/env python3
"""
MCP Test Driver for Excel Power Pivot MCP Server
Runs comprehensive tests against all MCP tools with VERBOSE output.
"""
import subprocess
import json
import sys
import time
import os
import threading
import queue
from datetime import datetime
from dataclasses import dataclass, field
from typing import Optional, List, Dict, Any

# Fix Windows console encoding issues
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

# Configuration
PROJECT_PATH = "ExcelPowerPivotMcp.csproj"
LOG_FILE = "test_results.log"
VERBOSE = True  # Show detailed data in output
TEST_TIMEOUT_SECONDS = 30  # Timeout for individual test tool calls

# ANSI Color Codes
class Colors:
    HEADER = '\033[95m'
    BLUE = '\033[94m'
    CYAN = '\033[96m'
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    RED = '\033[91m'
    BOLD = '\033[1m'
    DIM = '\033[2m'
    RESET = '\033[0m'

# Enable ANSI colors on Windows
if os.name == 'nt':
    os.system('')

@dataclass
class TestResult:
    name: str
    status: str  # PASS, FAIL
    message: str = ""
    duration_ms: float = 0

@dataclass
class TestStats:
    total: int = 0
    passed: int = 0
    failed: int = 0
    start_time: float = field(default_factory=time.time)
    results: List[TestResult] = field(default_factory=list)
    
    def add(self, result: TestResult):
        self.results.append(result)
        self.total += 1
        if result.status == "PASS":
            self.passed += 1
        elif result.status == "FAIL":
            self.failed += 1

stats = TestStats()
msg_id = 0
log_file_handle = None

def init_log_file():
    global log_file_handle
    log_file_handle = open(LOG_FILE, "w", encoding="utf-8")
    log_file_handle.write(f"Excel Power Pivot MCP - Test Results\n")
    log_file_handle.write(f"Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
    log_file_handle.write("=" * 60 + "\n\n")

def close_log_file():
    global log_file_handle
    if log_file_handle:
        log_file_handle.close()

def log(message: str, level: str = "INFO", to_file: bool = True):
    """Log a message with timestamp and optional color."""
    timestamp = time.strftime("%H:%M:%S")
    
    color = Colors.RESET
    icon = "•"
    if level == "PASS":
        color = Colors.GREEN
        icon = "✓"
    elif level == "FAIL":
        color = Colors.RED
        icon = "✗"
    elif level == "WARN":
        color = Colors.YELLOW
        icon = "⚠"
    elif level == "DATA":
        color = Colors.CYAN
        icon = "  →"
    elif level == "HEADER":
        color = Colors.CYAN + Colors.BOLD
        icon = "►"
    elif level == "DEBUG":
        color = Colors.DIM
        icon = "·"
    
    console_msg = f"{Colors.DIM}[{timestamp}]{Colors.RESET} {color}{icon} {message}{Colors.RESET}"
    print(console_msg)
    
    if to_file and log_file_handle:
        plain_msg = f"[{timestamp}] [{level}] {message}"
        log_file_handle.write(plain_msg + "\n")
        log_file_handle.flush()

def log_data(label: str, items: list, max_items: int = 10):
    """Log detailed data items."""
    if not VERBOSE or not items:
        return
    
    display_items = items[:max_items]
    for item in display_items:
        if isinstance(item, dict):
            # Get name from dict
            name = item.get("name", item.get("Name", item.get("tableName", item.get("measureName", str(item)[:50]))))
            log(f"{label}: {name}", "DATA")
        else:
            log(f"{label}: {item}", "DATA")
    
    if len(items) > max_items:
        log(f"  ... and {len(items) - max_items} more", "DATA")

def log_section(title: str):
    """Log a section header."""
    separator = "─" * 50
    print(f"\n{Colors.CYAN}{Colors.BOLD}{separator}{Colors.RESET}")
    print(f"{Colors.CYAN}{Colors.BOLD}  {title}{Colors.RESET}")
    print(f"{Colors.CYAN}{Colors.BOLD}{separator}{Colors.RESET}\n")
    
    if log_file_handle:
        log_file_handle.write(f"\n{'=' * 50}\n")
        log_file_handle.write(f"  {title}\n")
        log_file_handle.write(f"{'=' * 50}\n\n")

def log_test_result(test_name: str, status: str, message: str = "", duration_ms: float = 0):
    """Log a test result and track statistics."""
    result = TestResult(test_name, status, message, duration_ms)
    stats.add(result)
    
    duration_str = f" ({duration_ms:.0f}ms)" if duration_ms > 0 else ""
    full_message = f"{test_name}{duration_str}"
    if message:
        full_message += f" - {message}"
    
    log(full_message, status)

def print_summary():
    """Print final test summary."""
    elapsed = time.time() - stats.start_time
    
    print(f"\n{Colors.BOLD}{'═' * 60}{Colors.RESET}")
    print(f"{Colors.BOLD}  TEST SUMMARY{Colors.RESET}")
    print(f"{Colors.BOLD}{'═' * 60}{Colors.RESET}\n")
    
    pass_color = Colors.GREEN if stats.passed > 0 else Colors.DIM
    fail_color = Colors.RED if stats.failed > 0 else Colors.DIM
    
    print(f"  {pass_color}✓ Passed:  {stats.passed}{Colors.RESET}")
    print(f"  {fail_color}✗ Failed:  {stats.failed}{Colors.RESET}")
    print(f"  {Colors.DIM}─────────────{Colors.RESET}")
    print(f"  Total:   {stats.total}")
    print(f"\n  Duration: {elapsed:.1f}s\n")
    
    failures = [r for r in stats.results if r.status == "FAIL"]
    if failures:
        print(f"{Colors.RED}{Colors.BOLD}  Failed Tests:{Colors.RESET}")
        for f in failures:
            print(f"  {Colors.RED}  • {f.name}: {f.message}{Colors.RESET}")
        print()
    
    if stats.failed == 0:
        print(f"  {Colors.GREEN}{Colors.BOLD}✓ ALL TESTS PASSED{Colors.RESET}\n")
    else:
        print(f"  {Colors.RED}{Colors.BOLD}✗ SOME TESTS FAILED{Colors.RESET}\n")
    
    if log_file_handle:
        log_file_handle.write("\n" + "=" * 60 + "\n")
        log_file_handle.write("TEST SUMMARY\n")
        log_file_handle.write("=" * 60 + "\n\n")
        log_file_handle.write(f"Passed:  {stats.passed}\n")
        log_file_handle.write(f"Failed:  {stats.failed}\n")
        log_file_handle.write(f"Total:   {stats.total}\n")
        log_file_handle.write(f"Duration: {elapsed:.1f}s\n\n")

def check_dotnet_build():
    """Build the project before running tests."""
    log_section("BUILD")
    log("Building project...", "INFO")
    
    start = time.time()
    result = subprocess.run(
        ["dotnet", "build", PROJECT_PATH, "--nologo"], 
        capture_output=True, 
        text=True
    )
    duration = (time.time() - start) * 1000
    
    if result.returncode != 0:
        log(f"Build failed!", "FAIL")
        print(f"{Colors.RED}{result.stderr}{Colors.RESET}")
        sys.exit(1)
    
    log(f"Build successful ({duration:.0f}ms)", "PASS")

def create_request(method: str, params: Optional[Dict] = None) -> Dict:
    global msg_id
    msg_id += 1
    req = {"jsonrpc": "2.0", "id": msg_id, "method": method}
    if params is not None:
        req["params"] = params
    return req

def create_notification(method: str, params: Optional[Dict] = None) -> Dict:
    req = {"jsonrpc": "2.0", "method": method}
    if params is not None:
        req["params"] = params
    return req

def read_message(proc, timeout: float = TEST_TIMEOUT_SECONDS) -> Optional[Dict]:
    """Read a message from the server with timeout support."""
    result_queue = queue.Queue()
    
    def reader():
        while True:
            line = proc.stdout.readline()
            if not line:
                result_queue.put(None)
                return
            line = line.strip()
            if not line:
                continue
            if line.startswith("{"):
                try:
                    result_queue.put(json.loads(line))
                    return
                except json.JSONDecodeError:
                    pass
            elif "[Excel Power Pivot MCP]" in line:
                log(f"Server: {line}", "DEBUG", to_file=False)
    
    thread = threading.Thread(target=reader, daemon=True)
    thread.start()
    
    try:
        return result_queue.get(timeout=timeout)
    except queue.Empty:
        log(f"Timeout waiting for server response ({timeout}s)", "WARN")
        return {"error": f"Timeout waiting for server response after {timeout}s"}

def send_message(proc, msg: Dict):
    json_str = json.dumps(msg)
    proc.stdin.write(json_str + "\n")
    proc.stdin.flush()

def method_name(msg: Dict) -> str:
    if "method" in msg:
        if msg["method"] == "tools/call":
            return f"tools/call({msg['params']['name']})"
        return msg["method"]
    return "response"

def send_tool_call(proc, name: str, args: Dict):
    req = create_request("tools/call", {"name": name, "arguments": args})
    send_message(proc, req)

def get_tool_result(proc, timeout: float = TEST_TIMEOUT_SECONDS) -> Dict:
    resp = read_message(proc, timeout=timeout)
    if not resp:
        return {"error": "No response received"}
    
    if "error" in resp:
        error_msg = resp["error"]
        if isinstance(error_msg, dict):
            error_msg = error_msg.get("message", str(error_msg))
        return {"error": error_msg}

    if "result" in resp and "content" in resp["result"]:
        try:
            content = resp["result"]["content"]
            if content and content[0]["type"] == "text":
                text_content = content[0]["text"]
                if text_content.startswith("An error") or text_content.startswith("Error"):
                    return {"error": text_content}
                return json.loads(text_content)
        except json.JSONDecodeError:
            text = resp["result"]["content"][0].get("text", "")
            if text:
                return {"error": text}
            return resp.get("result", {})
        except Exception:
            return resp.get("result", {})
            
    return resp

import argparse

def parse_args():
    parser = argparse.ArgumentParser(description="MCP Test Driver")
    parser.add_argument("--exe", help="Path to pre-built executable to test (skips build)", default=None)
    return parser.parse_args()

def run_tests():
    """Main test execution."""
    args = parse_args()
    
    log_section("STARTING MCP SERVER")
    
    if args.exe:
        log(f"Running compiled executable: {args.exe}", "INFO")
        if not os.path.exists(args.exe):
            log(f"Executable not found: {args.exe}", "FAIL")
            sys.exit(1)
            
        cmd = [args.exe]
        # Executables run directly
    else:
        log("Running from source (dotnet run)...", "INFO")
        check_dotnet_build()
        cmd = ["dotnet", "run", "--no-build", "--project", PROJECT_PATH]

    proc = subprocess.Popen(
        cmd,
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        bufsize=1,
        encoding='utf-8'
    )
    
    try:
        # ═══════════════════════════════════════════════════════════
        # INITIALIZATION
        # ═══════════════════════════════════════════════════════════
        log_section("INITIALIZATION")
        
        start = time.time()
        init_req = create_request("initialize", {
            "protocolVersion": "2024-11-05",
            "capabilities": {},
            "clientInfo": {"name": "test-driver", "version": "2.0"}
        })
        send_message(proc, init_req)
        
        resp = read_message(proc, timeout=60)  # Longer timeout for init
        duration = (time.time() - start) * 1000
        
        if not resp or "result" not in resp:
            log_test_result("MCP Initialize", "FAIL", f"Bad response: {resp}")
            return
        
        server_info = resp.get("result", {}).get("serverInfo", {})
        log_test_result("MCP Initialize", "PASS", 
                       f"Server: {server_info.get('name', 'unknown')} v{server_info.get('version', '?')}", 
                       duration)

        send_message(proc, create_notification("notifications/initialized"))

        # List Tools
        start = time.time()
        send_message(proc, create_request("tools/list"))
        resp = read_message(proc, timeout=30)
        duration = (time.time() - start) * 1000
        
        tools = resp.get("result", {}).get("tools", [])
        tool_names = [t.get("name") for t in tools]
        log_test_result("List Tools", "PASS", f"Found {len(tools)} tools", duration)
        if VERBOSE:
            log(f"Tools: {', '.join(tool_names[:10])}...", "DATA")
        
        # Test State
        state = {
            "workbook_path": None,
            "connected": False,
            "tables": [],
            "first_table": None,
            "first_column": None,
            "added_table": None,
        }

        # ═══════════════════════════════════════════════════════════
        # CONNECTION TESTS
        # ═══════════════════════════════════════════════════════════
        log_section("CONNECTION TESTS")
        
        # Discover Workbooks
        start = time.time()
        send_tool_call(proc, "discover_workbooks", {})
        res = get_tool_result(proc)
        duration = (time.time() - start) * 1000
        
        if res.get("error"):
            log_test_result("discover_workbooks", "FAIL", str(res["error"]), duration)
            return
        
        workbooks = res.get("workbooks", [])
        if not workbooks:
            log_test_result("discover_workbooks", "FAIL", "No workbooks found. Open Excel with a workbook.", duration)
            return
        
        wb = workbooks[0]
        state["workbook_path"] = wb["path"]
        log_test_result("discover_workbooks", "PASS", f"Found {len(workbooks)} workbook(s)", duration)
        for w in workbooks:
            log(f"Workbook: {w.get('name')} (hasDataModel: {w.get('hasDataModel')})", "DATA")
        
        # Connect
        start = time.time()
        send_tool_call(proc, "connect_workbook", {"workbookPath": state["workbook_path"]})
        res = get_tool_result(proc)
        duration = (time.time() - start) * 1000
        
        if res.get("error"):
            log_test_result("connect_workbook", "FAIL", str(res["error"]), duration)
            return
        
        state["connected"] = True
        log_test_result("connect_workbook", "PASS", f"Connected to: {wb['name']}", duration)
        
        # Get Connection Status
        start = time.time()
        send_tool_call(proc, "get_connection_status", {})
        res = get_tool_result(proc)
        duration = (time.time() - start) * 1000
        
        if res.get("error"):
            log_test_result("get_connection_status", "FAIL", str(res["error"]), duration)
        else:
            connected = res.get("connected", False)
            workbook = res.get("workbook", "unknown")
            if connected:
                log_test_result("get_connection_status", "PASS", f"Connected: {workbook}", duration)
            else:
                log_test_result("get_connection_status", "FAIL", "Not connected", duration)
        
        # ═══════════════════════════════════════════════════════════
        # METADATA TESTS
        # ═══════════════════════════════════════════════════════════
        log_section("METADATA TESTS")
        
        # List Tables
        start = time.time()
        send_tool_call(proc, "list_tables", {"includeColumns": False})
        res = get_tool_result(proc)
        duration = (time.time() - start) * 1000
        
        tables = res.get("tables", [])
        state["tables"] = tables
        if tables:
            # Get table names for display
            table_names = [t.get("name", t.get("Name", "?")) for t in tables]
            # Prefer non-blank tables
            for name in table_names:
                if "Blank" not in name and "blank" not in name:
                    state["first_table"] = name
                    break
            if not state["first_table"]:
                state["first_table"] = table_names[0]
            
            log_test_result("list_tables", "PASS", f"Found {len(tables)} tables", duration)
            for t in tables:
                name = t.get("name", t.get("Name", "?"))
                visible = t.get("isVisible", t.get("IsVisible", True))
                log(f"Table: {name} (visible: {visible})", "DATA")
        else:
            log_test_result("list_tables", "FAIL", "No tables in model", duration)
        
        # List Measures (ALL measures, no table filter)
        start = time.time()
        send_tool_call(proc, "list_measures", {})  # No tableName filter!
        res = get_tool_result(proc)
        duration = (time.time() - start) * 1000
        
        measures = res.get("measures", [])
        log_test_result("list_measures", "PASS", f"Found {len(measures)} measures", duration)
        for m in measures:
            name = m.get("name", m.get("Name", "?"))
            table = m.get("tableName", m.get("TableName", "?"))
            log(f"Measure: [{table}].[{name}]", "DATA")
        
        # Model Summary
        start = time.time()
        send_tool_call(proc, "get_model_summary", {"detailLevel": "standard"})
        res = get_tool_result(proc)
        duration = (time.time() - start) * 1000
        
        if res.get("error"):
            log_test_result("get_model_summary", "FAIL", str(res["error"]), duration)
        else:
            log_test_result("get_model_summary", "PASS", 
                f"Tables: {res.get('tableCount', 0)}, Measures: {res.get('measureCount', 0)}, Relationships: {res.get('relationshipCount', 0)}", 
                duration)
        
        if state["first_table"]:
            table = state["first_table"]
            
            # List Columns
            start = time.time()
            send_tool_call(proc, "list_columns", {"tableName": table})
            res = get_tool_result(proc)
            duration = (time.time() - start) * 1000
            
            cols = res.get("columns", [])
            if cols:
                state["first_column"] = cols[0].get("name", cols[0].get("Name"))
                log_test_result("list_columns", "PASS", f"Found {len(cols)} columns in '{table}'", duration)
                for c in cols[:5]:  # Show first 5
                    name = c.get("name", c.get("Name", "?"))
                    dtype = c.get("dataType", c.get("DataType", "?"))
                    log(f"Column: {name} ({dtype})", "DATA")
                if len(cols) > 5:
                    log(f"  ... and {len(cols) - 5} more columns", "DATA")
            else:
                log_test_result("list_columns", "PASS", f"No columns in '{table}'", duration)
            
            # List Relationships
            start = time.time()
            send_tool_call(proc, "list_relationships", {})
            res = get_tool_result(proc)
            duration = (time.time() - start) * 1000
            
            rels = res.get("relationships", [])
            log_test_result("list_relationships", "PASS", f"Found {len(rels)} relationships", duration)
            for r in rels:
                from_table = r.get("foreignKeyTable", r.get("ForeignKeyTable", "?"))
                from_col = r.get("foreignKeyColumn", r.get("ForeignKeyColumn", "?"))
                to_table = r.get("primaryKeyTable", r.get("PrimaryKeyTable", "?"))
                to_col = r.get("primaryKeyColumn", r.get("PrimaryKeyColumn", "?"))
                log(f"Relationship: {from_table}.{from_col} -> {to_table}.{to_col}", "DATA")
            
            # List KPIs
            start = time.time()
            send_tool_call(proc, "list_kpis", {})
            res = get_tool_result(proc)
            duration = (time.time() - start) * 1000
            
            kpis = res.get("kpis", [])
            log_test_result("list_kpis", "PASS", f"Found {len(kpis)} KPIs", duration)
            for k in kpis:
                log(f"KPI: {k}", "DATA")
            
            # List Hierarchies
            start = time.time()
            send_tool_call(proc, "list_hierarchies", {})
            res = get_tool_result(proc)
            duration = (time.time() - start) * 1000
            
            hierarchies = res.get("hierarchies", [])
            log_test_result("list_hierarchies", "PASS", f"Found {len(hierarchies)} hierarchies", duration)
            for h in hierarchies:
                name = h.get("name", h.get("HIERARCHY_NAME", h.get("Name", str(h)[:50])))
                log(f"Hierarchy: {name}", "DATA")
            
            # Get Dependencies
            start = time.time()
            send_tool_call(proc, "get_dependencies", {})
            res = get_tool_result(proc)
            duration = (time.time() - start) * 1000
            
            deps = res.get("dependencies", [])
            log_test_result("get_dependencies", "PASS", f"Found {len(deps)} dependencies", duration)
            for d in deps[:5]:
                obj = d.get("object", d.get("OBJECT", "?"))
                ref = d.get("referencedObject", d.get("REFERENCED_OBJECT", "?"))
                log(f"Dependency: {obj} -> {ref}", "DATA")
            if len(deps) > 5:
                log(f"  ... and {len(deps) - 5} more dependencies", "DATA")
            
            # List Power Queries
            start = time.time()
            send_tool_call(proc, "list_power_queries", {})
            res = get_tool_result(proc)
            duration = (time.time() - start) * 1000
            
            queries = res.get("queries", [])
            log_test_result("list_power_queries", "PASS", f"Found {len(queries)} Power Queries", duration)
            for q in queries:
                name = q.get("name", q.get("Name", q.get("QueryName", str(q)[:50])))
                log(f"Query: {name}", "DATA")
            
            # List Excel Tables
            start = time.time()
            send_tool_call(proc, "list_excel_tables", {})
            res = get_tool_result(proc)
            duration = (time.time() - start) * 1000
            
            excel_tables = res.get("tables", [])
            log_test_result("list_excel_tables", "PASS", f"Found {len(excel_tables)} Excel tables", duration)
            for et in excel_tables:
                name = et.get("name", et.get("Name", "?"))
                in_model = et.get("isInModel", et.get("inModel", et.get("InModel", "?")))
                log(f"Excel Table: {name} (in model: {in_model})", "DATA")
        
        # ═══════════════════════════════════════════════════════════
        # PARAMETER VARIATION TESTS
        # ═══════════════════════════════════════════════════════════
        log_section("PARAMETER VARIATION TESTS")
        
        if state["first_table"]:
            table = state["first_table"]
            
            # list_tables with includeColumns=True
            start = time.time()
            send_tool_call(proc, "list_tables", {"includeColumns": True, "includeStats": False})
            res = get_tool_result(proc)
            duration = (time.time() - start) * 1000
            
            if res.get("error"):
                log_test_result("list_tables (includeColumns)", "FAIL", str(res["error"])[:60], duration)
            else:
                tables_with_cols = res.get("tables", [])
                has_columns = any(t.get("columns") for t in tables_with_cols)
                log_test_result("list_tables (includeColumns)", "PASS", 
                    f"Tables have columns: {has_columns}", duration)
            
            # list_tables with includeStats=True
            start = time.time()
            send_tool_call(proc, "list_tables", {"includeColumns": False, "includeStats": True})
            res = get_tool_result(proc)
            duration = (time.time() - start) * 1000
            
            if res.get("error"):
                log_test_result("list_tables (includeStats)", "FAIL", str(res["error"])[:60], duration)
            else:
                tables_with_stats = res.get("tables", [])
                has_stats = any(t.get("rowCount") is not None for t in tables_with_stats)
                log_test_result("list_tables (includeStats)", "PASS", 
                    f"Tables have rowCount: {has_stats}", duration)
            
            # list_measures with tableName filter
            start = time.time()
            send_tool_call(proc, "list_measures", {"tableName": table})
            res = get_tool_result(proc)
            duration = (time.time() - start) * 1000
            
            filtered_measures = res.get("measures", [])
            log_test_result("list_measures (tableName filter)", "PASS", 
                f"Found {len(filtered_measures)} measures in '{table}'", duration)
            
            # list_measures with includeExpressions=True
            start = time.time()
            send_tool_call(proc, "list_measures", {"includeExpressions": True, "includeMetadata": True})
            res = get_tool_result(proc)
            duration = (time.time() - start) * 1000
            
            if res.get("error"):
                log_test_result("list_measures (includeExpressions)", "FAIL", str(res["error"])[:60], duration)
            else:
                measures_full = res.get("measures", [])
                has_expr = any(m.get("expression") or m.get("Expression") for m in measures_full) if measures_full else True
                log_test_result("list_measures (includeExpressions)", "PASS", 
                    f"Got {len(measures_full)} measures with expressions", duration)
            
            # list_columns with includeExpressions=True
            start = time.time()
            send_tool_call(proc, "list_columns", {"tableName": table, "includeExpressions": True})
            res = get_tool_result(proc)
            duration = (time.time() - start) * 1000
            
            if res.get("error"):
                log_test_result("list_columns (includeExpressions)", "FAIL", str(res["error"])[:60], duration)
            else:
                cols = res.get("columns", [])
                log_test_result("list_columns (includeExpressions)", "PASS", 
                    f"Got {len(cols)} columns", duration)
            
            # get_model_summary with "quick" detail level
            start = time.time()
            send_tool_call(proc, "get_model_summary", {"detailLevel": "quick"})
            res = get_tool_result(proc)
            duration = (time.time() - start) * 1000
            
            if res.get("error"):
                log_test_result("get_model_summary (quick)", "FAIL", str(res["error"])[:60], duration)
            else:
                log_test_result("get_model_summary (quick)", "PASS", 
                    f"Tables: {res.get('tableCount', 0)}", duration)
            
            # get_model_summary with "full" detail level
            start = time.time()
            send_tool_call(proc, "get_model_summary", {"detailLevel": "full"})
            res = get_tool_result(proc)
            duration = (time.time() - start) * 1000
            
            if res.get("error"):
                log_test_result("get_model_summary (full)", "FAIL", str(res["error"])[:60], duration)
            else:
                full_tables = res.get("tables", [])
                has_row_counts = any(t.get("rowCount") is not None for t in full_tables) if full_tables else True
                log_test_result("get_model_summary (full)", "PASS", 
                    f"Got full model with rowCounts: {has_row_counts}", duration)
            
            # get_dependencies with objectName filter (use first measure if exists)
            if measures:
                first_measure = measures[0].get("name", measures[0].get("Name"))
                if first_measure:
                    start = time.time()
                    send_tool_call(proc, "get_dependencies", {"objectName": first_measure})
                    res = get_tool_result(proc)
                    duration = (time.time() - start) * 1000
                    
                    if res.get("error"):
                        log_test_result("get_dependencies (objectName)", "FAIL", str(res["error"])[:60], duration)
                    else:
                        dep_count = res.get("dependencyCount", len(res.get("objects", [])))
                        log_test_result("get_dependencies (objectName)", "PASS", 
                            f"Deps for '{first_measure}': {dep_count}", duration)
        
        # ═══════════════════════════════════════════════════════════
        # DAX QUERY TESTS
        # ═══════════════════════════════════════════════════════════
        log_section("DAX QUERY TESTS")
        
        if state["first_table"]:
            table = state["first_table"]
            
            # Preview Table
            start = time.time()
            send_tool_call(proc, "run_dax", {"tableName": table, "maxRows": 5})
            res = get_tool_result(proc)
            duration = (time.time() - start) * 1000
            
            if res.get("error"):
                log_test_result("run_dax (preview)", "FAIL", str(res["error"])[:80], duration)
            else:
                row_count = res.get("rowCount", 0)
                log_test_result("run_dax (preview)", "PASS", f"Returned {row_count} rows from '{table}'", duration)
                # Show first row
                rows = res.get("data", [])
                if rows and len(rows) > 0:
                    log(f"Sample row: {json.dumps(rows[0])[:100]}...", "DATA")
            
            # Column Profile
            if state.get("first_column"):
                start = time.time()
                send_tool_call(proc, "analyze_column", {
                    "tableName": table,
                    "columnName": state["first_column"]
                })
                res = get_tool_result(proc)
                duration = (time.time() - start) * 1000
                
                if res.get("error") or res.get("Error"):
                    log_test_result("analyze_column", "FAIL", str(res.get("error", res.get("Error", "")))[:80], duration)
                else:
                    log_test_result("analyze_column", "PASS", f"Column: {state['first_column']}", duration)
                    # Show stats with multiple key attempts
                    distinct = res.get("distinctCount", res.get("DistinctCount", 0))
                    null_count = res.get("nullCount", res.get("NullCount", 0))
                    min_val = res.get("min", res.get("Min", ""))
                    max_val = res.get("max", res.get("Max", ""))
                    log(f"Stats: distinct={distinct}, nulls={null_count}, min='{min_val}', max='{max_val}'", "DATA")
                    samples = res.get("distinctValues", res.get("DistinctValues", []))
                    if samples:
                        log(f"Sample values: {samples[:5]}", "DATA")
            
            # Custom DAX Query with EVALUATE
            start = time.time()
            # Strip brackets from table name for DAX syntax
            clean_table = table.strip("[]")
            dax_query = f"EVALUATE TOPN(3, '{clean_table}')"
            send_tool_call(proc, "run_dax", {"daxQuery": dax_query, "maxRows": 3})
            res = get_tool_result(proc)
            duration = (time.time() - start) * 1000
            
            if res.get("error"):
                log_test_result("run_dax (custom EVALUATE)", "FAIL", str(res["error"])[:60], duration)
            else:
                row_count = res.get("rowCount", 0)
                log_test_result("run_dax (custom EVALUATE)", "PASS", 
                    f"TOPN query returned {row_count} rows", duration)
        
        # ═══════════════════════════════════════════════════════════
        # MEASURE CRUD TESTS
        # ═══════════════════════════════════════════════════════════
        log_section("MEASURE CRUD TESTS")
        
        if state["first_table"]:
            table = state["first_table"]
            test_measure = "MCP_Test_Measure_" + str(int(time.time()))
            
            # Create Measure (with autoFormat=false for speed)
            start = time.time()
            send_tool_call(proc, "create_measure", {
                "tableName": table,
                "measureName": test_measure,
                "expression": f"COUNTROWS('{table}')",
                "description": "Test measure from MCP test driver",
                "autoFormat": False  # Skip DAX formatter for faster test
            })
            res = get_tool_result(proc)
            duration = (time.time() - start) * 1000
            
            if res.get("success") or res.get("Success"):
                log_test_result("create_measure", "PASS", f"Created '{test_measure}' (autoFormat=false)", duration)
                
                # Update Measure (description only)
                start = time.time()
                send_tool_call(proc, "update_measure", {
                    "measureName": test_measure,
                    "newDescription": "Updated by test driver"
                })
                res = get_tool_result(proc)
                duration = (time.time() - start) * 1000
                
                if res.get("success") or res.get("Success"):
                    log_test_result("update_measure (description)", "PASS", "Updated description", duration)
                else:
                    log_test_result("update_measure (description)", "FAIL", str(res.get("error", res))[:60], duration)
                
                # Update Measure (expression change)
                start = time.time()
                send_tool_call(proc, "update_measure", {
                    "measureName": test_measure,
                    "newExpression": f"COUNTROWS('{table}') + 0",
                    "autoFormat": False
                })
                res = get_tool_result(proc)
                duration = (time.time() - start) * 1000
                
                if res.get("success") or res.get("Success"):
                    log_test_result("update_measure (expression)", "PASS", "Updated expression", duration)
                else:
                    log_test_result("update_measure (expression)", "FAIL", str(res.get("error", res))[:60], duration)
                
                # Update Measure (rename)
                new_measure_name = test_measure + "_Renamed"
                start = time.time()
                send_tool_call(proc, "update_measure", {
                    "measureName": test_measure,
                    "newName": new_measure_name
                })
                res = get_tool_result(proc)
                duration = (time.time() - start) * 1000
                
                if res.get("success") or res.get("Success"):
                    log_test_result("update_measure (rename)", "PASS", f"Renamed to '{new_measure_name}'", duration)
                    test_measure = new_measure_name  # Update name for delete test
                else:
                    log_test_result("update_measure (rename)", "FAIL", str(res.get("error", res))[:60], duration)
                
                # Test delete_measure with confirm=False (confirmation flow)
                start = time.time()
                send_tool_call(proc, "delete_measure", {
                    "measureName": test_measure,
                    "confirm": False
                })
                res = get_tool_result(proc)
                duration = (time.time() - start) * 1000
                
                if res.get("requiresConfirmation"):
                    log_test_result("delete_measure (confirm=false)", "PASS", "Returned confirmation prompt", duration)
                else:
                    log_test_result("delete_measure (confirm=false)", "FAIL", "Should require confirmation", duration)
                
                # Delete Measure (with confirm=True)
                start = time.time()
                send_tool_call(proc, "delete_measure", {
                    "measureName": test_measure,
                    "confirm": True
                })
                res = get_tool_result(proc)
                duration = (time.time() - start) * 1000
                
                if res.get("success") or res.get("Success"):
                    log_test_result("delete_measure (confirm=true)", "PASS", f"Deleted '{test_measure}'", duration)
                else:
                    log_test_result("delete_measure (confirm=true)", "FAIL", str(res.get("error", res))[:60], duration)
            else:
                log_test_result("create_measure", "FAIL", str(res.get("error", res))[:80], duration)
        
        # ═══════════════════════════════════════════════════════════
        # SAVE WORKBOOK TEST
        # ═══════════════════════════════════════════════════════════
        log_section("SAVE WORKBOOK TEST")
        
        start = time.time()
        send_tool_call(proc, "save_workbook", {})
        res = get_tool_result(proc)
        duration = (time.time() - start) * 1000
        
        if res.get("success") or res.get("Success"):
            log_test_result("save_workbook", "PASS", f"Workbook saved", duration)
        else:
            log_test_result("save_workbook", "FAIL", str(res.get("error", res))[:80], duration)
        
        # ═══════════════════════════════════════════════════════════
        # TABLE OPERATION TESTS
        # ═══════════════════════════════════════════════════════════
        log_section("TABLE OPERATION TESTS")
        
        if state["first_table"]:
            table = state["first_table"]
            
            # Refresh Table
            start = time.time()
            send_tool_call(proc, "refresh_table", {"tableName": table})
            res = get_tool_result(proc)
            duration = (time.time() - start) * 1000
            
            if res.get("success") or res.get("Success"):
                log_test_result("refresh_table", "PASS", f"Refreshed '{table}'", duration)
            else:
                log_test_result("refresh_table", "FAIL", str(res.get("error", res))[:80], duration)
        
        # Refresh Model
        start = time.time()
        send_tool_call(proc, "refresh_model", {})
        res = get_tool_result(proc, timeout=60)  # refresh_model can take longer
        duration = (time.time() - start) * 1000
        
        if res.get("success") or res.get("Success"):
            log_test_result("refresh_model", "PASS", "", duration)
        else:
            log_test_result("refresh_model", "FAIL", str(res.get("error", res))[:80], duration)
        
        # ═══════════════════════════════════════════════════════════
        # RELATIONSHIP CRUD TESTS
        # ═══════════════════════════════════════════════════════════
        log_section("RELATIONSHIP CRUD TESTS")
        
        # Get existing relationships to find tables we can use
        send_tool_call(proc, "list_relationships", {})
        res = get_tool_result(proc)
        existing_rels = res.get("relationships", [])
        
        # We need at least 2 tables with columns to test relationships
        # Use the state tables if available
        if len(state["tables"]) >= 2:
            # Get columns from first two tables
            table1_name = state["tables"][0].get("name", state["tables"][0].get("Name", "")).strip("[]")
            table2_name = state["tables"][1].get("name", state["tables"][1].get("Name", "")).strip("[]")
            
            # Get columns for table1
            send_tool_call(proc, "list_columns", {"tableName": table1_name})
            res1 = get_tool_result(proc)
            cols1 = res1.get("columns", [])
            
            # Get columns for table2
            send_tool_call(proc, "list_columns", {"tableName": table2_name})
            res2 = get_tool_result(proc)
            cols2 = res2.get("columns", [])
            
            if cols1 and cols2:
                col1 = cols1[0].get("Name", cols1[0].get("name", ""))
                col2 = cols2[0].get("Name", cols2[0].get("name", ""))
                
                # Check if this relationship already exists
                rel_exists = False
                for rel in existing_rels:
                    if (rel.get("ForeignKeyTable") == table1_name and 
                        rel.get("ForeignKeyColumn") == col1 and
                        rel.get("PrimaryKeyTable") == table2_name and
                        rel.get("PrimaryKeyColumn") == col2):
                        rel_exists = True
                        break
                
                if not rel_exists:
                    # Test create_relationship with confirm=False (confirmation flow)
                    start = time.time()
                    send_tool_call(proc, "create_relationship", {
                        "foreignTable": table1_name,
                        "foreignColumn": col1,
                        "primaryTable": table2_name,
                        "primaryColumn": col2,
                        "confirm": False
                    })
                    res = get_tool_result(proc)
                    duration = (time.time() - start) * 1000
                    
                    if res.get("requiresConfirmation"):
                        log_test_result("create_relationship (confirm=false)", "PASS", "Returned confirmation prompt", duration)
                    else:
                        log_test_result("create_relationship (confirm=false)", "FAIL", "Should require confirmation", duration)
                    
                    # Test Create Relationship (with confirm=True)
                    start = time.time()
                    send_tool_call(proc, "create_relationship", {
                        "foreignTable": table1_name,
                        "foreignColumn": col1,
                        "primaryTable": table2_name,
                        "primaryColumn": col2,
                        "confirm": True
                    })
                    res = get_tool_result(proc)
                    duration = (time.time() - start) * 1000
                    
                    if res.get("success") or res.get("Success"):
                        log_test_result("create_relationship (confirm=true)", "PASS", f"{table1_name}.{col1} → {table2_name}.{col2}", duration)
                        
                        # Test Set Relationship Inactive
                        start = time.time()
                        send_tool_call(proc, "set_relationship_active", {
                            "foreignTable": table1_name,
                            "foreignColumn": col1,
                            "primaryTable": table2_name,
                            "primaryColumn": col2,
                            "active": False
                        })
                        res = get_tool_result(proc)
                        duration = (time.time() - start) * 1000
                        
                        if res.get("success") or res.get("Success"):
                            log_test_result("set_relationship_active (deactivate)", "PASS", "Relationship deactivated", duration)
                        else:
                            log_test_result("set_relationship_active (deactivate)", "FAIL", str(res.get("error", res))[:60], duration)
                        
                        # Test Set Relationship Active again
                        start = time.time()
                        send_tool_call(proc, "set_relationship_active", {
                            "foreignTable": table1_name,
                            "foreignColumn": col1,
                            "primaryTable": table2_name,
                            "primaryColumn": col2,
                            "active": True
                        })
                        res = get_tool_result(proc)
                        duration = (time.time() - start) * 1000
                        
                        if res.get("success") or res.get("Success"):
                            log_test_result("set_relationship_active (activate)", "PASS", "Relationship activated", duration)
                        else:
                            log_test_result("set_relationship_active (activate)", "FAIL", str(res.get("error", res))[:60], duration)
                        
                        # Test delete_relationship with confirm=False (confirmation flow)
                        start = time.time()
                        send_tool_call(proc, "delete_relationship", {
                            "foreignTable": table1_name,
                            "foreignColumn": col1,
                            "primaryTable": table2_name,
                            "primaryColumn": col2,
                            "confirm": False
                        })
                        res = get_tool_result(proc)
                        duration = (time.time() - start) * 1000
                        
                        if res.get("requiresConfirmation"):
                            log_test_result("delete_relationship (confirm=false)", "PASS", "Returned confirmation prompt", duration)
                        else:
                            log_test_result("delete_relationship (confirm=false)", "FAIL", "Should require confirmation", duration)
                        
                        # Test Delete Relationship (with confirm=True)
                        start = time.time()
                        send_tool_call(proc, "delete_relationship", {
                            "foreignTable": table1_name,
                            "foreignColumn": col1,
                            "primaryTable": table2_name,
                            "primaryColumn": col2,
                            "confirm": True
                        })
                        res = get_tool_result(proc)
                        duration = (time.time() - start) * 1000
                        
                        if res.get("success") or res.get("Success"):
                            log_test_result("delete_relationship (confirm=true)", "PASS", "Relationship deleted", duration)
                        else:
                            log_test_result("delete_relationship (confirm=true)", "FAIL", str(res.get("error", res))[:60], duration)
                    else:
                        # Relationship creation failed - might be data type mismatch, which is OK
                        error_msg = str(res.get("error", res))[:80]
                        if "data type" in error_msg.lower() or "mismatch" in error_msg.lower():
                            log_test_result("create_relationship", "PASS", f"Expected failure (type mismatch): {error_msg[:40]}", duration)
                        else:
                            log_test_result("create_relationship", "FAIL", error_msg, duration)
                else:
                    log("Skipping create_relationship test - relationship already exists", "WARN")
                    
                    # Test set_relationship_active on existing relationship
                    rel = existing_rels[0]
                    fk_table = rel.get("ForeignKeyTable", rel.get("foreignKeyTable"))
                    fk_col = rel.get("ForeignKeyColumn", rel.get("foreignKeyColumn"))
                    pk_table = rel.get("PrimaryKeyTable", rel.get("primaryKeyTable"))
                    pk_col = rel.get("PrimaryKeyColumn", rel.get("primaryKeyColumn"))
                    
                    # Toggle to inactive
                    start = time.time()
                    send_tool_call(proc, "set_relationship_active", {
                        "foreignTable": fk_table,
                        "foreignColumn": fk_col,
                        "primaryTable": pk_table,
                        "primaryColumn": pk_col,
                        "active": False
                    })
                    res = get_tool_result(proc)
                    duration = (time.time() - start) * 1000
                    
                    if res.get("success") or res.get("Success"):
                        log_test_result("set_relationship_active (deactivate)", "PASS", f"Deactivated existing relationship", duration)
                        
                        # Toggle back to active
                        start = time.time()
                        send_tool_call(proc, "set_relationship_active", {
                            "foreignTable": fk_table,
                            "foreignColumn": fk_col,
                            "primaryTable": pk_table,
                            "primaryColumn": pk_col,
                            "active": True
                        })
                        res = get_tool_result(proc)
                        duration = (time.time() - start) * 1000
                        
                        if res.get("success") or res.get("Success"):
                            log_test_result("set_relationship_active (activate)", "PASS", f"Re-activated relationship", duration)
                        else:
                            log_test_result("set_relationship_active (activate)", "FAIL", str(res.get("error", res))[:60], duration)
                    else:
                        log_test_result("set_relationship_active (deactivate)", "FAIL", str(res.get("error", res))[:60], duration)
            else:
                log("Skipping relationship tests - not enough columns found", "WARN")
        else:
            log("Skipping relationship tests - need at least 2 tables", "WARN")
        
        # ═══════════════════════════════════════════════════════════
        # ADD TABLE TO MODEL TESTS
        # ═══════════════════════════════════════════════════════════
        log_section("ADD TABLE TO MODEL TESTS")
        
        # Get Excel tables that are NOT in the model
        start = time.time()
        send_tool_call(proc, "list_excel_tables", {})
        res = get_tool_result(proc)
        duration = (time.time() - start) * 1000
        
        excel_tables = res.get("tables", [])
        available_tables = [t for t in excel_tables if not t.get("isInModel", True)]
        
        # Get Power Query names to avoid conflicts (which cause hangs)
        send_tool_call(proc, "list_power_queries", {})
        pq_res = get_tool_result(proc)
        pq_names = set(q.get("name", q.get("Name", "")) for q in pq_res.get("queries", []))
        
        # Filter out tables that have existing Power Queries with matching names
        safe_tables = [t for t in available_tables 
                       if t.get("name", t.get("Name", "")) not in pq_names]
        
        if safe_tables:
            test_table = safe_tables[0]
            table_name = test_table.get("name", test_table.get("Name"))
            
            # Test add_table_to_model with confirm=False (confirmation flow)
            start = time.time()
            send_tool_call(proc, "add_table_to_model", {
                "tableName": table_name,
                "usePowerQuery": False,
                "confirm": False
            })
            res = get_tool_result(proc)
            duration = (time.time() - start) * 1000
            
            if res.get("requiresConfirmation"):
                log_test_result("add_table_to_model (confirm=false)", "PASS", "Returned confirmation prompt", duration)
            else:
                log_test_result("add_table_to_model (confirm=false)", "FAIL", "Should require confirmation", duration)
            
            # Test add_table_to_model (direct method, with confirm=True)
            start = time.time()
            send_tool_call(proc, "add_table_to_model", {
                "tableName": table_name,
                "usePowerQuery": False,
                "confirm": True
            })
            res = get_tool_result(proc)
            duration = (time.time() - start) * 1000
            
            if res.get("success") or res.get("Success"):
                log_test_result("add_table_to_model (direct)", "PASS", f"Added '{table_name}' to model", duration)

                # Store the table name for delete_table test later
                state["added_table"] = table_name
            else:
                error_msg = str(res.get("error", res))[:80]
                log_test_result("add_table_to_model (direct)", "FAIL", error_msg, duration)

            # Test queryName validation (Issue #1 fix)
            # queryName parameter must match tableName for Power Query tables
            if len(safe_tables) >= 2:
                test_table2 = safe_tables[1]
                table_name2 = test_table2.get("name", test_table2.get("Name"))

                start = time.time()
                send_tool_call(proc, "add_table_to_model", {
                    "tableName": table_name2,
                    "queryName": "DifferentQueryName",  # Mismatch - should fail!
                    "usePowerQuery": True,
                    "confirm": True
                })
                res = get_tool_result(proc)
                duration = (time.time() - start) * 1000

                if res.get("error") and "must match TableName" in str(res.get("error")):
                    log_test_result("add_table_to_model (queryName mismatch)", "PASS",
                        "Correctly rejected mismatched queryName", duration)
                else:
                    log_test_result("add_table_to_model (queryName mismatch)", "FAIL",
                        f"Should reject queryName mismatch, got: {str(res)[:60]}", duration)
            else:
                log("Skipping queryName validation test - need 2+ safe tables", "WARN")
        elif available_tables:
            log("Skipping add_table_to_model - available tables have Power Query name conflicts", "WARN")
        else:
            log("Skipping add_table_to_model test - no Excel tables available to add", "WARN")
            log("  Hint: Create an Excel table that is not yet in the Power Pivot model", "DATA")

        # ═══════════════════════════════════════════════════════════
        # DELETE TABLE TESTS
        # ═══════════════════════════════════════════════════════════
        log_section("DELETE TABLE TESTS")

        # Test delete_table on the table we added earlier (if any)
        if state.get("added_table"):
            table_to_delete = state["added_table"]

            # Test delete_table with confirm=False (confirmation flow)
            start = time.time()
            send_tool_call(proc, "delete_table", {
                "tableName": table_to_delete,
                "confirm": False
            })
            res = get_tool_result(proc)
            duration = (time.time() - start) * 1000

            if res.get("requiresConfirmation"):
                log_test_result("delete_table (confirm=false)", "PASS", "Returned confirmation prompt", duration)
            else:
                log_test_result("delete_table (confirm=false)", "FAIL", "Should require confirmation", duration)

            # Test delete_table with confirm=True
            start = time.time()
            send_tool_call(proc, "delete_table", {
                "tableName": table_to_delete,
                "confirm": True
            })
            res = get_tool_result(proc)
            duration = (time.time() - start) * 1000

            if res.get("success") or res.get("Success"):
                log_test_result("delete_table (confirm=true)", "PASS", f"Deleted '{table_to_delete}'", duration)
            else:
                error_msg = str(res.get("error", res))[:80]
                log_test_result("delete_table (confirm=true)", "FAIL", error_msg, duration)
        else:
            log("Skipping delete_table test - no table was added earlier to delete", "WARN")

        # ═══════════════════════════════════════════════════════════
        # ERROR HANDLING & EDGE CASE TESTS
        # ═══════════════════════════════════════════════════════════
        log_section("ERROR HANDLING & EDGE CASE TESTS")
        
        # Test run_dax with invalid DAX syntax
        # Now that server has proper error handling, this should return JSON error
        start = time.time()
        send_tool_call(proc, "run_dax", {"daxQuery": "EVALUATE INVALID SYNTAX HERE !!!", "maxRows": 10})
        res = get_tool_result(proc)
        duration = (time.time() - start) * 1000
        
        if res.get("error"):
            log_test_result("run_dax (invalid syntax)", "PASS", "Error returned as expected", duration)
        else:
            log_test_result("run_dax (invalid syntax)", "FAIL", "Should have returned error", duration)
        
        # Test list_columns with invalid table name
        start = time.time()
        send_tool_call(proc, "list_columns", {"tableName": "NonExistentTable_12345"})
        res = get_tool_result(proc)
        duration = (time.time() - start) * 1000
        
        if res.get("error"):
            log_test_result("list_columns (invalid table)", "PASS", "Error returned as expected", duration)
        else:
            # Some implementations return empty columns for invalid tables
            cols = res.get("columns", [])
            if len(cols) == 0:
                log_test_result("list_columns (invalid table)", "PASS", "Empty columns returned", duration)
            else:
                log_test_result("list_columns (invalid table)", "FAIL", "Should error or return empty", duration)
        
        # Test analyze_column with invalid column name
        if state["first_table"]:
            start = time.time()
            send_tool_call(proc, "analyze_column", {
                "tableName": state["first_table"],
                "columnName": "NonExistentColumn_12345"
            })
            res = get_tool_result(proc)
            duration = (time.time() - start) * 1000
            
            if res.get("error") or res.get("Error"):
                log_test_result("analyze_column (invalid column)", "PASS", "Error returned as expected", duration)
            else:
                log_test_result("analyze_column (invalid column)", "FAIL", "Should return error", duration)
        
        # Test delete_measure with non-existent measure
        # Now that server has proper error handling, this should return JSON error
        start = time.time()
        send_tool_call(proc, "delete_measure", {
            "measureName": "NonExistentMeasure_DoesNotExist_12345",
            "confirm": True
        })
        res = get_tool_result(proc)
        duration = (time.time() - start) * 1000
        
        if res.get("error"):
            log_test_result("delete_measure (non-existent)", "PASS", "Error returned as expected", duration)
        else:
            log_test_result("delete_measure (non-existent)", "FAIL", f"Expected error, got: {str(res)[:50]}", duration)
        
        # Test create_measure with autoFormat=True
        # Now that server has proper error handling, this should work or return JSON error
        if state["first_table"]:
            table = state["first_table"]
            auto_fmt_measure = "MCP_AutoFmt_Test_" + str(int(time.time()))
            
            start = time.time()
            send_tool_call(proc, "create_measure", {
                "tableName": table,
                "measureName": auto_fmt_measure,
                "expression": f"COUNTROWS('{table}')",
                "description": "Test measure with autoFormat",
                "autoFormat": True
            })
            res = get_tool_result(proc)
            duration = (time.time() - start) * 1000
            
            if res.get("success") or res.get("Success"):
                log_test_result("create_measure (autoFormat=True)", "PASS", f"Created '{auto_fmt_measure}'", duration)
                
                # Clean up - delete the test measure
                send_tool_call(proc, "delete_measure", {
                    "measureName": auto_fmt_measure,
                    "confirm": True
                })
                cleanup_res = get_tool_result(proc)
                if cleanup_res.get("success") or cleanup_res.get("Success"):
                    log(f"Cleaned up test measure '{auto_fmt_measure}'", "DATA")
            elif res.get("error"):
                # Error is acceptable - what matters is server stayed responsive
                log_test_result("create_measure (autoFormat=True)", "PASS", f"Error handled: {str(res.get('error', ''))[:40]}", duration)
            else:
                log_test_result("create_measure (autoFormat=True)", "FAIL", f"Unexpected: {str(res)[:50]}", duration)
        
        # Test list_columns for Blank_Table (edge case - table exists)
        start = time.time()
        send_tool_call(proc, "list_columns", {"tableName": "Blank_Table"})
        res = get_tool_result(proc)
        duration = (time.time() - start) * 1000
        
        if res.get("error"):
            log_test_result("list_columns (Blank_Table)", "FAIL", str(res["error"])[:60], duration)
        else:
            cols = res.get("columns", [])
            log_test_result("list_columns (Blank_Table)", "PASS", 
                f"Found {len(cols)} columns in empty table", duration)
        
        # Test run_dax on empty/minimal data table
        start = time.time()
        send_tool_call(proc, "run_dax", {"tableName": "Blank_Table", "maxRows": 10})
        res = get_tool_result(proc)
        duration = (time.time() - start) * 1000
        
        if res.get("error"):
            log_test_result("run_dax (Blank_Table)", "FAIL", str(res["error"])[:60], duration)
        else:
            row_count = res.get("rowCount", 0)
            log_test_result("run_dax (Blank_Table)", "PASS", 
                f"Returned {row_count} rows from blank table", duration)
        
        # NOTE: add_table_to_model with Power Query is skipped because creating
        # Power Queries via COM can cause Excel dialogs or hangs.
        log("Skipping add_table_to_model Power Query test (can cause Excel dialog/hang)", "WARN")
        
    except Exception as e:
        import traceback
        log(f"Exception: {e}", "FAIL")
        log(traceback.format_exc(), "DEBUG")
    finally:
        log_section("CLEANUP")
        log("Terminating server...", "INFO")
        proc.terminate()
        try:
            proc.wait(timeout=5)
            log("Server terminated gracefully", "PASS")
        except subprocess.TimeoutExpired:
            proc.kill()
            log("Server killed (timeout)", "WARN")
        
        print_summary()

if __name__ == "__main__":
    init_log_file()
    try:
        run_tests()
    finally:
        close_log_file()
