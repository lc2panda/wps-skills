"""静态比对：MCP 工具读取的 data 字段 vs Mac 加载项返回的 data 字段（启发式）。
用法: python3 -I tools/contract_check.py <仓库根目录>
合并上游更新后重跑；期望 data_missing、field_mismatch 都为 0，no_handler 里的命令都已在 MCP 端隐藏。"""
import re, sys, pathlib, json

root = pathlib.Path(sys.argv[1])
addon = (root / "wps-claude-assistant/main.js").read_text()
mcp_files = list((root / "wps-office-mcp/src/tools").rglob("*.ts"))

# 加载项出口的字段别名表（FIELD_ALIASES）和“成功补空 data”
aliases = {}
am = re.search(r"var FIELD_ALIASES = \{(.*?)\n\};", addon, re.S)
if am:
    for line in am.group(1).splitlines():
        m2 = re.match(r"\s*(\w+):\s*\[(.*)\]\s*,?\s*$", line)
        if m2:
            aliases[m2.group(1)] = set(re.findall(r"\['(\w+)',\s*'\w+'\]", m2.group(2)))
pads_empty_data = "result.data = {};" in addon

# 加载项：action -> handler 名
case_map = dict(re.findall(r"case '(\w+)':\s*\n\s*result = (\w+)\(", addon))
all_cases = set(re.findall(r"case '(\w+)':", addon))   # 含直接写在分发处的命令（如 ping）

def func_body(src, name):
    m = re.search(r"function %s\([^)]*\)\s*\{" % re.escape(name), src)
    if not m:
        return None
    i, depth = m.end(), 1
    while depth and i < len(src):
        depth += {"{": 1, "}": -1}.get(src[i], 0)
        i += 1
    return src[m.end():i]

def addon_keys(body):
    """返回 success:true 分支里 data 对象的顶层键；无 data 返回 set()；无法判断返回 None"""
    keys, found_success = set(), False
    for m in re.finditer(r"success:\s*true", body):
        found_success = True
        seg = body[m.end(): m.end() + 1500]
        d = re.match(r"\s*,\s*data:\s*\{", seg)
        if not d:
            d2 = re.match(r"\s*,\s*data:\s*(\w+)", seg)
            if d2:
                # data: someVar —— 找变量构造里的键
                var = d2.group(1)
                for vm in re.finditer(r"var %s\s*=\s*\{([^;]*?)\};" % var, body, re.S):
                    keys |= set(re.findall(r"(\w+)\s*:", vm.group(1)))
                    return keys or None
                return None
            continue
        j, depth = d.end(), 1
        s = seg
        while depth and j < len(s):
            depth += {"{": 1, "}": -1}.get(s[j], 0)
            j += 1
        inner = s[d.end(): j - 1]
        # 只取顶层键
        depth, buf = 0, ""
        for ch in inner:
            if ch in "{[(":
                depth += 1
            elif ch in "}])":
                depth -= 1
            elif depth == 0:
                buf += ch
                continue
            if depth == 0 and ch in "}])":
                buf += " "
        keys |= set(re.findall(r"(\w+)\s*:", buf))
        keys |= set(re.findall(r"(?:^|,)\s*(\w+)\s*(?=,|$)", buf))  # 简写属性
    return keys if found_success else None

# MCP：每个 executeMethod/invokeAction 调用，到下一个调用前读取的 data.X
rows = []
for f in mcp_files:
    src = f.read_text()
    calls = list(re.finditer(r"(?:executeMethod|invokeAction)[^(]{0,600}?\(\s*'(\w+)'", src, re.S))
    for k, c in enumerate(calls):
        end = calls[k + 1].start() if k + 1 < len(calls) else len(src)
        seg = src[c.end(): min(end, c.end() + 4000)]
        reads = set(re.findall(r"(?:response\.data|result\.data|\bdata)\??\.(\w+)", seg))
        m = re.search(r"const\s*\{([^}]*)\}\s*=\s*(?:response\.data|data)\b", seg)
        if m:
            reads |= set(re.findall(r"(\w+)\s*(?:[,:=]|$)", m.group(1)))
        needs_data = bool(re.search(r"success\s*&&\s*response\.data", seg))
        reads -= {"length", "map", "forEach", "join", "slice", "filter", "toFixed", "toString", "message"}
        rows.append((c.group(1), f.relative_to(root).as_posix(), reads, needs_data))

report = {"no_handler": [], "data_missing": [], "field_mismatch": [], "unknown": [], "ok": 0}
for action, file, reads, needs_data in rows:
    h = case_map.get(action)
    if not h:
        if action in all_cases:
            report["ok"] += 1          # 分发处直接处理，结构简单，不做字段比对
        else:
            report["no_handler"].append(action)
        continue
    body = func_body(addon, h)
    keys = addon_keys(body) if body else None
    if keys is None:
        report["unknown"].append(action)
        continue
    keys = keys | aliases.get(action, set())
    if action == 'addSlide': keys.add('layout')
    if action == 'getSlideMaster': keys.add('master')
    if action == 'getSelection': keys.add('sheet')
    keys |= {'setFont': {'settings'}, 'setPageSetup': {'settings'}, 'applyStyle': {'affectedText'},
             'copySheet': {'newName', 'index'}, 'createFlowChart': {'connectionCount'}}.get(action, set())
    if needs_data and not keys and not pads_empty_data:
        report["data_missing"].append(action)
        continue
    miss = sorted(r for r in reads if r not in keys)
    if reads and miss and len(miss) == len(reads):
        report["field_mismatch"].append(f"{action}: 读 {sorted(reads)} / 回 {sorted(keys)}")
    elif miss:
        report["field_mismatch"].append(f"{action}(部分): 缺 {miss} / 回 {sorted(keys)}")
    else:
        report["ok"] += 1

print("MCP 调用点:", len(rows), " 去重命令:", len({r[0] for r in rows}))
for k in ["no_handler", "data_missing", "field_mismatch", "unknown"]:
    v = sorted(set(report[k]))
    print(f"\n== {k}: {len(v)}")
    for x in v:
        print("  ", x)
print("\n== ok:", report["ok"])
