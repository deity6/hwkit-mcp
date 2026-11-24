"""hwkit_mcp.pbipack / hwkit_mcp.modelcheck 自测（纯标准库，直接跑）。

    C:\\Python314\\python.exe tests\\test_pbipack.py            # 跑全部
    C:\\Python314\\python.exe tests\\test_pbipack.py --keep     # 保留临时目录

覆盖：
  1. inspect_pbix 打开真实 pbix（条目数 / 页面 / visual 数 / DataModel 字节数）
  2. extract_pbix 解包，文件数与 zip 条目数一致
  3. build_pbix 用真实 pbix 当模板原样回灌 Report，条目集合一致 + DataModel 哈希不变
     3b. pack_dir 解包目录重新打包（DataModel 取自模板 pbix）
  4. validate_tmdl 校验真实工程 TMDL（3 表 / 33 度量值 / 2 关系）
  5. 负向：人为插入非法属性 description，validate_tmdl 必须 ok=False 且 error 命中关键字
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
import sys
import tempfile
import time
import zipfile
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from hwkit_mcp.pbipack import (  # noqa: E402
    build_pbix,
    extract_pbix,
    inspect_pbix,
    list_entries,
    pack_dir,
)
from hwkit_mcp.modelcheck import (  # noqa: E402
    find_powershell,
    find_tabular_dll,
    validate_pbip,
    validate_tmdl,
)

PBI_DIR = Path(
    os.environ.get("HWKIT_TEST_ASSETS", "")
)
PBIX = PBI_DIR / "新新公司经营分析看板.pbix"
PBIP = PBI_DIR / "新新公司经营分析看板.pbip"
REPORT_DIR = PBI_DIR / "新新公司经营分析看板.Report"
SEMANTIC_DEF = PBI_DIR / "新新公司经营分析看板.SemanticModel" / "definition"

EXPECTED_DATAMODEL_BYTES = 161360
EXPECTED_TOTAL_MEASURES = 33
EXPECTED_MEASURES_BY_TABLE = {"日期": 0, "销售明细": 10, "财务月度": 23}

_RESULTS: list[tuple[bool, str, str]] = []


def check(cond: bool, label: str, info: str = "") -> bool:
    _RESULTS.append((bool(cond), label, info))
    print(f"  [{'PASS' if cond else 'FAIL'}] {label}" + (f"  — {info}" if info else ""))
    return bool(cond)


def section(title: str) -> None:
    print(f"\n=== {title} ===")


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def zip_entry_bytes(path: Path, name: str) -> bytes:
    with zipfile.ZipFile(path) as zf:
        return zf.read(name)


def zip_entry_compress_type(path: Path, name: str) -> int:
    with zipfile.ZipFile(path) as zf:
        return zf.getinfo(name).compress_type


def zipfile_is_intact(path: Path) -> bool:
    with zipfile.ZipFile(path) as zf:
        return zf.testzip() is None


def page_map(info: dict) -> dict[str, int]:
    return {p["name"]: p["visuals"] for p in info["pages"]}


# --------------------------------------------------------------------------- #
def test_1_inspect(tmp: Path) -> None:
    section("测试 1  inspect_pbix() 打开真实 pbix")
    info = inspect_pbix(str(PBIX))
    print("  entries=%d pages=%s version=%r datamodel_bytes=%d"
          % (len(info["entries"]), page_map(info), info["version"], info["datamodel_bytes"]))

    check(len(info["entries"]) >= 33, "条目数 >= 33", f"实际 {len(info['entries'])}")
    check(info["is_pbix"] is True and info["is_pbit"] is False, "is_pbix=True / is_pbit=False")
    check(info["has_datamodel"] is True, "has_datamodel=True")
    check(
        info["datamodel_bytes"] == EXPECTED_DATAMODEL_BYTES,
        f"datamodel_bytes == {EXPECTED_DATAMODEL_BYTES}",
        f"实际 {info['datamodel_bytes']}",
    )
    check(
        info["datamodel_compress_type"] == 0,
        "DataModel compress_type == 0 (stored)",
        f"实际 {info['datamodel_compress_type']}",
    )
    names = [p["name"] for p in info["pages"]]
    check(names == ["Finance", "Operations"], "pages == [Finance, Operations]", str(names))
    displays = {p["name"]: p["displayName"] for p in info["pages"]}
    check(
        displays.get("Finance") == "财务看板" and displays.get("Operations") == "经营看板",
        "displayName 正确（财务看板 / 经营看板）",
        str(displays),
    )
    pm = page_map(info)
    check(pm.get("Finance") == 10, "Finance visuals == 10", f"实际 {pm.get('Finance')}")
    check(pm.get("Operations") == 10, "Operations visuals == 10", f"实际 {pm.get('Operations')}")
    containers = {p["name"]: p["visual_containers"] for p in info["pages"]}
    check(containers == {"Finance": 11, "Operations": 11},
          "每页实际有 11 个 visual 容器（10 数据 + 1 hdr 文本框）", str(containers))
    types = {p["name"]: p["visual_types"] for p in info["pages"]}
    check(all(t.get("textbox") == 1 for t in types.values()),
          "每页恰有 1 个 textbox（hdr 装饰容器）",
          json.dumps(types, ensure_ascii=False))
    check(len(info["theme_files"]) == 1, "theme_files 有 1 个自定义主题", str(info["theme_files"]))
    check(info["version"] == "1.32", "version == '1.32'", repr(info["version"]))
    check(info["active_page"] == "Finance", "active_page == 'Finance'", str(info["active_page"]))

    entries = list_entries(str(PBIX))
    check(len(entries) == len(info["entries"]), "list_entries 与 inspect 条目数一致",
          f"{len(entries)} vs {len(info['entries'])}")
    dm = [e for e in entries if e["name"] == "DataModel"][0]
    print("  DataModel 条目: size=%d compressed=%d compress_type=%d"
          % (dm["size"], dm["compressed"], dm["compress_type"]))
    check(dm["size"] == dm["compressed"], "DataModel compressed == size（stored 不压缩）")


def test_2_extract(tmp: Path) -> Path:
    section("测试 2  extract_pbix()")
    out = tmp / "extract"
    res = extract_pbix(str(PBIX), str(out))
    entry_count = len(list_entries(str(PBIX)))
    print("  result=%s  条目数=%d" % ({k: res[k] for k in ("ok", "files", "entries")}, entry_count))
    check(res["ok"] is True, "extract ok")
    check(res["files"] == entry_count, "落盘文件数 == zip 条目数",
          f"{res['files']} vs {entry_count}")
    check((out / "Report" / "definition" / "report.json").is_file(), "写出 Report/definition/report.json")
    check((out / "DataModel").is_file(), "写出 DataModel")
    check((out / "DataModel").stat().st_size == EXPECTED_DATAMODEL_BYTES, "DataModel 落盘字节数一致")
    return out


def test_3_build(tmp: Path, extract_dir: Path) -> None:
    section("测试 3  build_pbix() 用真实 pbix 当模板原样回灌 Report")
    out = tmp / "rebuilt.pbix"
    res = build_pbix(str(extract_dir / "Report"), str(out), str(PBIX))
    print("  result ok=%s bytes=%d replaced=%d kept=%d written=%d ct_patched=%s dm_sha=%s"
          % (res.get("ok"), res.get("bytes", 0), res.get("entries_replaced", -1),
             res.get("entries_kept", -1), res.get("entries_written", -1),
             res.get("content_types_patched"), (res.get("datamodel_sha256") or "")[:16]))
    check(res["ok"] is True, "build ok", str(res.get("error") or ""))
    check(out.is_file(), "产出文件存在")

    orig_names = [e["name"] for e in list_entries(str(PBIX))]
    new_names = [e["name"] for e in list_entries(str(out))]
    check(set(orig_names) == set(new_names), "条目集合与原文件完全一致",
          f"原 {len(set(orig_names))} 个 / 新 {len(set(new_names))} 个"
          f" / 差集={sorted(set(orig_names) ^ set(new_names))}")
    check(len(new_names) == len(set(new_names)), "新文件无重复条目", f"{len(new_names)} 条")

    info = inspect_pbix(str(out))
    check(info["is_pbix"] is True, "新文件能被 inspect_pbix 正常解析")
    check(page_map(info) == {"Finance": 10, "Operations": 10}, "新文件页面与 visual 数正确",
          str(page_map(info)))
    check(info["datamodel_bytes"] == EXPECTED_DATAMODEL_BYTES, "新文件 DataModel 字节数不变",
          str(info["datamodel_bytes"]))
    check(info["datamodel_compress_type"] == 0, "新文件 DataModel 仍是 stored(0)",
          str(info["datamodel_compress_type"]))

    a = sha256_bytes(zip_entry_bytes(PBIX, "DataModel"))
    b = sha256_bytes(zip_entry_bytes(out, "DataModel"))
    print("  DataModel sha256 原=%s\n                  新=%s" % (a, b))
    check(a == b, "DataModel 字节级相同（sha256 一致）")

    # 非 Report 条目也必须逐字节一致（Version/Settings/Metadata/DiagramLayout/SecurityBindings/CT）
    others = [n for n in orig_names if not n.startswith("Report/")]
    bad = []
    for name in others:
        if name == "[Content_Types].xml":
            continue
        if zip_entry_bytes(PBIX, name) != zip_entry_bytes(out, name):
            bad.append(name)
    check(not bad, "非 Report 条目（%d 个）逐字节一致" % (len(others) - 1), str(bad))

    ct_orig = zip_entry_bytes(PBIX, "[Content_Types].xml")
    ct_new = zip_entry_bytes(out, "[Content_Types].xml")
    check(ct_orig == ct_new, "[Content_Types].xml 未被无谓改写（json Default 已存在）",
          "content_types_patched=%s" % res.get("content_types_patched"))

    orig_types = {e["name"]: e["compress_type"] for e in list_entries(str(PBIX))}
    new_types = {e["name"]: e["compress_type"] for e in list_entries(str(out))}
    diff = {
        n: (orig_types[n], new_types.get(n))
        for n in orig_types
        if not n.startswith("Report/") and orig_types[n] != new_types.get(n)
    }
    check(not diff, "非 Report 条目的 compress_type 全部与模板一致（ZipInfo 被复用）", str(diff))
    check(zipfile_is_intact(out), "产物 zip 结构完整（testzip 无坏条目）")
    print("  zip 大小：原 %d B -> 新 %d B（Report 内 json 重新 deflate，DataModel 原样 stored）"
          % (PBIX.stat().st_size, out.stat().st_size))

    # Report 侧内容必须与源目录逐字节一致（只有压缩方式可能不同）
    src_report = extract_dir / "Report"
    mismatch = []
    with zipfile.ZipFile(out) as zf:
        for name in zf.namelist():
            if not name.startswith("Report/"):
                continue
            rel = name[len("Report/"):]
            disk = src_report / rel.replace("/", "\\")
            if not disk.is_file() or zf.read(name) != disk.read_bytes():
                mismatch.append(name)
    check(not mismatch, "Report/** 内容与源目录逐字节一致", str(mismatch[:5]))


def test_3b_pack_dir(tmp: Path, extract_dir: Path) -> None:
    section("测试 3b  pack_dir() 把解包目录重新打包（DataModel 取自模板 pbix）")
    out = tmp / "repacked.pbix"
    res = pack_dir(str(extract_dir), str(out), datamodel_from=str(PBIX))
    print("  result ok=%s bytes=%d files=%d dm=%d" %
          (res.get("ok"), res.get("bytes", 0), res.get("files", -1), res.get("datamodel_bytes", 0)))
    check(res["ok"] is True, "pack_dir ok", str(res.get("error") or ""))
    info = inspect_pbix(str(out))
    check(page_map(info) == {"Finance": 10, "Operations": 10}, "pack_dir 产物可解析且页面正确",
          str(page_map(info)))
    a = sha256_bytes(zip_entry_bytes(PBIX, "DataModel"))
    b = sha256_bytes(zip_entry_bytes(out, "DataModel"))
    check(a == b, "pack_dir 产物 DataModel sha256 与模板一致")
    check(info["datamodel_compress_type"] == 0, "pack_dir 产物 DataModel 仍是 stored(0)",
          str(info["datamodel_compress_type"]))
    check(set(e["name"] for e in list_entries(str(out))) ==
          set(e["name"] for e in list_entries(str(PBIX))), "pack_dir 产物条目集合与原文件一致")


def test_3c_extra_resources(tmp: Path, extract_dir: Path) -> None:
    section("测试 3c  附加静态资源 + [Content_Types].xml 补 Default（正向分支）")
    extra = tmp / "extra_themes"
    (extra / "RegisteredResources").mkdir(parents=True, exist_ok=True)
    (extra / "RegisteredResources" / "MyTheme.json").write_text(
        json.dumps({"name": "MyTheme", "dataColors": ["#123456"]}, ensure_ascii=False),
        encoding="utf-8",
    )
    (extra / "logo.svg").write_text("<svg></svg>", encoding="utf-8")

    out = tmp / "with_extra.pbix"
    res = build_pbix(
        str(extract_dir / "Report"), str(out), str(PBIX), extra_static_resources_dir=str(extra)
    )
    print("  ok=%s written=%d ct_patched=%s" %
          (res.get("ok"), res.get("entries_written", -1), res.get("content_types_patched")))
    check(res["ok"] is True, "build_pbix(extra_static_resources_dir=...) ok", str(res.get("error") or ""))

    names = [e["name"] for e in list_entries(str(out))]
    reg = "Report/StaticResources/RegisteredResources/"
    check(reg + "MyTheme.json" in names, "RegisteredResources/ 下文件映射正确",
          str([n for n in names if "MyTheme" in n]))
    check(reg + "logo.svg" in names, "裸文件补进 RegisteredResources/", str([n for n in names if "logo." in n]))
    info = inspect_pbix(str(out))
    check(len(info["theme_files"]) == 3, "theme_files 现在有 3 个（原 1 + 新 2）",
          str(info["theme_files"]))

    ct = zip_entry_bytes(out, "[Content_Types].xml").decode("utf-8-sig")
    check('Extension="svg"' in ct, "[Content_Types].xml 补上了 svg 的 Default",
          re.findall(r'<Default[^>]*Extension="svg"[^>]*/>', ct) and "yes" or ct[-200:])
    check('Extension="json"' in ct and ct.count('Extension="json"') == 1,
          "json 的 Default 没有被重复插入")
    check(info["datamodel_compress_type"] == 0, "DataModel 仍 stored(0)")
    check(zipfile_is_intact(out), "产物 zip 结构完整")


def test_3d_pack_dir_local_datamodel(tmp: Path, extract_dir: Path) -> None:
    section("测试 3d  pack_dir() 不带 datamodel_from（用目录里现有的 DataModel）")
    out = tmp / "repacked_local.pbix"
    res = pack_dir(str(extract_dir), str(out))
    print("  ok=%s bytes=%d dm=%d" % (res.get("ok"), res.get("bytes", 0), res.get("datamodel_bytes", 0)))
    check(res["ok"] is True, "pack_dir(datamodel_from='') ok", str(res.get("error") or ""))
    a = sha256_bytes(zip_entry_bytes(PBIX, "DataModel"))
    b = sha256_bytes(zip_entry_bytes(out, "DataModel"))
    check(a == b, "目录里的 DataModel 与模板同哈希")
    check(zip_entry_compress_type(out, "DataModel") == 0, "DataModel 以 stored(0) 写入",
          str(zip_entry_compress_type(out, "DataModel")))
    info = inspect_pbix(str(out))
    check(info["is_pbix"] is True and page_map(info) == {"Finance": 10, "Operations": 10},
          "产物可被 inspect_pbix 正常解析", str(page_map(info)))


def test_4_validate_real(tmp: Path) -> None:
    section("测试 4  validate_tmdl() 校验真实工程 TMDL")
    print("  pwsh=%s" % (find_powershell(),))
    dll = find_tabular_dll()
    print("  tabular_dll=%s" % dll)
    check(dll is not None, "找到 Microsoft.PowerBI.Tabular.dll")

    t0 = time.time()
    res = validate_tmdl(str(SEMANTIC_DEF))
    dt = time.time() - t0
    counts = {t["name"]: t["measures"] for t in res["tables"]}
    total = sum(t["measures"] for t in res["tables"])
    print("  ok=%s CL=%s culture=%s 耗时=%.2fs error=%s"
          % (res["ok"], res["compatibility_level"], res["culture"], dt, res["error"]))
    print("  tables=%s" % json.dumps(res["tables"], ensure_ascii=False))
    print("  relationships=%s" % json.dumps(
        [{"from": r["from"], "to": r["to"]} for r in res["relationships"]], ensure_ascii=False))
    print("  detail=%s" % json.dumps(res["detail"], ensure_ascii=False))

    check(res["ok"] is True, "ok=True", str(res["error"]))
    check(res["compatibility_level"] == 1567, "compatibility_level == 1567",
          str(res["compatibility_level"]))
    check(res["culture"] == "zh-CN", "culture == zh-CN", str(res["culture"]))
    check(len(res["tables"]) == 3, "3 张表", str(list(counts)))
    check(counts == EXPECTED_MEASURES_BY_TABLE, "每表度量值 日期0/销售明细10/财务月度23", str(counts))
    check(total == EXPECTED_TOTAL_MEASURES, f"度量值合计 {EXPECTED_TOTAL_MEASURES}", str(total))
    check(len(res["relationships"]) == 2, "2 个关系",
          str([(r["from"], r["to"]) for r in res["relationships"]]))
    check(all(t["partitions"] >= 1 for t in res["tables"]), "每表至少 1 个分区",
          str({t["name"]: t["partitions"] for t in res["tables"]}))

    # validate_pbip 顺带验证
    pb = validate_pbip(str(PBIP))
    print("  validate_pbip: ok=%s definition_dir=%s" % (pb["ok"], pb.get("definition_dir")))
    check(pb["ok"] is True, "validate_pbip(.pbip) ok=True")
    check(pb.get("definition_dir") == str(SEMANTIC_DEF), "validate_pbip 解析到正确 definition 目录",
          str(pb.get("definition_dir")))
    check(sum(t["measures"] for t in pb["tables"]) == EXPECTED_TOTAL_MEASURES,
          "validate_pbip 度量值合计 33")


def test_5_negative(tmp: Path) -> None:
    section("测试 5  负向：插入非法属性 description，校验器必须报错")
    dst = tmp / "negative_definition"
    if dst.exists():
        shutil.rmtree(dst)
    shutil.copytree(SEMANTIC_DEF, dst)

    target = dst / "tables" / "销售明细.tmdl"
    lines = target.read_text(encoding="utf-8").splitlines(keepends=True)
    idx = next(i for i, ln in enumerate(lines) if ln.startswith("\tmeasure "))
    # 在该 measure 的第二行（formatString）之后插一行非法属性
    lines.insert(idx + 2, "\t\tdescription: 测试\n")
    target.write_text("".join(lines), encoding="utf-8")

    snippet = "".join(lines[idx : idx + 5])
    print("  插入位置(1-based) = %d，插入后该处为：" % (idx + 3))
    for ln in snippet.splitlines():
        print("    " + repr(ln))
    check(
        any(ln.startswith("\t\tdescription:") for ln in lines),
        "插入行确实以两个真实 TAB 开头",
        repr([ln for ln in lines if "description" in ln]),
    )

    res = validate_tmdl(str(dst))
    print("  ok=%s" % res["ok"])
    print("  error=%s" % json.dumps(res["error"], ensure_ascii=False))
    hit = bool(res["error"]) and (
        "description" in res["error"] or "UnknownKeyword" in res["error"]
    )
    check(res["ok"] is False, "ok=False（校验器不是走过场）")
    check(hit, "error 含 'description' 或 'UnknownKeyword'", str(res["error"])[:160])
    check(res["tables"] == [], "出错时不再返回假成功的数据", str(res["tables"]))

    # 对照组：同一份拷贝不插非法行时必须 ok=True，排除"复制本身弄坏了"
    clean = tmp / "negative_control"
    if clean.exists():
        shutil.rmtree(clean)
    shutil.copytree(SEMANTIC_DEF, clean)
    ctl = validate_tmdl(str(clean))
    check(ctl["ok"] is True and sum(t["measures"] for t in ctl["tables"]) == 33,
          "对照组：未改动的拷贝仍 ok=True / 33 度量值", str(ctl["error"]))


# --------------------------------------------------------------------------- #
def main(argv: list[str]) -> int:
    keep = "--keep" in argv
    if not PBIX.is_file():
        # 没有真实样本时优雅跳过（退出码 0），便于在任意机器 / CI 上跑
        print("SKIP: 缺少真实样本 pbix。", file=sys.stderr)
        print("      设环境变量 HWKIT_TEST_ASSETS 指向含 '新新公司经营分析看板.pbix' "
              "的目录即可跑完整用例。", file=sys.stderr)
        return 0

    tmp = Path(tempfile.mkdtemp(prefix="hwkit_pbipack_test_"))
    print(f"临时目录: {tmp}")
    print(f"python: {sys.version.split()[0]}  pbipack: {ROOT / 'src' / 'hwkit_mcp' / 'pbipack.py'}")
    t0 = time.time()
    try:
        test_1_inspect(tmp)
        extract_dir = test_2_extract(tmp)
        test_3_build(tmp, extract_dir)
        test_3b_pack_dir(tmp, extract_dir)
        test_3c_extra_resources(tmp, extract_dir)
        test_3d_pack_dir_local_datamodel(tmp, extract_dir)
        test_4_validate_real(tmp)
        test_5_negative(tmp)
    finally:
        dt = time.time() - t0
        failed = [r for r in _RESULTS if not r[0]]
        print("\n" + "=" * 72)
        print(f"合计 {len(_RESULTS)} 项断言，通过 {len(_RESULTS) - len(failed)}，失败 {len(failed)}，"
              f"耗时 {dt:.1f}s")
        if failed:
            for _, label, info in failed:
                print(f"  FAILED: {label}  {info}")
        print("=" * 72)
        if keep:
            print(f"已保留临时目录: {tmp}")
        else:
            shutil.rmtree(tmp, ignore_errors=True)
            print("临时目录已清理")
    return 0 if not [r for r in _RESULTS if not r[0]] else 1


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))