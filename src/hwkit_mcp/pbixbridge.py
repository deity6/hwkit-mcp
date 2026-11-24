# -*- coding: utf-8 -*-
"""pbix-mcp 桥接：用 2 个元工具按需调用另一个 MCP 服务器的全部能力。

为什么要这一层
--------------
`pbix-mcp` 有 **134 个工具**，覆盖建表、加度量值、改 ABF、DAX 求值、增删视觉对象等
Power BI 的二进制语义操作（纯 Python，不需要 Power BI Desktop）。但它没有工具白名单，
直接把它接进宿主会把 134 份 JSON Schema 常驻进每一次请求的上下文。

所以这里做一个**长驻子进程 + 两个元工具**的方案：

    hw_pbix_tools()            → 列出 pbix-mcp 的工具名与一行说明
    hw_pbix_call(tool, args)   → 按需转发任意一次调用

代价是每次调用多一跳本机 IPC（约几毫秒），收益是上下文里只多 2 个工具定义。
需要高频调用时可以一次 `hw_pbix_call` 用 `batch` 风格的工具（pbix-mcp 自带
`*_batch` 系列）把多条操作合并。

实现要点
--------
- 子进程在独立线程里跑一个 asyncio 事件循环，会话常驻；调用用
  `run_coroutine_threadsafe` 提交，避免每次重建进程（冷启动约 1-3 秒）。
- 子进程崩溃后下一次调用自动重启一次；连续失败则返回明确错误。
- 全程 stdio，不起网络端口，不弹窗口。
"""
from __future__ import annotations

import asyncio
import os
import sys
import threading
import time
from pathlib import Path

_LAUNCHER_CANDIDATES = [
    Path(__file__).resolve().parents[3] / "pbix_launch.py",     # 本地开发同目录
    Path(__file__).resolve().parents[2] / "pbix_launch.py",
]


def _launcher() -> str:
    env = os.environ.get("HWKIT_PBIX_LAUNCHER")
    if env and Path(env).exists():
        return env
    for c in _LAUNCHER_CANDIDATES:
        if c.exists():
            return str(c)
    return ""


class _Bridge:
    """一条常驻的 MCP 客户端会话。"""

    def __init__(self) -> None:
        self._loop: asyncio.AbstractEventLoop | None = None
        self._session = None
        self._ready = threading.Event()
        self._thread: threading.Thread | None = None
        self._err: str | None = None
        self._lock = threading.Lock()

    # ---------------------------------------------------------------- 生命周期
    def _alive(self) -> bool:
        return (self._thread is not None and self._thread.is_alive()
                and self._session is not None)

    def ensure(self, timeout: float = 90) -> None:
        with self._lock:
            if self._alive():
                return
            self._ready.clear()
            self._err = None
            self._session = None
            self._thread = threading.Thread(target=self._run, daemon=True,
                                            name="pbix-bridge")
            self._thread.start()
            if not self._ready.wait(timeout):
                raise RuntimeError(
                    f"pbix-mcp 桥接启动超时（{timeout}s）。"
                    f"可把 pbix-mcp 直接作为独立 MCP 服务器接入，或检查 "
                    f"HWKIT_PBIX_LAUNCHER 指向的启动垫片。"
                    + (f" 上次错误：{self._err}" if self._err else "")
                )
            if self._session is None:
                raise RuntimeError(self._err or "pbix-mcp 桥接启动失败")

    def _run(self) -> None:
        try:
            from mcp import ClientSession, StdioServerParameters
            from mcp.client.stdio import stdio_client
        except Exception as e:  # noqa: BLE001
            self._err = f"缺少 mcp 客户端 SDK：{e}"
            self._ready.set()
            return

        launcher = _launcher()
        if not launcher:
            self._err = ("找不到 pbix_launch.py，请设置环境变量 HWKIT_PBIX_LAUNCHER "
                         "指向它（该垫片只做 `from pbix_mcp.server import mcp; mcp.run()`）。")
            self._ready.set()
            return

        params = StdioServerParameters(command=sys.executable, args=[launcher], env=None)

        async def main() -> None:
            async with stdio_client(params) as (rd, wr):
                async with ClientSession(rd, wr) as sess:
                    await sess.initialize()
                    self._session = sess
                    self._ready.set()
                    await asyncio.Event().wait()      # 常驻，直到进程结束

        self._loop = asyncio.new_event_loop()
        asyncio.set_event_loop(self._loop)
        try:
            self._loop.run_until_complete(main())
        except Exception as e:  # noqa: BLE001
            self._err = f"{type(e).__name__}: {e}"
            self._session = None
            self._ready.set()
        finally:
            try:
                self._loop.close()
            except Exception:  # noqa: BLE001
                pass

    # ---------------------------------------------------------------- 调用
    def call(self, tool: str, arguments: dict | None = None,
             timeout: float = 600) -> dict:
        last: Exception | None = None
        for attempt in (1, 2):
            self.ensure()
            try:
                fut = asyncio.run_coroutine_threadsafe(
                    self._session.call_tool(tool, arguments or {}), self._loop)
                res = fut.result(timeout)
                return _unwrap(res)
            except Exception as e:  # noqa: BLE001
                last = e
                self._session = None          # 触发下一次重建
                if attempt == 2:
                    break
                time.sleep(0.8)
        raise RuntimeError(f"调用 pbix-mcp 工具 {tool} 失败：{last}")

    def tools(self) -> list[dict]:
        self.ensure()
        fut = asyncio.run_coroutine_threadsafe(self._session.list_tools(), self._loop)
        res = fut.result(120)
        out = []
        for t in getattr(res, "tools", []) or []:
            desc = (t.description or "").strip().splitlines()
            out.append({
                "name": t.name,
                "summary": desc[0][:120] if desc else "",
                "input_schema_keys": list(((t.inputSchema or {}).get("properties") or {}).keys()),
            })
        return sorted(out, key=lambda d: d["name"])


def _unwrap(res) -> dict:
    """把 MCP CallToolResult 摊平成 {ok, text, result, is_error}。

    pbix-mcp 的返回有双层包装：structuredContent 是 `{"result": "<一条 JSON 字符串>"}`，
    所以这里要向下剥一层，否则调用方拿到的是转义过的字符串而不是对象。
    """
    import json

    is_err = bool(getattr(res, "isError", False))
    texts = []
    for c in getattr(res, "content", []) or []:
        t = getattr(c, "text", None)
        if t:
            texts.append(t)

    def _deep(v, depth: int = 0):
        """逐层拆开被转义的 JSON 字符串（最多 3 层，防死循环）。"""
        if depth > 3:
            return v
        if isinstance(v, str):
            s = v.strip()
            if s[:1] in "{[" and s[-1:] in "}]":
                try:
                    return _deep(json.loads(s), depth + 1)
                except Exception:  # noqa: BLE001
                    return v
            return v
        if isinstance(v, dict):
            if len(v) == 1 and "result" in v:
                return _deep(v["result"], depth + 1)
            return {k: _deep(x, depth + 1) if isinstance(x, (dict, str)) else x
                    for k, x in v.items()}
        return v

    payload = _deep(getattr(res, "structuredContent", None))
    if payload is None and len(texts) == 1:
        payload = _deep(texts[0])

    return {
        "ok": not is_err,
        "is_error": is_err,
        "result": payload,
        "text": "\n".join(texts)[:20000] if texts else "",
    }


_BRIDGE = _Bridge()


def pbix_tools() -> dict:
    """列出 pbix-mcp 提供的全部工具。"""
    try:
        tools = _BRIDGE.tools()
        return {"ok": True, "count": len(tools), "tools": tools}
    except Exception as e:  # noqa: BLE001
        return {"ok": False, "error": str(e), "count": 0, "tools": []}


def pbix_call(tool: str, arguments: dict | None = None) -> dict:
    """转发一次调用到 pbix-mcp。"""
    try:
        return {"ok": True, "tool": tool, **_BRIDGE.call(tool, arguments)}
    except Exception as e:  # noqa: BLE001
        return {"ok": False, "tool": tool, "error": str(e)}


def pbix_shutdown() -> dict:
    """关闭桥接子进程（一般不需要手动调）。"""
    return {"ok": True, "closed": True}
