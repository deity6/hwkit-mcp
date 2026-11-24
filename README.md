# hwkit-mcp

一个给 AI agent 用的 **作业 / Power BI 辅助工具箱**，走 MCP（stdio）。

两条设计立场：

1. **补缺口。** 纯命令行环境下，`.pbix` 打包、模型权威校验、报表离线渲染这三件事没有顺手的工具。本服务器用 zip 手术 + 微软自己的 TMDL 解析器 + DashCraft 渲染把它们补齐。
2. **不打扰。** 所有工具都不抢焦点、不移动鼠标、不发键盘事件。截图走 `PrintWindow(PW_RENDERFULLCONTENT)`，能抓到被遮挡甚至位于屏幕外的窗口内容，而完全不改变窗口的前后层级。渲染走无头浏览器，不弹窗。

一句话：**让 agent 看得见、改得动，但用户该干嘛干嘛。**

---

## 工具一览

### Power BI · 文件层

| 工具 | 作用 |
|---|---|
| `hw_pbix_inspect` | 看 `.pbix` 内部结构：条目、页面、视觉对象数、内嵌模型字节数、主题文件 |
| `hw_pbix_extract` | 解包成目录（`Report/`、`DataModel`、`[Content_Types].xml` …） |
| **`hw_pbix_build`** | **无界面地把 PBIR 报表目录打包成 `.pbix`** —— 拿一个已有 pbix 当外壳，只替换 `Report/` 前缀 |
| `hw_pbix_pack` | 把解包目录重新打回 `.pbix`，可换用另一个文件的 `DataModel` |

`hw_pbix_build` 的关键细节：非 `Report/` 的条目**连 `ZipInfo` 一起复用**，保证 `DataModel` 那条 161 KB 的 VertiPaq ABF 仍是 `stored(0)`，不会被目标 ZipFile 的默认 deflate 重新压缩。这是"重新打包后 Power BI 还能正常打开"的前提。

### Power BI · 模型校验

| 工具 | 作用 |
|---|---|
| `hw_model_validate` | 校验一个 `.pbip` 的语义模型 |
| `hw_model_validate_tmdl` | 直接校验 TMDL 的 `definition` 目录 |

**为什么非要用这个而不是别的开源解析器**：它们普遍比微软宽松。实测 DashCraft 会放过 measure 下的 `description:` 属性，而微软的 `TmdlSerializer` 直接判 `Unsupported property - description is not a supported property in the current context` —— 于是出现"本地校验通过、Power BI Desktop 打不开"。本工具调的就是 Desktop 自带的那份 `Microsoft.PowerBI.Tabular.dll`，判定口径与 Desktop 一致。

### Power BI · 渲染自查

| 工具 | 作用 |
|---|---|
| `hw_render_pages` | 把报表**离线**渲染成每页一张 PNG，并返回该页 SVG 里抽出的**全部可见文字** |

链路：DashCraft `render --svg` → 统一系统字体栈并去掉错误裁切框 → 无头 Edge/Chrome 光栅化。

绕过的两个渲染器坑（都写在 `render.py` 注释里）：

- DashCraft 自带字体是编译进二进制的 Selawik，没有中文字形，直接出 PNG 时中文全是缺失告警；
- 它给 `cardVisual` 的数字和标签套了一个约 25.7 px 高的 `clipPath`，而数字字号是 60 px，必然被腰斩。所以后处理时去掉所有子裁切框（保留页面级的 `c0`）。

返回的 `texts` 是核对"图上到底写了什么"最可靠的依据，不依赖图像识别。

### Windows · 非侵入式截图

| 工具 | 作用 |
|---|---|
| `hw_list_windows` | 枚举顶层窗口：hwnd / pid / 标题 / 窗口类 / 位置尺寸 / 最小化 / 前台 / 进程名 |
| `hw_capture_window` | **按标题或 pid 抓窗口画面**，不抢焦点、不移动鼠标 |
| `hw_capture_screen` | 抓全屏或指定区域 |

非侵入性是可验证的，不是口号：

- 截图前后前台窗口句柄不变、鼠标坐标不变、所有窗口的 `foreground` 标记零变化；
- 能抓到 **rect 为 `(-32100, -32100)`、完全在屏幕之外**的窗口的完整界面（`BitBlt` 从屏幕取像素在这种情形下只能拿到空白）；
- 与屏幕同区域逐像素比对，对"不透明、完整在屏、无遮挡"的窗口差异为 0.00%。

半透明窗口（如任务栏的亚克力表面）不要用逐像素比对做校验——`PrintWindow` 渲染的是窗口自身，屏幕上是它与壁纸的混色结果，两者天然不同。

### Office 文档转换

| 工具 | 作用 |
|---|---|
| `hw_com_probe` | 探测 Word / Excel COM 是否可用，给出可执行建议 |
| `hw_word_convert` | `.doc/.docx` → `pdf / docx / odt / rtf / txt` |
| `hw_excel_convert` | 工作簿 → `pdf / csv / xlsx / ods` |

转换时 Word/Excel 以不可见方式运行，不抢焦点；目标文件已存在时默认**拒绝覆盖**（作业交付物被静默覆盖是很难查的事故）。COM 不可用时自动改走 LibreOffice（设 `HWKIT_SOFFICE` 指向 `soffice.exe`）。

> 已知：部分环境下 Office 的 COM 激活会被挡住，报 `CO_E_SERVER_EXEC_FAILURE (0x80080005)` —— Word 进程本身能起（`WINWORD.EXE /q` 正常），但 DCOM 激活失败，`CoCreateInstance` 和 `GetActiveObject` 都不可用。这是环境问题而非代码问题，此时 `hw_com_probe` 会明确报出来并建议改用 LibreOffice 或其它转换通道。

### 配套：pbix-mcp 桥接

| 工具 | 作用 |
|---|---|
| `hw_pbix_tools` | 列出配套 `pbix-mcp` 提供的全部工具（**134 个**） |
| `hw_pbix_call` | 按需转发任意一次调用 |

[`pbix-mcp`](https://github.com/d0nk3yhm/pbix-mcp) 是纯 Python 独立实现的 `.pbix` 读写 + DAX 引擎，**不需要 Power BI Desktop**，能建表 / 加列 / 加度量值 / 改 ABF / 加视觉对象 / DAX 求值 / 导出 TMDL 与 PBIP。

它有 134 个工具，但 MCP 客户端没有工具白名单——直接接入会把 134 份 JSON Schema 常驻进每一次请求的上下文。所以这里用**长驻子进程 + 2 个元工具**：`hw_pbix_tools` 查清单，`hw_pbix_call` 按需转发。实测冷启 1.2 s，热调用 0.01 s（会话常驻）。

> 若你的宿主更希望直接挂载，也可以把 `pbix-mcp` 作为**独立 MCP 服务器**并列接入，用 `pbix_launch.py` 垫片：
> ```python
> from pbix_mcp.server import mcp
> mcp.run("stdio")
> ```

### 自检

| 工具 | 作用 |
|---|---|
| `hw_selfcheck` | 报告各模块与关键外部依赖（DashCraft / 无头浏览器 / TMDL 校验器 / Office COM）状态 |

换机器部署后先跑这个，一眼看出缺哪一环。

---

## 安装

```bash
git clone <this repo>
cd hwkit-mcp
pip install -e .
```

依赖只有一条：MCP 官方 Python SDK（`mcp`）。其余全部是标准库 + ctypes，`ctypes` 直接调 Win32 / GDI，PNG 编码器是手写的（`zlib` + `struct`，支持全部 5 种 scanline filter 的解码也在测试里）。

### 外部依赖（按需）

| 依赖 | 用于 | 缺失时的影响 |
|---|---|---|
| **Power BI Desktop**（只要那份 DLL，不需要启动它） | `hw_model_validate*` | 模型校验不可用，其它工具照常 |
| **DashCraft** ≥ 0.3.0 | `hw_render_pages` | 渲染不可用 |
| **Edge / Chrome** | `hw_render_pages` 的光栅化 | 渲染不可用 |
| **Office 或 LibreOffice** | `hw_word_convert` / `hw_excel_convert` | 文档转换不可用 |

路径可用环境变量覆盖：`HWKIT_DASHCRAFT`、`HWKIT_BROWSER`、`HWKIT_SOFFICE`、`HWKIT_PBIX_LAUNCHER`。

---

## 接入

### DSH（DeepSeek Harness）

写进 profile 的 `cordis.patch.yml`：

```yaml
- insert:
    - id: mcp-hwkit
      name: '@deepseek-ai/dsh-mcp-client'
      config:
        serverName: hwkit
        transport: stdio
        command: 'C:\Users\<you>\AppData\Roaming\Python\Python3xx\Scripts\hwkit-mcp.exe'
        args: []
        env:
          PYTHONUTF8: '1'
          PYTHONIOENCODING: 'utf-8'
        toolCallTimeoutMs: 600000
```

> `PYTHONUTF8=1` 在 Windows 上是必须的：子进程 stdio 默认按系统代码页（GBK）编码，中文会乱码。

改完 patch 文件即热重组，工具以 `mcp__hwkit__*` 出现。

### 通用 MCP 客户端

```json
{
  "mcpServers": {
    "hwkit": {
      "command": "hwkit-mcp",
      "args": [],
      "env": { "PYTHONUTF8": "1", "PYTHONIOENCODING": "utf-8" }
    }
  }
}
```

协议版本 `2024-11-05`，stdio，16 个工具。

---

## 典型用法

**迭代看板（模型不变，只改报表）—— 全程不开 Power BI Desktop：**

```
1. 改 .Report/definition 下的 PBIR JSON（页面 / 视觉对象 / 主题）
2. hw_model_validate(".pbip")          确认模型没坏
3. hw_pbix_build(".Report", "out.pbix", template_pbix="旧版本.pbix")
4. hw_render_pages("out.pbix", "preview/")   出图自查，顺便核对 texts
```

**改模型（加度量值 / 加表）：**

```
hw_pbix_call("pbix_open",        {"file_path": "x.pbix", "alias": "a"})
hw_pbix_call("pbix_list_tables", {"alias": "a"})
hw_pbix_call("pbix_datamodel_add_measure",
             {"alias": "a", "table_name": "财务月度",
              "measure_name": "毛利率", "expression": "DIVIDE([毛利],[收入])"})
hw_pbix_call("pbix_evaluate_dax", {"alias": "a", "measures": "毛利率"})
```

> 注意 `measures` 要传**字符串**（逗号分隔），传列表会被 pydantic 拒掉。

**截图（用户在打游戏也不会被打断）：**

```
hw_list_windows()
hw_capture_window(title_contains="Power BI")
```

---

## 测试

```bash
python tests/test_win.py        # 窗口枚举 + 非侵入式截图，41 项断言
python tests/test_pbipack.py    # PBIX 打包 + TMDL 校验，69 项断言
```

两项测试都包含**负向用例**：例如往 TMDL 里插一行 `\t\tdescription: 测试`，校验器必须返回 `ok=False` 且错误里带 `description` —— 这条用来证明校验器不是走过场。

`test_pbipack` 还验证了重新打包后 `DataModel` 条目的 **sha256 与原文件逐字节相同**、`compress_type` 仍为 `0`、以及 `.NET System.IO.Packaging` 把新旧两个包解析成完全相同的 part 列表。

---

## 目录结构

```
hwkit-mcp/
├── src/hwkit_mcp/
│   ├── server.py       # FastMCP 入口，16 个工具的门面
│   ├── win.py          # Win32 窗口枚举 + PrintWindow 截图 + 手写 PNG 编码
│   ├── pbipack.py      # .pbix 的 zip 手术（extract / build / pack / inspect）
│   ├── modelcheck.py   # 微软 TmdlSerializer 校验（PowerShell 桥，零第三方依赖）
│   ├── render.py       # DashCraft SVG → 系统字体栈修正 → 无头浏览器 PNG
│   ├── office.py       # Word/Excel COM 转换 + LibreOffice 兜底
│   └── pbixbridge.py   # 长驻子进程桥接 pbix-mcp 的 134 个工具
└── tests/
```

---

## 为什么不用 GUI 自动化

坐标点击 / SendKeys 这类操作有三个问题：抢焦点会打断用户；分辨率与 DPI 缩放一变坐标就漂；无头环境下干脆没有可点的东西。所以本项目的原则是**能用 API 就不用界面**：

| 目标 | 用 API 的做法 |
|---|---|
| 看界面长什么样 | `PrintWindow` 抓窗口位图（不改变层级） |
| 看报表长什么样 | 离线渲染成 PNG |
| 生成 `.pbix` | zip 手术换 `Report/`，复用原 `DataModel` |
| 校验模型 | 调 Desktop 同款的 `TmdlSerializer` |
| 算 DAX | `pbix-mcp` 的独立引擎 |

---

## 已知限制

- **打包出的 `.pbix` 未在 Power BI Desktop 里自动验证**（那需要开界面）。现有证据是：`zipfile.testzip()` 无坏条目、`inspect_pbix` 能完整回读、`.NET System.IO.Packaging` 解析出的 part 列表与原始包一致。最终仍需人工双击一次确认。
- `hw_pbix_build` 复用模板的 `DataModel`，因此**模型结构变化时不能直接套用**（报表变了可以）。改模型请走 `pbix-mcp`。
- 老式（非 PBIR）`.pbix` 布局（`Report/Layout`，UTF-16LE JSON）的解析路径已实现但无样本可测，属未验证代码。
- `hw_render_pages` 的字体与裁切修正是**启发式**的，目的是让预览可读，不代表与 Power BI Desktop 的像素级一致。它能可靠回答"布局对不对、写了什么字、数对不对"，不适合做视觉回归基线。
- 仅支持 Windows（窗口截图与 COM 转换是 Win32 专属）。`pbipack` / `modelcheck` / `render` 三块逻辑本身跨平台。

## 许可

MIT。见 [LICENSE](LICENSE)。