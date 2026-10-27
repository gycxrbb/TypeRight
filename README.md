# TypeRight — Word 智能排版插件

> **指认一个样本 → 一键选中全文同类段落 → 统一改格式**；公式一键转公式对象；操作可存为方案复用。
> PRD v2.0（见根目录 PRD 文档）。设计原则：所见即所得、最小界面（Ribbon 3 按钮 + 右键菜单）、确定性签名匹配、每步独立撤销、规范兜底（西文自动 TNR / 公式 Cambria Math）。

## 快速开始

```powershell
# 安装（管理员：HKLM；非管理员：HKCU）
powershell -ExecutionPolicy Bypass -File tools\install.ps1
# 卸载
powershell -ExecutionPolicy Bypass -File tools\uninstall.ps1
```

安装后启动 Word → 功能区「**TypeRight**」选项卡：

1. **同类选择**（或右键菜单同名项）：光标点进任一段落 → 点按钮 → 同款格式的段落全部高亮命中 → 段落旁浮动条上选中文字体即批量应用；命中不够就在浮动条上勾「忽略字号/颜色/加粗」实时重算
2. **公式转换**：全文纯文本公式（`ax²+bx+c=0`）批量转 Word 公式对象，单行提示「已转换 N 处，失败 M 处」，失败原文保留并列出定位
3. **方案**：≥2 步批量操作后可「存为方案」；下一份同类文档打开 → 方案面板 → 应用 → 规则逐条回放 + 兜底规则 + 公式转换，一个撤销单元

诊断日志：`%LOCALAPPDATA%\TypeRight\addin-log.txt`（含同类匹配耗时实测）；方案目录：`%APPDATA%\TypeRight\profiles\`。

## 核心机制（PRD v2 §2）

- **格式签名**：段落首个非空 run 的有效格式元组 `(中文字体, 西文字体, 字号, 加粗, 斜体, 颜色, 对齐)`。只在首个可见字符的单字符区间上读取——段落级 Font 读取会触发 Word 全 run 聚合（长段落 ~15ms/段）
- **同类段落**：签名全等（严格）；浮动条勾选忽略维度后实时重算（宽松）
- **撤销粒度**：每个取值应用 = 1 个 `UndoRecord`；公式转换、应用方案各自独立
- **中西文联动（P0-3）**：选中文字体时 `NameAscii/NameOther` 自动写 Times New Roman（设置页可关/可改）
- **方案 Schema v2**：`rules[{order, match(签名+ignore), apply(null=保持源值)}]` + 末位 `catchAll` 兜底 + `convertFormulas`；字段级校验错误定位

## 架构

```
src/
├── TypeRight.Core/
│   ├── Signature/          # FormatSignature（提取/比较）SimilarSelector（匹配+高亮会话）BatchFormatter（批量应用+UndoRecord）
│   ├── Formatting/         # FormulaConverter（§6 公式引擎，规格不变）SampleDocBuilder WordFormatUtils
│   ├── Profiles/           # ProfileV2（Schema v2 + 校验 + 存储/导入导出）
│   ├── Scanning/ Classification/  # v1 语义引擎（R1-R12）——PRD §13.1 封存为 P2 资产
│   └── Settings.cs         # 三项设置 + 方案目录
├── TypeRight.Addin/        # Ribbon 3 按钮 + ContextMenuText + FloatBar（浮动格式条）+ Toast + 方案/设置窗口
└── TypeRight.SmokeTest/    # 真实 Word 黄金旅程（同类选择命中数/撤销粒度/公式计数/方案回放/幂等）
tests/                      # 40 个纯逻辑单测（签名比较/宽松维度/方案校验/JSON 往返）
tools/install.ps1           # 部署 %LOCALAPPDATA%\TypeRight\Addin + COM 注册
```

## 测试

```powershell
dotnet test TypeRight.sln                      # 40 单测
dotnet run --project src\TypeRight.SmokeTest   # 真实 Word 端到端（黄金旅程）
```

**性能**：冒烟测试经进程外 DCOM 运行（每 COM 调用 ~1ms 编组开销，~14ms/段）；插件实际运行在 Word 进程内（直调）。进程内实测看 `%LOCALAPPDATA%\TypeRight\addin-log.txt` 的「同类匹配: N 段 / Xms」日志行——PRD 验收口径以该日志为准。

## 已知边界（P1/P2）

- 匹配高亮用临时黄色高亮呈现（原本有高亮的段落不参与标记以免无法精确还原）；高亮组包为独立撤销记录「TypeRight 匹配标记」
- v1 语义识别引擎 / LLM / 预览面板：按 PRD §11 封存（代码保留在 Core/Scanning、Core/Classification）
- CTP 停靠任务窗格：v2 无任务窗格需求（PRD 明确无），相关代码已随 v1 移除
- MML2OMML.XSL 兜底管线已封装（`FormulaConverter.MathmlToOmml`），结构化 MathML 生成器属 P2
