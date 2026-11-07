# TypeRight — Word 指认式批量排版插件

> **光标点一下，全文同类段落一次改齐。**
> 同类选择（按格式签名全文匹配）· 纯文本公式一键转公式对象 · 操作存为方案下份文档一键重放

一个 COM 加载项，给讲义 / 试卷 / 笔记这类"大量段落格式一致"的 Word 文档用：不用宏、不逐段刷格式、所见即所得、每步可撤销。

**性能实测**（254 处公式的真实讲义，进程内）：公式转换 47s → **7.8s**；千段全文匹配 ~0.5s。

---

## ① 我只想装来用

### 环境要求

| | |
|---|---|
| ✔ 支持 | Windows 7 SP1+ / Win10 1903+ / Win11 + **Microsoft Word 桌面版** 2010 / 2013 / 2016 / 2019 / 2021 / 365（32/64 位均可） |
| ✔ 也支持 | **WPS Office**（个人版/教育版；安装时选"启用 WPS 支持"，实测同类选择/批量格式/浮动条/方案可用，公式转换实验性） |
| ✘ 不支持 | Word 网页版、Mac、Word 2007 及更早（平台限制） |
| 依赖 | .NET Framework 4.8（Win10 1903+ 一般自带；Win7/8 需手动安装，安装器会给出下载链接） |

### 安装（免管理员，3 步）

1. 到本仓库 **Releases** 页下载最新 `TypeRight-vX.Y.Z.zip`（或从群里拿）
2. 关闭所有 Word/WPS 窗口，解压后双击 **install.bat**（中途会问是否启用 WPS 支持，按需选 y）
3. 启动 Word → 功能区出现「TypeRight」选项卡即成功（首次启动有三步用法提示）

### 三步上手

1. 光标点进一个段落（或选中一段文字）
2. 功能区「TypeRight」→ **同类选择**（右键菜单也有）
3. 浮动条上改字体 / 字号（A⁻ A⁺）/ 加粗 / 斜体 / 下划线 / 删除线 / 颜色——**全部同类段落同时生效**
   - `Esc` 或点击别处关闭浮动条；「**下一处 ⤵**」把命中处设为真选区，可直接用 Word 原生「开始」选项卡 / `Ctrl+D` 调该处
   - 「匹配」行可切宽松模式（忽略字号 / 颜色 / 加粗后实时重算）
   - **公式转换**：`ax²+bx+c=0` 这类纯文本全文批量转可编辑公式对象，失败片段原样保留
   - **方案**：改过 2 步以上可整套存为方案，下一份同类文档一键回放

### 常见问题

- **Word 2010 / 2013 上能用吗？** 能，功能全部可用；仅 Word 2010 的撤销粒度降级——一次批量修改需多按几次 `Ctrl+Z` 才能完全回退（2010 的对象模型没有自定义撤销记录 API，首次启动会提示一次）。
- **功能区没有 TypeRight 选项卡？** 关掉所有 Word，双击包里的 `repair.bat`，再开 Word（系统清理工具会抹加载项注册、Word 崩溃会自动禁用加载项，repair 一键恢复）。
- **360 / 火绒弹窗？** 选"信任 / 允许"（未做商业代码签名，首次询问属正常；插件完全离线，不联网不收集数据）。
- **想反馈问题？** 选项卡「使用说明」→「**复制诊断信息**」→ 粘贴发给作者（自动带版本 + 最近日志）。
- **卸载**：双击 `uninstall.bat`。

---

## ② 我是开发者（源码二次开发 / 学习）

### 技术栈与架构

- **.NET Framework 4.8 + NetOffice**（COM Add-in，免 VSTO 运行时）+ WPF（浮动条/窗口）+ C#
- `src/TypeRight.Core`——无 UI 的核心库：
  - `Signature/` **格式签名匹配**：5 锚点模型（首字符 / 25%/50%/75% 中段 / 尾字符 × 字体·字号·加粗·颜色），未知值不与任何值相等（宁少勿错）；引导句感知单元匹配（★ 导读框这类"行首标题句 + 正文"同段混合体拆成两个匹配单元）；全程公式区感知（锚点跳过 OMML 区，公式化后的文档匹配不退化）
  - `Signature/BatchFormatter.cs` **批量格式执行**：幂等写入（值没变不写）、字号护栏、每批一个 UndoRecord、区间自动避开公式区
  - `Formatting/FormulaConverter.cs` + `Formulas/` **公式转换三阶段流水线**：全文规划（纯位置切分）→ 全部公式合并一个 docx、隐藏草稿一次装载、降序批量 `FormattedText` 注入 → 文档级 O(1) 终检、失败逐项回退（MathML → Office 自带 MML2OMML.XSL → OMML）
  - `Profiles/` 方案存回放；`Classification/`、`Roles/` 语义分析资产（P2 封存）
- `src/TypeRight.Addin`——Ribbon / 右键菜单 / 浮动格式条 / 主题跟随宿主 Word（屏幕实测检测）/ 会话管理
- 测试资产同步：`Word智能排版插件_PRD.md`（产品规格）、`tests/TypeRight_测试验收规格.md`（21 个验收用例的口径）

### 构建与开发安装

```powershell
# 单元测试（45 个）
dotnet test tests\TypeRight.Core.Tests\TypeRight.Core.Tests.csproj

# 开发安装（构建 + 发布到 %LOCALAPPDATA%\TypeRight\Addin + HKCU 注册）
powershell -ExecutionPolicy Bypass -File tools\install.ps1

# 全量验收（21 用例，真实讲义 + 黄金文件；DCOM 环境约 30-40 分钟）
tests\TypeRight.Acceptance\bin\Debug\TypeRight.Acceptance.exe

# 单独取证：p9 = 公式化后匹配保真 + 字号零漏网（Program.cs 里有 p1-p9 全套）
tests\TypeRight.Acceptance\bin\Debug\TypeRight.Acceptance.exe p9
```

### 发布流程（自动）

推一个 `v*` tag 即自动构建并发布：

```powershell
git tag v0.2.1 && git push origin v0.2.1   # → GitHub Actions 构建 → Release 挂载安装包 zip
```

流水线：`.github/workflows/release.yml`（windows runner 上跑 `tools/package-release.ps1` → 产物 `TypeRight-v*.zip` 附到 Release）。手动出包也行：`powershell -File tools\package-release.ps1`。

### 这个项目里值得看的东西（Word COM 避坑实录）

踩过的坑都写在代码注释与提交历史里，挑几个狠的：

- `Range.Information(wdWithInTable)` 在进程内部分场景返回**假值**（匹配归零）→ 改按 `doc.Tables` 区间位置判定
- `Font.Bold/Size` 存在间歇性垃圾值、`Font.Name` 对部分区间读取抖动 → 重读稳定化 + 中段锚点不比字体
- 跨段落/单元格结束标记设置高亮会**扩散到整段** → 区间排除结束标记
- 每个公式 `Documents.Open` 一次临时文档 = 窗口频闪 + ~150ms/个 → 隐藏草稿文档一次装载复用
- `Paragraphs[i]` 索引器 O(n) 重扫（千段 O(n²)≈9s）→ 一律枚举器
- COM 自动化语境下跨线程调用不可靠 → `SetTimer` 挂 Word 窗口、`WM_TIMER` 回调落在 Word UI 线程

### 已知限制（设计决定）

- 中段锚点不比对字体名（Word 读取抖动，为不误判的取舍）；段落中间个别字符的字体差异不参与同类判定
- 方案回放按整段匹配（引导句感知目前只在"同类选择"生效）
- 公式转换中括号不配对等不可解析片段原样保留为文本

---

## 目录导览

```
src/TypeRight.Core/          核心库（签名匹配 / 批量格式 / 公式管线 / 方案）
src/TypeRight.Addin/         COM 加载项（Ribbon / 浮条 / 会话 / 主题）
tests/TypeRight.Core.Tests/  45 个单元测试
tests/TypeRight.Acceptance/  21 个验收用例 + DiagP1-P9 取证工具
tools/                       install / uninstall / package-release / dist-scripts
Word智能排版插件_PRD.md      产品规格（v2.0）
```

仅供学习交流。二次开发请注明出处。
