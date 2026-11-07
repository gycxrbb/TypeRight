# Product

<!-- impeccable:product-schema 1 -->

## Platform

windows-desktop（WPF，宿主为 Microsoft Word / WPS Office 的 COM 加载项；非 Web）

## Users

教师/家教（作者本人即首批用户）与需要批量整理讲义、试卷、笔记的学生群体；在 Word/WPS 中编辑大量"同格式段落"文档的真实场景。

## Product Purpose

TypeRight＝指认式批量排版：光标点进样本段 → 一键"选中"全文同格式段落 → 浮动条上批量改字体/字号/样式 → 操作可存为方案回放。成功＝十几步的逐段刷格式变成两三步。

## Positioning

"格式签名匹配"是核心差异：按首/中/尾锚点签名找同类段落，引导句感知、公式区感知；邻接产品（Word 查找替换/样式刷）无此机制。

## Operating Context

- 宿主：Word 2010~365（32/64 位）、WPS Office；Win7 SP1~Win11
- 交互入口：功能区 TypeRight 选项卡 + 右键菜单；命中后弹出**浮动格式条（FloatBar）**
- 浮条跟随宿主深浅主题（屏幕实测检测）；Esc/点击别处关闭；会话高亮为临时态

## Capabilities and Constraints

同类选择（严格/宽松）· 批量格式（中西文联动）· 跳转导航（下一处 n/N + 强调色）· 公式转换（OMML 注入管线）· 方案存回放 · 诊断导出
约束：COM 加载项免管理员安装（WPS 支持除外）；不动 Word 原生 UI 只做增量。

## Brand Commitments

功能一切照旧的前提下，UI 对齐宿主原生观感（用户明确要求：浮动条不能与 Word/WPS 割裂；参考 Word 自带 Mini Toolbar）。中文字体必须用复合字体避免宋体回退。

## Evidence on Hand

- 仓库：`Word智能排版插件_PRD.md`（产品规格 v2）、`README.md`、`tests/TypeRight_测试验收规格.md`
- 真实资产：初三数学讲义测试文档（tests/）；21 用例验收套件全绿基线
- 截图取证链：`tests/output/ui-float-*.png`（三主题像素断言，FloatBar 旧观感的真实记录）

## Product Principles

1. 指认式：用户指一个样本，插件找同类——不猜、不全自动
2. 宿主原生感：插件是 Word/WPS 的延伸，不是贴上去的第三方窗
3. 每步可撤销、临时态绝不污染文档
4. 性能即体验：千段匹配亚秒级

## Accessibility & Inclusion

教师文档中文为主；浮条控件需在深浅两主题下均可辨（现有主题跟随机制即为此服务）。
