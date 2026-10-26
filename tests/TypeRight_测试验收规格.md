# TypeRight 测试验收规格：讲义黄金样本

> 版本 v1.0 ｜ 对应 PRD：TypeRight（VSTO）
> 本文档是验收测试的唯一事实来源。所有断言均已硬编码，可直接翻译为测试代码，禁止测试时临时目测判断。

---

## 1. 目的

验证 TypeRight 对**纯直接格式堆叠的学术讲义**的端到端能力：

> 输入 `tests/` 下的原始讲义 → 套用测试 Profile 一键智能排版 → 产出在字体与公式维度上与黄金文件 `初三数学代数终极讲义_字体统一版.docx` 一致，且内容、结构、强调格式零损失。

黄金文件本身即由「三轮人工确认的批量排版（字体统一 + 公式转 OMML）」产生，视为业务正确性的定义。

---

## 2. 测试资产与目录约定

```
TypeRight/
└── tests/
    ├── 初三数学代数终极讲义_一元二次方程与二次函数大一统.docx   # 输入（原始版，1,016 KB）
    ├── 初三数学代数终极讲义_字体统一版.docx                       # 黄金（期望输出，1,026 KB）
    ├── profiles/
    │   └── golden-lecture.json                                   # 测试专用 Profile（见 §3）
    └── TypeRight.Tests/                                          # 测试工程
```

- 所有断言以 `tests/` 相对路径寻址，测试输出写入 `tests/output/`，禁止改动两份资产原件
- 输入文件存在性为测试前置条件，缺失即 fail 并提示文件名

---

## 3. 测试专用 Profile（golden-lecture.json）

黄金文件只改了**字体族**和**公式形态**，字号/加粗/颜色全部保留源值。因此测试 Profile 使用 null 语义：

```json
{
  "id": "golden-lecture",
  "name": "黄金样本测试",
  "roles": {
    "title":     { "eastAsiaFont": "楷体", "latinFont": "Times New Roman", "sizePt": null, "bold": null, "color": null },
    "heading1":  { "eastAsiaFont": "楷体", "latinFont": "Times New Roman", "sizePt": null, "bold": null, "color": null },
    "heading2":  { "eastAsiaFont": "黑体", "latinFont": "Times New Roman", "sizePt": null, "bold": null, "color": null },
    "body":      { "eastAsiaFont": "宋体", "latinFont": "Times New Roman", "sizePt": null, "bold": null, "color": null },
    "caption":   { "eastAsiaFont": "宋体", "latinFont": "Times New Roman", "sizePt": null, "bold": null, "color": null },
    "listItem":  { "eastAsiaFont": "宋体", "latinFont": "Times New Roman", "sizePt": null, "bold": null, "color": null },
    "tableText": { "eastAsiaFont": "宋体", "latinFont": "Times New Roman", "sizePt": null, "bold": null, "color": null }
  },
  "formula": { "convertPlainText": true, "mathFont": "Cambria Math" },
  "mountWordStyles": false
}
```

**null 语义（必须实现）**：字段为 null = 不触碰该属性的源值。本 Profile 下插件只准改字体族与公式，任何字号/加粗/颜色变动都视为 bug。
`mountWordStyles: false`：黄金文件未挂样式，本测试不要求挂载（挂载能力由独立用例覆盖，见 TC-09）。

---

## 4. 黄金文件事实表（断言的硬编码依据）

### 4.1 总量指标

| 指标 | 输入文件 | 黄金文件 |
|---|---|---|
| 非空段落数 | 95 | 95 |
| 表格数 | 11 | 11 |
| 图片数（word/media/*） | 6 | 6 |
| `m:oMath` 公式对象数 | 0 | **255** |
| 其中 `m:f`（分式） | 0 | 25 |
| 其中 `m:rad`（根式） | 0 | 7 |
| 上下标结构 `m:sSub`+`m:sSup`+`m:sSubSup` | 0 | 201 |
| 普通文本 run（w:t）中 Unicode 上下标/根号字符（`⁰¹²³⁴⁵⁶⁷⁸⁹ⁿ₀₁₂₃₄√`）总数 | 194 | **0** |

### 4.2 角色判定锚点（按段落文本前缀搜索定位）

| 锚点：段落文本前缀 | 期望角色 | 期望中文字体（NameFarEast） | 字号/颜色（应保持源值） |
|---|---|---|---|
| `初三数学代数终极大统领：` | title | 楷体 | 20pt / 加粗 / #1A365D |
| `模块一：一元二次方程的本质` ~ `模块七：九年级上册代数综合`（共 7 段，前缀 `模块`） | heading1 | 楷体 | 15pt / 加粗 / #1A365D |
| `1.1 一元二次方程的标准定义` 等全部 `^\d\.\d\s` 段（共 16 段：1.1 1.2 2.1 2.2 3.1 4.1 4.2 4.3 5.1 5.2 5.3 6.1 6.2 6.3 6.4 6.5） | heading2 | 黑体 | 12.5pt / 加粗 / #2B6CB0 |
| `【判定准则】一元二次方程必须同时满足` | body | 宋体 | 源值：段首加粗 / #1E1E1E |
| `解一元二次方程的灵魂是` | body | 宋体 | 源值：常规 |
| `【图解例图】` 前缀段（共 6 段） | caption | 宋体 | 9pt / #718096 |
| 全部 11 个表格内所有 run | tableText | 宋体 | 源值 |

角色计数断言：**楷体段落 = 8（1 主标题 + 7 模块），黑体段落 = 16，其余正文类均为宋体。**

### 4.3 公式结构锚点（在产出 docx 的 document.xml 上断言）

| 锚点段落（文本前缀） | 必须存在的 OMML 结构 |
|---|---|
| `① 不等实根：` | `m:f`，其 `m:num` 含 `m:rad`（radicand 为 `Δ`）与 `±`，`m:den` 文本为 `2a` |
| `【公式核心】` | `m:sSub` 结构表达 `x₁`、`x₂`；两个 `m:f`，den 分别为 `a`、`a` |
| `2. 差方公式：` | `m:sSubSup`（底 x，sub 1，sup 2）；`m:f` 的 num 为 `m:rad`(Δ)、den 含 `|a|` |
| `【弦长秒杀公式】` | `m:rad` 的 radicand 内含 `m:sSup`（`(x₁+x₂)²`）；`m:f`（√Δ / \|a\|） |
| 表格「公式法」行（含文本 `当 Δ ≥ 0`） | 单元格内存在 `m:f`，num 含 `m:rad`(Δ) |

### 4.4 不得转换的内容（负向断言）

以下片段**必须保持普通文本**，不得变成 oMath：
- 小节编号 `1.1`~`6.5` 本体
- `减少 70%`、`90% 都在考查`、`平均耗时 15 秒`、`压缩至 2 分钟` 中的纯数字与百分数
- `（地面原点 vs 顶点原点）` 中的 `vs`
- 流程箭头 `➔`（出现在 `审题 ➔ 设元 ➔ 列式`）、圈号 `①②③`

---

## 5. 测试用例

> 测试工程建议：`TypeRight.Tests` 控制台或 MSTest；排版动作通过插件引擎（或直接 Interop 调用插件 API）执行后**另存**到 `tests/output/actual.docx`；文件级断言用 Open XML SDK 读取，运行态断言用 Interop。

### TC-01 角色识别准确率
**Given** 原始讲义 **When** 仅运行角色识别（不排版）**Then** §4.2 全部锚点段落的角色判定与期望一致；角色计数：title=1、heading1=7、heading2=16、caption=6；错判段落数 = 0。

### TC-02 字体族应用（核心）
**Given** 原始讲义 **When** 套 golden-lecture 一键排版并另存 **Then** 对产出文件逐 run 断言：
- §4.2 各锚点段落所有 run：`Font.NameFarEast` = 期望中文字体，且 `Font.NameAscii` = `Times New Roman`（XML 层断言：`w:rFonts/@w:eastAsia`、`@w:ascii`、`@w:hAnsi`）
- 11 个表格内**每一个** run：`eastAsia=宋体`、`ascii=hAnsi=Times New Roman`
- 全部 255 个 oMath 内的 run：`ascii=hAnsi=Cambria Math`

### TC-03 公式转换
**Then** 产出文件满足 §4.1 总量指标：oMath = 255、m:f = 25、m:rad = 7、上下标结构 = 201（允许 ±2 的实现差异容差，超出即 fail 并输出差异清单）；w:t 中 Unicode 上下标/根号字符数 = 0；§4.3 全部结构锚点命中；§4.4 负向断言全部通过。

### TC-04 内容零损失
**Then** 产出与输入对比：非空段落数 95 不变；表格 11 不变；图片 6 不变；提取全文中文字符流（剔除公式 run 与空白）逐字相等——**一个汉字都不许丢**。

### TC-05 源格式保持（null 语义）
**Then** §4.2 各锚点的字号/加粗/颜色与输入文件逐段一致（20pt/15pt/12.5pt、#1A365D/#2B6CB0/#1E1E1E/#718096）；段首加粗标签（如 `【判定准则】`、`① 整式性：` 所在的 43 处 #1E1E1E 加粗 run）加粗与颜色全部保留。

### TC-06 单次撤销
**When** 排版完成后执行一次 `Documents.Undo`（或 Ctrl+Z）**Then** 文档回到排版前状态：抽查 ¶0 的 NameFarEast 恢复为 `Microsoft YaHei`、oMath 计数回到 0。**整次排版必须是一个撤销单元。**

### TC-07 幂等
**When** 对产出文件再次应用同一 Profile **Then** 报告改动数 = 0，且二次产出与一次产出在 document.xml 层面无差异（规范化后比较）。

### TC-08 性能
**Then** 从触发到另存完成 < 10s（本样本 11 页 / 255 公式，属小规模；另构造 100 页放大样本时同限）。

### TC-09 样式挂载（独立 Profile）
**Given** `mountWordStyles: true` 的 Profile **Then** 产出中 heading1 段落 `Paragraph.Style` = `标题 1`、heading2 = `标题 2`，导航窗格大纲可见 7+16 个条目；且样式定义本身携带对应中西文字体（改样式即全文联动）。

### TC-10 失败安全
**When** 人为构造一个无法 BuildUp 的公式片段（如括号不闭合 `y = a(x - h`）**Then** 该片段原样保留为文本、计入统计的「转换失败」计数，其余公式照常转换，无异常抛出。

---

## 6. 断言实现指引

| 断言目标 | 推荐途径 |
|---|---|
| 字体（eastAsia/ascii） | Open XML SDK 读 `w:rPr/w:rFonts` 属性；运行态用 `Range.Font.NameFarEast / NameAscii` |
| OMML 结构 | 解压产出 docx，对 `word/document.xml` 做 XPath（命名空间 `m` = `http://schemas.openxmlformats.org/officeDocument/2006/math`），如 `//m:oMath[.//m:rad and .//m:f]` |
| 中文字符流对比 | 两文件各取 `w:t` 文本拼接，正则剔除公式 run 与空白后 `SequenceEqual` |
| 段落定位 | 一律用「文本前缀匹配」（§4.2 锚点列），禁止用段落索引（空段落增删会漂移） |
| 颜色断言 | `w:color/@w:val`，大写十六进制无 # 前缀（如 `1A365D`） |

---

## 7. 明确非验收项

- **不要求**产出与黄金文件字节级一致（docx 打包时间戳、rsid 等必然不同）
- **不要求**分页结果一致（字体度量受系统影响）
- **不断言** LibreOffice/WPS 下的渲染效果（渲染差异是字体替代问题，非插件问题）
- 黄金文件中 oMath 的具体内部嵌套形态不强制逐节点相等（BuildUp 与手工 OMML 结构合法等价即可），仅强制 §4.1 总量与 §4.3 结构锚点

## 8. 通过标准

TC-01 ~ TC-10 全绿 = 黄金样本验收通过。任一用例 fail 时输出：失败断言、期望/实际值、锚点定位（段落前缀或 XPath），禁止只报 "assert failed"。
