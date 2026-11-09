# 系统架构

## 总体结构

```
┌─────────────────────────────────────────────────────┐
│                   Word / WPS 宿主进程                 │
│                                                     │
│  ┌──────────────┐    ┌──────────────────────────┐  │
│  │  Ribbon XML   │    │  TypeRight.Addin (COM)   │  │
│  │  (功能区/右键)  │──▶│  ┌────────────────────┐  │  │
│  └──────────────┘    │  │ TypeRightController │  │  │
│                      │  │ (编排器：所有命令入口) │  │  │
│  ┌──────────────┐    │  └───────┬────────────┘  │  │
│  │  浮条/进度窗   │◀───│          │               │  │
│  │  (WPF Layered)│    │  ┌───────▼────────────┐  │  │
│  └──────────────┘    │  │ TypeRight.Core      │  │  │
│                      │  │ (签名/公式/方案/设置) │  │  │
│                      │  └────────────────────┘  │  │
│                      └──────────────────────────┘  │
└─────────────────────────────────────────────────────┘
```

## 数据流：同类选择

```
用户光标点入段落
  │
  ▼
Controller.SimilarSelect()
  ├── 1. 读样本段签名（FormatSignatureReader.Read）
  │      └── 5锚点：首字/25%/50%/75%/尾字 × 中文字体/西文字体/字号/加粗/颜色
  │      └── 公式区感知：锚点跳过 OMML（MathZoneRanges）
  │
  ├── 2. 引导句边界检测（DetectLeadBoundary）
  │      └── 顺序扫描找首个"非首锚格式"字符（不用二分——岛状格式不满足单调性）
  │      └── 混合段拆"引导句+正文"两个匹配单元
  │
  ├── 3. 全文匹配（SimilarSelector.MatchUnits）
  │      └── 单遍枚举所有段落（WordCompat.ParagraphsOf 保 O(n)）
  │      └── 表格分区隔离（正文样本不匹配表格段）
  │      └── 严格/宽松模式（忽略字号/颜色/加粗）
  │
  ├── 4. 高亮应用（ApplySessionHighlights）
  │      └── 黄色 HighlightColorIndex 写入命中区间（排除段落/单元格标记）
  │      └── 包在一个 UndoRecord 内
  │
  └── 5. 浮条弹出（FloatBar.ShowBar）
         └── 主题跟随（OfficeTheme.Detect → ApplyTheme）
         └── 显示匹配计数 + 样本签名
```

## 数据流：文本同类（v0.3.0 新增）

```
用户选中/光标点入一个词（如「正确」）
  │
  ▼
Controller.TextSimilarSelect()
  ├── 1. 采样模式：选区文本 或 光标处向两侧扩展连续词字符
  │
  ├── 2. 全文搜索（TextMatcher.FindOccurrences）
  │      └── 每段 Text 一次读取，C# IndexOf 全部命中（非重叠）
  │      └── 否定前缀排除（不/非/未 前一个字符 → 跳过）
  │      └── 公式区跳过
  │
  └── 3. 会话创建 + 高亮 + 浮条（文本模式：匹配行切为「正确」× N + 排否定开关）
```

## 数据流：公式转换

```
用户点击「公式转换」
  │
  ▼
Controller.ConvertFormulas()
  ├── 即刻反馈：等待光标 + 状态栏 + 进度窗显示
  │
  ├── FormulaConverter.ConvertDocument()
  │   ├── 阶段1 规划（纯读）：逐段 Text → 纯位置切分 gap → FindSegments → PrepareMml
  │   │   └── 公式区（OMaths）作为"墙"切出非公式子区间
  │   │
  │   ├── 阶段2 注入：
  │   │   ├── 全部公式合并一个 docx（BuildMinimalDocx 多段落）
  │   │   ├── 隐藏草稿 Documents.Add(Visible:=false) InsertFile 一次装载
  │   │   └── 降序逐个 targets[i].FormattedText = src.FormattedText（目标位置稳定）
  │   │
  │   └── 阶段3 终检（O(1)）：doc.OMaths.Count 增量 == 声明数？不符才逐项回退
  │
  ├── 取消（Esc）：GetAsyncKeyState 实时检测 → 注入每16处/回退每8处检查
  │   └── 已完成部分保持有效，整程同一 UndoRecord
  │
  └── 完成反馈：进度窗变完成态 1.8s 淡出
```

## 兼容层（WordCompat）

```
WordCompat.ParagraphsOf(collection)
  ├── Word: 直接返回 NetOffice 枚举结果
  └── WPS: NetOffice 枚举器抛 FactoryException → 底层 IEnumVARIANT + 手工 new Paragraph(null, item)

WordCompat.OMathsOf(collection)
  ├── Word: 索引器（快路径，196公式 697ms vs 枚举 3484ms）
  └── WPS: 索引器也走工厂 → 枚举兜底

WordCompat.BeginUndo / EndUndo
  ├── Word 2013+: Application.UndoRecord.StartCustomRecord
  └── Word 2010: 无 UndoRecord → null（撤销粒度退化为多条）
```

## 主题检测（OfficeTheme.Detect）

```
PrintWindow(Word窗口) → 位图 → 顶部3%/7%/13%行采样 → 亮度判 Light/Dark/DarkGray
                        ↑
              PrintWindow 抓宿主窗口自身（遮挡免疫）
              （旧方案 CopyFromScreen 被遮挡时采到别的窗口 → WPS亮色出深色浮条的根因）
```

## 部署结构

```
%LOCALAPPDATA%\TypeRight\Addin\     ← DLL 部署目录
  TypeRight.Addin.dll               ← 主加载项
  TypeRight.Core.dll                ← 核心库
  NetOffice.dll / WordApi.dll / OfficeApi.dll / VBIDEApi.dll
  Newtonsoft.Json.dll / stdole.dll

%LOCALAPPDATA%\TypeRight\
  settings.json                     ← 用户设置
  addin-log.txt                     ← 生命周期日志
  profiles\                         ← 方案 JSON
  first-run-shown.txt               ← 首次提示标记

注册表:
  HKCU\Software\Microsoft\Office\Word\Addins\TypeRight.Connect
    LoadBehavior=3, FriendlyName, Description
  HKCU\Software\Classes\CLSID\{6C7C...}          ← 64位 COM 类
  HKLM\Software\Classes\Wow6432Node\CLSID\{...}  ← 32位 COM 类（WPS 用，需管理员）
```
