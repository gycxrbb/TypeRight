# 功能规格：公式转换

## 概述

把 `ax²+bx+c=0` 这类纯文本公式全文批量转为 Word 可编辑公式对象（OMML），失败片段原样保留。

## 三阶段流水线

### 阶段 1：规划（纯读，无编辑）

```
逐段 Text 读取 → 公式区(OMaths)作为"墙"切出非公式子区间(gap)
→ gap 内 FindSegments 找公式片段（括号配对/UnicodeMath 上下标）
→ PrepareMml 生成 MathML（trim/balance/map/build）
→ PlannedSeg { Target=活Range, Mml=预备好的MathML, SegText, ... }
```

### 阶段 2：批量注入

```
全部公式合并一个 docx（BuildMinimalDocx 多段落 OMML）
→ 隐藏草稿 Documents.Add(Visible:=false) InsertFile 一次装载
→ 降序逐个 targets[i].FormattedText = src.FormattedText
   （降序 = 未处理目标位置稳定；活 Range 自动跟随编辑）
```

性能关键：一个 docx + 一次装载 + 降序注入，代替逐公式 Documents.Open（频闪 + ~150ms/个）。

### 阶段 3：终检（O(1)）→ 回退

```
zonesAfter - zonesBefore == claimed ?  ✓整批成功
  : 逐项定位 plan[i].Target.OMaths.Count == 0 → 标记失败
  → 失败项走单发回退（OMaths.Add + BuildUp 或 MML/XSL 注入）
  → 仍失败的保留原文本
```

## 表格 rPr 保护（TC-20）

公式构建会重写单元格 run 的 rPr（丢失 w:sz/w:b/w:color）。批量注入段（FormattedText 粘贴）**不剥离** rPr（TC-20 裁决）；只有"预备失败必走单发路径"的表格段落才做 ForceExplicitRunProps 回写。

## 三层反馈（v0.2.6+）

| 层 | 实现 |
|---|---|
| 即刻 | 等待光标 + 状态栏文字 |
| 过程 | 进度窗（Fluent 卡，视口居中）：阶段文案 + 计数 + 进度条 + Esc 提示 |
| 完成 | 进度窗就地变完成态 1.8s 淡出：「✓ 已转换 254 处 · 5.2 秒」 |

## 取消（Esc）

- `GetAsyncKeyState(VK_ESCAPE)` 实时读硬件按键状态（UI 线程被占满时唯一可靠信号）
- 检查点：规划每 32 段 / 注入每 16 处 / 回退每 8 处
- 语义：已完成部分保持有效 + 计数准确 + 整程同一 UndoRecord（Ctrl+Z 可整体回退）

## 性能参考

| 场景 | 耗时 |
|---|---|
| 254 处公式（真实讲义，进程内） | ~2s |
| 含公式区构建的同类选择（196 公式文档） | ~300ms |
| 取消后部分完成（注入早期 Esc） | ~1s |

## 已知限制

- WPS：OMath 对象缺 BuildUp → 线性片段（无结构兜底路径）转换失败，保留原文本（主管线 docx 注入可用）
- 括号不配对等不可解析片段原样保留
- 中段锚点不比对字体名（Word 读取抖动，宁少勿错）
