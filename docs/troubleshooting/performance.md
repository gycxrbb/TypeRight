# 排障手册：性能问题

## 同类选择耗时过长

### 已修复的性能陷阱

| 问题 | 症状 | 根因 | 修复 |
|---|---|---|---|
| OMaths 枚举慢路径 | 公式文档同类选择 6~9s | NetOffice 枚举器 IEnumVARIANT 在 Word 的 OMaths 上是病态慢路径（196 公式 3484ms vs 索引器 697ms） | OMathsOf 索引器优先 + 枚举兜底 |
| 公式区重复构建 | 同上 | MathZoneRanges 在样本预检 + 全文扫描各建一次 ≈ 6s | 预检 zones=null |
| 数百命中卡顿 | 会话中连续打字/点击明显延迟 | HandleSelectionChange 每次 O(n) 次 COM 活比对 | 二分 + 相邻活校验（O(log n) + ≤3 次 COM） |
| 逐公式 Documents.Open | 窗口频闪 + ~150ms/个 | 每个公式开一个临时文档 | 全部合并一个 docx + 隐藏草稿一次装载 |

### 性能诊断

addin-log 中每条匹配自带耗时分解：
```
同类匹配: 57 段 / 3045ms（进程内） ignore=None 样本单元=段落/正文
耗时分解: 规划91ms(片段扫描5) 预备9ms 草稿装载686ms 公式复制483ms ...
```

## 主题检测错误（WPS 亮色出深色浮条）

### 根因

旧方案 `CopyFromScreen` 抓屏幕像素——宿主窗口被遮挡时采到**别的窗口**，回退注册表读的是 **Word** 的主题键（WPS 里也读 Word 的 → 跨宿主污染）。

### 修复

采样改 `PrintWindow(PW_RENDERFULLCONTENT)` 抓宿主窗口自身位图（遮挡免疫）+ 回退按宿主分流（WPS 默认 Light）。

### 验证

WPS 被完全遮挡时浮条仍正确显示浅色 `#FAFAFA`（旧代码此场景错跟 Word 暗色）。

## 功能区图标消失

### 根因

`GetHicon()+PICTYPE_ICON` 在主题翻转 → Invalidate 后部分机器不被重渲染。

### 修复

换 `CreateDIBSection 32bpp 预乘ARGB + PICTYPE_BITMAP`（Excel-DNA 等一线加载项验证多年的标准路径）。
