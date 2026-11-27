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

## 「同类选择」点击后发顿（2026-11 实测与修复）

### 症状

点「同类选择」要等约 1 秒浮条才出来，整体感觉"重"、不跟手。

### 实测拆解（145 段讲义，43 处命中）

| 阶段 | 修复前 | 修复后 |
|---|---|---|
| 关旧会话 + 建会话 | ~7ms | ~7ms |
| 匹配 `MatchUnits` | 320~450ms | 320~450ms（未变，见下） |
| **高亮染色** | **380~780ms** | **0ms（移到后台）** |
| 浮条弹出 | 40~180ms | 40~180ms |
| **合计** | **~1100ms** | **~460~630ms** |

### 三个根因

**① 高亮染色在点击路径上同步执行（最大头，约占一半）**

Word COM 往返实测 **4.2~4.6ms/次**。每处命中 3 次往返（造 Range / 读原值 / 写高亮），
43 处 = 129 次 ≈ 550ms。→ 改为 `Dispatcher.BeginInvoke(..., DispatcherPriority.Background)`
排到消息循环空档：浮条立刻弹出，黄标随后补上。

> ⚠️ **坑：COM 加载项里 `System.Windows.Application.Current` 通常为 null**
> （没有 WPF Application 实例，只有 Word 的 STA 线程跑着 WPF 窗口）。
> 用 `Application.Current.Dispatcher` 会掉进同步回退分支，把 500ms 又塞回点击路径。
> 必须用浮条自身的 `Dispatcher`。

**② 主题检测每次浮条打开都强制执行（~110~200ms）**

`RefreshRibbonTheme()` 原本用 `OfficeTheme.Detect(app, force: true)` 绕过缓存，
而该检测是 `PrintWindow` 抓**整个** Word 窗口（实测 164ms/次）。

→ 两处修：`force: true` 改为走 5 秒缓存；`PrintWindow` 只抓顶部 16% 窄带。

**③ 异步高亮需要防过期**

排队期间会话可能被关掉、或被新的匹配替换。用 `ReferenceEquals` 同时比对
**会话快照**和 `MatchedUnits` 引用，任一变化就跳过这次染色。

### 仍存在的成本（未优化）

`MatchUnits` 320~450ms 是**真实计算成本**：145 段每段都要读签名
（`FormatSignatureReader.Read` ≈ 10 次 COM 往返），1450 次 × 4.3ms ≈ 300ms，与实测吻合。

试过但**无效**（已回退，别重复踩）：
- 循环外预筛空段/装饰段 → 这些段本来就不是空段，多读一次 `Range.Text` 反而更慢
- 合并相邻命中为连续段 → 本机文档 43 处是 **43 个完全分散的引导句**，合并率 0%
  （同段多单元的文档仍有收益，逻辑保留）

进一步方向（未做）：把"每段 10 次往返"改成批量读
（`Paragraphs.Range` 一次取文本 + 只对候选段读字体），预计可压到 100ms 以内。

## 浮条"糊糊的"、文字发虚（2026-11 实测与修复）

### 症状

浮条整体观感不锐利，文字边缘发灰、1px 边框有虚边。用户主观描述为"糊糊的""不丝滑"。

> 注意：这与上一节的**点击延迟**是两回事——那个是"慢"，这个是"糊"。两者可以同时存在。

### 本机环境（复现所需的全部前提）

| 项 | 值 |
|---|---|
| 物理分辨率 | 1920×1080 |
| 逻辑分辨率 | 1536×864 |
| 缩放 | **125%** |
| Word.exe DPI awareness | **System-DPI aware(1)** |

### 根因一：文字渲染设置缺失（可修，已修）

`FloatBar.xaml` 的 `<Window>` 上**三个属性一个都没有**：

| 缺失项 | 后果 |
|---|---|
| `TextOptions.TextRenderingMode="ClearType"` | 文字退化为灰度抗锯齿 → 边缘发灰 |
| `UseLayoutRounding="True"` | 布局落在小数像素上 → 边缘模糊 |
| `SnapsToDevicePixels="True"` | 1px 边框被拉成 1.25px 虚边 |

修复即补上这三行。**量化验证**（Laplacian 方差，越高越锐利）：

| 区域 | 修复前 | 修复后 | 提升 |
|---|---|---|---|
| 全图 | 1000.5 | **3132.6** | 3.1× |
| 文字行 | 2571.7 | **7637.2** | 3.0× |

复现测量方法：`CopyFromScreen` 抓 1920×1080 物理像素 → 按 `(x,y,w,h) × 1.25` 裁出浮条
→ 对灰度图算 `4*p[x,y] - p[x-1,y] - p[x+1,y] - p[x,y-1] - p[x,y+1]` 的方差。

> `AllowsTransparency="True"`（圆角+投影需要）本身会让 WPF 禁用 ClearType 子像素抗锯齿、
> 退化为灰度抗锯齿。显式声明 ClearType 只能恢复其中一部分。若愿意放弃圆角/阴影改直角矩形，
> 关掉它可拿到更锐的文字——属可选优化，未做。

### 根因二：DPI 虚拟化（宿主级，**改不了**）

实测证据链：

```
GetDpiForWindow(浮条) = 120          ← 系统知道是 125%
GetWindowRect(浮条)   = 648×155      ← 但物理像素只有 648（应为 810×194）
VirtualScreen         = 1701×1764    ← = 物理 ÷ 1.25，DPI 虚拟化的典型特征
```

含义：**Word 按 96 DPI 渲染，再由 DWM 整体拉伸 1.25 倍**。浮条跑在 Word 进程里、创建的是
子窗口，所以继承同一套渲染环境。换句话说 —— **Word 自己的界面在这台机器上也是同样糊的**。

### 三条走过的死路（**别再试**）

| 方案 | 结果 | 原因 |
|---|---|---|
| 给 `TypeRight.Addin.dll` 加 `app.manifest` 声明 `dpiAware/PerMonitorV2` | ❌ 无效 | 只有 **exe** 的 manifest 才决定进程 awareness；DLL 的 manifest 不参与 |
| `SetProcessDpiAwarenessContext(PerMonitorV2)` / `SetProcessDpiAwareness(2)` / `SetProcessDPIAware()` | ❌ 无效 | 进程 awareness 在 Word 创建**第一个窗口**时就已冻结。事后调用返回"成功"但被忽略（实测只走到 `SystemDpiAware` 降级分支，窗口尺寸毫无变化） |
| `SetThreadDpiAwarenessContext(PerMonitorV2)`（线程级覆盖） | ❌ 无效 | API 返回成功（日志 `thread -> PerMonitorV2 OK`），但 .NET Framework 4.8 的 WPF 不响应线程级 DPI 上下文变化，窗口仍是 648×155 |

> 结论：**加载项无法单方面改变宿主的 DPI 感知级别**。这不是插件 bug。
> 若将来确实要根治，只能让用户在 Windows 显示设置里把 Word 设为"替代高 DPI 缩放行为：应用程序"
> （兼容性属性），或者把系统缩放调到 100%。

### 残余风险提示

上面的 DPI 证据链对**未来排查很有用**：一旦看到
`GetWindowRect(浮条) == Width`（DIP 值）而不是 `Width × scale`，就说明进程处于 DPI 虚拟化状态。

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
