# 功能规格：填表防崩套件（照片严丝合缝入格 + PPT 式悬浮填表框）

> 状态：已上线（v0.4.0 · 2026-11-10）  
> 目标场景：高校各类申请表（优秀团员/发展对象推荐表、三好学生、奖学金、课题申报）、政企登记表/审批表、单页简历等**“版面严格固定、绝对不能撑破表格或顶乱格式”**的文档填写场景。

---

## 1. 设计哲学与核心取舍

Word 的底层排版引擎是**流式排版（Flow Layout，牵一发而动全身）**，而用户在填写申请表、贴证件照、做单页简历时，心理预期是 **PowerPoint 式的容器/画布排版（Box/Canvas Layout，格子固定、互不挤压）**。

在功能落地过程中，我们确立了三条核心准则：
1. **解决用户真痛点**（照片贴入对齐难、打字撑大单元格/顶走落款/顶歪表头下划线）；
2. **用起来零阻碍**（一键直达，不改变用户习惯）；
3. **绝不破坏原本的文档内容与视觉美感**（摒弃全表暴力改行高或局部缩小字号导致字体忽大忽小的做法，统一用 PPT 式的浮动图层思路由点及面解决格式挤压问题）。

因此，功能区「**填表防崩**」分组收敛为两个高频、纯粹、零破坏的核心按钮：
- **📸 照片入格**（`PhotoCellFitter`）：一键将证件照/图片严丝合缝嵌入目标单元格，四边紧贴表格内壁（留出 `0.75pt` 边框线保护余量），不裁切画面、不撑破表格。
- **📦 悬浮填表框**（`CellTextBoxPlacer`）：在当前表格单元格或表头下划线处铺设无边框、全透明的 PPT 式悬浮文本框（`wdWrapFront`），完全保留格内原有文字与底部签字栏位置，怎么打字都绝不撑大表格或顶歪右侧文字。

---

## 2. 功能一：📸 照片入格（Smart Photo Fit-to-Cell）

### 2.1 三种自然操作路径
1. **剪贴板直贴**：复制图片/照片文件后点击「照片入格」→ 自动贴入当前格（若光标不在照片格则自动识别表格中的照片格），Toast 提供「换选文件…」快捷替换入口。
2. **选本地文件贴入**：光标点进照片格后点击「照片入格」（剪贴板无图时）→ 弹出文件选择框选图贴入。
3. **已粘贴图片救场**：已将图片粘贴进单元格（`InlineShape` 把行高撑烂）或选中了浮动图片（`Shape`）→ 点击「照片入格」自动转悬浮、恢复行高并严丝合缝入格。

### 2.2 核心算法：两步等比缩放 + 尺寸微调对齐（零画面裁切）

代码落点：[`PhotoCellFitter.cs`](../../src/TypeRight.Core/Forms/PhotoCellFitter.cs)、[`FormGeometry.cs`](../../src/TypeRight.Core/Forms/FormGeometry.cs)

1. **清空提示字并归零顶端对齐（同一 `UndoRecord` 内）**
   - 清空照片格内预置的 `"近期一寸\r免冠照片"` 提示文字（保留单元格结束符 `\r\a`），并将单元格设为顶端对齐（`wdCellAlignVerticalTop`）+ 段落左对齐，确保测高起点精确落在单元格上边框。
2. **第 ① 步：测定目标单元格与图片的宽高及宽高比**
   - **非合并单元格**：直接读取 `targetCell.Width` 与 `targetCell.Height`（零 `Range.Information` 调用）。
   - **纵向合并单元格（跨多行照片格，如跨 1~3 行）**：
     - Word 对纵向合并单元格读取 `Cell.Height` 会返回哨兵值 `9999999`，且被合并的第 3 行常因多行文字折行（`wdRowHeightAtLeast`）导致实际渲染高度大于名义行高（例如《发展对象申请表》第 3 行名义高 `22pt`，实际被 3 行文字撑到 `39.55pt`）。
     - **精确测高算法（[`FormGeometry.MeasureMergedPhotoCellHeightFast`](../../src/TypeRight.Core/Forms/FormGeometry.cs)）**：
       - 读取已归零照片格的起始垂直坐标 `startTopY = GetPageY(tStart)`；
       - 从 `targetCol` 索引处向后扫描紧邻下方的首个宽格（`cw > targetW * 1.4f`）或同列格 `belowCell`，读取其首字符垂直坐标 `belowTopY = GetPageY(bStart)`；
       - **扣除下方单元格行距偏移（`extraLeading`）**：若下方单元格（如「团支部意见」正文格）设置了固定行距（如 `20pt`）或段前距（`SpaceBefore`），其首字符 `GetPageY` 会比标准单倍行距下移 `(LineSpacing - 14.3pt) * 0.5`。扣除该偏移后计算 `(belowTopY - extraLeading) - startTopY`，即可得到合并格的**精确外框高度**（误差 `< 0.5pt`）。
   - **扣除边框保护内缩（`BorderInsetPt = 0.75pt`）**：
     - `targetW = cellW - 1.5f`，`targetH = cellH - 1.5f`，计算单元格宽高比 `cellRatio = targetW / targetH`。
   - **原图零损加载**：本地文件或剪贴板文件直接交由 `doc.Shapes.AddPicture` 原生加载（不经过 GDI+ 裁切或重编码，保留 100% 完整画面内容），读取图片原始宽高 `imgW`、`imgH` 与宽高比 `imgRatio = imgW / imgH`。
3. **第 ② 步：对比宽高比，先等比缩放、再按需微调尺寸做到四边严丝合缝（[`PhotoCellFitter.ScaleAndAlignPhotoShape`](../../src/TypeRight.Core/Forms/PhotoCellFitter.cs)）**
   - **绑定单元格四边外框原点**：
     - 当图片锚定在单元格内（`LayoutInCell = -1`）时，Word 默认的水平基准面是 `wdRelativeHorizontalPositionColumn`（起点含 `LeftPadding = 5.4pt`，会导致整张图向右偏移 `5.4pt`，左边留白、右边冲出表格右边框），默认垂直基准面是 `wdRelativeVerticalPositionParagraph`（向下偏移 `~1.3pt`）。
     - 显式设置 `RelativeHorizontalPosition = wdRelativeHorizontalPositionPage` 与 `RelativeVerticalPosition = wdRelativeVerticalPositionPage`（配合 `LayoutInCell = -1`），使 `(0, 0)` 原点精确落在**单元格四条外边框**上。
   - **先等比例缩放**：开启 `LockAspectRatio = MsoTriState.msoTrue`，将图片等比例缩放到与单元格宽度对齐（`shape.Width = targetW`）。
   - **按需微调对齐**：
     - 若图片与单元格原本宽高比一致（`|imgRatio - cellRatio| <= 0.008`），无需改变比例，直接移入 `(0.75pt, 0.75pt)` 即完成；
     - 若宽高比不一致，解锁纵横比（`LockAspectRatio = MsoTriState.msoFalse`），将尺寸微调至 `Width = targetW`、`Height = targetH`，并定位于 `Left = 0.75pt`、`Top = 0.75pt`，实现四边严丝合缝且四周表格黑线完整露出。

---

## 3. 功能二：📦 悬浮填表框（PPT-Style Floating Form TextBox）

### 3.1 两大适用位置
1. **表格单元格内（场景 A）**：
   - 光标点进任意单元格（如「团支部意见」「学院意见」或普通信息格）→ 点击「悬浮填表框」；
   - 生成几乎撑满整个所在单元格（四周内缩 `1.0pt`）的无边框、全透明悬浮文本框，光标自动进入框内闪烁，直接打字即可。
2. **表头下划线处（场景 B）**：
   - 光标点在表头下划线填空区（如 `推荐单位：____202X级____（全称）____团支部`）→ 点击「悬浮填表框」；
   - 自动识别当前光标所在的下划线区间起止坐标，在下划线正上方铺设等宽透明悬浮文本框（自动继承原下划线处字体字号并居中），打字绝不顶歪右侧文字。

### 3.2 关键防坑设计（零破坏、零位移、零抖动）

代码落点：[`CellTextBoxPlacer.cs`](../../src/TypeRight.Core/Forms/CellTextBoxPlacer.cs)、[`FormGeometry.cs`](../../src/TypeRight.Core/Forms/FormGeometry.cs)

1. **绝不删改单元格内原有示例或说明文字**
   - 插入悬浮填表框时，仅做只读几何测绘，**不对单元格自身的 `Range.Text` 做任何删除或修改**，完整保留模板自带的填写示例与底部落款。
2. **表外同页就近锚定（彻底解决底端对齐签名上移、左侧灰条与视口滚动抖动）**
   - **痛点 1（签名上移与左侧灰条）**：在《发展对象申请表》中，「学院团委或直属团支部意见」等单元格设置了 `VerticalAlignment = wdCellAlignVerticalBottom`（底端对齐），依靠顶部的空行把「负责人（签字盖章）/ 年 月 日」留在单元格底部。若将浮动 `Textbox` 锚定在该单元格内部的段落上，Word 会重算单元格内容布局，导致底部签字行瞬间向上跳移，且在单元格左侧边缘产生一条灰色锚定指示条。
   - **痛点 2（视口滚动抖动）**：若所有表外锚点一律挂在文档开头 `doc.Range(0, 0)`，当用户滚动到页面下半部操作时，调用 `TextFrame.TextRange.Select()` 会触发 Word 将视口向页首滚动再弹回，产生肉眼可见的画面抖动。
   - **解法（[`FormGeometry.GetPageAnchorOutsideTable`](../../src/TypeRight.Core/Forms/FormGeometry.cs)）**：
     - 浮动文本框统一锚定在**表格外部**的段落上（从而 `LayoutInCell = 0`，与单元格内部流式排版完全隔离，底端对齐签名纹丝不动、零灰条）；
     - **上下半页智能就近锚定**：上半页单元格（`Top < 320pt`）优先锚定在表头首段 `doc.Range(0, 0)`，下半页单元格（`Top >= 320pt`）优先锚定在表格紧邻下方的表尾段 `doc.Range(table.Range.End, table.Range.End)`，并锁定 `ActiveWindow.VerticalPercentScrolled`，彻底消除视口滚动抖动。
3. **非破坏性单元格左上角坐标反推（[`FormGeometry.MeasureCellTopLeftFast`](../../src/TypeRight.Core/Forms/FormGeometry.cs)）**
   - 无需修改单元格对齐方式，仅通过 **2 次** `Range.Information` 调用反推单元格精确左上角页面坐标：
     - **水平坐标**：`cellLeft = GetPageX(startRng) - LeftPadding(5.4pt) - FirstLineIndent`；
     - **垂直坐标**：
       - 若为顶端对齐（`wdCellAlignVerticalTop`）：`cellTop = GetPageY(startRng)`；
       - 若为底端对齐（`wdCellAlignVerticalBottom`）：直接读取单元格末尾字符的 `yEnd = GetPageY(endRng)`，反推 `cellTop = (yEnd + 15.6f) - nominalHeight`（省去首字符 Y 查询，且不受底部多行落款影响）；
       - 若为单行垂直居中（`wdCellAlignVerticalCenter`）：`cellTop = GetPageY(startRng) - (nominalHeight - 15.6f) * 0.5f`。

---

## 4. 性能优化记录（从 ~2s 降至进程内 ~100–200ms）

在真机 Word 16.0 上对《首都经济贸易大学优秀团员作党的发展对象申请表.docx》进行分步打点剖析后，实施了以下关键优化：

| 瓶颈环节 | 优化前耗时 | 根因分析 | 优化方案 | 优化后耗时 |
| :--- | :---: | :--- | :--- | :---: |
| **文本框属性配置 (`props`)** | `1438 ms` | 创建 `Textbox` 后连续执行了 **28 次** COM 属性写入（含大量已是 Word 默认值的 `Margin*`、`WordWrap`、`AutoSize`、`SpaceBefore/After`、`FirstLineIndent` 等） | 仅保留 **6 次**非默认核心属性写入（`WrapFormat.Type`、`Line.Visible`、`Fill.Visible`、`AlternativeText`、`Font.Name`、`Font.Size`） | **`62 ms`** (-95%) |
| **合并照片格测高 (`measure`)** | `1346 ms` | 遍历全表 26 个单元格逐一读取 `c.RowIndex` 与 `c.Width` 跨进程 COM 属性 | 直接从 `targetCol` 索引位置向后扫描首个宽格/同列格，仅需 2 次 `GetPageY` | **`67 ms`** (-95%) |
| **选区图形预检 (`selSh`)** | `11–180 ms` | 普通文本光标下直接访问 `sel.ShapeRange` 会触发 Word 内部 COM 异常开销 | 先判断 `sel.Type == WdSelectionType.wdSelectionShape` 再访问 `sel.ShapeRange` | **`0–1 ms`** |
| **原图加载 (`prep`)** | `66 ms` | 每次均用 GDI+ 解码图片、裁切并重编码保存临时 PNG 文件 | 本地/剪贴板文件直接传给 `Shapes.AddPicture` 原生加载，零重编码 | **`0 ms`** |
