# 排障手册：功能区（选项卡）消失

## 症状

Word 启动后 TypeRight 选项卡不出现，右键菜单也没有"同类选择"。插件本体正常加载（日志有 OnConnection/GetCustomUI）。

## 诊断判据

日志看 `GetCustomUI → N 字符` 后有无 `Ribbon onLoad`：
- **有 onLoad** = XML 被采纳，选项卡应可见
- **无 onLoad** = XML 被静默拒绝（整块丢弃）

## 已知诱因（按发现顺序）

| # | 诱因 | 版本 | 修复 |
|---|---|---|---|
| 1 | `size="small"` 属性 | v0.2.2 | btnHelp 改 `size="large"` |
| 2 | XML 属性重复（insertBeforeMso/getImage 写两遍） | v0.3.1 | 删重复行 |

## 预防

编辑 `Ribbon.xml` 后**必须**跑：

```powershell
python -c "import xml.etree.ElementTree as ET; ET.Parse(r'src\TypeRight.Addin\Ribbon.xml'); print('OK')"
```

ElementTree 会报重复属性。`size="small"` 需人工检查（XML 合法但 Word 2610 拒绝）。

## 排障配方（二分法）

1. 部署目录放 `Ribbon.Debug.xml`（v0.2.2+ 的诊断旁路：存在则优先读取）
2. 从零回调 XML 开始：零回调 → 加 onLoad → 加 onAction → 加 getImage → 加 contextMenus → 对正式 XML 逐块剥离
3. 每轮一个 Word 实例 + UIA 枚举 TabItem + 日志看 onLoad
4. 本配方 12 轮内可定位到单个属性
