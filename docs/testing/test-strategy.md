# 测试策略

## 双层测试体系

| 层 | 工具 | 数量 | 覆盖 |
|---|---|---|---|
| 单元测试 | xUnit（TypeRight.Core.Tests） | 110 | 签名模型/签名读取/匹配/格式写入/公式片段切分/MathML 预备/序列重排（标记解析+切流+断序）/填表几何 |
| 验收测试 | 真机 COM（TypeRight.Acceptance） | 21 | 全流程黄金文件比对（TC-01~18）+ 修复验收（TC-19~21） |

## 运行方式

```powershell
# 单元测试（秒级）
dotnet test tests\TypeRight.Core.Tests\TypeRight.Core.Tests.csproj

# 全量验收（约 30-40 分钟）
tests\TypeRight.Acceptance\bin\Debug\TypeRight.Acceptance.exe

# 单独取证
tests\...\TypeRight.Acceptance.exe p9    # 公式化后匹配保真 + 字号零漏网
tests\...\TypeRight.Acceptance.exe tc21  # 指定用例
```

## ⚠️ 套件必须单例运行

两条验收链并发 = 互杀 Word 进程 + 互写 tests/output。症状：从某 TC 起全部"用例异常终止"+ zero 断言不匹配。启动前确认无残留 WINWORD。

## 验收资产

- 原始文档：`tests/初三数学代数终极讲义_一元二次方程与二次函数大一统.docx`（真实讲义，145 非空段+11 表格+6 图片）
- 黄金文件：`tests/初三数学代数终极讲义_字体统一版.docx`
- 测试规格：`tests/TypeRight_测试验收规格.md`

## 用例概要

| 用例 | 覆盖 |
|---|---|
| TC-01~05 | 严格/宽松模式命中数（7/8/16/21+64+7/0） |
| TC-06~09 | 中西文联动/表格覆盖/撤销粒度/幂等 |
| TC-10~14 | 公式转换总量/失败安全/黄金旅程/内容零损失/源格式保持 |
| TC-15~18 | 方案存/应用/幂等/性能 |
| TC-19 | 分式内部结构完整（MML 管线） |
| TC-20 | 表格内字号不丢失（rPr 深拷贝） |
| TC-21 | 公式化后匹配保真 + 批量字号零漏网 |

## 自动化钩子

加载项暴露 `COMAddIn.Object` 接口，测试脚本经 DCOM 调用：
- `AutomationSimilarSelect()` / `AutomationCloseSession()` / `AutomationJumpNext()`
- `AutomationConvertFormulas()` / `AutomationTextSimilarSelect()` / `AutomationRenumberSequences()`
- `AutomationForceFloatTheme(string)` / `AutomationShowFloatBar()` / `AutomationRefreshRibbon()`

钩子经 `SetTimer` 挂 Word 窗口 → WM_TIMER 回调落在 Word UI 线程（可靠语境）。
