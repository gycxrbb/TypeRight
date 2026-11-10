# TypeRight — Word 指认式批量排版插件

> **光标点一下，全文同类段落一次改齐。**
> 同类选择（格式签名匹配）· 文本同类（字义匹配）· 纯文本公式一键转公式对象 · 填表防崩（照片严丝合缝入格 + PPT 式悬浮填表框）· 操作存为方案下份文档一键重放

一个 COM 加载项，给讲义 / 试卷 / 笔记 / 申请表这类 Word/WPS 文档用：不用宏、不逐段刷格式、所见即所得、每步可撤销。

**性能实测**（254 处公式的真实讲义）：公式转换 **2s**（进度窗 + Esc 可取消）；全文匹配 **<0.5s**；文本同类 **<0.1s**。

---

## ① 我只想装来用

### 环境要求

| | |
|---|---|
| ✔ 支持 | Windows 7 SP1+ / Win10 / Win11 + **Microsoft Word** 2010~365（32/64 位）+ **WPS Office**（个人版/教育版） |
| ✘ 不支持 | Word 网页版、Mac、Word 2007 及更早 |
| 依赖 | .NET Framework 4.8（Win10+ 一般自带；Win7/8 需手动安装，安装器自动检测） |

### 安装（3 步）

1. 下载最新 `TypeRight-vX.Y.Z.zip`（[Releases 页](https://github.com/gycxrbb/TypeRight/releases)）
2. 关闭 Word/WPS，解压双击 **install.bat**（用 WPS 的选 y 启用支持）
3. 启动 Word → 功能区出现「TypeRight」选项卡即成功

### 五步上手

| 步骤 | 操作 |
|---|---|
| 1 | 光标点进一个段落（或选中一段文字） |
| 2 | 「TypeRight」→ **同类选择**（按格式匹配）或 **文本同类**（按文字内容匹配） |
| 3 | 浮条上改字体 / 字号 / 加粗 / 斜体 / 下划线 / 删除线 / **着重号** / 颜色（标准色板 + 系统取色器） |
| 4 | 「下一处 ▼」逐处跳转（滚动到命中处 + 当处强调色区分）；点色块换强调色 |
| 5 | **公式转换**：进度窗 + Esc 可取消；**序列重排**：插入/删除条目后一键重排全文序号；**照片入格 / 悬浮填表框**：贴照片严丝合缝、打字绝不撑坏申请表；**方案**：整套操作存为方案下份文档一键回放 |

### 常见问题

- **功能区没有 TypeRight？** 关掉 Word 双击 `repair.bat`
- **Word 2010 有什么不同？** 撤销粒度降级（一次批量修改需多按几次 Ctrl+Z）
- **安全软件弹窗？** 选信任/允许（未做商业签名，插件完全离线不联网）
- **反馈问题？** 选项卡「使用说明」→「复制诊断信息」
- **卸载**：双击 `uninstall.bat`

---

## ② 我是开发者

### 架构

```
src/TypeRight.Core/              无 UI 核心库
├── Signature/                   格式签名匹配 + 字义匹配 + 批量执行
├── Formatting/                  公式转换三阶段流水线
├── Formulas/                    MathML → OMML 批量注入
├── Forms/                       填表防崩套件（照片严丝合缝入格 + PPT式悬浮填表框）
├── Sequences/                   序列重排（1. ① （1） 一、 断序一键重排）
├── Profiles/                    方案存回放
├── Settings.cs                  用户设置
└── WordCompat.cs                宿主兼容层（老Word降级 + WPS枚举）

src/TypeRight.Addin/             COM 加载项
├── Controllers/                 编排器（所有命令入口）
├── UI/                          浮条 / 进度窗 / Toast / 图标 / 主题检测
└── Ribbon.xml                   功能区 + 右键菜单定义
```

### 构建与开发安装

```powershell
# 单元测试（110 个）
dotnet test tests\TypeRight.Core.Tests\TypeRight.Core.Tests.csproj

# 开发安装
powershell -ExecutionPolicy Bypass -File tools\install.ps1

# 全量验收（21 用例，30-40 分钟）
tests\TypeRight.Acceptance\bin\Debug\TypeRight.Acceptance.exe
```

### 发布

```powershell
git tag v0.4.0 && git push origin v0.4.0   # → Actions → Release
```

### 更多文档

| 文档 | 内容 |
|---|---|
| [AGENTS.md](AGENTS.md) | 开发规则与 AI 助手上下文 |
| [docs/architecture/](docs/architecture/) | 系统架构与数据流 |
| [docs/features/](docs/features/) | 功能规格（同类选择/文本同类/公式转换） |
| [docs/compatibility/](docs/compatibility/) | Word/WPS/Windows 兼容矩阵 |
| [docs/troubleshooting/](docs/troubleshooting/) | 排障手册（功能区消失/性能/主题检测） |
| [docs/testing/](docs/testing/) | 测试策略与验收口径 |
| [docs/release/](docs/release/) | 发布流程与版本记录 |
| [DESIGN.md](DESIGN.md) | 设计系统（令牌/圆角/状态语义） |
| [PRODUCT.md](PRODUCT.md) | 产品上下文 |

---

仅供学习交流。二次开发请注明出处。
