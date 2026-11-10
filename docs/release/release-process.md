# 发布流程

## 自动发布（推荐）

```powershell
# 1. 版本号已更新（TypeRight.Addin.csproj <Version>）
# 2. 提交推送
git add -A && git commit -m "v0.4.0" && git push origin main
# 3. 打 tag（触发 GitHub Actions 自动构建+发 Release）
git tag v0.4.0 && git push origin v0.4.0
# 4. 等 Actions 跑完（~5 分钟），产物自动挂 Release
```

流水线：`.github/workflows/release.yml`（windows runner → package-release.ps1 → zip → softprops 挂 Release）

## 手动出包

```powershell
powershell -File tools\package-release.ps1
# 产物: dist\TypeRight-v0.4.0.zip
```

## 安装包内容（18 个文件）

```
install.bat / repair.bat / uninstall.bat       ← 用户入口（ASCII 包装）
dist-scripts\install-core.ps1                  ← 安装逻辑（BOM 必须有）
dist-scripts\repair-core.ps1                   ← 修复（LoadBehavior 重置为 3）
dist-scripts\uninstall-core.ps1                ← 卸载
dist-scripts\install-wps-elevated.ps1          ← WPS 32位注册（UAC，可选）
files\TypeRight.Addin.dll 等 10 个 DLL          ← 编译产物
使用说明.txt                                     ← 用户手册（BOM）
```

## 发给朋友的直达链接

```
https://github.com/gycxrbb/TypeRight/releases/download/v0.4.0/TypeRight-v0.4.0.zip
```

> ⚠️ Releases 页面另有 GitHub 自动附带的 "Source code (zip)"——发直达链接避免混淆。

## 版本记录

| 版本 | 日期 | 主线 |
|---|---|---|
| 0.1.0 | 2026-10 | MVP |
| 0.2.0 | 2026-11-04 | 离线安装包 + 帮助窗 + 首次提示 |
| 0.2.1 | 2026-11-05 | Word 2010/2013 + Win7 兼容 |
| 0.2.2 | 2026-11-06 | Office 2610 size="small" 根因修复 |
| 0.2.3 | 2026-11-06 | WPS 支持 |
| 0.2.4 | 2026-11-07 | 跳转滚动 + 强调色 + 字体/图标 |
| 0.2.5 | 2026-11-07 | Fluent 换皮 + 主题跟随 + 性能修复 |
| 0.2.6 | 2026-11-08 | 公式转换三层反馈 + Esc 取消 |
| 0.2.7 | 2026-11-08 | 进度窗居中 |
| 0.3.0 | 2026-11-09 | 文本同类（字义匹配）+ 着重号 |
| 0.3.1 | 2026-11-09 | XML 属性重复修复 |
| 0.4.0 | 2026-11-10 | 填表防崩套件（照片严丝合缝入格 + PPT 式悬浮填表框） |
| 0.5.0 | 2026-11-10 | 序列一键重排 + 字体颜色板扩展（标准色 + 系统取色器） |
