# TypeRight Design System

> 浮动格式条（FloatBar）· 对齐宿主 Word/WPS 原生 Fluent · 2026-11-07 换皮重构（v0.2.4+）

## 设计立场

浮条是宿主（Word/WPS）Mini Toolbar 的成年扩展，不是插件的自我表达画布。全部视觉语言交给 Fluent 中性灰阶 + 单一品牌蓝；圆角、阴影、悬停、选中态均按 Word 原生控件行为。

## 令牌（ThemePalette，15 项 × 3 主题）

代码真值：`src/TypeRight.Addin/UI/OfficeTheme.cs`；XAML 经 DynamicResource 消费；宿主主题由 `OfficeTheme.Detect`（屏幕实测）决定，ShowBar 时热切换。

| 令牌 | Light（Word 浅色） | Dark（Word 深色） | DarkGray | 用途 |
|---|---|---|---|---|
| Panel | #FAFAFA | #2B2B2F | #44444C | 面板底 |
| PanelBorder | #D9D9D9 | #3F3F44 | #5A5A63 | 面板发丝边 |
| Text | #1F1F1F | #F3F3F3 | #F5F5F5 | 主文字 |
| TextSecondary | #616161 | #ADADB3 | #C2C2C8 | 次级文字/标签 |
| InputBg / InputBorder | #FFFFFF / #C8C8C8 | #38383E / #57575E | #50505A / #6E6E78 | 下拉框/复选框 |
| Hover / Pressed | #ECECEC / #DEDEDE | #37373D / #42424A | #4E4E57 / #5A5A64 | 按钮悬停/按下 |
| Accent / AccentHover | #2B579A / #2F63AC | #6CA0F5 / #84B0F7 | #7FA8F0 / #93BAF4 | 唯一强调色（聚焦边/选中/色块） |
| PopupBg / PopupBorder | #FFFFFF / #E1E1E1 | #333338 / #45454B | #4A4A52 / #62626B | 弹层 |
| ItemHover | #E8F1FB（蓝晕） | #42424A | #55555F | 菜单项悬停/BIUS 选中底 |
| Separator | #EAEAEA | #3A3A3F | #55555C | 分隔线 |
| Disabled | #A6A6A6 | #77777D | #9E9EA6 | 禁用 |

## 形状与深度

- 面板圆角 **8px**，控件/下拉 **4px**，弹层 **6px**；分隔线 1px
- 阴影：面板 Blur 14 / Depth 3 / Op 0.28；弹层 Blur 10 / Depth 3 / Op 0.25（偏移+柔焦，非光晕）
- 字体：`Segoe UI, Microsoft YaHei UI` 复合（西文 Segoe、中文雅黑，杜绝宋体回退）

## 状态语义

- BIUS 选中 = **ItemHover 浅底 + Accent 文字**（Fluent 语言的浅底选中，非饱和实心块）
- 复选框选中 = Accent 实底 + 白勾（原生行为）
- 悬停灰、按下深灰；禁用 45% 透明度
- 菜单悬停：浅色主题用 #E8F1FB 蓝晕（Word 菜单同款），深色用中性灰

## 交互（不属换皮，但同系统约束）

- 强调色（当前跳转处底纹）：柔和 4 色（蜜橙/天蓝/薄荷/淡紫），与命中黄高亮协调，WPS 兜底青
- 动效：弹层 Fade 150ms 级；无装饰动效（Operate 纪律）

## 已知边界

- PrintWindow 截图对分层窗口有青色伪影——像素取证判 G/B 通道（见 tools/ui-verify.ps1）
- 高 DPI 下 WPF DIP 渲染一致；Word/WPS 主题检测失败时回退 Light
