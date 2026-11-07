# Surface brief: FloatBar（浮动格式条）

## Scope & Mode
TypeRight 浮动格式条（WPF Layered Window，622×自适应，三行：匹配信息/字体字号样式/匹配条件）及其弹出层（字号、强调色、字体下拉）。Operate 模式。用户已选：换皮对齐（布局与功能零改动）。

## Audience / Job
教师在 Word/WPS 中批量改讲义格式；浮条是高频操作面，必须像宿主原生工具而非第三方贴片。

## Direction contract

THESIS：浮条不是"插件的脸"，是 Word/WPS 自身 Mini Toolbar 的成年扩展——所有视觉判断交给宿主语言，插件零自我表达。

OWN-WORLD：Fluent 中性灰体系。浅色：#FAFAFA 面板 + #DADADA 发丝边 + #1F1F1F/#616161 双级文字 + #2B579A 唯一强调色；深色：#2B2B2F 面板 + #3F3F42 隐边 + #6CA0F5 强调。8px 面板圆角、4px 控件圆角、6px 弹层圆角；偏移柔和阴影；选中态=强调色浅底而非饱和实心块；菜单悬停浅色用 #EFF6FC 蓝晕。

STORY：用户看到浮条的第一反应应是"Word 自己弹的"；第二反应才是"它比原生多懂格式签名"。

FIRST VIEWPORT：622px 宽圆角浮条悬于文档上方；首行"⠿ 已匹配 N 段 + 下一处 n/N ▼ + 色块"左、✕ 右；第二行字体/字号±/BIUS/颜色（Word 字体组同构）；第三行匹配条件；底部西文联动提示。全部控件 Fluent 灰阶，强调色只出现在：选中态、聚焦边、色块本身。

FORM：brief-pinned（用户指定对齐宿主原生，参考 Word Mini Toolbar）；seed: none (unavailable — 方向由用户钉定，未掷骰)。

FINISH：unreviewed and undocumented is unfinished; this build ends with the finish review, the verdict, DESIGN.md, and every shipping raster carrying its provenance.
