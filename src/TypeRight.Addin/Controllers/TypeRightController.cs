using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using NetOffice.WordApi;
using NetOffice.WordApi.Enums;
using TypeRight.Addin.UI;
using TypeRight.Formatting;
using TypeRight.Forms;
using TypeRight.Profiles;
using TypeRight.Signature;

namespace TypeRight.Addin.Controllers
{
    /// <summary>一次同类选择会话（PRD v2 P0-1/P0-2）。</summary>
    internal sealed class SelectSession
    {
        public Document Doc;
        public string DocFullName;
        public Paragraph SampleParagraph;
        /// <summary>光标在样本段内的偏移（引导句感知：决定样本单元是引导句还是正文）。</summary>
        public int SampleOffset;
        /// <summary>样本段落的原始签名（应用格式前采集，方案规则用它匹配）。</summary>
        public FormatSignature SampleSignature;
        /// <summary>浮条显示用的签名：格式模式 = SampleSignature；
        /// 文本同类模式没有格式签名，取首个命中处的实际格式（用户要知道"这段现在是什么字体字号"）。</summary>
        public FormatSignature DisplaySignature;
        /// <summary>样本是否引导句单元。</summary>
        public bool SampleIsLead;
        /// <summary>样本段落是否在表格内（正文/表格分区，按表格区间位置判定）。</summary>
        public bool SampleInTable;
        public FormatSignature.IgnoreDims Ignore;
        /// <summary>命中单元（引导句感知）：高亮与格式应用都作用于单元区间。</summary>
        public List<MatchUnit> MatchedUnits = new List<MatchUnit>();
        /// <summary>本会话的"目标字符类型"：只对该类型的字符高亮/改格式，段内其他类型保持原样。
        /// None = 不做字符类型过滤（旧行为，整段刷）。
        /// 取值来源：用户光标/选区落在哪个字符上——点汉字就只动汉字，点数字就只动数字
        /// （用户 2026-11 口径："汉字归汉字、数字归数字"）。</summary>
        public CharClass TargetClass = CharClass.None;
        public SimilarSelector.HighlightSession Highlights;
        /// <summary>文本同类模式（字义匹配）：单元=文本出现处而非格式段落；忽略维度等开关不适用。</summary>
        public bool TextMode;
        public string TextPattern;
        public bool TextExcludeNegation = true;
    }

    /// <summary>一次批量格式操作（沉淀为方案规则的原材料）。</summary>
    public sealed class OpRecord
    {
        public FormatSignature SampleSignature;
        public FormatSignature.IgnoreDims Ignore;
        public FormatChanges Applied;
        /// <summary>样本是否在表格内（方案规则要区分正文/表格两个文本域）。</summary>
        public bool SampleInTable;
    }

    /// <summary>
    /// 指认式排版编排器（PRD v2 §3 用户旅程）：
    /// 同类选择 → 浮动条改格式（每步独立撤销）→ 公式转换 → 存为方案 → 下一份文档一键回放。
    /// </summary>
    public sealed class TypeRightController
    {
        private readonly TypeRightAddin _addin;
        private AppSettings _settings = SettingsStore.Load();

        private SelectSession _session;
        private FloatBar _floatBar;
        private ToastWindow _toast;
        private FormulaFailuresWindow _failuresWindow;

        private readonly List<OpRecord> _docOps = new List<OpRecord>();
        private int _lastMatchCount;

        /// <summary>当前会话命中数（浮条跳转进度显示用）。</summary>
        public int GetMatchedCount() { return _lastMatchCount; }

        /// <summary>Esc 是否处于按下态（GetAsyncKeyState 读硬件状态，与焦点无关；
        /// 转换在 Word UI 线程上跑、无消息循环，这是唯一可靠的取消信号）。</summary>
        private bool IsEscapePressed()
        {
            try { return (GetAsyncKeyState(0x1B) & 0x8000) != 0; }
            catch { return false; }
        }

        [DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(int vKey);

        /// <summary>公式转换完成主文案（人话）。</summary>
        private static string FormulaDoneMainText(FormulaStats s)
        {
            if (s.Cancelled) return "已取消 · 本次已转换 " + s.Converted + " 处（Ctrl+Z 可整体撤销）";
            if (s.Converted == 0 && s.Failed == 0) return "没有找到可转换的文本公式";
            return "✓ 已转换 " + s.Converted + " 处公式" + (s.ElapsedMs > 0 ? " · " + (s.ElapsedMs / 1000.0).ToString("0.0") + " 秒" : "");
        }

        private static string FormulaDoneSubText(FormulaStats s)
        {
            if (s.Failed > 0) return s.Failed + " 处无法解析，已保留原文本（多为括号不配对）";
            return s.Cancelled ? "" : "Esc 关闭浮动条可见完整统计";
        }

        /// <summary>设置跳转强调色（十六进制 RRGGBB）；持久化，会话进行中立即重涂当前处。</summary>
        public void SetEmphasisColor(string hex)
        {
            try
            {
                _settings.EmphasisColor = hex;
                SettingsStore.Save(_settings);
                if (_session != null) _session.Highlights.ReapplyCurrent(_settings.EmphasisColorRgb);
            }
            catch (Exception ex) { TypeRightAddin.Log("SetEmphasisColor: " + ex.Message); }
        }

        /// <summary>保存设置（浮条"其他颜色"自定义色等轻量持久化入口）。</summary>
        public void SaveSettings()
        {
            try { SettingsStore.Save(_settings); }
            catch (Exception ex) { TypeRightAddin.Log("SaveSettings: " + ex.Message); }
        }
        private TimeSpan _lastMatchDuration = TimeSpan.Zero;
        private string _opsDocFullName;
        private bool _saveOffered;

        public TypeRightController(TypeRightAddin addin)
        {
            if (addin == null) throw new ArgumentNullException("addin");
            _addin = addin;
        }

        public Application App { get { return _addin.WordApplication; } }

        public AppSettings Settings { get { return _settings; } }

        // ───────────────────────── P0-1 同类选择 ─────────────────────────

        /// <summary>入口：光标段落（或选区首段）为样本，匹配全文同类。</summary>
        public void SimilarSelect()
        {
            var swAll = System.Diagnostics.Stopwatch.StartNew();
            TypeRightAddin.Log("SimilarSelect: enter");
            try
            {
                Document doc = ActiveDocument();
                if (doc == null)
                {
                    TypeRightAddin.Log("SimilarSelect: no active document");
                    ShowToast("请先打开一个文档");
                    return;
                }

                Paragraph sample = null;
                try { sample = App.Selection.Range.Paragraphs[1]; }
                catch (Exception ex) { TypeRightAddin.Log("SimilarSelect: sample read failed: " + ex.Message); }
                if (sample == null)
                {
                    ShowToast("光标不在正文段落中");
                    return;
                }

                // 预检只判有无可见文本：zones=null 省一次全套公式区构建（真实带区读取在 MatchUnits 内）
                var sig = FormatSignatureReader.Read(sample, null);
                TypeRightAddin.Log("SimilarSelect: sample sig = " + sig);
                if (sig.IsEmpty)
                {
                    ShowToast("当前段落没有可见文本，换一个段落试试");
                    return;
                }

                CloseSession();

                int sampleOffset = 0;
                try { sampleOffset = Math.Max(0, App.Selection.Range.Start - sample.Range.Start); } catch { }

                // 目标字符类型 = 用户点中/选中的那个字符的类型。
                // 点汉字 → 只动汉字（日期里的数字保持原样）；点数字 → 只动数字。
                // 落到标点/空白上（None）时退回不过滤，避免"点了逗号就什么都选不到"。
                CharClass target = CharClass.None;
                try
                {
                    Range sel = App.Selection.Range;
                    int probeStart = Math.Max(sample.Range.Start, Math.Min(sel.Start, sample.Range.End - 1));
                    string probeText = doc.Range(probeStart, Math.Min(probeStart + 1, sample.Range.End)).Text ?? "";
                    if (probeText.Length > 0) target = CharClassDetector.ClassOf(probeText[0]);
                }
                catch { }
                TypeRightAddin.Log("同类选择: 目标字符类型=" + CharClassDetector.Name(target));

                _session = new SelectSession
                {
                    Doc = doc,
                    DocFullName = SafeFullName(doc),
                    SampleParagraph = sample,
                    SampleOffset = sampleOffset,
                    SampleSignature = sig,
                    DisplaySignature = sig,
                    TargetClass = target,
                    SampleInTable = SimilarSelector.IsInTable(doc, sample),
                    Ignore = _settings.StrictByDefault ? FormatSignature.IgnoreDims.None : FormatSignature.IgnoreDims.None,
                    Highlights = new SimilarSelector.HighlightSession(doc)
                };
                long tSession = swAll.ElapsedMilliseconds;

                if (!string.Equals(_opsDocFullName, _session.DocFullName, StringComparison.OrdinalIgnoreCase))
                {
                    _docOps.Clear();
                    _opsDocFullName = _session.DocFullName;
                    _saveOffered = false;
                }

                RecomputeMatches();
                long tMatch = swAll.ElapsedMilliseconds;
                ShowFloatBar(sample);
                long tBar = swAll.ElapsedMilliseconds;
                // 保留一条汇总：pre=关旧会话 session=建会话 match=匹配(含高亮排队) bar=浮条弹出
                // 高亮已异步，所以 match 段里不再包含 500ms 染色时间
                TypeRightAddin.Log("SimilarSelect: " + _lastMatchCount + " 处 · 匹配 " + (tMatch - tSession)
                    + "ms · 浮条 " + (tBar - tMatch) + "ms · 合计 " + tBar + "ms");
            }
            catch (Exception ex)
            {
                TypeRightAddin.Log("SimilarSelect: " + ex);
                ShowToast("同类选择失败：" + ex.Message);
            }
        }

        /// <summary>读某个命中单元区间的实际格式（浮条显示用：字体 + 中文字号）。
        /// 失败返回空签名——浮条按"未知"显示 —，绝不让 0 漏到界面。</summary>
        private static FormatSignature ReadUnitDisplaySignature(Document doc, MatchUnit unit)
        {
            try
            {
                if (unit == null || doc == null) return default(FormatSignature);
                int end = unit.End > unit.Start ? unit.End : unit.Start + 1;
                return FormatSignatureReader.ReadRange(doc.Range(unit.Start, end));
            }
            catch { return default(FormatSignature); }
        }

        /// <summary>把高亮排到消息循环低优先级队列执行（不阻塞"同类选择"的可见反馈）。
        /// 与 ApplySessionHighlights 的区别只有时机：内容完全一致，仍在同一个 UndoRecord 内。
        ///
        /// 注意：COM 加载项里 System.Windows.Application.Current 通常为 null（没有 WPF Application
        /// 实例，只有 Word 的 STA 线程跑着 WPF 窗口），所以不能用 Application.Current.Dispatcher
        /// —— 那样会掉进同步回退分支，把 500ms 高亮重新塞回点击路径。用浮条自身的 Dispatcher。</summary>
        private void QueueHighlight(UnitMatchResult result)
        {
            if (_session == null || _session.Highlights == null) return;
            var sessionSnapshot = _session;

            System.Windows.Threading.Dispatcher dispatcher = null;
            try { if (_floatBar != null) dispatcher = _floatBar.Dispatcher; } catch { }
            if (dispatcher == null)
            {
                try { dispatcher = System.Windows.Threading.Dispatcher.CurrentDispatcher; } catch { }
            }
            if (dispatcher == null) { ApplySessionHighlights(result); return; }

            dispatcher.BeginInvoke(new Action(() =>
            {
                // 排队期间会话可能已被用户关掉/换掉 —— 用快照比对，不做过期高亮
                if (!ReferenceEquals(_session, sessionSnapshot)) return;
                // 会话没变但已经被新的匹配替换了命中集 → 这次排队的高亮已过期，跳过
                if (!ReferenceEquals(sessionSnapshot.MatchedUnits, result.Units)) return;
                long t0 = System.Diagnostics.Stopwatch.GetTimestamp();
                ApplySessionHighlights(result, sessionSnapshot);
                TypeRightAddin.Log(string.Format("高亮应用（延迟）: {0} 处 / 合并成 {1} 段 / {2:0}ms",
                    result.Count, sessionSnapshot.Highlights.LastRunCount,
                    (System.Diagnostics.Stopwatch.GetTimestamp() - t0) * 1000.0
                        / System.Diagnostics.Stopwatch.Frequency));
            }), System.Windows.Threading.DispatcherPriority.Background);
        }

        /// <summary>把命中单元集标为黄色高亮（格式/文本两种模式共用；一个撤销记录 + 关屏幕刷新）。
        /// explicitSession：异步调用时传快照，避免读到已被替换的 _session。</summary>
        private void ApplySessionHighlights(UnitMatchResult result, SelectSession explicitSession = null)
        {
            SelectSession session = explicitSession ?? _session;
            if (session == null || session.Highlights == null) return;
            try
            {
                var undo = WordCompat.BeginUndo(App, "TypeRight 匹配标记");
                bool screenOff = false;
                try { App.ScreenUpdating = false; screenOff = true; } catch { }
                try { session.Highlights.ApplyUnits(result, session.TargetClass); }
                finally
                {
                    WordCompat.EndUndo(undo);
                    if (screenOff) { try { App.ScreenUpdating = true; } catch { } }
                }
            }
            catch { }
        }

        /// <summary>按当前忽略维度重算命中单元集并刷新高亮/计数（引导句感知）。</summary>
        private void RecomputeMatches()
        {
            if (_session == null) return;
            var result = SimilarSelector.MatchUnits(_session.Doc, _session.SampleParagraph, _session.SampleOffset, _session.Ignore);
            _session.MatchedUnits = result.Units;
            if (!result.SampleSignature.IsEmpty)
            {
                _session.SampleSignature = result.SampleSignature;
                _session.DisplaySignature = result.SampleSignature;   // 浮条显示跟随样本真值（引导句感知）
            }
            _session.SampleIsLead = result.SampleIsLead;

            TypeRightAddin.Log(string.Format("同类匹配: {0} 处 / {1}ms（进程内） ignore={2} 样本单元={3}",
                result.Count, result.Duration.TotalMilliseconds.ToString("0"), _session.Ignore,
                result.SampleIsLead ? "引导句" : "段落/正文"));

            if (result.Count == 0)
            {
                // 样本单元自身必然命中，0 命中=异常：记录首几段状态便于现场诊断
                try
                {
                    var diagZones = SimilarSelector.MathZoneRanges(_session.Doc);
                    int diag = 0;
                    foreach (Paragraph dp in WordCompat.ParagraphsOf(_session.Doc))
                    {
                        if (diag >= 6) break;
                        FormatSignature ds = FormatSignatureReader.Read(dp, diagZones);
                        if (ds.IsEmpty) continue;
                        TypeRightAddin.Log("  [0命中诊断] 段" + diag + " sig=" + ds);
                        diag++;
                    }
                }
                catch { }
            }

            // 高亮：先弹浮条让用户马上看到反馈，高亮随后在消息循环里补上。
            // 43 处命中 = 86 次 COM 往返 ≈ 500ms（实测 4.5ms/往返），同步做会让点击发顿。
            // 把"高亮"从用户可见的关键路径上挪走：浮条立刻出现，黄标在 ~0ms 后跟上
            // （同一 UI 线程排队执行，用户看到的是"先弹条、再染色"，感知延迟从 500ms 降到 ~0）。
            _lastMatchCount = result.Count;
            _lastMatchDuration = result.Duration;
            if (_floatBar != null)
            {
                _floatBar.UpdateMatched(result.Count, result.Duration, result.SampleIsLead);
                _floatBar.UpdateSampleInfo(_session != null ? _session.DisplaySignature : result.SampleSignature);
            }

            QueueHighlight(result);

            // 命中摘要：现场诊断"为什么被判同类"（样本签名 + 命中单元签名对比）
            try
            {
                var summary = new System.Text.StringBuilder();
                for (int i = 0; i < Math.Min(6, result.Units.Count); i++)
                {
                    try { summary.Append(" | ").Append(result.Units[i].Signature); } catch { }
                }
                TypeRightAddin.Log("  [命中摘要] sample=" + result.SampleSignature + summary);
            }
            catch { }

        }

        /// <summary>浮动条：忽略维度变化（PRD：实时重算）。</summary>
        public void OnIgnoreChanged(FormatSignature.IgnoreDims ignore)
        {
            if (_session == null) return;
            _session.Ignore = ignore;
            RecomputeMatches();
        }

        /// <summary>文本同类（字义匹配，2026-11 出卷场景）：
        /// 选中文字或光标点入一个词 → 全文所有出现处成为命中单元（不看格式）→
        /// 浮条批量改格式（如着重号）。默认排除否定前缀（不/非/未），公式区内不匹配。</summary>
        public void TextSimilarSelect()
        {
            try
            {
                Document doc = ActiveDocument();
                if (doc == null) { ShowToast("请先打开一个文档"); return; }

                string pattern = GetTextPattern();
                if (string.IsNullOrEmpty(pattern))
                {
                    ShowToast("光标处没有词：请选中一段文字（如「正确」）再点，或把光标点进词内");
                    return;
                }

                CloseSession(offerSave: false);
                var units = TypeRight.Signature.TextMatcher.FindOccurrences(doc, pattern, true);
                TypeRightAddin.Log("TextSimilar: pattern=" + pattern + " units=" + units.Count);
                if (units.Count == 0) { ShowToast("全文没有找到「" + pattern + "」"); return; }

                _session = new SelectSession
                {
                    Doc = doc,
                    DocFullName = SafeFullName(doc),
                    SampleParagraph = units[0].Paragraph,
                    TextMode = true,
                    TextPattern = pattern,
                    TextExcludeNegation = true,
                    MatchedUnits = units,
                    Highlights = new SimilarSelector.HighlightSession(doc)
                };
                // 文本模式显示首个命中处的真实字体字号（不是格式签名——这里不按格式匹配）
                _session.DisplaySignature = ReadUnitDisplaySignature(doc, units[0]);
                ApplySessionHighlights(new UnitMatchResult { Units = _session.MatchedUnits });
                _lastMatchCount = units.Count;
                _lastMatchDuration = TimeSpan.FromMilliseconds(0);
                _saveOffered = false;
                ShowFloatBar(units[0].Paragraph);
                _floatBar.SetTextMode(pattern, units.Count, true);
            }
            catch (Exception ex)
            {
                TypeRightAddin.Log("TextSimilarSelect: " + ex);
            }
        }

        /// <summary>最近一次序列重排的机读摘要（验收断言用）："changes=N;runs=M"。</summary>
        public string LastRenumberSummary { get; private set; }

        /// <summary>序列一键重排：全文扫描段首序号（1. ① （1） 1） 一、 等），
        /// 只把断序的序列（缺号/重号/乱序）重排为 1..n，连续的绝不碰。
        /// 有选区时只处理选区涉及的序列；整体包一个撤销记录。</summary>
        public void RenumberSequences()
        {
            try
            {
                Document doc = ActiveDocument();
                if (doc == null) { ShowToast("请先打开一个文档"); return; }

                Range scope = null;
                try
                {
                    Selection sel = App.Selection;
                    if (sel != null && sel.Start < sel.End) scope = sel.Range;
                }
                catch { }

                var sw = System.Diagnostics.Stopwatch.StartNew();
                var plan = TypeRight.Sequences.SequenceRenumberer.Plan(doc, scope);
                TypeRightAddin.Log("序列重排: 计划 " + plan.Changes.Count + " 处 / " + plan.Summaries.Count + " 序列 / " +
                                   sw.ElapsedMilliseconds + "ms" + (scope != null ? "（选区）" : ""));
                LastRenumberSummary = "changes=" + plan.Changes.Count + ";runs=" + plan.Summaries.Count +
                                       ";skipped=" + plan.Skipped.Count + ";ms=" + sw.ElapsedMilliseconds;

                if (plan.Changes.Count == 0)
                {
                    string skipNote = plan.Skipped.Count > 0 ? "（" + string.Join("；", plan.Skipped) + "）" : "";
                    ShowToast("序号全部连续，无需重排 ✓" + skipNote, 6000);
                    return;
                }

                var undo = WordCompat.BeginUndo(App, "TypeRight 序列重排");
                try { TypeRight.Sequences.SequenceRenumberer.Apply(doc, plan); }
                finally { WordCompat.EndUndo(undo); }

                string detail = string.Join("、", plan.Summaries.ToArray());
                if (detail.Length > 60) detail = detail.Substring(0, 60) + "…";
                string skip = plan.Skipped.Count > 0 ? "（跳过: " + string.Join("；", plan.Skipped.ToArray()) + "）" : "";
                ShowToast("已重排 " + plan.Changes.Count + " 处序号：" + detail + skip + "　Ctrl+Z 可撤销", 7000);
            }
            catch (Exception ex)
            {
                TypeRightAddin.Log("RenumberSequences: " + ex);
                ShowToast("序列重排失败：" + ex.Message, 6000);
            }
        }

        /// <summary>采样文本模式：优先选区文本；否则光标处向两侧扩展连续词字符（CJK/字母/数字）。</summary>
        private string GetTextPattern()
        {
            Selection sel = App.Selection;
            if (sel == null) return null;
            try
            {
                string selText = (sel.Text ?? "").Replace("\r", "").Replace("\a", "").Trim();
                if (selText.Length >= 2 && selText.Length <= 32) return selText;
                if (selText.Length > 32) return selText.Substring(0, 32);
            }
            catch { }
            try
            {
                Range r = sel.Range;
                Paragraph p = r.Paragraphs[1];
                Range pr = p.Range;
                int pStart = pr.Start;
                string text = (pr.Text ?? "").TrimEnd('\r');
                int caret = Math.Max(0, Math.Min(text.Length, r.Start - pStart));
                if (caret >= text.Length) return null;
                int a = caret, b = caret + 1;
                while (a > 0 && IsWordChar(text[a - 1])) a--;
                while (b < text.Length && IsWordChar(text[b])) b++;
                if (b - a < 1 || b - a > 32) return b - a <= 32 ? text.Substring(a, b - a) : null;
                return text.Substring(a, b - a);
            }
            catch { return null; }
        }

        private static bool IsWordChar(char c)
        {
            return char.IsLetterOrDigit(c) && !char.IsPunctuation(c);
        }

        /// <summary>文本模式：切换"排除否定前缀"后实时重算（格式匹配的忽略维度开关在文本模式无意义）。</summary>
        public void SetTextExcludeNegation(bool exclude)
        {
            if (_session == null || !_session.TextMode) return;
            _session.TextExcludeNegation = exclude;
            RecomputeTextMatches();
        }

        private void RecomputeTextMatches()
        {
            try
            {
                var units = TypeRight.Signature.TextMatcher.FindOccurrences(_session.Doc, _session.TextPattern, _session.TextExcludeNegation);
                _session.MatchedUnits = units;
                ApplySessionHighlights(new UnitMatchResult { Units = _session.MatchedUnits });
                _lastMatchCount = units.Count;
                _floatBar.UpdateMatched(units.Count, TimeSpan.FromMilliseconds(0), false);
                if (_floatBar != null) _floatBar.SetTextMode(_session.TextPattern, units.Count, _session.TextExcludeNegation);
            }
            catch (Exception ex) { TypeRightAddin.Log("RecomputeTextMatches: " + ex); }
        }

        /// <summary>把下一处命中设为 Word 真选区（循环跳转）——最接近"真选中"的桥：
        /// Word 只支持单一连续选区，无法多重选中，跳过去之后即可用原生"开始"选项卡/Ctrl+D 调该处。
        /// 返回跳到的显示序号（1 起），无会话返回 0。</summary>
        public int JumpToNextMatch()
        {
            if (_session == null || _session.MatchedUnits.Count == 0) return 0;
            try
            {
                int cur;
                try { cur = App.Selection.Range.Start; } catch { cur = int.MinValue; }
                int next = -1;
                for (int i = 0; i < _session.MatchedUnits.Count; i++)
                {
                    int s = _session.MatchedUnits[i].Start;
                    if (s > cur) { next = i; break; }
                }
                if (next < 0) next = 0;   // 循环
                MatchUnit u = _session.MatchedUnits[next];
                int end = u.End > u.Start + 1 ? u.End - 1 : u.End;   // 排除段落标记
                App.Selection.SetRange(u.Start, end);
                // 当前处青色突出（其余保持黄）：一次跳转 = 一组撤销记录，不刷 Ctrl+Z 栈
                var undoMark = WordCompat.BeginUndo(App, "TypeRight 跳转标记");
                try { _session.Highlights.MarkCurrent(u.Start, end, _settings.EmphasisColorRgb); }
                finally { WordCompat.EndUndo(undoMark); }
                // 跳转滚动：目标段落滚到视口顶部（WPS 无此 API 时静默退化为选区定位）
                try { App.ActiveWindow.ScrollIntoView(App.Selection.Range, true); } catch { }
                return next + 1;
            }
            catch (Exception ex)
            {
                TypeRightAddin.Log("JumpToNextMatch: " + ex);
                return 0;
            }
        }

        /// <summary>浮动条：应用格式（每个取值 = 一个撤销单元）。</summary>
        public void ApplyChanges(FormatChanges changes)        {
            if (_session == null || changes == null || !changes.HasAny) return;

            // P0-3 中西文联动：设置中文字体时西文/数字自动 TNR（可在设置关闭）
            if (changes.EastAsiaFont != null && _settings.LatinLinkageEnabled && changes.LatinFont == null)
                changes.LatinFont = _settings.DefaultLatinFont;

            try
            {
                var stats = BatchFormatter.ApplyUnits(App, _session.Doc, _session.MatchedUnits, changes,
                    "同类格式（" + _session.MatchedUnits.Count + " 处）",
                    onlyClass: _session.TargetClass);

                _docOps.Add(new OpRecord
                {
                    SampleSignature = _session.SampleSignature,   // 原始签名：回放时按它匹配
                    Ignore = _session.Ignore,
                    Applied = changes,
                    SampleInTable = _session.SampleInTable
                });
                TypeRightAddin.Log("ApplyChanges: " + _session.MatchedUnits.Count + " 处，实际改动 " + stats.Changed);
            }
            catch (Exception ex)
            {
                TypeRightAddin.Log("ApplyChanges: " + ex);
                ShowToast("应用失败：" + ex.Message);
            }
            // 应用后命中集保持（PRD P0-2），不重算
        }

        /// <summary>结束会话：还原高亮、关浮动条，视情况提示存方案。</summary>
        public void CloseSession(bool offerSave = true)
        {
            if (_session != null)
            {
                try
                {
                    var undo = WordCompat.BeginUndo(App, "TypeRight 匹配标记还原");
                    try { _session.Highlights.Clear(); }
                    finally { WordCompat.EndUndo(undo); }
                }
                catch { }
                _session = null;
            }
            if (_floatBar != null) _floatBar.HideBar();

            if (offerSave && !_saveOffered && _docOps.Count >= 1)
            {
                // 门槛从「>= 2 次操作」降到 1 次：最常见的用法就是"选中同类正文 → 只调一下字号"，
                // 一次改动就该能存。旧门槛让这条最主流的路径永远等不到提示（2026-11 实测）。
                _saveOffered = true;
                string what = DescribePendingOps();
                ShowToast("刚才把" + what + "。存成方案，下次点一下就还原。", "存为方案",
                    () => OpenProfiles(saveMode: true));
            }
        }

        /// <summary>把本次待存的操作说成人话（Toast 用），例如「正文里 13 段的字号 四号 → 三号」。
        /// 用户要知道自己存的到底是什么，否则不会点那个按钮。</summary>
        private string DescribePendingOps()
        {
            try
            {
                OpRecord last = _docOps[_docOps.Count - 1];
                string where = last.SampleInTable ? "表格里" : "正文里";
                string change = ProfileDescriber.DescribeChange(
                    ToMatchSpec(last.SampleSignature, last.Ignore, last.SampleInTable),
                    ToApplySpec(last.Applied));

                string head = where + " " + _lastMatchCount + " 段的";
                if (_docOps.Count > 1)
                    return head + "格式（" + _docOps.Count + " 处修改）";
                return head + (change ?? "格式改了");
            }
            catch
            {
                return "同类段落的格式";
            }
        }

        /// <summary>关闭文档前：若活动会话在该文档上，先清高亮再放行（防止临时黄标被存盘）。</summary>
        public void ClearSessionIfDoc(Document doc)
        {
            if (_session == null || doc == null) return;
            try
            {
                if (!ReferenceEquals(_session.Doc?.UnderlyingObject, doc.UnderlyingObject) &&
                    !string.Equals(SafeFullName(doc), _session.DocFullName, StringComparison.OrdinalIgnoreCase)) return;
            }
            catch { return; }
            CloseSession(offerSave: false);
        }

        /// <summary>光标移到命中集之外 → 关闭浮动条（PRD P0-2：点击文档其他处关闭）。</summary>
        public void HandleSelectionChange(Selection selection)
        {
            if (_session == null || selection == null) return;
            try
            {
                string docName = SafeFullName(ActiveDocument());
                if (!string.Equals(docName, _session.DocFullName, StringComparison.OrdinalIgnoreCase))
                {
                    CloseSession(offerSave: false);
                    return;
                }

                // 会话内快速判定（2026-11-07 性能修复）：数百命中时旧写法每次光标移动都对全部
                // 命中做 O(n) 次活 COM 比对（u.Paragraph.Range.Start × N）——输入/点击明显卡顿的根源。
                // 改为：对缓存的单元起点（文档序天然有序）二分定位，仅对相邻 ±1 单元做活区间校验：
                // O(log n) + ≤3 次 COM 调用。活对象跟随编辑，位置漂移可容忍。
                int selStart;
                try { selStart = selection.Range.Start; } catch { return; }
                var units = _session.MatchedUnits;
                int lo = 0, hi = units.Count - 1, idx = -1;
                while (lo <= hi)
                {
                    int mid = (lo + hi) / 2;
                    if (units[mid].Start <= selStart) { idx = mid; lo = mid + 1; } else hi = mid - 1;
                }
                for (int k = Math.Max(0, idx - 1); k <= Math.Min(units.Count - 1, idx + 1); k++)
                {
                    try
                    {
                        var r = units[k].Paragraph.Range;
                        if (r.Start <= selStart && selStart < r.End) return; // 仍在命中段内
                    }
                    catch { }
                }
                CloseSession(offerSave: true);
            }
            catch { }
        }

        // ───────────────────────── P0-4 公式转换 ─────────────────────────

        public void ConvertFormulas()
        {
            try
            {
                Document doc = ActiveDocument();
                if (doc == null)
                {
                    ShowToast("请先打开一个文档");
                    return;
                }
                CloseSession(offerSave: false);

                // 三层反馈（2026-11 体验优化）：A 即刻层——等待光标 + 状态栏（Word 原生长任务语言）
                var progress = new UI.ConvertProgressWindow();
                progress.ShowAt(App);
                try { App.System.Cursor = WdCursorType.wdCursorWait; } catch { }
                try { App.StatusBar = "TypeRight：正在转换公式…（按 Esc 取消）"; } catch { }

                bool screenOff = false;
                try { App.ScreenUpdating = false; screenOff = true; } catch { }
                var undo = WordCompat.BeginUndo(App, "公式转换");
                FormulaStats stats;
                try
                {
                    var converter = new FormulaConverter(doc, null, FormulaConverter.ResolveMml2OmmlPath(App))
                    {
                        Progress = (phase, done, total) => progress.Update(phase, done, total),
                        CancelPending = IsEscapePressed
                    };
                    stats = converter.ConvertDocument();
                }
                finally
                {
                    WordCompat.EndUndo(undo);
                    if (screenOff) { try { App.ScreenUpdating = true; } catch { } }
                }

                // 还原光标/状态栏；C 完成层：进度窗就地变完成态（视线内，替代右下角 toast）
                try { App.System.Cursor = WdCursorType.wdCursorNormal; } catch { }
                try { App.StatusBar = ""; } catch { }
                progress.ShowDone(FormulaDoneMainText(stats), FormulaDoneSubText(stats));
                if (stats.Failed > 0)
                {
                    if (_failuresWindow == null) _failuresWindow = new FormulaFailuresWindow();
                    _failuresWindow.ShowResults(stats);
                }
                TypeRightAddin.Log("ConvertFormulas: " + stats.SummaryText());
                if (stats.TimingText != null) TypeRightAddin.Log("ConvertFormulas: " + stats.TimingText);
            }
            catch (Exception ex)
            {
                TypeRightAddin.Log("ConvertFormulas: " + ex);
                ShowToast("公式转换失败：" + ex.Message);
            }
        }

        // ───────────────────────── P1-1 方案 ─────────────────────────

        /// <summary>把本次文档的操作沉淀为方案。返回落盘文件的完整路径。
        /// 同名/同 id 不覆盖已有方案：自动避让成 -2、-3（要覆盖原方案请用面板里的「重命名」，
        /// 而不是再存一次）。</summary>
        public string SaveProfile(string name, bool withCatchAll, string catchAllFont, bool convertFormulas)
        {
            ProfileV2 profile = BuildPendingProfile(name, withCatchAll, catchAllFont, convertFormulas);
            if (profile == null) return null;

            var store = new ProfileStoreV2(_settings.ProfilesDirectory);
            profile.Id = ProfileId.MakeUnique(name,
                id => System.IO.File.Exists(System.IO.Path.Combine(store.Directory, id + ".json")));
            return store.Save(profile);
        }

        /// <summary>把当前待存的操作组装成 ProfileV2（不落盘）。
        /// 「另存为」窗口用它把"这次改了什么"先摆给用户看，确认后再真的保存。</summary>
        public ProfileV2 BuildPendingProfile(string name, bool withCatchAll, string catchAllFont, bool convertFormulas)
        {
            if (_docOps.Count == 0) return null;

            var profile = new ProfileV2
            {
                Id = ProfileId.Make(name),
                Name = name,
                ConvertFormulas = convertFormulas,
                Rules = new List<ProfileRule>()
            };

            int order = 1;
            var rawRules = new List<ProfileRule>();
            foreach (OpRecord op in _docOps)
            {
                rawRules.Add(new ProfileRule
                {
                    Order = order++,
                    Match = ToMatchSpec(op.SampleSignature, op.Ignore, op.SampleInTable),
                    Apply = ToApplySpec(op.Applied)
                });
            }
            // 同一次会话里连着改字体→字号→加粗，会产生若干条 match 完全相同的规则。
            // 回放是「每段只被首条命中规则处理」，不合并的话后面几条命中 0 段，用户改的字号就被吞了。
            profile.Rules = OpNormalizer.Normalize(rawRules);

            if (withCatchAll && !string.IsNullOrEmpty(catchAllFont))
            {
                profile.Rules.Add(new ProfileRule
                {
                    Order = profile.Rules.Count + 1,
                    Match = new MatchSpec { CatchAll = true },
                    Apply = new ApplySpec { EastAsiaFont = catchAllFont, LatinFont = _settings.DefaultLatinFont }
                });
            }
            return profile;
        }

        private static MatchSpec ToMatchSpec(FormatSignature sig, FormatSignature.IgnoreDims ignore, bool sampleInTable)
        {
            var spec = new MatchSpec
            {
                EastAsiaFont = sig.EastAsiaFont,
                LatinFont = sig.LatinFont,
                Ignore = new List<string>()
            };
            if (!double.IsNaN(sig.SizePt)) spec.SizePt = sig.SizePt;
            if (!sig.BoldMixed) spec.Bold = sig.Bold;
            if (sig.ColorRgb >= 0) spec.Color = "#" + sig.ColorRgb.ToString("X6");
            if ((ignore & FormatSignature.IgnoreDims.SizePt) != 0) spec.Ignore.Add("sizePt");
            if ((ignore & FormatSignature.IgnoreDims.Color) != 0) spec.Ignore.Add("color");
            if ((ignore & FormatSignature.IgnoreDims.Bold) != 0) spec.Ignore.Add("bold");
            if (spec.Ignore.Count == 0) spec.Ignore = null;
            spec.InTable = sampleInTable;
            return spec;
        }

        private static ApplySpec ToApplySpec(FormatChanges changes)
        {
            return new ApplySpec
            {
                EastAsiaFont = changes.EastAsiaFont,
                LatinFont = changes.LatinFont,
                SizePt = changes.SizePt,
                Bold = changes.Bold,
                Color = changes.ColorRgb.HasValue ? "#" + changes.ColorRgb.Value.ToString("X6") : null
            };
        }

        /// <summary>应用方案：委托 ProfileApplier（与验收测试共用同一实现）。
        ///
        /// 这里原先有一份内联的重复实现，与 ProfileApplier 已经不一致：内联版没有
        /// "已应用态识别"，导致方案对同一文档应用第二次时，末尾兜底规则会把前面已经
        /// 格式化好的段落再刷一遍——PRD 要求的「应用两次第二次 0 改动」在这条路径上不成立。
        /// 现在只保留一份真值。</summary>
        public void ApplyProfile(ProfileV2 profile)
        {
            Document doc = ActiveDocument();
            if (doc == null) { ShowToast("请先打开一个文档"); return; }
            CloseSession(offerSave: false);

            try
            {
                ProfileApplyReport result = ProfileApplier.Apply(App, doc, profile);
                // 面向用户说人话："规则1" 这类内部编号换成人话描述
                string text = BuildApplyToast(profile, result);
                ShowToast(text, 8000);
                TypeRightAddin.Log("ApplyProfile[" + profile.Name + "]: " + string.Join("; ", result.Lines));
            }
            catch (Exception ex)
            {
                TypeRightAddin.Log("ApplyProfile: " + ex);
                ShowToast("应用方案失败：" + ex.Message, 7000);
            }
        }

        /// <summary>应用结果 → 人话提示（"规则1命中 8 段" 用户看不懂，换成方案卡片同一套描述）。</summary>
        private static string BuildApplyToast(ProfileV2 profile, ProfileApplyReport report)
        {
            var parts = new List<string>();
            ProfileSummary summary = ProfileDescriber.Describe(profile);
            for (int i = 0; i < summary.Rules.Count; i++)
            {
                RuleSummary rs = summary.Rules[i];
                // report.Lines 与规则一一对应（公式行在末尾，不参与）
                if (i < report.Lines.Count)
                {
                    string line = report.Lines[i];
                    int hit = ExtractIntAfter(line, "命中 ");
                    string who = rs.IsCatchAll ? "其余段落" : rs.MatchText;
                    parts.Add(who + "：" + (rs.ChangeText ?? "格式已统一")
                        + (hit >= 0 ? "（" + hit + " 段）" : ""));
                }
            }
            if (profile.ConvertFormulas)
                parts.Add("公式 " + report.FormulaConverted + " 处");

            string head = "已应用「" + profile.Name + "」：";
            string body = string.Join("；", parts.ToArray());
            string tail = report.Unmatched > 0 ? "　另有 " + report.Unmatched + " 段未匹配任何规则，未改动。" : "";
            return head + body + tail;
        }

        /// <summary>从报告行里抠出 "命中 N 段" 的 N；取不到返回 -1。</summary>
        private static int ExtractIntAfter(string line, string marker)
        {
            if (string.IsNullOrEmpty(line)) return -1;
            int at = line.IndexOf(marker, StringComparison.Ordinal);
            if (at < 0) return -1;
            int i = at + marker.Length;
            int start = i;
            while (i < line.Length && char.IsDigit(line[i])) i++;
            if (i == start) return -1;
            int v;
            return int.TryParse(line.Substring(start, i - start), out v) ? v : -1;
        }

        private static int? ParseColor(string hex)
        {
            var v = WordFormatUtils.ParseHexColor(hex);
            return v;
        }

        // ───────────────────────── 界面 ─────────────────────────

        private void ShowFloatBar(Paragraph sample)
        {
            var sw = System.Diagnostics.Stopwatch.StartNew();
            if (_floatBar == null)
            {
                _floatBar = new FloatBar(this);
                _floatBar.SetFonts(FontListProvider.GetFontList());
            }
            long tCtor = sw.ElapsedMilliseconds;
            _addin.RefreshRibbonTheme();   // 宿主主题变化 → 刷新功能区图标配色
            long tTheme = sw.ElapsedMilliseconds;
            _floatBar.PositionNear(App, sample);
            long tPos = sw.ElapsedMilliseconds;
            _floatBar.ShowBar();
            long tShow = sw.ElapsedMilliseconds;
            TypeRightAddin.Log("ShowFloatBar 分段: ctor=" + tCtor + " theme=" + (tTheme - tCtor)
                + " pos=" + (tPos - tTheme) + " show=" + (tShow - tPos) + " total=" + tShow + "ms");
            // 首次会话浮条是刚创建的（RecomputeMatches 时还不存在），用会话里的真值补刷计数
            if (_session != null)
            {
                _floatBar.UpdateMatched(_lastMatchCount, _lastMatchDuration, _session != null && _session.SampleIsLead);
                _floatBar.UpdateSampleInfo(_session.DisplaySignature);
                if (_session.TextMode)
                    _floatBar.SetTextMode(_session.TextPattern, _lastMatchCount, _session.TextExcludeNegation);
            }
            TypeRightAddin.Log("ShowFloatBar 全部完成: " + sw.ElapsedMilliseconds + "ms");
        }

        /// <summary>空闲时预建浮条（隐藏）：把 WPF 窗口初始化 + 字体列表装载
        /// 从"同类选择"首次点击挪到 Word 启动后的空闲期。</summary>
        public void PrewarmFloatBar()
        {
            if (_floatBar != null) return;
            try
            {
                // ⚠️ 不要在这里调 SetThreadDpiAwarenessContext（曾经加过，已移除）。
                // 实测后果：Word 的 UI 线程被切到 PerMonitorV2，而浮条窗口是在进程级
                // System-DPI aware 上下文里创建的 → 窗口的**视觉位置**与**鼠标命中区**错位。
                // 用户表现：看到 ✕ 在那里，正正点击没反应；必须先让鼠标划过让 WPF 产生
                // hover 高亮（说明鼠标其实进了另一个位置），才能点中。
                // 该 API 对 .NET Framework 4.8 的 WPF 本来就不生效（窗口尺寸不变），
                // 唯一效果就是制造上面这个错位——所以彻底不用。
                _floatBar = new FloatBar(this);
                _floatBar.SetFonts(FontListProvider.GetFontList());
                TypeRightAddin.Log("PrewarmFloatBar: ok");
            }
            catch (Exception ex)
            {
                TypeRightAddin.Log("PrewarmFloatBar: " + ex.Message);
                try { if (_floatBar != null) { _floatBar.Close(); _floatBar = null; } } catch { }
            }
        }

        /// <summary>自动化验证：模拟在浮条字号框手输（返回显示快照）。</summary>
        public string TypeSizeInFloatBar(string text)
        {
            if (_floatBar == null) return "(no bar)";
            try { return _floatBar.AutomationTypeSize(text); }
            catch (Exception ex) { return "(error: " + ex.Message + ")"; }
        }

        /// <summary>自动化验证：模拟在浮条字体框手输（返回显示快照）。</summary>
        public string TypeFontInFloatBar(string text)
        {
            if (_floatBar == null) return "(no bar)";
            try { return _floatBar.AutomationTypeFont(text); }
            catch (Exception ex) { return "(error: " + ex.Message + ")"; }
        }

        /// <summary>自动化验证：✕ 按钮的物理坐标（命中区标定用）。</summary>
        public string GetCloseButtonPhysical()
        {
            if (_floatBar == null) return "(no bar)";
            try { return _floatBar.ProbeCloseButtonPhysical(); }
            catch (Exception ex) { return "(error: " + ex.Message + ")"; }
        }

        /// <summary>自动化验证：标定 WPF Left/Top ↔ 物理矩形的映射（坐标系诊断用）。</summary>
        public string ProbeFloatBarCoords(double left, double top)
        {
            if (_floatBar == null) return "(no bar)";
            try { return _floatBar.ProbeCoordinateMapping(left, top); }
            catch (Exception ex) { return "(error: " + ex.Message + ")"; }
        }

        /// <summary>自动化验证：读浮条关闭按钮的屏幕坐标（真机点击测试用）。</summary>
        public string GetFloatBarCloseProbe()
        {
            if (_floatBar == null) return "(no bar)";
            try { return _floatBar.GetCloseButtonProbe(); }
            catch (Exception ex) { return "(error: " + ex.Message + ")"; }
        }

        /// <summary>自动化验证：读回悬浮条当前显示的字体/字号文本（"宋体|四号"）。</summary>
        public string GetFloatBarDisplaySnapshot()
        {
            if (_floatBar == null) return "(no bar)";
            try { return _floatBar.GetDisplaySnapshot(); }
            catch (Exception ex) { return "(error: " + ex.Message + ")"; }
        }

        /// <summary>自动化验证：ShowActivated=false 重显悬浮条（后台无前台权时不触发失焦隐藏）。</summary>
        public void ShowFloatBarForAutomation()
        {
            if (_floatBar == null) return;
            _floatBar.ShowActivated = false;
            _floatBar.ShowBar();
        }

        /// <summary>自动化验证：强制悬浮条主题（"dark"/"darkgray"/"light"，其他=自动检测）。仅 UI 显示，不触碰文档。</summary>
        public void ForceFloatBarTheme(string theme)
        {
            if (_floatBar == null) return;
            OfficeThemeKind kind;
            switch ((theme ?? "").Trim().ToLowerInvariant())
            {
                case "dark": kind = OfficeThemeKind.Dark; break;
                case "darkgray": kind = OfficeThemeKind.DarkGray; break;
                case "light": kind = OfficeThemeKind.Light; break;
                default: kind = OfficeTheme.Detect(App); break;
            }
            // 主题只动 UI 资源；调用线程未知时经浮条自身 Dispatcher 封送
            var bar = _floatBar;
            if (bar.Dispatcher.CheckAccess()) bar.ApplyTheme(kind);
            else bar.Dispatcher.BeginInvoke(new Action(() => bar.ApplyTheme(kind)));
        }

        public void ShowToast(string message, int durationMs = 5000)
        {
            ShowToast(message, null, null, durationMs);
        }

        public void ShowToast(string message, string actionLabel, Action action, int durationMs = 5000)
        {
            if (_toast == null) _toast = new ToastWindow();
            _toast.Show(App, message, actionLabel, action, durationMs);
        }

        /// <summary>打开方案窗口。saveMode=true → 只负责"存刚才那次改动"；
        /// false → 只负责"挑一个用到当前文档"。两个语境分窗，不再混在一屏。</summary>
        public void OpenProfiles(bool saveMode = false)
        {
            try
            {
                if (saveMode)
                {
                    if (_docOps.Count == 0)
                    {
                        ShowToast("还没有可保存的排版操作。先用「同类选择」调整一处格式再回来。", 6000);
                        return;
                    }
                    new SaveProfileWindow(this).Show();
                }
                else
                {
                    new ProfileManagerWindow(this).Show();
                }
            }
            catch (Exception ex)
            {
                TypeRightAddin.Log("OpenProfiles: " + ex);
                ShowToast("打开方案窗口失败：" + ex.Message);
            }
        }

        public void OpenSettings()
        {
            try
            {
                var win = new SettingsWindow(SettingsStore.Load());
                win.ShowDialog();
                _settings = SettingsStore.Load();
            }
            catch (Exception ex)
            {
                ShowToast("打开设置失败：" + ex.Message);
            }
        }

        /// <summary>使用说明窗（三步用法 + 诊断日志 + 版本）。</summary>
        public void ShowHelp()
        {
            try
            {
                var win = new HelpWindow(this);
                win.Show();
                TypeRightAddin.Log("ShowHelp: v" + typeof(HelpWindow).Assembly.GetName().Version.ToString(3));
            }
            catch (Exception ex)
            {
                TypeRightAddin.Log("ShowHelp: " + ex);
                ShowToast("打开使用说明失败：" + ex.Message);
            }
        }

        // ───────────────────────── 填表与版面防崩套件 ─────────────────────────

        /// <summary>📸 证件照一键严丝合缝入格（支持剪贴板/选文件/已粘贴图片转悬浮）。</summary>
        public void FitPhotoToCell(string forceFilePath = null)
        {
            try
            {
                CloseSession(false);
                var sw = System.Diagnostics.Stopwatch.StartNew();
                PhotoFitResult res = PhotoCellFitter.FitPhoto(App, forceFilePath);
                sw.Stop();
                TypeRightAddin.Log("FitPhotoToCell: elapsed=" + sw.ElapsedMilliseconds + "ms (" + res.TimingText + ") success=" + res.Success + " clip=" + res.FromClipboard + " msg=" + res.Message);
                if (string.IsNullOrEmpty(res.Message)) return;

                if (res.Success && res.FromClipboard)
                {
                    // 从剪贴板直接贴入后，提供“换选文件…”按钮，方便用户改从磁盘挑选照片替换
                    ShowToast(res.Message, "换选文件…", () =>
                    {
                        string path = PhotoCellFitter.PromptSelectImageFile();
                        if (!string.IsNullOrEmpty(path))
                            FitPhotoToCell(path);
                    }, 7000);
                }
                else
                {
                    ShowToast(res.Message, 6000);
                }
            }
            catch (Exception ex)
            {
                TypeRightAddin.Log("FitPhotoToCell 异常: " + ex);
                ShowToast("照片入格出错：" + ex.Message);
            }
        }

        /// <summary>📦 单元格铺设 PPT 式透明悬浮填表框（自动避让底部签字/盖章/日期落款）。</summary>
        public void PlaceCellTextBox()
        {
            try
            {
                CloseSession(false);
                var sw = System.Diagnostics.Stopwatch.StartNew();
                CellTextBoxResult res = CellTextBoxPlacer.PlaceInCurrentCell(App);
                sw.Stop();
                TypeRightAddin.Log("PlaceCellTextBox: elapsed=" + sw.ElapsedMilliseconds + "ms (" + res.TimingText + ") success=" + res.Success + " sig=" + res.ReservedSignature + " msg=" + res.Message);
                if (!string.IsNullOrEmpty(res.Message))
                    ShowToast(res.Message, 6000);
            }
            catch (Exception ex)
            {
                TypeRightAddin.Log("PlaceCellTextBox 异常: " + ex);
                ShowToast("铺设悬浮填表框出错：" + ex.Message);
            }
        }

        /// <summary>📏 锁死表格行高列宽 + 溢出智能缩字 + 死守一页。</summary>
        public void LockAndGuardTable()
        {
            try
            {
                CloseSession(false);
                TableGuardResult res = TableGuard.LockAndGuardTable(App);
                TypeRightAddin.Log("LockAndGuardTable: success=" + res.Success + " rows=" + res.LockedRows + " shrunk=" + res.ShrunkCells + " msg=" + res.Message);
                if (!string.IsNullOrEmpty(res.Message))
                    ShowToast(res.Message, 6500);
            }
            catch (Exception ex)
            {
                TypeRightAddin.Log("LockAndGuardTable 异常: " + ex);
                ShowToast("锁表防崩出错：" + ex.Message);
            }
        }

        /// <summary>✍️ 表头下划线定宽零位移填空/自动吞空格对齐。</summary>
        public void FillOrBalanceUnderline()
        {
            try
            {
                CloseSession(false);
                UnderlineFillResult res = UnderlineFiller.FillOrBalance(App, UnderlinePromptWindow.Prompt);
                TypeRightAddin.Log("FillOrBalanceUnderline: success=" + res.Success + " count=" + res.BalancedCount + " msg=" + res.Message);
                if (!string.IsNullOrEmpty(res.Message))
                    ShowToast(res.Message, 6000);
            }
            catch (Exception ex)
            {
                TypeRightAddin.Log("FillOrBalanceUnderline 异常: " + ex);
                ShowToast("下划线填空出错：" + ex.Message);
            }
        }

        public int PendingOpCount { get { return _docOps.Count; } }

        // ───────────────────────── 基础 ─────────────────────────

        private Document ActiveDocument()
        {
            try { return App.ActiveDocument; }
            catch { return null; }
        }

        private static string SafeFullName(Document doc)
        {
            try { return doc.FullName; }
            catch { return null; }
        }

        public void Shutdown()
        {
            try { if (_floatBar != null) _floatBar.Close(); } catch { }
            try { if (_toast != null) _toast.Close(); } catch { }
            try { if (_failuresWindow != null) _failuresWindow.Close(); } catch { }
        }
    }
}
