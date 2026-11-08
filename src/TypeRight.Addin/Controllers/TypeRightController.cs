using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using NetOffice.WordApi;
using NetOffice.WordApi.Enums;
using TypeRight.Addin.UI;
using TypeRight.Formatting;
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
        /// <summary>样本是否引导句单元。</summary>
        public bool SampleIsLead;
        /// <summary>样本段落是否在表格内（正文/表格分区，按表格区间位置判定）。</summary>
        public bool SampleInTable;
        public FormatSignature.IgnoreDims Ignore;
        /// <summary>命中单元（引导句感知）：高亮与格式应用都作用于单元区间。</summary>
        public List<MatchUnit> MatchedUnits = new List<MatchUnit>();
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

                _session = new SelectSession
                {
                    Doc = doc,
                    DocFullName = SafeFullName(doc),
                    SampleParagraph = sample,
                    SampleOffset = sampleOffset,
                    SampleSignature = sig,
                    SampleInTable = SimilarSelector.IsInTable(doc, sample),
                    Ignore = _settings.StrictByDefault ? FormatSignature.IgnoreDims.None : FormatSignature.IgnoreDims.None,
                    Highlights = new SimilarSelector.HighlightSession(doc)
                };

                if (!string.Equals(_opsDocFullName, _session.DocFullName, StringComparison.OrdinalIgnoreCase))
                {
                    _docOps.Clear();
                    _opsDocFullName = _session.DocFullName;
                    _saveOffered = false;
                }

                RecomputeMatches();
                ShowFloatBar(sample);
            }
            catch (Exception ex)
            {
                TypeRightAddin.Log("SimilarSelect: " + ex);
                ShowToast("同类选择失败：" + ex.Message);
            }
        }

        /// <summary>把命中单元集标为黄色高亮（格式/文本两种模式共用；一个撤销记录 + 关屏幕刷新）。</summary>
        private void ApplySessionHighlights(UnitMatchResult result)
        {
            if (_session == null || _session.Highlights == null) return;
            try
            {
                var undo = WordCompat.BeginUndo(App, "TypeRight 匹配标记");
                bool screenOff = false;
                try { App.ScreenUpdating = false; screenOff = true; } catch { }
                try { _session.Highlights.ApplyUnits(result); }
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
            if (!result.SampleSignature.IsEmpty) _session.SampleSignature = result.SampleSignature;
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

            // 高亮包成一个撤销记录，避免刷屏 Ctrl+Z 栈；批量写期间关屏幕刷新防闪烁卡顿
            ApplySessionHighlights(result);

            _lastMatchCount = result.Count;
            _lastMatchDuration = result.Duration;
            if (_floatBar != null)
            {
                _floatBar.UpdateMatched(result.Count, result.Duration, result.SampleIsLead);
                _floatBar.UpdateSampleInfo(result.SampleSignature);
            }

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
                    "同类格式（" + _session.MatchedUnits.Count + " 处）");

                _docOps.Add(new OpRecord
                {
                    SampleSignature = _session.SampleSignature,   // 原始签名：回放时按它匹配
                    Ignore = _session.Ignore,
                    Applied = changes
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

            if (offerSave && !_saveOffered && _docOps.Count >= 2)
            {
                _saveOffered = true;
                ShowToast("这次操作可以存为方案，下份同类文档一键完成。", "存为方案", () => OpenProfiles(saveMode: true));
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

        /// <summary>把本次文档的操作沉淀为方案。</summary>
        public string SaveProfile(string name, bool withCatchAll, string catchAllFont, bool convertFormulas)
        {
            if (_docOps.Count == 0) return null;
            var profile = new ProfileV2
            {
                Id = MakeId(name),
                Name = name,
                ConvertFormulas = convertFormulas,
                Rules = new List<ProfileRule>()
            };

            int order = 1;
            foreach (OpRecord op in _docOps)
            {
                profile.Rules.Add(new ProfileRule
                {
                    Order = order++,
                    Match = ToMatchSpec(op.SampleSignature, op.Ignore),
                    Apply = ToApplySpec(op.Applied)
                });
            }
            if (withCatchAll && !string.IsNullOrEmpty(catchAllFont))
            {
                profile.Rules.Add(new ProfileRule
                {
                    Order = order++,
                    Match = new MatchSpec { CatchAll = true },
                    Apply = new ApplySpec { EastAsiaFont = catchAllFont, LatinFont = _settings.DefaultLatinFont }
                });
            }

            var store = new ProfileStoreV2(_settings.ProfilesDirectory);
            store.Save(profile);
            return store.Directory;
        }

        private static MatchSpec ToMatchSpec(FormatSignature sig, FormatSignature.IgnoreDims ignore)
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

        /// <summary>应用方案：按序执行规则（每段只被首条命中规则处理）+ 兜底 + 公式，整程一个撤销单元。</summary>
        public void ApplyProfile(ProfileV2 profile)
        {
            Document doc = ActiveDocument();
            if (doc == null) { ShowToast("请先打开一个文档"); return; }
            CloseSession(offerSave: false);

            var report = new List<string>();
            var consumed = new HashSet<int>();
            List<int[]> mathZones = SimilarSelector.MathZoneRanges(doc);   // 公式区感知（公式化过的文档方案回放不再错配）

            bool screenOff = false;
            try { App.ScreenUpdating = false; screenOff = true; } catch { }
            var undo = WordCompat.BeginUndo(App, "应用方案「" + profile.Name + "」");
            try
            {
                foreach (ProfileRule rule in profile.Rules.OrderBy(r => r.Order))
                {
                    var hit = new List<Paragraph>();
                    Paragraphs paragraphs = doc.Paragraphs;
                    int count = paragraphs.Count;
                    for (int i = 1; i <= count; i++)
                    {
                        Paragraph p = null;
                        try { p = paragraphs[i]; } catch { continue; }
                        if (p == null) continue;
                        int key;
                        try { key = p.Range.Start; } catch { continue; }
                        if (consumed.Contains(key)) continue;
                        var sig = FormatSignatureReader.Read(p, mathZones);
                        if (sig.IsEmpty) continue;
                        if (rule.Match.MatchesPara(sig)) hit.Add(p);
                    }

                    if (hit.Count > 0)
                    {
                        foreach (Paragraph p in hit)
                        {
                            try { consumed.Add(p.Range.Start); } catch { }
                        }
                        var changes = new FormatChanges
                        {
                            EastAsiaFont = rule.Apply.EastAsiaFont,
                            LatinFont = rule.Apply.LatinFont,
                            SizePt = rule.Apply.SizePt,
                            Bold = rule.Apply.Bold,
                            ColorRgb = ParseColor(rule.Apply.Color)
                        };
                        BatchFormatter.Apply(App, doc, hit, changes, null, wrapUndo: false);
                    }
                    report.Add(string.Format("规则{0}{1}命中 {2} 段", rule.Order,
                        rule.Match.CatchAll == true ? "（其余段落）" : " ", hit.Count));
                }

                if (profile.ConvertFormulas)
                {
                    var converter = new FormulaConverter(doc, null, FormulaConverter.ResolveMml2OmmlPath(App));
                    FormulaStats stats = converter.ConvertDocument();
                    report.Add("公式 " + stats.Converted + " 处" + (stats.Failed > 0 ? "（失败 " + stats.Failed + " 处已保留）" : ""));
                }
            }
            finally
            {
                WordCompat.EndUndo(undo);
                if (screenOff) { try { App.ScreenUpdating = true; } catch { } }
            }

            int unmatched = 0;
            try
            {
                foreach (Paragraph p in WordCompat.ParagraphsOf(doc))
                {
                    try
                    {
                        var sig = FormatSignatureReader.Read(p, mathZones);
                        if (!sig.IsEmpty && !consumed.Contains(p.Range.Start)) unmatched++;
                    }
                    catch { }
                }
            }
            catch { }
            report.Add("未命中 " + unmatched + " 段（未改动）。");

            ShowToast(string.Join(" ｜ ", report), 8000);
            TypeRightAddin.Log("ApplyProfile[" + profile.Name + "]: " + string.Join("; ", report));
        }

        private static int? ParseColor(string hex)
        {
            var v = WordFormatUtils.ParseHexColor(hex);
            return v;
        }

        // ───────────────────────── 界面 ─────────────────────────

        private void ShowFloatBar(Paragraph sample)
        {
            if (_floatBar == null)
            {
                _floatBar = new FloatBar(this);
                _floatBar.SetFonts(FontListProvider.GetFontList());
            }
            _addin.RefreshRibbonTheme();   // 宿主主题变化 → 刷新功能区图标配色
            _floatBar.PositionNear(App, sample);
            _floatBar.ShowBar();
            // 首次会话浮条是刚创建的（RecomputeMatches 时还不存在），用会话里的真值补刷计数
            if (_session != null)
            {
                _floatBar.UpdateMatched(_lastMatchCount, _lastMatchDuration, _session != null && _session.SampleIsLead);
                _floatBar.UpdateSampleInfo(_session.SampleSignature);
            }
        }

        /// <summary>空闲时预建浮条（隐藏）：把 WPF 窗口初始化 + 字体列表装载
        /// 从"同类选择"首次点击挪到 Word 启动后的空闲期。</summary>
        public void PrewarmFloatBar()
        {
            if (_floatBar != null) return;
            try
            {
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

        public void OpenProfiles(bool saveMode = false)
        {
            var win = new ProfilesWindow(this, saveMode);
            win.Show();
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

        private static string MakeId(string name)
        {
            var sb = new System.Text.StringBuilder();
            bool lastDash = false;
            foreach (char c in (name ?? "profile").Trim())
            {
                if (char.IsLetterOrDigit(c) || c == '-') { sb.Append(c); lastDash = false; }
                else if (!lastDash) { sb.Append('-'); lastDash = true; }
            }
            string id = sb.ToString().Trim('-').ToLowerInvariant();
            return id.Length > 0 ? id : "profile";
        }

        public void Shutdown()
        {
            try { if (_floatBar != null) _floatBar.Close(); } catch { }
            try { if (_toast != null) _toast.Close(); } catch { }
            try { if (_failuresWindow != null) _failuresWindow.Close(); } catch { }
        }
    }
}
