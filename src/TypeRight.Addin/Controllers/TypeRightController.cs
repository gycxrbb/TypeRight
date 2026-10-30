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
        /// <summary>样本段落的原始签名（应用格式前采集，方案规则用它匹配）。</summary>
        public FormatSignature SampleSignature;
        /// <summary>样本段落是否在表格内（正文/表格分区，按表格区间位置判定）。</summary>
        public bool SampleInTable;
        /// <summary>最近一次匹配计数/耗时（浮条首显即用真值，不等下一次重算）。</summary>
        public int LastMatchCount;
        public TimeSpan LastMatchDuration;
        public FormatSignature.IgnoreDims Ignore;
        public List<Paragraph> Matched = new List<Paragraph>();
        public SimilarSelector.HighlightSession Highlights;
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

                var sig = FormatSignatureReader.Read(sample);
                TypeRightAddin.Log("SimilarSelect: sample sig = " + sig);
                if (sig.IsEmpty)
                {
                    ShowToast("当前段落没有可见文本，换一个段落试试");
                    return;
                }

                CloseSession();

                _session = new SelectSession
                {
                    Doc = doc,
                    DocFullName = SafeFullName(doc),
                    SampleParagraph = sample,
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

        /// <summary>按当前忽略维度重算命中集并刷新高亮/计数。</summary>
        private void RecomputeMatches()
        {
            if (_session == null) return;
            var result = SimilarSelector.Match(_session.Doc, _session.SampleSignature, _session.Ignore, _session.SampleInTable);
            if (result.Count == 0)
            {
                // 样本段自身必然命中（同签名同分区），0 命中=异常：记录首几段状态便于现场诊断
                try
                {
                    int diag = 0;
                    foreach (Paragraph dp in _session.Doc.Paragraphs)
                    {
                        if (diag >= 6) break;
                        FormatSignature ds = FormatSignatureReader.Read(dp);
                        if (ds.IsEmpty) continue;
                        TypeRightAddin.Log("  [0命中诊断] 段" + diag + " sig=" + ds +
                            " 匹配=" + ds.Matches(_session.SampleSignature, _session.Ignore));
                        diag++;
                    }
                }
                catch { }
            }
            _session.Matched = result.Paragraphs;
            long ticksPerMs = System.Diagnostics.Stopwatch.Frequency / 1000;
            TypeRightAddin.Log(string.Format(
                "同类匹配: {0} 段 / {1}ms（进程内） ignore={2} | 分解: 枚举{3}ms 区间{4}ms 文本{5}ms 探针{6}ms 字体{7}ms 对齐{8}ms",
                result.Count, result.Duration.TotalMilliseconds.ToString("0"), _session.Ignore,
                FormatSignatureReader.TicksEnumerator / ticksPerMs, FormatSignatureReader.TicksRange / ticksPerMs,
                FormatSignatureReader.TicksText / ticksPerMs, FormatSignatureReader.TicksProbe / ticksPerMs,
                FormatSignatureReader.TicksFont / ticksPerMs, FormatSignatureReader.TicksAlign / ticksPerMs));

            // 高亮包成一个撤销记录，避免刷屏 Ctrl+Z 栈；批量写期间关屏幕刷新防闪烁卡顿
            try
            {
                var undo = App.UndoRecord;
                undo.StartCustomRecord("TypeRight 匹配标记");
                bool screenOff = false;
                try { App.ScreenUpdating = false; screenOff = true; } catch { }
                try { _session.Highlights.Apply(result); }
                finally
                {
                    undo.EndCustomRecord();
                    if (screenOff) { try { App.ScreenUpdating = true; } catch { } }
                }
            }
            catch { }

            _session.LastMatchCount = result.Count;
            _session.LastMatchDuration = result.Duration;
            if (_floatBar != null)
                _floatBar.UpdateMatched(result.Count, result.Duration);
        }

        /// <summary>浮动条：忽略维度变化（PRD：实时重算）。</summary>
        public void OnIgnoreChanged(FormatSignature.IgnoreDims ignore)
        {
            if (_session == null) return;
            _session.Ignore = ignore;
            RecomputeMatches();
        }

        /// <summary>浮动条：应用格式（每个取值 = 一个撤销单元）。</summary>
        public void ApplyChanges(FormatChanges changes)
        {
            if (_session == null || changes == null || !changes.HasAny) return;

            // P0-3 中西文联动：设置中文字体时西文/数字自动 TNR（可在设置关闭）
            if (changes.EastAsiaFont != null && _settings.LatinLinkageEnabled && changes.LatinFont == null)
                changes.LatinFont = _settings.DefaultLatinFont;

            try
            {
                var stats = BatchFormatter.Apply(App, _session.Doc, _session.Matched, changes,
                    "同类格式（" + _session.Matched.Count + " 段）");

                _docOps.Add(new OpRecord
                {
                    SampleSignature = _session.SampleSignature,   // 原始签名：回放时按它匹配
                    Ignore = _session.Ignore,
                    Applied = changes
                });
                TypeRightAddin.Log("ApplyChanges: " + _session.Matched.Count + " 段，实际改动 " + stats.Changed);
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
                    var undo = App.UndoRecord;
                    undo.StartCustomRecord("TypeRight 匹配标记还原");
                    try { _session.Highlights.Clear(); }
                    finally { undo.EndCustomRecord(); }
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

                Paragraph p = null;
                try { p = selection.Range.Paragraphs[1]; }
                catch { }
                if (p == null) return;
                int start = p.Range.Start;
                foreach (Paragraph m in _session.Matched)
                {
                    try
                    {
                        if (m.Range.Start == start) return; // 仍在命中集内
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

                bool screenOff = false;
                try { App.ScreenUpdating = false; screenOff = true; } catch { }
                var undo = App.UndoRecord;
                undo.StartCustomRecord("公式转换");
                FormulaStats stats;
                try
                {
                    var converter = new FormulaConverter(doc, null, FormulaConverter.ResolveMml2OmmlPath(App));
                    stats = converter.ConvertDocument();
                }
                finally
                {
                    try { undo.EndCustomRecord(); } catch { }
                    if (screenOff) { try { App.ScreenUpdating = true; } catch { } }
                }

                ShowToast(stats.SummaryText());
                if (stats.Failed > 0)
                {
                    if (_failuresWindow == null) _failuresWindow = new FormulaFailuresWindow();
                    _failuresWindow.ShowResults(stats);
                }
                TypeRightAddin.Log("ConvertFormulas: " + stats.SummaryText());
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

            bool screenOff = false;
            try { App.ScreenUpdating = false; screenOff = true; } catch { }
            var undo = App.UndoRecord;
            undo.StartCustomRecord("应用方案「" + profile.Name + "」");
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
                        var sig = FormatSignatureReader.Read(p);
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
                try { undo.EndCustomRecord(); } catch { }
                if (screenOff) { try { App.ScreenUpdating = true; } catch { } }
            }

            int unmatched = 0;
            try
            {
                foreach (Paragraph p in doc.Paragraphs)
                {
                    try
                    {
                        var sig = FormatSignatureReader.Read(p);
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
                _floatBar.UpdateMatched(_session.LastMatchCount, _session.LastMatchDuration);
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
