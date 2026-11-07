using System;
using System.Collections.Generic;
using System.Linq;
using NetOffice.WordApi;
using TypeRight.Formatting;
using TypeRight.Signature;

namespace TypeRight.Profiles
{
    /// <summary>方案应用报告（PRD P1-1：按规则命中数统计）。</summary>
    public sealed class ProfileApplyReport
    {
        public List<string> Lines = new List<string>();
        public int Unmatched;
        public int FormulaConverted;
        public int FormulaFailed;
        public int TotalChanged;
    }

    /// <summary>
    /// 方案执行器：按 order 顺序逐规则匹配（每段只被首条命中规则处理）→ 兜底 → 公式转换，
    /// 整程一个撤销单元。控制器与验收测试共用。
    /// </summary>
    public static class ProfileApplier
    {
        public sealed class Options
        {
            public Action<int, int> Progress;   // (ruleOrder, hitCount)
        }

        private sealed class ParaEntry
        {
            public Paragraph Paragraph;
            public FormatSignature Sig;
            public bool InTable;
            public int Start;
        }


        /// <summary>段落是否已处于规则的 apply 结果状态（字体相等，字号/加粗如指定也相等）。</summary>
        private static bool IsAlreadyApplied(FormatSignature sig, ApplySpec apply)
        {
            if (apply == null) return false;
            if (apply.EastAsiaFont != null && !string.Equals(sig.EastAsiaFont, apply.EastAsiaFont, StringComparison.OrdinalIgnoreCase)) return false;
            if (apply.LatinFont != null && !string.Equals(sig.LatinFont, apply.LatinFont, StringComparison.OrdinalIgnoreCase)) return false;
            if (apply.SizePt.HasValue && (double.IsNaN(sig.SizePt) || Math.Abs(sig.SizePt - apply.SizePt.Value) >= 0.01)) return false;
            if (apply.Bold.HasValue && sig.Bold != apply.Bold.Value) return false;
            if (apply.EastAsiaFont == null && apply.LatinFont == null && !apply.SizePt.HasValue && !apply.Bold.HasValue) return false;
            return true;
        }

        private static List<ParaEntry> CollectParas(Document doc, HashSet<int> consumed)
        {
            var list = new List<ParaEntry>();
            var mathZones = TypeRight.Signature.SimilarSelector.MathZoneRanges(doc);   // 公式区感知：公式化过的文档回放不再错配
            foreach (Paragraph p in WordCompat.ParagraphsOf(doc))
            {
                try
                {
                    var e = new ParaEntry { Paragraph = p };
                    e.Start = p.Range.Start;
                    e.InTable = Convert.ToBoolean(p.Range.get_Information(NetOffice.WordApi.Enums.WdInformation.wdWithInTable));
                    e.Sig = FormatSignatureReader.Read(p, mathZones);
                    if (!e.Sig.IsEmpty) list.Add(e);
                }
                catch { }
            }
            return list;
        }

        public static ProfileApplyReport Apply(NetOffice.WordApi.Application app, Document doc, ProfileV2 profile,
            double? mathSizePt = null, Options options = null)
        {
            if (doc == null) throw new ArgumentNullException("doc");
            if (profile == null) throw new ArgumentNullException("profile");

            var report = new ProfileApplyReport();
            var consumed = new HashSet<int>();

            bool screenOff = false;
            try { app.ScreenUpdating = false; screenOff = true; } catch { }
            var undo = WordCompat.BeginUndo(app, "应用方案「" + profile.Name + "」");
            try
            {
                // 单遍收集段落（签名 + 表格分区），规则在内存中匹配（性能：避免逐规则全文档扫描）
                var allParas = CollectParas(doc, consumed);

                foreach (ProfileRule rule in profile.Rules.OrderBy(r => r.Order))
                {
                    var hit = new List<Paragraph>();
                    foreach (var entry in allParas)
                    {
                        try
                        {
                            if (consumed.Contains(entry.Start)) continue;
                            if (rule.Match.CatchAll != true && rule.Match.InTable.HasValue && rule.Match.InTable.Value != entry.InTable) continue;
                            // 已应用态识别（幂等关键）：段落当前字体已等于本规则 apply 目标 → 视为本规则属地，
                            // 防止后续 catchAll 把规则已格式化的段落（楷体/黑体）误刷为兜底字体
                            if (rule.Match.MatchesPara(entry.Sig) || IsAlreadyApplied(entry.Sig, rule.Apply)) hit.Add(entry.Paragraph);
                        }
                        catch { }
                    }

                    foreach (Paragraph p in hit)
                    {
                        try { consumed.Add(p.Range.Start); } catch { }
                    }

                    int changed = 0;
                    if (hit.Count > 0)
                    {
                        var changes = new FormatChanges
                        {
                            EastAsiaFont = rule.Apply.EastAsiaFont,
                            LatinFont = rule.Apply.LatinFont,
                            SizePt = rule.Apply.SizePt,
                            Bold = rule.Apply.Bold,
                            ColorRgb = Formatting.WordFormatUtils.ParseHexColor(rule.Apply.Color)
                        };
                        var stats = BatchFormatter.Apply(app, doc, hit, changes, null, wrapUndo: false);
                        changed = stats.Changed;
                    }
                    report.TotalChanged += changed;
                    report.Lines.Add(string.Format("规则{0}{1}命中 {2} 段（改动 {3}）",
                        rule.Order, rule.Match.CatchAll == true ? "（其余段落）" : " ", hit.Count, changed));
                    if (options != null && options.Progress != null) options.Progress(rule.Order, hit.Count);
                }

                if (profile.ConvertFormulas)
                {
                    var converter = new FormulaConverter(doc, mathSizePt, FormulaConverter.ResolveMml2OmmlPath(app));
                    FormulaStats stats = converter.ConvertDocument();
                    report.FormulaConverted = stats.Converted;
                    report.FormulaFailed = stats.Failed;
                    report.Lines.Add("公式 " + stats.Converted + " 处" +
                        (stats.Failed > 0 ? "（失败 " + stats.Failed + " 处已保留）" : ""));
                }

                int unmatched = 0;
                foreach (var entry in allParas)
                {
                    if (!consumed.Contains(entry.Start)) unmatched++;
                }
                report.Unmatched = unmatched;
                report.Lines.Add("未命中 " + report.Unmatched + " 段（未改动）。");
            }
            finally
            {
                WordCompat.EndUndo(undo);
                if (screenOff) { try { app.ScreenUpdating = true; } catch { } }
            }
            return report;
        }
    }

    /// <summary>操作录制 → 方案规则（控制器与验收测试共用，PRD P1-1 存为方案）。</summary>
    public static class OpRecording
    {
        public sealed class Op
        {
            public FormatSignature SampleSignature;
            public FormatSignature.IgnoreDims Ignore;
            public FormatChanges Applied;
            public bool SampleInTable;
            public bool IsCatchAll;
        }

        public static ProfileV2 BuildProfile(string id, string name, IList<Op> ops, bool withCatchAll,
            string catchAllFont, string defaultLatinFont, bool convertFormulas)
        {
            var profile = new ProfileV2
            {
                Id = id,
                Name = name,
                ConvertFormulas = convertFormulas,
                Rules = new List<ProfileRule>()
            };
            int order = 1;
            foreach (Op op in ops)
            {
                if (op.IsCatchAll)
                {
                    profile.Rules.Add(new ProfileRule
                    {
                        Order = order++,
                        Match = new MatchSpec { CatchAll = true },
                        Apply = ToApplySpec(op.Applied)
                    });
                }
                else
                {
                    profile.Rules.Add(new ProfileRule
                    {
                        Order = order++,
                        Match = ToMatchSpec(op.SampleSignature, op.Ignore, op.SampleInTable),
                        Apply = ToApplySpec(op.Applied)
                    });
                }
            }
            if (withCatchAll && !string.IsNullOrEmpty(catchAllFont))
            {
                profile.Rules.Add(new ProfileRule
                {
                    Order = order++,
                    Match = new MatchSpec { CatchAll = true },
                    Apply = new ApplySpec { EastAsiaFont = catchAllFont, LatinFont = defaultLatinFont }
                });
            }
            return profile;
        }

        public static MatchSpec ToMatchSpec(FormatSignature sig, FormatSignature.IgnoreDims ignore, bool sampleInTable = false)
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

        public static ApplySpec ToApplySpec(FormatChanges changes)
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
    }
}
