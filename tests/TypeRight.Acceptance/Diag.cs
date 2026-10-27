using System;
using System.Linq;
using System.Text;
using NetOffice.WordApi;
using NetOffice.WordApi.Enums;
using TypeRight.Signature;

namespace TypeRight.Acceptance
{
    /// <summary>一次性诊断：签名普查 + 命中清单 + 分式/根式构建探针。dotnet run --project tests/TypeRight.Acceptance -- diag</summary>
    internal static class Diag
    {
        public static void Run(TestEnv env)
        {
            Console.OutputEncoding = Encoding.UTF8;
            Document doc = env.OpenCopy("diag");
            try
            {
                Console.WriteLine("══ 签名普查（关键锚点）══");
                Dump(doc, env, "初三数学");
                Dump(doc, env, "模块一");
                Dump(doc, env, "模块二");
                Dump(doc, env, "1.1 ");
                Dump(doc, env, "解一元二次方程的灵魂是");
                Dump(doc, env, "【图解例图】");
                Dump(doc, env, "——涵盖");
                Dump(doc, env, "★");

                Console.WriteLine();
                Console.WriteLine("══ 模块样本 忽略字号 命中清单（TC-02 的 16 段）══");
                Paragraph module = env.FindParagraph(doc, "模块一");
                var sig = FormatSignatureReader.Read(module);
                var hit = SimilarSelector.Match(doc, sig, FormatSignature.IgnoreDims.SizePt);
                foreach (Paragraph p in hit.Paragraphs)
                {
                    string t = (p.Range.Text ?? "").TrimEnd('\r', '\a', '\v');
                    bool inTable;
                    try { inTable = Convert.ToBoolean(p.Range.get_Information(WdInformation.wdWithInTable)); }
                    catch { inTable = false; }
                    var s = FormatSignatureReader.Read(p);
                    Console.WriteLine("  [" + (inTable ? "表" : "文") + "] " + t.Substring(0, Math.Min(20, t.Length)) + "  → " + s);
                }

                Console.WriteLine();
                Console.WriteLine("══ 图注样本 忽略字号+颜色 命中清单（TC-04③ 的 6 段）══");
                Paragraph cap = env.FindParagraph(doc, "【图解例图】");
                var sigC = FormatSignatureReader.Read(cap);
                foreach (Paragraph p in SimilarSelector.Match(doc, sigC, FormatSignature.IgnoreDims.SizePt | FormatSignature.IgnoreDims.Color).Paragraphs)
                {
                    string t = (p.Range.Text ?? "").TrimEnd('\r', '\a', '\v');
                    Console.WriteLine("  " + t.Substring(0, Math.Min(20, t.Length)) + "  → " + FormatSignatureReader.Read(p));
                }

                var sub = env.FindParagraph(doc, "——涵盖");
                if (sub != null)
                {
                    var subSig = FormatSignatureReader.Read(sub);
                    bool direct = subSig.Matches(sigC, FormatSignature.IgnoreDims.SizePt | FormatSignature.IgnoreDims.Color);
                    Console.WriteLine("副标题直判 Matches(caption, 忽略字号+颜色) = " + direct);
                    Console.WriteLine("  caption: 东=[" + sigC.EastAsiaFont + "] 西=[" + sigC.LatinFont + "] 斜raw=" + sigC.ItalicRaw + " 粗raw=" + sigC.BoldRaw);
                    Console.WriteLine("  副标题:  东=[" + subSig.EastAsiaFont + "] 西=[" + subSig.LatinFont + "] 斜raw=" + subSig.ItalicRaw + " 粗raw=" + subSig.BoldRaw);
                }

                Console.WriteLine();
                Console.WriteLine("══ 正文严格命中 27 段清单（完整签名）══");
                {
                    var bodyP2 = env.FindParagraph(doc, "解一元二次方程的灵魂是");
                    int n = 0;
                    foreach (Paragraph p2 in SimilarSelector.Match(doc, bodyP2, FormatSignature.IgnoreDims.None).Paragraphs)
                    {
                        n++;
                        string t2 = (p2.Range.Text ?? "").TrimEnd((char)13, (char)7, (char)11);
                        var s3 = FormatSignatureReader.Read(p2);
                        Console.WriteLine("  [" + n + "] " + t2.Substring(0, Math.Min(14, t2.Length)) +
                            "  东=" + (s3.EastAsiaFont ?? "?") + " 西=" + (s3.LatinFont ?? "?") +
                            " 号=" + s3.SizePt.ToString("0.#") + " 粗raw=" + s3.BoldRaw + " 斜raw=" + s3.ItalicRaw +
                            " 色=" + s3.ColorRgb.ToString("X6") + " 齐=" + s3.Alignment);
                    }
                }

                Console.WriteLine();
                Console.WriteLine("══ 斜体真值诊断 ══");
                var bodyP = env.FindParagraph(doc, "解一元二次方程的灵魂是");
                var bodySig2 = FormatSignatureReader.Read(bodyP);
                foreach (Paragraph p in SimilarSelector.Match(doc, bodyP, FormatSignature.IgnoreDims.Bold | FormatSignature.IgnoreDims.Color | (FormatSignature.IgnoreDims)0).Paragraphs)
                {
                    string t = (p.Range.Text ?? "").TrimEnd((char)13, (char)7, (char)11);
                    // 逐字符读斜体（前3个可见字符）
                    var chars = p.Range.Characters;
                    string italics = "";
                    for (int ci = 1; ci <= Math.Min(3, chars.Count); ci++)
                    {
                        try { italics += chars[ci].Font.Italic + "/"; } catch { italics += "?/"; }
                    }
                    var ps = FormatSignatureReader.Read(p);
                    Console.WriteLine("  [" + t.Substring(0, Math.Min(16, t.Length)) + "] 首字斜体串=" + italics + " 签名斜raw=" + ps.ItalicRaw + " 色=" + ps.ColorRgb.ToString("X6"));
                }
                var capP = env.FindParagraph(doc, "【图解例图】");
                {
                    var chars = capP.Range.Characters;
                    string italics = "";
                    for (int ci = 1; ci <= Math.Min(4, chars.Count); ci++)
                    {
                        try { italics += chars[ci].Font.Italic + "/"; } catch { italics += "?/"; }
                    }
                    Console.WriteLine("  图注前4字符斜体: " + italics);
                }

                Console.WriteLine();
                Console.WriteLine("══ 结构顺序矩阵 ─═");
                foreach (int mode in new[] { 3 })
                {
                    TypeRight.Formatting.FormulaConverter.StructureOrderMode = mode;
                    Document md = env.OpenCopy("matrix-" + mode);
                    try
                    {
                        var ms = new TypeRight.Formatting.FormulaConverter(md, null, null).ConvertDocument();
                        string saveMsg;
                        try { md.SaveAs2(env.CopyPath("matrix-save-" + mode)); saveMsg = "保存OK"; }
                        catch (Exception ex3) { saveMsg = "保存失败:" + ex3.Message.Substring(0, Math.Min(30, ex3.Message.Length)); }
                        md.Close(false);
                        using (var mx = DocXml.Load(env.CopyPath("matrix-save-" + mode)))
                        {
                            if (mx.OMathCount() == 0)
                                Console.WriteLine("  mode" + mode + ": 转换" + ms.Converted + " 失败" + ms.Failed + " | " + saveMsg + " | (文件级统计不可用)");
                            else
                                Console.WriteLine("  mode" + mode + ": 转换" + ms.Converted + " 失败" + ms.Failed + " | " + saveMsg +
                                    " | oMath=" + mx.OMathCount() + " m:f=" + mx.MathElementCount("f") + " m:rad=" + mx.MathElementCount("rad") +
                                    " 上下标=" + (mx.MathElementCount("sSub") + mx.MathElementCount("sSup") + mx.MathElementCount("sSubSup")));
                        }
                    }
                    catch (Exception ex4) { Console.WriteLine("  mode" + mode + " 外层异常: " + ex4.Message); try { md.Close(false); } catch { } }
                }
                TypeRight.Formatting.FormulaConverter.StructureOrderMode = 0;

                Console.WriteLine();
                Console.WriteLine("══ 公式缺口普查（真实讲义）══");
                {
                    Document cd = env.OpenCopy("diag-census");
                    try
                    {
                        // 转换前：全文候选片段清单
                        var candidates = new System.Collections.Generic.List<string>();
                        foreach (Paragraph cp in cd.Paragraphs)
                        {
                            string txt = (cp.Range.Text ?? "").TrimEnd((char)13, (char)7, (char)11);
                            foreach (var seg in TypeRight.Formulas.MathSegmenter.FindSegments(txt))
                                candidates.Add(seg.Text);
                        }
                        int before = candidates.Count;

                        var conv = new TypeRight.Formatting.FormulaConverter(cd, null, null);
                        var st = conv.ConvertDocument();
                        int after = cd.OMaths.Count;
                        Console.WriteLine("  候选片段=" + before + " 转换=" + st.Converted + " 失败=" + st.Failed + " oMath=" + after);
                        foreach (string fb in TypeRight.Formatting.FormulaConverter.FallbackLog.Take(20))
                            Console.WriteLine("    回退: " + fb);
                        Console.WriteLine("    回退总数: " + TypeRight.Formatting.FormulaConverter.FallbackLog.Count);

                        try
                        {
                            cd.SaveAs2(env.CopyPath("census-save"));
                            Console.WriteLine("  保存验证: OK mode=" + TypeRight.Formatting.FormulaConverter.StructureOrderMode);
                        }
                        catch (Exception ex2) { Console.WriteLine("  保存验证失败 mode=" + TypeRight.Formatting.FormulaConverter.StructureOrderMode + ": " + ex2.Message); }

                        // 我的公式区清单（与黄金差集用）：运行态读线性文本
                        {
                            int zi = 0;
                            int ztotal = cd.OMaths.Count;
                            for (int zi2 = 1; zi2 <= Math.Min(80, ztotal); zi2++)
                            {
                                string zt;
                                try { zt = cd.OMaths[zi2].Range.Text ?? ""; } catch { zt = "?"; }
                                Console.WriteLine("    myzone[" + zt.TrimEnd((char)13, (char)7) + "]");
                            }
                            Console.WriteLine("    myzone total=" + ztotal);
                        }
                        // 转换后仍以纯文本存在的片段（未被覆盖者）
                        var remaining = new System.Collections.Generic.List<string>();
                        foreach (Paragraph cp in cd.Paragraphs)
                        {
                            string txt = (cp.Range.Text ?? "").TrimEnd((char)13, (char)7, (char)11);
                            foreach (var seg in TypeRight.Formulas.MathSegmenter.FindSegments(txt))
                                remaining.Add(seg.Text);
                        }
                        Console.WriteLine("  转换后残留纯文本片段=" + remaining.Count);
                        foreach (string r in remaining.Take(25))
                            Console.WriteLine("    残留: [" + r + "]");
                        foreach (string fl in st.FailedLocations.Take(10))
                            Console.WriteLine("    失败: " + fl);
                        // 失败片段原文上下文
                        foreach (string fl in st.FailedLocations.Take(8))
                        {
                            string needle = fl.Substring(fl.IndexOf("  ") + 2);
                            needle = needle.Replace("…", "").Trim();
                            if (needle.Length < 4) needle = needle;
                            var f = cd.Content;
                            f.Find.ClearFormatting();
                            f.Find.Text = needle.Substring(0, Math.Min(10, needle.Length));
                            f.Find.Forward = true;
                            f.Find.Wrap = NetOffice.WordApi.Enums.WdFindWrap.wdFindStop;
                            if (f.Find.Execute())
                            {
                                var ctx = f.Duplicate;
                                ctx.Start = Math.Max(0, f.Start - 14);
                                ctx.End = Math.Min(cd.Content.End, f.End + 20);
                                Console.WriteLine("    ctx: [" + (ctx.Text ?? "").Replace("\r", " ").Replace("\a", " ") + "]");
                            }
                        }
                    }
                    finally { cd.Close(false); }
                }

                Console.WriteLine();
                Console.WriteLine("══ BuildUp 变体探针 ══");
                string[] variants =
                {
                    "V1 a/b + c",
                    "V2 (a)/(b) + c",
                    "V3 x = (-b±√(b²-4ac))/(2a)",
                    "V4 √(b²-4ac) ≥ 0",
                    "V5 a∕b（U+2215）",
                    "V6 x₁² = p",
                };
                Document probe = env.App.Documents.Add();
                try
                {
                    var rng = probe.Range(0, 0);
                    foreach (string v in variants)
                    {
                        rng.InsertAfter(v);
                        rng.InsertParagraphAfter();
                    }
                    var conv = new TypeRight.Formatting.FormulaConverter(probe, null, null);
                    var stats = conv.ConvertDocument();
                    Console.WriteLine("  转换: " + stats.SummaryText() + " 失败: " + stats.Failed);
                    probe.SaveAs2(env.CopyPath("diag-probe"));
                    probe.Close(false);
                }
                catch (Exception ex) { Console.WriteLine("  probe ex: " + ex.Message); try { probe.Close(false); } catch { } }

                using (var xml = DocXml.Load(env.CopyPath("diag-probe")))
                {
                    Console.WriteLine("  probe: oMath=" + xml.OMathCount() +
                        " m:f=" + xml.MathElementCount("f") +
                        " m:rad=" + xml.MathElementCount("rad") +
                        " sSub=" + xml.MathElementCount("sSub") +
                        " sSup=" + xml.MathElementCount("sSup") +
                        " sSubSup=" + xml.MathElementCount("sSubSup"));
                    int i = 0;
                    foreach (var om in xml.Root.Descendants(DocXml.M + "oMath"))
                    {
                        i++;
                        if (i <= 8)
                            Console.WriteLine("    zone" + i + " [" + DocXml.MathText(om) + "]" +
                                " f=" + om.Descendants(DocXml.M + "f").Count() +
                                " rad=" + om.Descendants(DocXml.M + "rad").Count());
                    }
                }
            }
            finally { doc.Close(false); }
        }

        private static void Dump(Document doc, TestEnv env, string prefix)
        {
            Paragraph p = env.FindParagraph(doc, prefix);
            if (p == null) { Console.WriteLine("  [缺失] " + prefix); return; }
            string t = (p.Range.Text ?? "").TrimEnd('\r', '\a', '\v');
            bool inTable;
            try { inTable = Convert.ToBoolean(p.Range.get_Information(WdInformation.wdWithInTable)); }
            catch { inTable = false; }
            var s = FormatSignatureReader.Read(p);
            Console.WriteLine("  [" + (inTable ? "表" : "文") + "] " + t.Substring(0, Math.Min(18, t.Length)) +
                "  东=" + (s.EastAsiaFont ?? "?") + " 西=" + (s.LatinFont ?? "?") +
                " 号=" + (double.IsNaN(s.SizePt) ? "?" : s.SizePt.ToString()) +
                " 粗=" + s.BoldRaw + " 斜=" + s.ItalicRaw +
                " 色=" + (s.ColorRgb < 0 ? "?" : s.ColorRgb.ToString("X6")) +
                " 齐=" + s.Alignment);
        }
    }
}
