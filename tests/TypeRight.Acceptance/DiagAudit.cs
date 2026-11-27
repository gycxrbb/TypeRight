using System;
using System.Collections.Generic;
using NetOffice.WordApi;
using TypeRight.Signature;

namespace TypeRight.Acceptance
{
    /// <summary>
    /// 真实文档诊断：逐段打印签名（含主导字符类型），用于定位"数字段被误判同类"。
    /// 用法：TypeRight.Acceptance.exe audit &lt;docxPath&gt;
    /// </summary>
    internal static class DiagAudit
    {
        public static int Run(TestEnv env, string path)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            // NetOffice 的 Open 重载不接受命名参数；位置参数：(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles)
            Document doc = env.App.Documents.Open(path, false, true, false);
            try
            {
                Console.WriteLine();
                Console.WriteLine("── 逐段签名审计 ──────────────────────");
                Console.WriteLine("段号 字符数 首字 字符类型 Ea字体      Latin字体       字号   加粗 文本前16字");
                Console.WriteLine("──────────────────────────────────────────────────────────────────────");

                int idx = 0;
                var sigs = new List<FormatSignature>();
                var texts = new List<string>();
                foreach (Paragraph p in WordCompat.ParagraphsOf(doc))
                {
                    idx++;
                    string raw = "";
                    try { raw = p.Range.Text ?? ""; } catch { }
                    string visible = raw.Replace("\r", "").Replace("\a", "").Replace("\v", "").Trim();
                    if (visible.Length == 0) continue;

                    FormatSignature s = FormatSignatureReader.Read(p);
                    sigs.Add(s);
                    texts.Add(visible);

                    string head = visible.Length > 16 ? visible.Substring(0, 16) : visible;
                    int digit = 0, latin = 0, cjk = 0;
                    foreach (char c in visible)
                    {
                        if (CharClassDetector.IsDigit(c)) digit++;
                        else if (CharClassDetector.IsLatin(c)) latin++;
                        else if (CharClassDetector.IsCjk(c)) cjk++;
                    }
                    Console.WriteLine(string.Format("{0,-4} {1,-6} {2,-4} {3,-8} {4,-11} {5,-16} {6,-6} B{7}  {8}",
                        idx, visible.Length, visible[0], CharClassDetector.Name(s.CharClass),
                        s.EastAsiaFont ?? "?", s.LatinFont ?? "?",
                        double.IsNaN(s.SizePt) ? "?" : s.SizePt.ToString("0.#"),
                        s.BoldRaw == -1 ? 1 : 0, head,
                        ""));
                    if (digit + latin > 0)
                        Console.WriteLine(string.Format("      └ 字符构成: 汉字{0} 字母{1} 数字{2}", cjk, latin, digit));
                }
            }
            finally { try { doc.Close(false); } catch { } }
            return 0;
        }
    }
}
