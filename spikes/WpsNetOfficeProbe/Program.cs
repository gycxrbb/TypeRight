// ============================================================================
// NetOffice 包装层 vs dynamic 直调 —— 在 WPS 上的 A/B 对照探针（spike）
// 用途：解释"SimilarSelect 在 WPS 里样本签名全 ?/?"的现象——
//       定位是 NetOffice 的调用方式问题，还是 WPS 的 Font 对象本身问题。
// ============================================================================
using System;
using NetOffice.WordApi;

namespace WpsNetOfficeProbe
{
    internal static class Program
    {
        [STAThread]
        private static int Main(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            string docx = args.Length > 0 ? args[0] :
                System.IO.Path.Combine(System.IO.Path.GetTempPath(), "wps-smoke.docx");

            // ── dynamic 激活 WPS ──
            Type t = Type.GetTypeFromProgID("KWPS.Application");
            dynamic rawApp = Activator.CreateInstance(t);
            rawApp.Visible = false;
            Console.WriteLine("WPS 激活: Version=" + rawApp.Version);

            // ── 包成 NetOffice ──
            Application app = null;
            try
            {
                app = new Application(null, rawApp);   // NetOffice attach 构造 (ICOMObject parent, object comProxy)
                Console.WriteLine("NetOffice 包装: OK  Version=" + app.Version);
            }
            catch (Exception ex)
            {
                Console.WriteLine("NetOffice 包装失败: " + ex.Message);
            }

            dynamic rawDoc = rawApp.Documents.Open(docx, Type.Missing, false);
            Console.WriteLine("文档打开: " + rawDoc.FullName);
            Document doc = null;
            if (app != null) { try { doc = new Document(app, rawDoc); Console.WriteLine("文档包装: OK"); } catch (Exception ex) { Console.WriteLine("文档包装失败: " + ex.Message); } }
            int pos = (int)rawDoc.Paragraphs[30].Range.Start;

            Console.WriteLine();
            Console.WriteLine("═══ A. dynamic 直调（基准，探针已证明可用）═══");
            try
            {
                dynamic f = rawDoc.Range(pos, pos + 1).Font;
                Console.WriteLine("  NameFarEast = " + f.NameFarEast);
                Console.WriteLine("  Size        = " + f.Size);
                Console.WriteLine("  Bold        = " + f.Bold);
                Console.WriteLine("  Color       = " + f.Color);
            }
            catch (Exception ex) { Console.WriteLine("  失败: " + ex.GetType().Name + ": " + ex.Message); }

            if (doc != null)
            {
                Console.WriteLine();
                Console.WriteLine("═══ B. NetOffice 包装层（产品的真实路径）═══");
                try
                {
                    Range r = doc.Range(pos, pos + 1);
                    Console.WriteLine("  Range Start/End = " + r.Start + "/" + r.End);
                    NetOffice.WordApi.Font f = r.Font;
                    Console.WriteLine("  Font 对象获取: OK (" + f.GetType().Name + ")");
                    Console.WriteLine("  NameFarEast = [" + TryRead(() => f.NameFarEast) + "]");
                    Console.WriteLine("  NameAscii   = [" + TryRead(() => f.NameAscii) + "]");
                    Console.WriteLine("  Size        = [" + TryRead(() => f.Size) + "]");
                    Console.WriteLine("  Bold        = [" + TryRead(() => f.Bold) + "]");
                    Console.WriteLine("  Color       = [" + TryRead(() => f.Color) + "]");
                }
                catch (Exception ex) { Console.WriteLine("  失败: " + ex.GetType().Name + ": " + ex.Message); }

                Console.WriteLine();
                Console.WriteLine("═══ C. NetOffice 写入 + 回读 ═══");
                try
                {
                    Range r = doc.Range(pos, pos + 1);
                    r.Font.NameFarEast = "楷体";
                    r.Font.Size = 15.0f;
                    Console.WriteLine("  写入调用未抛异常");
                    Range r2 = doc.Range(pos, pos + 1);
                    Console.WriteLine("  回读 NameFarEast = [" + TryRead(() => r2.Font.NameFarEast) + "]");
                    Console.WriteLine("  回读 Size        = [" + TryRead(() => r2.Font.Size) + "]");
                }
                catch (Exception ex) { Console.WriteLine("  失败: " + ex.GetType().Name + ": " + ex.Message); }
            }

            if (doc != null)
            {
                Console.WriteLine();
                Console.WriteLine("═══ D. 段落枚举：底层 COM 集合 + 手工包装（拟用于产品的兼容路径）═══");
                try
                {
                    int count = 0; string firstText = null; float firstSize = 0;
                    var rawEnum = (System.Collections.IEnumerable)doc.Paragraphs.UnderlyingObject;
                    foreach (object item in rawEnum)
                    {
                        var para = new NetOffice.WordApi.Paragraph(null, item);
                        if (count == 0)
                        {
                            firstText = para.Range.Text; 
                            firstSize = para.Range.Font.Size;
                        }
                        count++;
                        if (count >= 200) break;
                    }
                    Console.WriteLine("  枚举段落数: " + count);
                    Console.WriteLine("  首段文本: " + (firstText == null ? "null" : firstText.Substring(0, Math.Min(20, firstText.Length))));
                    Console.WriteLine("  首段字号: " + firstSize);
                }
                catch (Exception ex) { Console.WriteLine("  失败: " + ex.GetType().Name + ": " + ex.Message); }
            }

            if (doc != null)
            {
                Console.WriteLine();
                Console.WriteLine("═══ E. NetOffice Shading 写入（对 Word/WPS 同样适用）═══");
                try
                {
                    var rr = doc.Range(0, 5);
                    Console.WriteLine("  写前 Bg=" + rr.Shading.BackgroundPatternColor);
                    rr.Shading.Texture = NetOffice.WordApi.Enums.WdTextureIndex.wdTextureNone;
                    rr.Shading.BackgroundPatternColor = (NetOffice.WordApi.Enums.WdColor)4961256;
                    var rr2 = doc.Range(0, 5);
                    Console.WriteLine("  写后 Bg=" + rr2.Shading.BackgroundPatternColor + " (期望 4961256)");
                }
                catch (Exception ex) { Console.WriteLine("  失败: " + ex.GetType().Name + ": " + ex.Message); }
            }

            try { rawDoc.Close(false); } catch { }
            try { rawApp.Quit(); } catch { }
            Console.WriteLine("完成");
            return 0;
        }

        private static string TryRead(Func<object> f)
        {
            try { var v = f(); return v == null ? "null" : v.ToString(); }
            catch (Exception ex) { return "异常:" + ex.GetType().Name + "(" + ex.Message.Substring(0, Math.Min(60, ex.Message.Length)) + ")"; }
        }
    }
}
