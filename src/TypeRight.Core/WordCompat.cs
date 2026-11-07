using System.Collections.Generic;
using NetOffice.WordApi;

namespace TypeRight
{
    /// <summary>宿主兼容层。
    /// ① 老 Word（2010/2013）：Application.UndoRecord 与 Window.Hwnd 都是 2013 (15.0) 才进对象模型；
    /// ② WPS：NetOffice 集合枚举器经工厂按类名解析 Word.* coclass（FactoryException: Class does not exist），
    /// 改为枚举底层 COM 集合 + 手工包装（Word/WPS 行为一致，保 O(n) 单遍）。
    /// 业务代码不再各自感知宿主差异。</summary>
    public static class WordCompat
    {
        /// <summary>宿主主版本号（"16.0"→16；WPS 自报 12.0；读取失败返回 0）。</summary>
        public static int MajorVersion(Application app)
        {
            try { return int.Parse(app.Version.Split('.')[0]); }
            catch { return 0; }
        }

        /// <summary>是否运行在 WPS 内（WPS 自报 12.0；真 Word 2007 同报 12.0 但已在支持线之外）。</summary>
        public static bool IsWps(Application app)
        {
            try { return app.Version.StartsWith("12"); }
            catch { return false; }
        }

        /// <summary>开启自定义撤销记录。Word 2010/老宿主没有 UndoRecord → 返回 null：
        /// 撤销粒度退化为多条记录（Ctrl+Z 多按几次），功能不受损，绝不阻断主流程。</summary>
        public static UndoRecord BeginUndo(Application app, string name)
        {
            try
            {
                UndoRecord record = app.UndoRecord;
                if (record == null) return null;
                record.StartCustomRecord(name);
                return record;
            }
            catch { return null; }
        }

        /// <summary>结束撤销记录（null 安全；收尾失败吞掉——记录已尽力收口）。</summary>
        public static void EndUndo(UndoRecord record)
        {
            if (record == null) return;
            try { record.EndCustomRecord(); } catch { }
        }

        // ── 集合兼容枚举（WPS 必需；Word 行为等价）──
        // 原理：NetOffice 的 GetEnumerator 对每个元素走 Core.CreateObjectFromComProxy 按类名
        // （如 "Word.Paragraph"）查注册表包装类型——WPS 进程内没有 Word 的 coclass 注册，抛
        // FactoryException。底层 IEnumVARIANT 枚举本身可用（dynamic foreach 实测 183 段全通），
        // 手工 new Paragraph(parent, proxy) 是纯代理包装、不查注册表。

        /// <summary>段落集合兼容枚举（替代 foreach (Paragraph p in xxx.Paragraphs)）。</summary>
        public static List<Paragraph> ParagraphsOf(Paragraphs paragraphs)
        {
            var list = new List<Paragraph>();
            try
            {
                foreach (object item in (System.Collections.IEnumerable)paragraphs.UnderlyingObject)
                    list.Add(new Paragraph(null, item));
            }
            catch { }
            return list;
        }

        /// <summary>文档级段落（常用重载）。</summary>
        public static List<Paragraph> ParagraphsOf(Document doc)
        {
            return ParagraphsOf(doc.Paragraphs);
        }

        /// <summary>表格集合兼容枚举。</summary>
        public static List<Table> TablesOf(Tables tables)
        {
            var list = new List<Table>();
            try
            {
                foreach (object item in (System.Collections.IEnumerable)tables.UnderlyingObject)
                    list.Add(new Table(null, item));
            }
            catch { }
            return list;
        }

        /// <summary>公式集合兼容枚举（MathZoneRanges 用；原索引器在 WPS 同样走工厂包装）。</summary>
        public static List<OMath> OMathsOf(OMaths omaths)
        {
            var list = new List<OMath>();
            try
            {
                foreach (object item in (System.Collections.IEnumerable)omaths.UnderlyingObject)
                    list.Add(new OMath(null, item));
            }
            catch { }
            return list;
        }
    }
}
