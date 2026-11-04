using NetOffice.WordApi;

namespace TypeRight
{
    /// <summary>宿主 Word 版本兼容层（Word 2010/2013 支持专项）。
    /// 背景：Application.UndoRecord 与 Window.Hwnd 都是 Word 2013 (15.0) 才进入对象模型的，
    /// 2010 上访问直接抛 COMException。这里统一降级，业务代码不再各自感知版本。</summary>
    public static class WordCompat
    {
        /// <summary>宿主 Word 主版本号（"16.0"→16；读取失败返回 0，视为未知）。</summary>
        public static int MajorVersion(Application app)
        {
            try { return int.Parse(app.Version.Split('.')[0]); }
            catch { return 0; }
        }

        /// <summary>开启自定义撤销记录。Word 2010 没有 UndoRecord → 返回 null：
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
    }
}
