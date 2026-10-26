using System;
using System.Globalization;
using System.Windows.Data;
using TypeRight.Addin.Controllers;
using TypeRight.Roles;

namespace TypeRight.Addin.Panes
{
    public partial class PreviewControl : System.Windows.Controls.UserControl
    {
        public PreviewViewModel ViewModel { get; private set; }

        public PreviewControl()
        {
            InitializeComponent();
            ViewModel = new PreviewViewModel();
            DataContext = ViewModel;
        }

        /// <summary>绑定控制器（窗格与回落窗口共用）。</summary>
        public void Bind(FormattingController controller)
        {
            ViewModel.LoadDefaults(controller);
        }
    }

    /// <summary>ParagraphRole → 中文标签。</summary>
    public sealed class RoleLabelConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is ParagraphRole) return ((ParagraphRole)value).ToLabel();
            return value == null ? "" : value.ToString();
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotSupportedException();
        }
    }
}
