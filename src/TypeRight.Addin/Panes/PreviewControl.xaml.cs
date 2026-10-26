using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;
using System.Windows.Media;
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

        /// <summary>绑定控制器（任务窗格与浮动窗口共用）。</summary>
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

    /// <summary>PanePage + 参数（页面名）→ Visibility。</summary>
    public sealed class PageToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (!(value is PanePage) || parameter == null) return Visibility.Collapsed;
            return value.ToString().Equals(parameter.ToString(), StringComparison.OrdinalIgnoreCase)
                ? Visibility.Visible : Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotSupportedException();
        }
    }

    /// <summary>置信类型 → 徽标底色。</summary>
    public sealed class ConfidenceBrushConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is ConfidenceKind)
            {
                switch ((ConfidenceKind)value)
                {
                    case ConfidenceKind.Rule: return new SolidColorBrush(Color.FromRgb(0x2F, 0x85, 0x5A));
                    case ConfidenceKind.Guess: return new SolidColorBrush(Color.FromRgb(0xB4, 0x53, 0x09));
                    case ConfidenceKind.Manual: return new SolidColorBrush(Color.FromRgb(0x2B, 0x5C, 0xB8));
                }
            }
            return new SolidColorBrush(Colors.Gray);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotSupportedException();
        }
    }
}
