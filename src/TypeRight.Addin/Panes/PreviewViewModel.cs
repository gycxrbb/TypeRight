using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Windows.Data;
using System.Windows.Input;
using NetOffice.WordApi;
using TypeRight.Addin.Controllers;
using TypeRight.Formulas;
using TypeRight.Profiles;
using TypeRight.Roles;
using TypeRight.Scanning;

namespace TypeRight.Addin.Panes
{
    public sealed class ParagraphRowVM : INotifyPropertyChanged
    {
        private ParagraphRole _role;

        public int Index { get; set; }
        public string Preview { get; set; }
        public bool LowConfidence { get; set; }
        public string RuleLabel { get; set; }
        public bool UserChanged { get; set; }
        public bool InTable { get; set; }

        /// <summary>下拉可选角色（含 skip）。</summary>
        public static readonly ParagraphRole[] RoleChoices =
        {
            ParagraphRole.Title, ParagraphRole.Heading1, ParagraphRole.Heading2, ParagraphRole.Heading3,
            ParagraphRole.Body, ParagraphRole.Caption, ParagraphRole.ListItem,
            ParagraphRole.TableText, ParagraphRole.FormulaLine, ParagraphRole.Skip
        };

        public ParagraphRole Role
        {
            get { return _role; }
            set
            {
                if (_role != value)
                {
                    _role = value;
                    UserChanged = true;
                    OnPropertyChanged("Role");
                    OnPropertyChanged("RoleLabel");
                }
            }
        }

        public string RoleLabel { get { return _role.ToLabel(); } }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged(string name)
        {
            var h = PropertyChanged;
            if (h != null) h(this, new PropertyChangedEventArgs(name));
        }
    }

    /// <summary>预览窗格视图模型（PRD §7.2）：统计 + 可改角色的逐段清单 + 应用选项。</summary>
    public sealed class PreviewViewModel : INotifyPropertyChanged
    {
        private static readonly Dictionary<string, Profile> EmptyProfiles = new Dictionary<string, Profile>();

        private FormattingController _controller;
        private List<ParagraphFeatures> _features;
        private Range _restrict;                       // 选区模式的有效区间
        private Dictionary<string, Profile> _profileMap = EmptyProfiles;

        public ObservableCollection<Profile> Profiles { get; private set; }

        private Profile _selectedProfile;
        public Profile SelectedProfile
        {
            get { return _selectedProfile; }
            set
            {
                if (!ReferenceEquals(_selectedProfile, value))
                {
                    _selectedProfile = value;
                    OnPropertyChanged("SelectedProfile");
                    OnPropertyChanged("ProfileDescription");
                    if (_controller != null && value != null) _controller.SaveLastProfile(value.Id);
                }
            }
        }

        public string ProfileDescription
        {
            get { return _selectedProfile == null ? "" : (_selectedProfile.Description ?? ""); }
        }

        private string _summaryText = "尚未扫描。点击“一键智能排版”开始。";
        public string SummaryText
        {
            get { return _summaryText; }
            private set { _summaryText = value; OnPropertyChanged("SummaryText"); }
        }

        private string _noteText;
        public string NoteText
        {
            get { return _noteText; }
            private set { _noteText = value; OnPropertyChanged("NoteText"); }
        }

        private bool _selectionMode;
        public bool SelectionMode
        {
            get { return _selectionMode; }
            private set { _selectionMode = value; OnPropertyChanged("SelectionMode"); OnPropertyChanged("ModeText"); }
        }

        public string ModeText { get { return _selectionMode ? "（选区模式）" : ""; } }

        public ObservableCollection<ParagraphRowVM> Rows { get; private set; }

        private bool _showOnlyLowConfidence;
        public bool ShowOnlyLowConfidence
        {
            get { return _showOnlyLowConfidence; }
            set
            {
                _showOnlyLowConfidence = value;
                OnPropertyChanged("ShowOnlyLowConfidence");
                RefreshFilter();
            }
        }

        private bool _mountStyles = true;
        public bool MountStyles
        {
            get { return _mountStyles; }
            set { _mountStyles = value; OnPropertyChanged("MountStyles"); }
        }

        private bool _convertFormulas = true;
        public bool ConvertFormulas
        {
            get { return _convertFormulas; }
            set { _convertFormulas = value; OnPropertyChanged("ConvertFormulas"); }
        }

        public ICommand ApplyCommand { get; private set; }
        public ICommand RescanCommand { get; private set; }
        public ICommand HideCommand { get; private set; }

        public bool HasScan { get { return _features != null; } }

        public PreviewViewModel()
        {
            Profiles = new ObservableCollection<Profile>();
            Rows = new ObservableCollection<ParagraphRowVM>();
            ApplyCommand = new RelayCommand(_ => Apply(), _ => HasScan && SelectedProfile != null);
            RescanCommand = new RelayCommand(_ => Rescan());
            HideCommand = new RelayCommand(_ => Hide());
        }

        // ---- 由控制器填充 ----

        /// <summary>窗格初次显示时设置默认选项（未扫描状态）。</summary>
        public void LoadDefaults(FormattingController controller)
        {
            if (controller == null) return;
            if (controller.Settings != null) MountStyles = controller.Settings.MountStylesDefault;
            _controller = controller;
        }

        public void LoadScan(FormattingController controller, ScanResult scan, bool selectionMode, string note, Range restrict)
        {
            if (controller == null) throw new ArgumentNullException("controller");
            _controller = controller;
            _features = scan.Features;
            _restrict = restrict;
            SelectionMode = selectionMode;
            NoteText = note ?? (controller.ProfileErrors != null && controller.ProfileErrors.Count > 0
                ? "方案加载警告：" + string.Join("；", controller.ProfileErrors)
                : null);

            // 方案列表
            Profiles.Clear();
            _profileMap = new Dictionary<string, Profile>();
            foreach (var lp in controller.Profiles)
            {
                if (lp == null || lp.Profile == null) continue;
                if (!_profileMap.ContainsKey(lp.Profile.Id))
                {
                    _profileMap.Add(lp.Profile.Id, lp.Profile);
                    Profiles.Add(lp.Profile);
                }
            }
            Profile pick = null;
            string lastId = controller.LastProfileId;
            if (lastId != null) _profileMap.TryGetValue(lastId, out pick);
            SelectedProfile = pick ?? Profiles.FirstOrDefault();

            if (controller.Settings != null)
            {
                MountStyles = controller.Settings.MountStylesDefault;
            }

            RebuildRows(scan);
        }

        private void RebuildRows(ScanResult scan)
        {
            Rows.Clear();
            var counts = new Dictionary<ParagraphRole, int>();
            int formulaSegments = 0;

            foreach (var f in scan.Features)
            {
                counts[f.Role] = counts.ContainsKey(f.Role) ? counts[f.Role] + 1 : 1;

                bool listable = !f.IsEmpty || f.Role == ParagraphRole.Skip;
                if (listable)
                {
                    Rows.Add(new ParagraphRowVM
                    {
                        Index = f.Index,
                        Preview = MakePreview(f),
                        Role = f.Role,
                        LowConfidence = !f.HighConfidence && f.Role != ParagraphRole.Skip,
                        RuleLabel = "R" + (f.RulePriority > 0 ? f.RulePriority.ToString() : "?"),
                        InTable = f.InTable
                    });
                }

                if (IsConvertible(f.Role))
                {
                    if (f.Role == ParagraphRole.FormulaLine) formulaSegments++;
                    else if (!f.IsEmpty && !f.HasFields)
                        formulaSegments += MathSegmenter.FindSegments(f.Text).Count;
                }
            }

            var sb = new StringBuilder();
            sb.Append("识别结果：");
            foreach (var kv in counts.OrderByDescending(k => k.Value))
            {
                if (kv.Key == ParagraphRole.Skip || kv.Value == 0) continue;
                sb.Append("\r\n  ").Append(kv.Key.ToLabel()).Append(" ").Append(kv.Value).Append(" 段");
            }
            if (scan.TableCount > 0) sb.Append("\r\n  表格 ").Append(scan.TableCount).Append(" 张");
            if (formulaSegments > 0) sb.Append("\r\n  公式片段 ").Append(formulaSegments).Append(" 处（将转公式对象）");
            sb.Append("\r\n  跳过 ").Append(counts.ContainsKey(ParagraphRole.Skip) ? counts[ParagraphRole.Skip] : 0).Append(" 段");
            sb.Append("\r\n  扫描用时 ").Append(scan.Duration.TotalMilliseconds.ToString("0")).Append(" ms");
            SummaryText = sb.ToString();

            RefreshFilter();
        }

        private static bool IsConvertible(ParagraphRole role)
        {
            switch (role)
            {
                case ParagraphRole.Body:
                case ParagraphRole.FormulaLine:
                case ParagraphRole.ListItem:
                case ParagraphRole.Caption:
                case ParagraphRole.TableText:
                case ParagraphRole.TableHeader:
                    return true;
                default:
                    return false;
            }
        }

        private static string MakePreview(ParagraphFeatures f)
        {
            string text = f.Text;
            if (text.Length > 44) text = text.Substring(0, 44) + "…";
            return "¶" + f.Index + "  " + text;
        }

        private void RefreshFilter()
        {
            var view = CollectionViewSource.GetDefaultView(Rows);
            if (view != null)
            {
                view.Filter = ShowOnlyLowConfidence
                    ? (Predicate<object>)(o => o is ParagraphRowVM r && (r.LowConfidence || r.UserChanged))
                    : null;
            }
        }

        // ---- 命令 ----

        private void Apply()
        {
            if (_controller == null || SelectedProfile == null) return;

            // 校验方案（预览加载时已校验过；双保险）
            var errors = TypeRight.Profiles.ProfileValidator.Validate(SelectedProfile);
            if (errors.Count > 0)
            {
                System.Windows.MessageBox.Show("方案校验未通过：\r\n" + string.Join("\r\n", errors), "TypeRight");
                return;
            }

            var assignments = new Dictionary<int, ParagraphRole>();
            foreach (var row in Rows)
                assignments[row.Index] = row.Role;

            _controller.Apply(assignments, SelectedProfile, MountStyles, ConvertFormulas, _restrict);
            Rescan();
        }

        private void Rescan()
        {
            if (_controller == null) return;
            _controller.Run(_selectionMode);
        }

        private void Hide()
        {
            _controller.HidePreviewPane();
        }

        /// <summary>选区模式时记录有效区间（应用阶段还原）。</summary>
        public void SetRestrict(Range restrict)
        {
            _restrict = restrict;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged(string name)
        {
            var h = PropertyChanged;
            if (h != null) h(this, new PropertyChangedEventArgs(name));
        }
    }

    /// <summary>轻量命令。</summary>
    public sealed class RelayCommand : ICommand
    {
        private readonly Action<object> _execute;
        private readonly Predicate<object> _canExecute;

        public RelayCommand(Action<object> execute, Predicate<object> canExecute = null)
        {
            _execute = execute;
            _canExecute = canExecute;
        }

        public bool CanExecute(object parameter)
        {
            return _canExecute == null || _canExecute(parameter);
        }

        public void Execute(object parameter)
        {
            _execute(parameter);
        }

        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }
    }
}
