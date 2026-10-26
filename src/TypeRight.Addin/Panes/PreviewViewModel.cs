using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using NetOffice.WordApi;
using TypeRight.Addin.Controllers;
using TypeRight.Formulas;
using TypeRight.Profiles;
using TypeRight.Roles;
using TypeRight.Scanning;

namespace TypeRight.Addin.Panes
{
    public enum PanePage { Welcome, Preview, Result }

    public enum ConfidenceKind { Rule, Guess, Manual }

    public sealed class StatChip
    {
        public string Label { get; set; }
        public int Count { get; set; }
        public Brush Dot { get; set; }
        public string ToolTip { get; set; }
    }

    public sealed class ParagraphRowVM : INotifyPropertyChanged
    {
        private ParagraphRole _role;

        public int Index { get; set; }
        public string Preview { get; set; }
        public bool LowConfidence { get; set; }
        public bool InTable { get; set; }
        public bool UserChanged { get; set; }

        public static readonly ParagraphRole[] RoleChoices =
        {
            ParagraphRole.Title, ParagraphRole.Heading1, ParagraphRole.Heading2, ParagraphRole.Heading3,
            ParagraphRole.Body, ParagraphRole.Caption, ParagraphRole.ListItem,
            ParagraphRole.TableHeader, ParagraphRole.TableText, ParagraphRole.FormulaLine, ParagraphRole.Skip
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
                    OnPropertyChanged("ConfidenceLabel");
                    OnPropertyChanged("ConfidenceKind");
                }
            }
        }

        public string RoleLabel { get { return _role.ToLabel(); } }

        /// <summary>置信标签：规则（绿）/ 推测（琥珀）/ 手动（蓝）。</summary>
        public ConfidenceKind ConfidenceKind
        {
            get
            {
                if (UserChanged) return Panes.ConfidenceKind.Manual;
                return LowConfidence ? Panes.ConfidenceKind.Guess : Panes.ConfidenceKind.Rule;
            }
        }

        public string ConfidenceLabel
        {
            get
            {
                switch (ConfidenceKind)
                {
                    case Panes.ConfidenceKind.Manual: return "手动";
                    case Panes.ConfidenceKind.Guess: return "推测";
                    default: return "规则";
                }
            }
        }

        /// <summary>推测项的提示（tooltip 用）。</summary>
        public string ConfidenceTip
        {
            get
            {
                switch (ConfidenceKind)
                {
                    case Panes.ConfidenceKind.Manual: return "你手动指定了角色";
                    case Panes.ConfidenceKind.Guess: return "按格式聚类推测，建议核对（橙色行）";
                    default: return "按编号/样式等明确规则识别";
                }
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged(string name)
        {
            var h = PropertyChanged;
            if (h != null) h(this, new PropertyChangedEventArgs(name));
        }
    }

    /// <summary>
    /// 预览面板视图模型：欢迎引导 → 确认角色 → 应用结果，三状态产品流。
    /// </summary>
    public sealed class PreviewViewModel : INotifyPropertyChanged
    {
        private static readonly Dictionary<string, Profile> EmptyProfiles = new Dictionary<string, Profile>();

        private FormattingController _controller;
        private List<ParagraphFeatures> _features;
        private Range _restrict;
        private bool _selectionMode;
        private Dictionary<string, Profile> _profileMap = EmptyProfiles;

        // ---- 状态 ----

        private PanePage _page = PanePage.Welcome;
        public PanePage Page
        {
            get { return _page; }
            private set { _page = value; OnPropertyChanged("Page"); OnPropertyChanged("IsWelcome"); OnPropertyChanged("IsPreview"); OnPropertyChanged("IsResult"); }
        }

        public bool IsWelcome { get { return _page == PanePage.Welcome; } }
        public bool IsPreview { get { return _page == PanePage.Preview; } }
        public bool IsResult { get { return _page == PanePage.Result; } }

        private string _docName;
        public string DocName
        {
            get { return _docName; }
            private set { _docName = value; OnPropertyChanged("DocName"); }
        }

        public string ModeText { get { return _selectionMode ? "选区模式 · 仅处理选中文本" : ""; } }

        // ---- 方案 ----

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

        public string ProfileDescription { get { return _selectedProfile == null ? "" : (_selectedProfile.Description ?? ""); } }

        // ---- 统计 ----

        public ObservableCollection<StatChip> Chips { get; private set; }

        private string _scanMeta;
        public string ScanMeta
        {
            get { return _scanMeta; }
            private set { _scanMeta = value; OnPropertyChanged("ScanMeta"); }
        }

        private string _noteText;
        public string NoteText
        {
            get { return _noteText; }
            private set { _noteText = value; OnPropertyChanged("NoteText"); OnPropertyChanged("HasNote"); }
        }

        public bool HasNote { get { return !string.IsNullOrEmpty(_noteText); } }

        // ---- 列表 ----

        public ObservableCollection<ParagraphRowVM> Rows { get; private set; }

        private string _searchText = "";
        public string SearchText
        {
            get { return _searchText; }
            set { _searchText = value; OnPropertyChanged("SearchText"); RefreshFilter(); }
        }

        private bool _showOnlyLowConfidence;
        public bool ShowOnlyLowConfidence
        {
            get { return _showOnlyLowConfidence; }
            set { _showOnlyLowConfidence = value; OnPropertyChanged("ShowOnlyLowConfidence"); OnPropertyChanged("LowConfidenceText"); RefreshFilter(); }
        }

        public string LowConfidenceText
        {
            get
            {
                int n = Rows == null ? 0 : Rows.Count(r => r.LowConfidence && !r.UserChanged);
                return n > 0 ? string.Format("仅看需确认（{0}）", n) : "仅看需确认";
            }
        }

        // ---- 选项 ----

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

        // ---- 结果页 ----

        private bool _resultOk;
        public bool ResultOk
        {
            get { return _resultOk; }
            private set { _resultOk = value; OnPropertyChanged("ResultOk"); }
        }

        private ObservableCollection<string> _resultLines = new ObservableCollection<string>();
        public ObservableCollection<string> ResultLines { get { return _resultLines; } }

        // ---- 命令 ----

        public ICommand ScanNowCommand { get; private set; }
        public ICommand CreateSampleCommand { get; private set; }
        public ICommand RescanCommand { get; private set; }
        public ICommand ApplyCommand { get; private set; }
        public ICommand CancelCommand { get; private set; }
        public ICommand FinishCommand { get; private set; }

        public bool HasScan { get { return _features != null && Rows.Count > 0; } }

        public PreviewViewModel()
        {
            Profiles = new ObservableCollection<Profile>();
            Chips = new ObservableCollection<StatChip>();
            Rows = new ObservableCollection<ParagraphRowVM>();
            ScanNowCommand = new RelayCommand(_ => ScanNow());
            CreateSampleCommand = new RelayCommand(_ => CreateSample());
            RescanCommand = new RelayCommand(_ => ScanNow());
            ApplyCommand = new RelayCommand(_ => Apply(), _ => HasScan && SelectedProfile != null);
            CancelCommand = new RelayCommand(_ => Hide());
            FinishCommand = new RelayCommand(_ => Hide());
        }

        // ---- 初始化 ----

        /// <summary>面板初次显示时设置默认选项（未扫描状态 → 欢迎页）。</summary>
        public void LoadDefaults(FormattingController controller)
        {
            if (controller == null) return;
            _controller = controller;
            if (controller.Settings != null) MountStyles = controller.Settings.MountStylesDefault;
            if (Page == PanePage.Welcome)
            {
                DocName = controller.ActiveDocumentName();
                OnPropertyChanged("CanScanCurrent");
            }
        }

        /// <summary>显示欢迎页（无文档/首次打开）。</summary>
        public void ShowWelcome(FormattingController controller)
        {
            _controller = controller;
            DocName = controller == null ? null : controller.ActiveDocumentName();
            OnPropertyChanged("CanScanCurrent");
            Page = PanePage.Welcome;
        }

        public bool CanScanCurrent { get { return !string.IsNullOrEmpty(DocName); } }

        private void ScanNow()
        {
            if (_controller == null) return;
            _controller.Run(_selectionMode);
        }

        private void CreateSample()
        {
            if (_controller == null) return;
            _controller.CreateSampleDocument();
        }

        private void Hide()
        {
            if (_controller == null) return;
            _controller.HidePreviewPane();
        }

        // ---- 扫描结果载入 ----

        public void LoadScan(FormattingController controller, ScanResult scan, bool selectionMode, string note, Range restrict)
        {
            if (controller == null) throw new ArgumentNullException("controller");
            _controller = controller;
            _features = scan.Features;
            _restrict = restrict;
            _selectionMode = selectionMode;
            OnPropertyChanged("ModeText");

            DocName = controller.ActiveDocumentName();
            NoteText = note ?? (controller.ProfileErrors != null && controller.ProfileErrors.Count > 0
                ? "方案加载警告：" + string.Join("；", controller.ProfileErrors)
                : null);

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

            if (controller.Settings != null) MountStyles = controller.Settings.MountStylesDefault;

            RebuildRows(scan);
            Page = PanePage.Preview;
        }

        private void RebuildRows(ScanResult scan)
        {
            Rows.Clear();
            Chips.Clear();

            var counts = new Dictionary<ParagraphRole, int>();
            int formulaSegments = 0;
            int skipped = 0;
            bool hasHeadings = false;

            foreach (var f in scan.Features)
            {
                counts[f.Role] = counts.ContainsKey(f.Role) ? counts[f.Role] + 1 : 1;

                if (f.Role == ParagraphRole.Skip)
                {
                    skipped++;
                    if (f.IsEmpty) continue; // 空段不进清单
                }
                if (f.Role.IsHeading() || f.Role == ParagraphRole.Title) hasHeadings = true;

                Rows.Add(new ParagraphRowVM
                {
                    Index = f.Index,
                    Preview = MakePreview(f),
                    Role = f.Role,
                    LowConfidence = !f.HighConfidence && f.Role != ParagraphRole.Skip,
                    InTable = f.InTable
                });

                if (IsConvertible(f.Role))
                {
                    if (f.Role == ParagraphRole.FormulaLine) formulaSegments++;
                    else if (!f.IsEmpty && !f.HasFields)
                        formulaSegments += MathSegmenter.FindSegments(f.Text).Count;
                }
            }

            AddChip(ParagraphRole.Title, "主标题");
            AddChip(ParagraphRole.Heading1, "一级标题");
            AddChip(ParagraphRole.Heading2, "二级标题");
            AddChip(ParagraphRole.Heading3, "三级标题");
            AddChip(ParagraphRole.Body, "正文");
            AddChip(ParagraphRole.Caption, "图注/表注");
            AddChip(ParagraphRole.ListItem, "列表项");
            AddChip(ParagraphRole.TableHeader, "表头");
            AddChip(ParagraphRole.TableText, "表格文字");
            AddChip(ParagraphRole.FormulaLine, "公式行");

            if (formulaSegments > 0)
                Chips.Add(new StatChip { Label = "公式片段将转换", Count = formulaSegments, Dot = Brushes("#DB2777"), ToolTip = "识别出的纯文本公式，应用时将转为可编辑的公式对象" });
            if (scan.TableCount > 0)
                Chips.Add(new StatChip { Label = "表格", Count = scan.TableCount, Dot = Brushes("#64748B"), ToolTip = "表格结构不会被修改，仅调整字体" });

            _hasHeadingsAfterApply = hasHeadings;
            _roleCounts = counts;

            var meta = new List<string>();
            meta.Add("段落 " + scan.Features.Count);
            if (skipped > 0) meta.Add("跳过 " + skipped);
            meta.Add("耗时 " + scan.Duration.TotalMilliseconds.ToString("0") + "ms");
            ScanMeta = string.Join(" · ", meta);

            OnPropertyChanged("LowConfidenceText");
            RefreshFilter();
        }

        private bool _hasHeadingsAfterApply;
        private Dictionary<ParagraphRole, int> _roleCounts;

        private static Brush Brushes(string hex)
        {
            var converter = new BrushConverter();
            return (Brush)converter.ConvertFromString(hex);
        }

        private void AddChip(ParagraphRole role, string label)
        {
            int count;
            if (_roleCounts == null || !_roleCounts.TryGetValue(role, out count) || count == 0) return;
            string dotHex;
            if (!_RoleDotColors.TryGetValue(role, out dotHex)) dotHex = "#6B7280";
            Chips.Add(new StatChip { Label = label, Count = count, Dot = Brushes(dotHex), ToolTip = label });
        }

        private static readonly Dictionary<ParagraphRole, string> _RoleDotColors = new Dictionary<ParagraphRole, string>
        {
            { ParagraphRole.Title, "#7C3AED" },
            { ParagraphRole.Heading1, "#1D4ED8" },
            { ParagraphRole.Heading2, "#2563EB" },
            { ParagraphRole.Heading3, "#3B82F6" },
            { ParagraphRole.Body, "#6B7280" },
            { ParagraphRole.Caption, "#0D9488" },
            { ParagraphRole.ListItem, "#CA8A04" },
            { ParagraphRole.TableHeader, "#64748B" },
            { ParagraphRole.TableText, "#64748B" },
            { ParagraphRole.FormulaLine, "#DB2777" }
        };

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
            if (text.Length > 40) text = text.Substring(0, 40) + "…";
            return text;
        }

        private void RefreshFilter()
        {
            var view = CollectionViewSource.GetDefaultView(Rows);
            if (view == null) return;
            string keyword = (SearchText ?? "").Trim();
            bool lowOnly = ShowOnlyLowConfidence;
            view.Filter = o =>
            {
                var r = o as ParagraphRowVM;
                if (r == null) return false;
                if (keyword.Length > 0)
                    return r.Preview != null && r.Preview.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0;
                if (lowOnly)
                    return r.LowConfidence || r.UserChanged;
                return true;
            };
        }

        // ---- 应用 ----

        private void Apply()
        {
            if (_controller == null || SelectedProfile == null) return;

            var errors = ProfileValidator.Validate(SelectedProfile);
            if (errors.Count > 0)
            {
                ResultLines.Clear();
                foreach (var e in errors) ResultLines.Add(e);
                ResultOk = false;
                Page = PanePage.Result;
                return;
            }

            var assignments = new Dictionary<int, ParagraphRole>();
            foreach (var row in Rows)
                assignments[row.Index] = row.Role;

            var outcome = _controller.Apply(assignments, SelectedProfile, MountStyles, ConvertFormulas, _restrict);

            ResultLines.Clear();
            if (outcome.Error != null)
            {
                ResultOk = false;
                ResultLines.Add(outcome.Error);
                ResultLines.Add("文档未被修改。");
            }
            else
            {
                ResultOk = true;
                var s = outcome.Stats;
                var roleParts = s.ParagraphsByRole.Where(kv => kv.Value > 0).Select(kv => kv.Key + " " + kv.Value);
                ResultLines.Add(string.Join("，", roleParts));
                if (s.StyleDefinitionsUpdated > 0)
                    ResultLines.Add("更新了 " + s.StyleDefinitionsUpdated + " 个样式定义（后续新段落自动跟随）");
                if (s.FormulasConverted > 0)
                    ResultLines.Add("公式转换 " + s.FormulasConverted + " 处" + (s.FormulasFailed > 0 ? "，" + s.FormulasFailed + " 处失败已保留原文" : ""));
                if (_hasHeadingsAfterApply)
                    ResultLines.Add("大纲已生成：视图 → 导航窗格 可查看目录结构");
            }
            Page = PanePage.Result;
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
