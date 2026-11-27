using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using TypeRight.Addin.Controllers;
using TypeRight.Profiles;

namespace TypeRight.Addin.UI
{
    /// <summary>
    /// 「另存为方案」窗口 —— 只做一件事：把刚才那次格式修改存下来。
    ///
    /// 设计依据（用户原话）：
    ///   「我把当前『仿宋 黑色 四号 无加粗』的字体类字号转变成了三号，
    ///     存下来之后，下一份文档直接点这个方案，就自动把同类文字的字号变成三号。」
    ///
    /// 所以这个窗口只需要：
    ///   1. 把这次改动原样复述一遍给用户看（确认存对了）
    ///   2. 给个名字
    ///   3. 存
    ///
    /// 这里**不出现**：方案列表、应用到当前文档、导入导出、兜底段落、公式转换开关。
    /// 那些属于「方案管理」，是另一个语境（用户此时的心思是"存下来"，不是"挑一个用"）。
    /// </summary>
    public sealed class SaveProfileWindow : ThemedWindow
    {
        private readonly TypeRightController _controller;
        private TextBox _txtName;
        private TextBlock _status;
        private Button _btnSave;

        public SaveProfileWindow(TypeRightController controller)
        {
            _controller = controller;
            Title = "存为方案";
            Width = 560;
            SizeToContent = SizeToContent.Height;
            ResizeMode = ResizeMode.NoResize;

            Build();
            Loaded += (s, e) => { _txtName.Focus(); _txtName.SelectAll(); };
        }

        /// <summary>主题要跟随宿主 Word，所以必须把 Application 交出去。</summary>
        protected override NetOffice.WordApi.Application ResolveApp()
        {
            try { return _controller.App; } catch { return null; }
        }

        private Border _previewCard;

        private void Build()
        {
            var root = new StackPanel { Margin = new Thickness(20, 18, 20, 18) };

            // ── 标题 ──
            root.Children.Add(TitleText("存为方案"));
            var sub = CaptionText("下次遇到同类文档，点一下这个方案就能还原这套格式。");
            sub.Margin = new Thickness(0, 4, 0, 14);
            root.Children.Add(sub);

            // ── 这次改了什么（最重要的一块：让用户确认）──
            Border preview = BuildPreview();
            if (preview != null)
            {
                preview.Margin = new Thickness(0, 0, 0, 16);
                root.Children.Add(preview);
            }

            // ── 名字 ──
            var nameLabel = CaptionText("给这个方案起个名字");
            nameLabel.Margin = new Thickness(0, 0, 0, 6);
            root.Children.Add(nameLabel);

            _txtName = StyledTextBox(DefaultName());
            _txtName.KeyDown += (s, e) =>
            {
                if (e.Key == Key.Enter) { e.Handled = true; DoSave(); }
            };
            root.Children.Add(_txtName);

            // ── 按钮 ──
            var buttons = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 16, 0, 0)
            };
            var cancel = SecondaryButton("取消");
            cancel.IsCancel = true;
            cancel.Margin = new Thickness(0, 0, 8, 0);
            cancel.Click += (s, e) => Close();
            buttons.Children.Add(cancel);

            _btnSave = PrimaryButton("存下来");
            _btnSave.IsDefault = true;
            _btnSave.Click += (s, e) => DoSave();
            buttons.Children.Add(_btnSave);
            root.Children.Add(buttons);

            _status = CaptionText("");
            _status.Margin = new Thickness(0, 10, 0, 0);
            root.Children.Add(_status);

            Content = root;
        }

        /// <summary>把这次操作复述给用户：「正文里 13 段 · 仿宋 黑色 四号 不加粗 → 字号改成三号」。
        /// 用户看到这句才能确认自己存的是不是想要的那件事。</summary>
        private Border BuildPreview()
        {
            ProfileV2 draft;
            try
            {
                draft = _controller.BuildPendingProfile("(预览)", false, null, false);
            }
            catch { return null; }
            if (draft == null || draft.Rules == null || draft.Rules.Count == 0) return null;

            var inner = new StackPanel();

            var head = new TextBlock
            {
                Text = "这次改了",
                FontSize = 11.5,
                FontWeight = FontWeights.SemiBold,
                Margin = new Thickness(0, 0, 0, 8)
            };
            inner.Children.Add(head);

            foreach (ProfileRule rule in draft.Rules)
            {
                RuleSummary rs = ProfileDescriber.DescribeRule(rule);
                if (string.IsNullOrEmpty(rs.MatchText)) continue;

                var line = new StackPanel { Margin = new Thickness(0, 0, 0, 8) };

                // 「挑中什么样的段落」
                var matchLine = new TextBlock
                {
                    Text = rs.MatchText,
                    FontSize = 12.5,
                    TextWrapping = TextWrapping.Wrap
                };
                UseResource(matchLine, TextBlock.ForegroundProperty, "TextBrush");
                line.Children.Add(matchLine);

                // 「改成了什么」——用强调色，这是用户最关心的一句
                if (!string.IsNullOrEmpty(rs.ChangeText))
                {
                    var changeLine = new TextBlock
                    {
                        Text = "→ " + rs.ChangeText,
                        FontSize = 13,
                        FontWeight = FontWeights.SemiBold,
                        TextWrapping = TextWrapping.Wrap,
                        Margin = new Thickness(0, 3, 0, 0)
                    };
                    UseResource(changeLine, TextBlock.ForegroundProperty, "AccentBrush");
                    line.Children.Add(changeLine);
                }
                inner.Children.Add(line);
            }

            _previewCard = Card(inner);
            return _previewCard;
        }

        private string DefaultName()
        {
            try
            {
                var doc = _controller.App != null ? _controller.App.ActiveDocument : null;
                if (doc != null)
                {
                    string name = System.IO.Path.GetFileNameWithoutExtension(doc.Name);
                    if (!string.IsNullOrWhiteSpace(name)) return name;
                }
            }
            catch { }
            return "我的方案 " + DateTime.Now.ToString("MMdd");
        }

        private void DoSave()
        {
            string name = (_txtName.Text ?? "").Trim();
            if (name.Length == 0)
            {
                _status.Text = "给它起个名字，下次才找得到。";
                _status.Foreground = B("DangerBrush");
                _txtName.Focus();
                return;
            }

            try
            {
                // 这里不带兜底、不带公式开关——那是方案管理里的事，另存为只忠实记录这次改动
                string file = _controller.SaveProfile(name, false, null, false);
                if (file == null)
                {
                    _status.Text = "没有可保存的内容。先做一次「同类选择」改格式再回来。";
                    _status.Foreground = B("DangerBrush");
                    return;
                }

                _btnSave.IsEnabled = false;
                string savedAs = System.IO.Path.GetFileNameWithoutExtension(file);
                bool renamed = savedAs != ProfileId.Make(name);

                _status.Text = "已存好「" + name + "」"
                    + (renamed ? "（已有同名，这份存为 " + savedAs + "）" : "")
                    + "，可以关掉这个窗口了。";
                _status.Foreground = B("SuccessBrush");

                // 存完自动关闭：这件事办完了，不该让用户再找"关闭"按钮
                var timer = new System.Windows.Threading.DispatcherTimer
                {
                    Interval = TimeSpan.FromMilliseconds(900)
                };
                timer.Tick += (s, e) => { timer.Stop(); Close(); };
                timer.Start();
            }
            catch (ProfileValidationException ex)
            {
                _status.Text = string.Join("；", ex.Errors);
                _status.Foreground = B("DangerBrush");
            }
            catch (Exception ex)
            {
                _status.Text = "保存失败：" + ex.Message;
                _status.Foreground = B("DangerBrush");
            }
        }
    }
}
