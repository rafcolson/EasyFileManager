using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using WinFormsLib;
using static WinFormsLib.Constants;
using static WinFormsLib.DateTimeExtensions;
using static WinFormsLib.Forms;
using static WinFormsLib.ToolStripMenuItems;
using static WinFormsLib.Utils;

namespace EasyFileManager
{
    public partial class Main : Form
    {
        #region Init

        private const string BACKUP_PREFIX = "_EFM_BACKUP_";
        private const string DUPLICATES_PREFIX = "_EFM_DUPLICATES_";

        private static readonly List<string> _history = [];
        private static readonly List<string> _selectedPaths = [];

        private static CancellationTokenSource _cancellationTokenSource = new();
        private static Task _task = Task.CompletedTask;
        private static int _historyIndex = -1;

        private static int _customizeSplitContainerHeight;
        private static int _renameSplitContainerHeight;
        private static int _copyMoveSplitContainerHeight;
        private static int _finalizeSplitContainerHeight;

        private static bool _isClosing;
        private static bool _isUpdating;
        private static bool _isApplying;

        private static EditTextContextMenuStrip? _editTextBoxContextMenuStrip;
        private static EditTextContextMenuStrip? _editPathContextMenuStrip;
        private static EditTextContextMenuStrip? _editPropsContextMenuStrip;
        private static EditTextContextMenuStrip? _editExplorerContextMenuStrip;

        private readonly string[] _args;
        private readonly System.Windows.Forms.Timer _previewFormattingTimer = new() { Interval = 500 };

        public static AdaptiveFileSystemWatcher FolderWatcher { get; } = new();
        public static BackgroundWorker ProgressBackgroundWorker { get; } = new()
        {
            WorkerReportsProgress = true
        };
        public static EasyProgress Progress { get; set; } = new(x =>
        {
            if (_isClosing) { Exit(); return; }
            ProgressBackgroundWorker.ReportProgress(x);
        });
        public static EasyLogger Logger { get; private set; } = new();
        public static EasyOptions Options { get; private set; } = new();
        public static EasyFolder Folder { get; private set; } = new();
        public static EasyFolders Folders { get; } = [];
        public static EasyFiles Files { get; } = [];
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public bool IsUpdating
        {
            get => _isUpdating;
            set
            {
                _isUpdating = value;
                Cursor = _isUpdating ? Cursors.WaitCursor : Cursors.Default;
            }
        }
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public bool IsApplying
        {
            get => _isApplying;
            set
            {
                _isApplying = value;
                FolderWatcher.EnableRaisingEvents = !_isApplying;
                Cursor = _isApplying ? Cursors.WaitCursor : Cursors.Default;
            }
        }

        public Main(string[] args)
        {
            _args = args;
            StartInitialization();
        }

        private void Main_Load(object? sender, EventArgs e) => FinishInitialization();

        protected override void OnDpiChanged(DpiChangedEventArgs e)
        {
            base.OnDpiChanged(e);
            _ = BeginInvoke((MethodInvoker)delegate
            {
                InitializeRowHeaders();
            });
        }

        private void InitializeRowHeaders()
        {
            ExplorerDataGridView.ResizeRowHeaders(DeviceDpi);
            PropsDataGridView.ResizeRowHeaders(DeviceDpi);
        }

        private static void InitializeLanguage()
        {
            Language language = (Language)Properties.Settings.Default.LanguageIndex;
            if (language.GetValue<string>() is string twoLetterISOLanguageName)
            {
                CultureInfo ci = CultureInfo.GetCultureInfo(twoLetterISOLanguageName);
                Thread.CurrentThread.CurrentCulture = ci;
                Thread.CurrentThread.CurrentUICulture = ci;
            }
        }

        private void InitializeFields()
        {
            _customizeSplitContainerHeight = CustomizeSplitContainer.Height;
            _renameSplitContainerHeight = RenameSplitContainer.Height;
            _copyMoveSplitContainerHeight = CopyMoveSplitContainer.Height;
            _finalizeSplitContainerHeight = FinalizeSplitContainer.Height;

            _editTextBoxContextMenuStrip = new(Font);
            _editPathContextMenuStrip = new(Font);
            if (GetToolStripMenuItem(Globals.Explore, PathToolStripMenuItem_Click) is ToolStripMenuItem exploreMenuItem)
            {
                _ = _editPathContextMenuStrip.Items.Add(exploreMenuItem);
            }
            _editPropsContextMenuStrip = new(Font, EditMenuItems.CopySelectAll);
            _editExplorerContextMenuStrip = new(Font, EditMenuItems.CopySelectAll);
            ToolStripMenuItem? explorerMenuItem = GetToolStripMenuItem(Globals.Explore, (o, e) => ExplorerToolStripMenuItem_Click(o, e));
            if (explorerMenuItem is not null)
            {
                _ = _editExplorerContextMenuStrip.Items.Add(explorerMenuItem);
            }
            ToolStripMenuItem? openMenuItem = GetToolStripMenuItem(Globals.Open, (o, e) => ExplorerToolStripMenuItem_Click(o, e, true));
            if (openMenuItem is not null)
            {
                _ = _editExplorerContextMenuStrip.Items.Add(openMenuItem);
            }
            _editExplorerContextMenuStrip.Opening += EditExplorerContextMenuStrip_Opening;
            _previewFormattingTimer.Tick += PreviewFormattingTimer_TickAsync;
        }

        private void InitializeMainEventHandlers()
        {
            Load += Main_Load;
            FormClosing += Main_FormClosing;
            FormClosed += Main_FormClosed;
            Disposed += Main_Disposed;
        }

        private void InitializeLayout()
        {
            ShowOutputToolStripMenuItem.Checked = Options.Show.HasFlag(EasyShow.Output);
            ShowThumbnailToolStripMenuItem.Checked = Options.Show.HasFlag(EasyShow.Thumbnail);
            ShowPropertiesToolStripMenuItem.Checked = Options.Show.HasFlag(EasyShow.Properties);
            ShowMillisecondsToolStripMenuItem.Checked = Options.Show.HasFlag(EasyShow.Milliseconds);
            ShowHiddenItemsToolStripMenuItem.Checked = Options.Show.HasFlag(EasyShow.HiddenItems);

            ExtractEmbeddedImageToolStripMenuItem.Checked = Options.ExtractEmbeddedMetadata.HasFlag(ExtractEmbeddedMetadata.Image);
            ExtractEmbeddedVideoToolStripMenuItem.Checked = Options.ExtractEmbeddedMetadata.HasFlag(ExtractEmbeddedMetadata.Video);
            ExtractEmbeddedAudioToolStripMenuItem.Checked = Options.ExtractEmbeddedMetadata.HasFlag(ExtractEmbeddedMetadata.Audio);
            ExtractEmbeddedDocumentToolStripMenuItem.Checked = Options.ExtractEmbeddedMetadata.HasFlag(ExtractEmbeddedMetadata.Document);
            ConvertEmbeddedImageToolStripMenuItem.Checked = Options.ConvertEmbeddedToEasyMetadata.HasFlag(ConvertEmbeddedToEasyMetadata.Image);
            ConvertEmbeddedVideoToolStripMenuItem.Checked = Options.ConvertEmbeddedToEasyMetadata.HasFlag(ConvertEmbeddedToEasyMetadata.Video);
            ConvertEmbeddedAudioToolStripMenuItem.Checked = Options.ConvertEmbeddedToEasyMetadata.HasFlag(ConvertEmbeddedToEasyMetadata.Audio);
            ConvertEmbeddedDocumentToolStripMenuItem.Checked = Options.ConvertEmbeddedToEasyMetadata.HasFlag(ConvertEmbeddedToEasyMetadata.Document);
            PreserveDateCreatedToolStripMenuItem.Checked = Options.PreserveDateCreated;
            PreserveDateModifiedToolStripMenuItem.Checked = Options.PreserveDateModified;
            LogApplicationEventsToolStripMenuItem.Checked = Properties.Settings.Default.LogApplicationEvents;
            ShutDownUponCompletionToolStripMenuItem.Checked = Properties.Settings.Default.ShutdownUponCompletion;

            PathTextBox.ContextMenuStrip = _editPathContextMenuStrip;
            ExplorerDataGridView.ContextMenuStrip = _editExplorerContextMenuStrip;
            PropsDataGridView.ContextMenuStrip = _editPropsContextMenuStrip;

            CustomizeCheckBox.Checked = Options.CustomizeEnabled;

            TitleCheckBox.Checked = Options.TitleEnabled;
            TitleTextBox.Enabled = Options.TitleEnabled;
            TitleTextBox.Text = Options.Title;
            TitleTextBox.ContextMenuStrip = _editTextBoxContextMenuStrip;

            SubjectCheckBox.Checked = Options.SubjectEnabled;
            SubjectTextBox.Enabled = Options.SubjectEnabled;
            SubjectTextBox.Text = Options.Subject;
            SubjectTextBox.ContextMenuStrip = _editTextBoxContextMenuStrip;

            CommentCheckBox.Checked = Options.CommentEnabled;
            CommentTextBox.Enabled = Options.CommentEnabled;
            CommentTextBox.Text = Options.Comment;
            CommentTextBox.ContextMenuStrip = _editTextBoxContextMenuStrip;

            KeywordsCheckBox.CheckState = Options.KeywordsState;
            UpdateKeywordsControlsA();

            GeolocationCheckBox.Checked = Options.GeolocationEnabled;
            WriteGPSCoordsCheckBox.Checked = Options.WriteGPSCoords;
            WriteGPSAreaInfoCheckBox.Checked = Options.WriteGPSAreaInfo;
            WriteGPSEasyMetadataCheckBox.Checked = Options.WriteGPSEasyMetadata;
            GPSSourceComboBox.Items.Clear();
            GPSSourceComboBox.Items.AddRange([.. Enum.GetValues<EasyMetadataSource>().Select(x => x.GetEasyGlobalStringValue())]);
            GPSSourceComboBox.Dock = DockStyle.None;
            GPSSourceComboBox.UpdateDropDownList();
            GPSSourceComboBox.Dock = DockStyle.Fill;
            GPSSourceComboBox.SelectedIndex = Options.GPSSourceIndex;
            UpdateGPSControlsA();

            EasyMetadataCheckBox.Checked = Options.EasyMetadataEnabled;
            UpdateEasyMetadataControlsA();

            DateCheckBox.CheckState = Options.DateState;
            WriteDateCreatedCheckBox.Checked = Options.WriteDateCreated;
            WriteDateTakenOrEncodedCheckBox.Checked = Options.WriteDateCamera;
            WriteDateModifiedCheckBox.Checked = Options.WriteDateModified;
            WriteDateEasyMetadataCheckBox.Checked = Options.WriteDateEasyMetadata;
            UpdateDateControlsA();

            RenameCheckBox.Checked = Options.RenameEnabled;

            PrefixTextBox.Text = Options.Prefix;
            PrefixTextBox.ContextMenuStrip = _editTextBoxContextMenuStrip;

            ReplaceWithCheckBox.CheckState = Options.ReplaceWithState;
            ReplaceTextBox.Text = Options.Replace;
            ReplaceTextBox.ContextMenuStrip = _editTextBoxContextMenuStrip;
            WithTextBox.Text = Options.With;
            WithTextBox.ContextMenuStrip = _editTextBoxContextMenuStrip;
            UpdateReplaceWithControls();

            DateFormatCheckBox.Checked = Options.DateFormatEnabled;
            DateFormatComboBox.Enabled = Options.DateFormatEnabled;
            DateFormatComboBox.Items.Clear();
            DateFormatComboBox.Items.AddRange([.. Enum.GetValues<DateTimeFormat>().Select(x => x.GetValue<string>()).Cast<object>()]);
            DateFormatComboBox.Dock = DockStyle.None;
            DateFormatComboBox.UpdateDropDownList();
            DateFormatComboBox.Dock = DockStyle.Fill;
            DateFormatComboBox.SelectedItem = Options.DateFormat;

            SuffixTextBox.Text = Options.Suffix;
            SuffixTextBox.ContextMenuStrip = _editTextBoxContextMenuStrip;

            CopyMoveCheckBox.CheckState = Options.CopyMoveState;
            UpdateCopyMoveCheckBox();

            BackupFolderCheckBox.Checked = Options.BackupFolderEnabled;
            BackupFolderTextBox.Text = Options.BackupFolderPath;
            UpdateBackupControls();

            TopFolderCheckBox.Checked = Options.TopFolderEnabled;
            TopFolderTextBox.Text = Options.TopFolderPath;
            SubfoldersCheckBox.Checked = Options.SubfoldersEnabled;
            UpdateFolderControlsA();

            FilterCheckBox.Checked = Options.FilterEnabled;
            FilterNameComboBox.Items.Clear();
            FilterNameComboBox.Items.AddRange([.. Enum.GetValues<EasyNameFilter>().Select(x => x.GetEasyGlobalStringValue())]);
            FilterNameComboBox.Dock = DockStyle.None;
            FilterNameComboBox.UpdateDropDownList();
            FilterNameComboBox.Dock = DockStyle.Fill;
            FilterNameComboBox.SelectedIndex = (int)Options.NameFilter;
            FilterStringTextBox.Text = Options.FilterString;
            UpdateFilterControlsA();

            FinalizeCheckBox.Checked = Options.FinalizeEnabled;
            CleanUpCheckBox.Checked = Options.CleanUpEnabled;
            UpdateCleanUpControlsA();

            DuplicatesCheckBox.CheckState = Options.DuplicatesState;
            DuplicatesFolderTextBox.Text = Options.DuplicatesFolderPath;
            ShowDuplicatesCompareDialogCheckBox.Checked = Options.ShowDuplicatesCompareDialog;
            UpdateDuplicatesControlsA();

            UpdateCustomizePanel();
            UpdateRenamePanel();
            UpdateCopyMovePanel();
            UpdateFinalizePanel();

            UpdateEditPanel(true, true, true, true);

            _ = Update(Properties.Settings.Default.StartupDirectory, true);

            ActiveControl = null;
        }

        private void InitializeToolTips()
        {
            PreviousFolderButton.AddToolTip(Globals.PreviousFolder);
            NextFolderButton.AddToolTip(Globals.NextFolder);
            OneLevelUpButton.AddToolTip(Globals.OneLevelUp);
            SelectFolderButton.AddToolTip(Globals.SelectFolder);
            SelectTopFolderButton.AddToolTip(Globals.SelectTopFolder);
            KeywordsEditButton.AddToolTip(Globals.EditKeywords);
            GPSEditButton.AddToolTip(Globals.EditGPS);
            EasyMetadataEditButton.AddToolTip(Globals.EditEasyMetadata);
            SelectBackupFolderButton.AddToolTip(Globals.SelectBackupFolder);
            SubfoldersEditButton.AddToolTip(Globals.EditSubfolders);
            FilterEditButton.AddToolTip(Globals.EditFilter);
            SelectDuplicatesFolderButton.AddToolTip(Globals.SelectDuplicatesFolder);
            DuplicatesCompareEditButton.AddToolTip(Globals.EditComparisonProfile);

            WriteGPSCoordsCheckBox.AddToolTip(Globals.WriteGPSCoords);
            WriteGPSAreaInfoCheckBox.AddToolTip(Globals.WriteGPSAreaInfo);
            WriteGPSEasyMetadataCheckBox.AddToolTip(Globals.WriteGPSEasyMetadata);
            WriteDateModifiedCheckBox.AddToolTip(Globals.WriteDateModified);
            WriteDateCreatedCheckBox.AddToolTip(Globals.WriteDateCreated);
            WriteDateTakenOrEncodedCheckBox.AddToolTip(Globals.WriteDateTakenOrEncoded);
            WriteDateEasyMetadataCheckBox.AddToolTip(Globals.WriteDateEasyMetadata);
            ShowDuplicatesCompareDialogCheckBox.AddToolTip(Globals.ShowComparisonDialog);
        }

        private void InitializeMinorEventHandlers()
        {
            ProgressBackgroundWorker.ProgressChanged += ProgressBackgroundWorker_ProgressChanged;
            FolderWatcher.Changed += FolderWatcher_Changed;

            ImportToolStripMenuItem.Click += ImportToolStripMenuItem_Click;
            ExportToolStripMenuItem.Click += ExportToolStripMenuItem_Click;
            ExitToolStripMenuItem.Click += ExitToolStripMenuItem_Click;

            ShowOutputToolStripMenuItem.Click += ShowOutputToolStripMenuItem_Click;
            ShowThumbnailToolStripMenuItem.Click += ShowThumbnailToolStripMenuItem_Click;
            ShowHiddenItemsToolStripMenuItem.Click += ShowHiddenItemsToolStripMenuItem_Click;
            ShowPropertiesToolStripMenuItem.Click += ShowPropertiesToolStripMenuItem_Click;
            ShowMillisecondsToolStripMenuItem.Click += ShowMillisecondsToolStripMenuItem_Click;
            LanguageToolStripMenuItem.Click += LanguageToolStripMenuItem_Click;

            ExtractEmbeddedImageToolStripMenuItem.Click += ExtractEmbeddedImageToolStripMenuItem_Click;
            ExtractEmbeddedVideoToolStripMenuItem.Click += ExtractEmbeddedVideoToolStripMenuItem_Click;
            ExtractEmbeddedAudioToolStripMenuItem.Click += ExtractEmbeddedAudioToolStripMenuItem_Click;
            ExtractEmbeddedDocumentToolStripMenuItem.Click += ExtractEmbeddedDocumentToolStripMenuItem_Click;
            ConvertEmbeddedImageToolStripMenuItem.Click += ConvertEmbeddedImageToolStripMenuItem_Click;
            ConvertEmbeddedVideoToolStripMenuItem.Click += ConvertEmbeddedVideoToolStripMenuItem_Click;
            ConvertEmbeddedAudioToolStripMenuItem.Click += ConvertEmbeddedAudioToolStripMenuItem_Click;
            ConvertEmbeddedDocumentToolStripMenuItem.Click += ConvertEmbeddedDocumentToolStripMenuItem_Click;
            PreserveDateCreatedToolStripMenuItem.Click += PreserveDateCreatedToolStripMenuItem_Click;
            PreserveDateModifiedToolStripMenuItem.Click += PreserveDateModifiedToolStripMenuItem_Click;
            LogApplicationEventsToolStripMenuItem.Click += LogApplicationEventsToolStripMenuItem_Click;
            ShutDownUponCompletionToolStripMenuItem.Click += ShutDownUponCompletionToolStripMenuItem_Click;

            SplashScreenToolStripMenuItem.Click += SplashScreenToolStripMenuItem_Click;
            RestoreToDefaultToolStripMenuItem.Click += RestoreDefaultToolStripMenuItem_Click;
            AboutEasyFileManagerToolStripMenuItem.Click += AboutEasyFileManagerToolStripMenuItem_Click;

            PreviousFolderButton.Click += PreviousFolderButton_Click;
            NextFolderButton.Click += NextFolderButton_Click;
            OneLevelUpButton.Click += OneLevelUpButton_Click;
            SelectFolderButton.Click += SelectFolderButton_Click;
            GPSEditButton.Click += GPSEditButton_Click;
            DateEditButton.Click += DateEditButton_Click;
            EasyMetadataEditButton.Click += EasyMetadataEditButton_Click;
            ApplyButton.Click += ApplyButton_ClickAsync;

            ExplorerDataGridView.SelectionChanged += ExplorerDataGridView_SelectionChanged;
            ExplorerDataGridView.CellDoubleClick += ExplorerDataGridView_CellDoubleClick;

            foreach (Button btn in RotateTableLayoutPanel.GetControls<Button>())
            {
                btn.Click += RotateButton_Click;
            }
            foreach (DataGridView dgv in this.GetControls<DataGridView>())
            {
                dgv.CellMouseDown += DataGridView_CellMouseDown;
            }
            foreach (TextBox tb in this.GetControls<TextBox>(true))
            {
                tb.KeyPress += TextBox_KeyPress;
                tb.MouseDown += TextBox_MouseDown;
            }
            PathTextBox.Leave += PathTextBox_Leave;

            CustomizeCheckBox.CheckedChanged += CustomizeCheckBox_CheckedChanged;
            RenameCheckBox.CheckedChanged += RenameCheckBox_CheckedChanged;
            CopyMoveCheckBox.CheckStateChanged += CopyMoveCheckBox_CheckStateChanged;
            FinalizeCheckBox.CheckedChanged += FinalizeCheckBox_CheckedChanged;

            TitleCheckBox.CheckedChanged += TitleCheckBox_CheckedChanged;
            TitleTextBox.Leave += TitleTextBox_Leave;

            SubjectCheckBox.CheckedChanged += SubjectCheckBox_CheckedChanged;
            SubjectTextBox.Leave += SubjectTextBox_Leave;

            CommentCheckBox.CheckedChanged += CommentCheckBox_CheckedChanged;
            CommentTextBox.Leave += CommentTextBox_Leave;

            KeywordsCheckBox.CheckStateChanged += KeywordsCheckBox_CheckStateChanged;
            KeywordsComboBox.SelectedIndexChanged += KeywordsComboBox_SelectedIndexChanged;
            KeywordsEditButton.Click += KeywordsEditButton_Click;

            GeolocationCheckBox.CheckedChanged += GeolocationCheckBox_CheckedChanged;
            GPSSourceComboBox.SelectedIndexChanged += GPSSourceComboBox_SelectedIndexChanged;
            GPSComboBox.SelectedIndexChanged += GPSComboBox_SelectedIndexChanged;
            WriteGPSCoordsCheckBox.CheckedChanged += WriteGPSCoordsCheckBox_CheckedChanged;
            WriteGPSAreaInfoCheckBox.CheckedChanged += WriteGPSAreaInfoCheckBox_CheckedChanged;
            WriteGPSEasyMetadataCheckBox.CheckedChanged += WriteGPSEasyMetadataCheckBox_CheckedChanged;

            DateCheckBox.CheckStateChanged += DateCheckBox_CheckStateChanged;
            DateComboBox.SelectedIndexChanged += DateComboBox_SelectedIndexChanged;
            WriteDateModifiedCheckBox.CheckedChanged += WriteDateModifiedCheckBox_CheckedChanged;
            WriteDateCreatedCheckBox.CheckedChanged += WriteDateCreatedCheckBox_CheckedChanged;
            WriteDateTakenOrEncodedCheckBox.CheckedChanged += WriteDateTakenOrEncodedCheckBox_CheckedChanged;
            WriteDateEasyMetadataCheckBox.CheckedChanged += WriteDateEasyMetadataCheckBox_CheckedChanged;

            EasyMetadataCheckBox.CheckedChanged += EasyMetadataCheckBox_CheckedChanged;
            EasyMetadataComboBox.SelectedIndexChanged += EasyMetadataComboBox_SelectedIndexChanged;

            PrefixTextBox.Leave += PrefixTextBox_Leave;
            ReplaceTextBox.Leave += ReplaceTextBox_Leave;
            WithTextBox.Leave += WithTextBox_Leave;
            SuffixTextBox.Leave += SuffixTextBox_Leave;

            ReplaceWithCheckBox.CheckStateChanged += ReplaceWithCheckBox_CheckStateChanged;

            DateFormatCheckBox.CheckedChanged += DateFormatCheckBox_CheckedChanged;
            DateFormatComboBox.SelectedIndexChanged += DateFormatComboBox_SelectedIndexChanged;

            BackupFolderCheckBox.CheckedChanged += BackupFolderCheckBox_CheckedChanged;
            TopFolderCheckBox.CheckedChanged += TopFolderCheckBox_CheckedChanged;
            SubfoldersCheckBox.CheckedChanged += SubfoldersCheckBox_CheckedChanged;
            FilterCheckBox.CheckedChanged += FilterCheckBox_CheckedChanged;

            SelectBackupFolderButton.Click += SelectBackupFolderButton_Click;
            SelectTopFolderButton.Click += SelectTopFolderButton_Click;
            SubfoldersEditButton.Click += SubfoldersEditButton_Click;
            FilterEditButton.Click += FilterEditButton_Click;

            FilterNameComboBox.SelectedIndexChanged += FilterNameComboBox_SelectedIndexChanged;
            FilterStringTextBox.Leave += FilterStringTextBox_Leave;

            CleanUpCheckBox.CheckedChanged += CleanUpCheckBox_CheckedChanged;
            DuplicatesCheckBox.CheckStateChanged += DuplicatesCheckBox_CheckStateChanged;
            ShowDuplicatesCompareDialogCheckBox.CheckedChanged += ShowDuplicatesCompareDialogCheckBox_CheckedChanged;

            CleanUpEditButton.Click += CleanUpEditButton_Click;
            SelectDuplicatesFolderButton.Click += SelectDuplicatesFolderButton_Click;
            DuplicatesCompareEditButton.Click += DuplicatesCompareEditButton_Click;
        }

        private void InitializeWatchers()
        {
            FolderWatcher.SynchronizingObject = this;
        }

        private void ShowSplash(bool hideMainForm = false, TimeSpan? timeSpan = null)
        {
            if (hideMainForm) { Visible = false; }
            using Splash splash = new(timeSpan);
            _ = splash.ShowDialog();
            if (hideMainForm) { Visible = true; }
        }

        private void StartInitialization()
        {
            Debug.WriteLine($"Initializing...");

            InitializeLanguage();
            InitializeComponent();
            Icon = Properties.Resources.ApplicationIcon;
            InitializeFields();
            //RestoreToDefault();
            Options = new(Properties.Settings.Default.EasyOptions);
            if (_args.Length == 1)
            {
                _ = UpdateOptions(_args[0]);
            }

            this.SuspendDrawing();
            InitializeMainEventHandlers();
            ShowSplash(true, TimeSpan.FromSeconds(2));
        }

        private void FinishInitialization()
        {
            InitializeWatchers();
            InitializeLayout();
            InitializeToolTips();
            _ = BeginInvoke((MethodInvoker)delegate
            {
                InitializeMinorEventHandlers();
                InitializeRowHeaders();
                this.ResumeDrawing();

                Debug.WriteLine($"Initialized '{Name}'.");
            });
        }

        #endregion

        #region Exit/Restart

        private void Main_FormClosing(object? sender, FormClosingEventArgs e)
        {
            SaveOptions();

            _isClosing = true;
            if (Tag is not CloseReason.ApplicationExitCall)
            {
                Tag = e.CloseReason;

                Debug.WriteLine($"Closing ({Tag}) '{Name}'.");
            }
            WaitUntilTaskCompleted();
        }

        private void Main_FormClosed(object? sender, FormClosedEventArgs e)
        {
            DisposeFields();
            DisposeMinorEventHandlers();

            Debug.WriteLine($"Closed ({Tag}) '{Name}'.");
        }

        private static void Exit()
        {
            WaitUntilTaskCompleted();

            string n = GetAssemblyName();
            CloseForms(true, CloseReason.ApplicationExitCall);
            Application.Exit();

            Debug.WriteLine($"Exited '{n}'.");

            Environment.Exit(0);
        }

        private static void Restart()
        {
            string n = GetAssemblyName();
            CloseForms(true, CloseReason.UserClosing);

            Debug.WriteLine($"Restarting '{n}'.");

            Application.Restart();
        }

        #endregion

        #region Dispose

        private void Main_Disposed(object? sender, EventArgs e)
        {
            DisposeMainEventHandlers();
            this.ClearEventHandlers();

            DisposeProperties();

            DisposeWinFormsLib();

            Debug.WriteLine($"Disposed.");
        }

        private void DisposeFields()
        {
            Debug.WriteLine($"Disposing '{Name}': fields...");

            _task.Dispose();
            _cancellationTokenSource.Dispose();
            _editTextBoxContextMenuStrip?.Dispose();
            _editPathContextMenuStrip?.Dispose();
            _editPropsContextMenuStrip?.Dispose();
            _editExplorerContextMenuStrip?.Dispose();
            _previewFormattingTimer.Dispose();
        }

        private void DisposeMinorEventHandlers()
        {
            Debug.WriteLine($"Disposing '{Name}': minor event handlers...");

            ProgressBackgroundWorker.ProgressChanged -= ProgressBackgroundWorker_ProgressChanged;
            FolderWatcher.Changed -= FolderWatcher_Changed;

            ImportToolStripMenuItem.Click -= ImportToolStripMenuItem_Click;
            ExportToolStripMenuItem.Click -= ExportToolStripMenuItem_Click;
            ExitToolStripMenuItem.Click -= ExitToolStripMenuItem_Click;

            ShowOutputToolStripMenuItem.Click -= ShowOutputToolStripMenuItem_Click;
            ShowThumbnailToolStripMenuItem.Click -= ShowThumbnailToolStripMenuItem_Click;
            ShowHiddenItemsToolStripMenuItem.Click -= ShowHiddenItemsToolStripMenuItem_Click;
            ShowPropertiesToolStripMenuItem.Click -= ShowPropertiesToolStripMenuItem_Click;
            ShowMillisecondsToolStripMenuItem.Click -= ShowMillisecondsToolStripMenuItem_Click;
            LanguageToolStripMenuItem.Click -= LanguageToolStripMenuItem_Click;

            ExtractEmbeddedImageToolStripMenuItem.Click -= ExtractEmbeddedImageToolStripMenuItem_Click;
            ExtractEmbeddedVideoToolStripMenuItem.Click -= ExtractEmbeddedVideoToolStripMenuItem_Click;
            ExtractEmbeddedAudioToolStripMenuItem.Click -= ExtractEmbeddedAudioToolStripMenuItem_Click;
            ExtractEmbeddedDocumentToolStripMenuItem.Click -= ExtractEmbeddedDocumentToolStripMenuItem_Click;
            ConvertEmbeddedImageToolStripMenuItem.Click -= ConvertEmbeddedImageToolStripMenuItem_Click;
            ConvertEmbeddedVideoToolStripMenuItem.Click -= ConvertEmbeddedVideoToolStripMenuItem_Click;
            ConvertEmbeddedAudioToolStripMenuItem.Click -= ConvertEmbeddedAudioToolStripMenuItem_Click;
            ConvertEmbeddedDocumentToolStripMenuItem.Click -= ConvertEmbeddedDocumentToolStripMenuItem_Click;
            PreserveDateCreatedToolStripMenuItem.Click -= PreserveDateCreatedToolStripMenuItem_Click;
            PreserveDateModifiedToolStripMenuItem.Click -= PreserveDateModifiedToolStripMenuItem_Click;
            LogApplicationEventsToolStripMenuItem.Click -= LogApplicationEventsToolStripMenuItem_Click;
            ShutDownUponCompletionToolStripMenuItem.Click -= ShutDownUponCompletionToolStripMenuItem_Click;

            SplashScreenToolStripMenuItem.Click -= SplashScreenToolStripMenuItem_Click;
            RestoreToDefaultToolStripMenuItem.Click -= RestoreDefaultToolStripMenuItem_Click;
            AboutEasyFileManagerToolStripMenuItem.Click -= AboutEasyFileManagerToolStripMenuItem_Click;

            PreviousFolderButton.Click -= PreviousFolderButton_Click;
            NextFolderButton.Click -= NextFolderButton_Click;
            OneLevelUpButton.Click -= OneLevelUpButton_Click;
            SelectFolderButton.Click -= SelectFolderButton_Click;
            GPSEditButton.Click -= GPSEditButton_Click;
            DateEditButton.Click -= DateEditButton_Click;
            EasyMetadataEditButton.Click -= EasyMetadataEditButton_Click;
            ApplyButton.Click -= ApplyButton_ClickAsync;

            ExplorerDataGridView.SelectionChanged -= ExplorerDataGridView_SelectionChanged;
            ExplorerDataGridView.CellDoubleClick -= ExplorerDataGridView_CellDoubleClick;

            foreach (Button btn in RotateTableLayoutPanel.GetControls<Button>())
            {
                btn.Click -= RotateButton_Click;
            }
            foreach (DataGridView dgv in this.GetControls<DataGridView>())
            {
                dgv.CellMouseDown -= DataGridView_CellMouseDown;
            }
            foreach (TextBox textBox in this.GetControls<TextBox>(true))
            {
                textBox.KeyPress -= TextBox_KeyPress;
                textBox.MouseDown -= TextBox_MouseDown;
            }
            PathTextBox.Leave -= PathTextBox_Leave;

            CustomizeCheckBox.CheckedChanged -= CustomizeCheckBox_CheckedChanged;
            RenameCheckBox.CheckedChanged -= RenameCheckBox_CheckedChanged;
            CopyMoveCheckBox.CheckStateChanged -= CopyMoveCheckBox_CheckStateChanged;
            FinalizeCheckBox.CheckedChanged -= FinalizeCheckBox_CheckedChanged;

            TitleCheckBox.CheckedChanged -= TitleCheckBox_CheckedChanged;
            TitleTextBox.Leave -= TitleTextBox_Leave;

            SubjectCheckBox.CheckedChanged -= SubjectCheckBox_CheckedChanged;
            SubjectTextBox.Leave -= SubjectTextBox_Leave;

            CommentCheckBox.CheckedChanged -= CommentCheckBox_CheckedChanged;
            CommentTextBox.Leave -= CommentTextBox_Leave;

            KeywordsCheckBox.CheckStateChanged -= KeywordsCheckBox_CheckStateChanged;
            KeywordsComboBox.SelectedIndexChanged -= KeywordsComboBox_SelectedIndexChanged;
            KeywordsEditButton.Click -= KeywordsEditButton_Click;

            GeolocationCheckBox.CheckedChanged -= GeolocationCheckBox_CheckedChanged;
            GPSSourceComboBox.SelectedIndexChanged -= GPSSourceComboBox_SelectedIndexChanged;
            GPSComboBox.SelectedIndexChanged -= GPSComboBox_SelectedIndexChanged;
            WriteGPSCoordsCheckBox.CheckedChanged -= WriteGPSCoordsCheckBox_CheckedChanged;
            WriteGPSAreaInfoCheckBox.CheckedChanged -= WriteGPSAreaInfoCheckBox_CheckedChanged;
            WriteGPSEasyMetadataCheckBox.CheckedChanged -= WriteGPSEasyMetadataCheckBox_CheckedChanged;

            DateCheckBox.CheckedChanged -= DateCheckBox_CheckStateChanged;
            DateComboBox.SelectedIndexChanged -= DateComboBox_SelectedIndexChanged;
            WriteDateModifiedCheckBox.CheckedChanged -= WriteDateModifiedCheckBox_CheckedChanged;
            WriteDateCreatedCheckBox.CheckedChanged -= WriteDateCreatedCheckBox_CheckedChanged;
            WriteDateTakenOrEncodedCheckBox.CheckedChanged -= WriteDateTakenOrEncodedCheckBox_CheckedChanged;
            WriteDateEasyMetadataCheckBox.CheckedChanged -= WriteDateEasyMetadataCheckBox_CheckedChanged;

            EasyMetadataCheckBox.CheckStateChanged -= EasyMetadataCheckBox_CheckedChanged;
            EasyMetadataComboBox.SelectedIndexChanged -= EasyMetadataComboBox_SelectedIndexChanged;

            PrefixTextBox.Leave -= PrefixTextBox_Leave;
            ReplaceTextBox.Leave -= ReplaceTextBox_Leave;
            WithTextBox.Leave -= WithTextBox_Leave;
            SuffixTextBox.Leave -= SuffixTextBox_Leave;

            ReplaceWithCheckBox.CheckStateChanged -= ReplaceWithCheckBox_CheckStateChanged;

            DateFormatCheckBox.CheckedChanged -= DateFormatCheckBox_CheckedChanged;
            DateFormatComboBox.SelectedIndexChanged -= DateFormatComboBox_SelectedIndexChanged;

            BackupFolderCheckBox.CheckedChanged -= BackupFolderCheckBox_CheckedChanged;
            TopFolderCheckBox.CheckedChanged -= TopFolderCheckBox_CheckedChanged;
            SubfoldersCheckBox.CheckedChanged -= SubfoldersCheckBox_CheckedChanged;
            FilterCheckBox.CheckedChanged -= FilterCheckBox_CheckedChanged;

            FilterNameComboBox.SelectedIndexChanged -= FilterNameComboBox_SelectedIndexChanged;
            FilterStringTextBox.Leave -= FilterStringTextBox_Leave;

            SelectBackupFolderButton.Click -= SelectBackupFolderButton_Click;
            SelectTopFolderButton.Click -= SelectTopFolderButton_Click;
            SubfoldersEditButton.Click -= SubfoldersEditButton_Click;
            FilterEditButton.Click -= FilterEditButton_Click;

            CleanUpCheckBox.CheckedChanged -= CleanUpCheckBox_CheckedChanged;
            DuplicatesCheckBox.CheckStateChanged -= DuplicatesCheckBox_CheckStateChanged;
            ShowDuplicatesCompareDialogCheckBox.CheckedChanged -= ShowDuplicatesCompareDialogCheckBox_CheckedChanged;

            CleanUpEditButton.Click -= CleanUpEditButton_Click;
            SelectDuplicatesFolderButton.Click -= SelectDuplicatesFolderButton_Click;
            DuplicatesCompareEditButton.Click -= DuplicatesCompareEditButton_Click;
        }

        private void DisposeMainEventHandlers()
        {
            Debug.WriteLine($"Disposing '{Name}': main event handlers...");

            Load -= Main_Load;
            FormClosing -= Main_FormClosing;
            FormClosed -= Main_FormClosed;
            Disposed -= Main_Disposed;
        }

        private static void DisposeProperties()
        {
            FolderWatcher.Dispose();
            ProgressBackgroundWorker.Dispose();
        }

        private static void DisposeWinFormsLib()
        {
            GeoLocator.Dispose();
            ExifWrapper.Dispose();
            Utils.Dispose();
        }

        private static void RestoreToDefault(string startupDirectory = EMPTY_STRING)
        {
            if (string.IsNullOrEmpty(startupDirectory))
            {
                Properties.Settings.Default.StartupDirectory = STARTUP_DIRECTORY_DEFAULT;
            }
            Options = new();
            SaveOptions();
        }

        #endregion
    }
}
