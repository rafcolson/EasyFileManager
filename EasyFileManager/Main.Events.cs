using System.ComponentModel;
using System.Diagnostics;

using WinFormsLib;

using static WinFormsLib.Utils;
using static WinFormsLib.Forms;
using static WinFormsLib.Buttons;
using static WinFormsLib.Constants;
using static WinFormsLib.GeoLocator;

namespace EasyFileManager
{
    public partial class Main
    {
        #region FileSystemWatcher

        private void FolderWatcher_RenamedOrDeleted(object? sender, FileSystemEventArgs e)
        {
            FolderWatcher.EnableRaisingEvents = false;
            Update(e.ChangeType == WatcherChangeTypes.Renamed ? e.FullPath : string.Empty, true);
            FolderWatcher.EnableRaisingEvents = true;
        }

        private void FolderContentsWatcher_ChangedRenamedDeletedOrCreated(object? sender, FileSystemEventArgs e)
        {
            FolderContentsWatcher.EnableRaisingEvents = false;
            Update(Folder.Path, true);
            FolderContentsWatcher.EnableRaisingEvents = true;
        }

        #endregion

        #region BackgroundWorker

        private void ProgressBackgroundWorker_ProgressChanged(object? sender, ProgressChangedEventArgs e)
        {
            int i = e.ProgressPercentage;
            if (i == 0)
            {
                MainProgressBar.Style = ProgressBarStyle.Blocks;
            }
            else if (i == 99)
            {
                MainProgressBar.Style = ProgressBarStyle.Marquee;
            }
            else if (i == 100)
            {
                Delay(() => { if (MainProgressBar.Value != 0) { ProgressBackgroundWorker.ReportProgress(0); } }, 750 );
            }
            MainProgressBar.Value = i;
            MainProgressBar.Visible = i != 0;
            MainProgressStatusLabel.Text = Progress.Info;
        }

        #endregion

        #region DataGridView

        private void ExplorerDataGridView_SelectionChanged(object? sender, EventArgs e) => UpdateSelection();

        private void ExplorerDataGridView_CellDoubleClick(object? sender, DataGridViewCellEventArgs e)
        {
            DataGridViewRow row = ExplorerDataGridView.Rows[e.RowIndex];
            if (row.Tag is EasyFolder ed)
            {
                Update(ed.Path);
            }
            else if (row.Tag is EasyFile ef)
            {
                Process.Start(EXPLORER_EXE, ef.Path);
            }
        }

        private void DataGridView_CellMouseDown(object? sender, DataGridViewCellMouseEventArgs e)
        {
            if (sender is DataGridView dgv)
            {
                if (e.Button == MouseButtons.Right)
                {
                    if (e.RowIndex != -1 && !dgv.SelectedRows.Contains(dgv.Rows[e.RowIndex]))
                    {
                        foreach (DataGridViewRow row in dgv.Rows)
                        {
                            if (row.Index == e.RowIndex)
                            {
                                dgv.CurrentCell = row.Cells[e.ColumnIndex == -1 ? 0 : e.ColumnIndex];
                                row.Selected = true;
                            }
                            else
                            {
                                row.Selected = false;
                            }
                        }
                    }
                    dgv.Focus();
                }
            }
        }

        #endregion

        #region ContextMenuStrip

        private void EditExplorerContextMenuStrip_Opening(object? sender, CancelEventArgs e)
        {
            if (sender is ContextMenuStrip cms && cms.SourceControl is DataGridView dgv)
            {
                DataGridViewSelectedRowCollection rows = dgv.SelectedRows;
                bool isValidSelection = rows.Count == 1 && rows[0].Tag is EasyPath ep && !ep.Type.HasFlag(EasyType.System);
                string[] openExplore = new[] { Globals.Open, Globals.Explore };
                foreach (ToolStripItem item in cms.Items)
                {
                    if (openExplore.Contains(item.Text))
                    {
                        item.Enabled = isValidSelection;
                    }
                }
            }
        }

        #endregion

        #region MenuItem

        private void ExportToolStripMenuItem_Click(object? sender, EventArgs e)
        {
            string initialDirectory = Properties.Settings.Default.ImportExportDirectory;
            if (!Directory.Exists(initialDirectory))
            {
                initialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyComputer);
            }
            ExportSettingsDialog.InitialDirectory = initialDirectory;
            if (ExportSettingsDialog.ShowDialog() == DialogResult.OK)
            {
                Properties.Settings.Default.ImportExportDirectory = ImportSettingsDialog.InitialDirectory;
                string s = Options.ToString();
                Properties.Settings.Default.EasyOptions = s;
                Properties.Settings.Default.Save();
                new FileInfo(ExportSettingsDialog.FileName).WriteText(s);
            }
        }

        private void ImportToolStripMenuItem_Click(object? sender, EventArgs e)
        {
            string initialDirectory = Properties.Settings.Default.ImportExportDirectory;
            if (!Directory.Exists(initialDirectory))
            {
                initialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyComputer);
            }
            ImportSettingsDialog.InitialDirectory = initialDirectory;
            if (ImportSettingsDialog.ShowDialog() == DialogResult.OK)
            {
                Properties.Settings.Default.ImportExportDirectory = ImportSettingsDialog.InitialDirectory;
                UpdateOptions(ImportSettingsDialog.FileName);

                WaitUntilTaskCompleted();

                Folders.Clear();
                Files.Clear();
                InitializeLayout();
            }
        }

        private void ExitToolStripMenuItem_Click(object? sender, EventArgs e) => Exit();

        private void ShowThumbnailToolStripMenuItem_Click(object? sender, EventArgs e)
        {
            bool b = !ShowThumbnailToolStripMenuItem.Checked;
            Properties.Settings.Default.ShowThumbnail = b;
            ShowThumbnailToolStripMenuItem.Checked = b;
            UpdateThumbnailPropsView();
        }

        private void ShowPropertiesToolStripMenuItem_Click(object? sender, EventArgs e)
        {
            bool b = !ShowPropertiesToolStripMenuItem.Checked;
            Properties.Settings.Default.ShowProperties = b;
            ShowPropertiesToolStripMenuItem.Checked = b;
            UpdateThumbnailPropsView();
        }

        private void ShowMillisecondsToolStripMenuItem_Click(object? sender, EventArgs e)
        {
            bool b = !ShowMillisecondsToolStripMenuItem.Checked;
            Properties.Settings.Default.ShowMilliseconds = b;
            ShowMillisecondsToolStripMenuItem.Checked = b;
            UpdateThumbnailPropsView();
        }

        private void ShowHiddenItemsToolStripMenuItem_Click(object? sender, EventArgs e)
        {
            bool b = !ShowHiddenItemsToolStripMenuItem.Checked;
            Properties.Settings.Default.ShowHiddenItems = b;
            ShowHiddenItemsToolStripMenuItem.Checked = b;
            UpdateExplorerView();
            UpdateThumbnailPropsView();
        }

        private void ShowVideoMetadataToolStripMenuItem_Click(object? sender, EventArgs e)
        {
            bool b = !ShowVideoMetadataToolStripMenuItem.Checked;
            Properties.Settings.Default.ShowVideoMetadata = b;
            ShowVideoMetadataToolStripMenuItem.Checked = b;
            UpdateExplorerView();
            UpdateThumbnailPropsView();
        }

        private void LanguageToolStripMenuItem_Click(object? sender, EventArgs e)
        {

            string[] sa = Enum.GetValues<Language>().Select(x => $"{x.GetGlobalStringValue()} ({x.GetValue<string>()})").ToArray();
            using DropDownListDialog ddld = new(sa, Properties.Settings.Default.LanguageIndex);
            if (ddld.ShowDialog() == DialogResult.OK)
            {
                Properties.Settings.Default.LanguageIndex = ddld.SelectedIndex;
                Restart();
            }
        }

        private void UseEasyMetadataWithVideoToolStripMenuItem_Click(object? sender, EventArgs e)
        {
            bool b = !UseEasyMetadataWithVideoToolStripMenuItem.Checked;
            Options.UseEasyMetadataWithVideo = b;
            UseEasyMetadataWithVideoToolStripMenuItem.Checked = b;
            UpdateExplorerView();
            UpdateThumbnailPropsView();
        }

        private void PreserveDateCreatedToolStripMenuItem_Click(object? sender, EventArgs e)
        {
            bool b = !PreserveDateCreatedToolStripMenuItem.Checked;
            Options.PreserveDateCreated = b;
            PreserveDateCreatedToolStripMenuItem.Checked = b;
        }

        private void PreserveDateModifiedToolStripMenuItem_Click(object? sender, EventArgs e)
        {
            bool b = !PreserveDateModifiedToolStripMenuItem.Checked;
            Options.PreserveDateModified = b;
            PreserveDateModifiedToolStripMenuItem.Checked = b;
        }

        private void LogApplicationEventsToolStripMenuItem_Click(object? sender, EventArgs e)
        {
            bool b = !LogApplicationEventsToolStripMenuItem.Checked;
            Options.LogApplicationEvents = b;
            LogApplicationEventsToolStripMenuItem.Checked = b;
        }

        private void ShutDownUponCompletionToolStripMenuItem_Click(object? sender, EventArgs e)
        {
            bool b = !ShutDownUponCompletionToolStripMenuItem.Checked;
            Options.ShutdownUponCompletion = b;
            ShutDownUponCompletionToolStripMenuItem.Checked = b;
        }

        private void SplashScreenToolStripMenuItem_Click(object? sender, EventArgs e) => ShowSplash();

        private void AboutEasyFileManagerToolStripMenuItem_Click(object? sender, EventArgs e)
        {
            using MessageDialog md = new(Globals.AboutMessage);
            md.ShowDialog();
        }

        private async void RestoreDefaultToolStripMenuItem_ClickAsync(object? sender, EventArgs e)
        {
            RestoreToDefault();

            await UpdateFormattingAsync();
        }

        private static void PathToolStripMenuItem_Click(object? sender, EventArgs e)
        {
            if (sender is ToolStripMenuItem tsmi && tsmi.Owner is ContextMenuStrip cms && cms.SourceControl is TextBox tb)
            {
                string s = tb.Text;
                if (IsValidDirectoryPath(s))
                {
                    Process.Start(EXPLORER_EXE, GetParentPath(s));
                }
            }
        }

        private static void ExplorerToolStripMenuItem_Click(object? sender, EventArgs e, bool open = false)
        {
            if (sender is ToolStripMenuItem tsmi && tsmi.Owner is ContextMenuStrip cms && cms.SourceControl is DataGridView dgv)
            {
                if (dgv.SelectedRows.Count == 1 && dgv.SelectedRows[0].Tag is EasyPath ep)
                {
                    Process.Start(EXPLORER_EXE, open ? ep.Path : ep.ParentPath);
                }
            }
        }

        #endregion

        #region Button

        private void RotateButton_Click(object? sender, EventArgs e)
        {
            if (!IsUpdating && !IsApplying && sender is Button button)
            {
                IsApplying = true;

                int i = Array.IndexOf(RotateTableLayoutPanel.GetControls<Button>(), button);
                RotateImage((EasyOrientation)i);

                IsApplying = false;
                
                UpdateThumbnailPropsView();
            }
            PathTableLayoutPanel.Focus();
        }

        private void KeywordsEditButton_Click(object? sender, EventArgs e)
        {
            List<object> items = new(KeywordsComboBox.Items.Cast<object>());
            using EditListDialog eld = new(ref items, font: Font, editButtons: EditButtons.AddRemoveMoveUpMoveDownSortEdit);
            eld.OnAddItem = () =>
            {
                if (GetKeywords() is EasyList<string> el)
                {
                    eld.SelectedIndex = eld.Items.Add(el);
                }
                return Task.CompletedTask;
            };
            eld.OnEditItem = () =>
            {
                if (GetKeywords(eld.SelectedItem) is EasyList<string> el)
                {
                    int i = eld.SelectedIndex;
                    eld.Items.RemoveAt(i);
                    eld.Items.Insert(i, el);
                    eld.SelectedIndex = i;
                }
                return Task.CompletedTask;
            };
            if (eld.ShowDialog() == DialogResult.OK)
            {
                Options.Keywords.Replace(items.Cast<EasyList<string>>().ToArray());
                Options.KeywordsIndex = eld.SelectedIndex;

                UpdateKeywordsControlsA();
            }
            KeywordsGroupBox.Focus();
        }

        private void GPSEditButton_Click(object? sender, EventArgs e)
        {
            List<object> items = new(GPSComboBox.Items.Cast<object>());
            using EditListDialog eld = new(ref items, font: Font, editButtons: EditButtons.AddRemoveMoveUpMoveDownSortEdit);
            eld.OnAddItem = async () =>
            {
                if (await GetGPSAsync() is KeyValuePair<string, GeoCoordinates?> kvp)
                {
                    if (!eld.Items.Cast<object>().Select(x => ((KeyValuePair<string, GeoCoordinates?>)x).Key).Contains(kvp.Key))
                    {
                        eld.SelectedIndex = eld.Items.Add(kvp);
                    }
                    else
                    {
                        using MessageDialog md = new(string.Format(Globals.FormatAlreadyExists, kvp.Key), font: Font);
                        md.ShowDialog();
                    }
                }
            };
            eld.OnEditItem = async () =>
            {
                if (eld.SelectedItem is KeyValuePair<string, GeoCoordinates?> kvp)
                {
                    if (await GetGPSAsync(eld.SelectedItem) is KeyValuePair<string, GeoCoordinates?> kvpOut)
                    {
                        string k = kvpOut.Key;
                        if (kvp.Key == k || !eld.Items.Cast<object>().Select(x => ((KeyValuePair<string, GeoCoordinates?>)x).Key).Contains(k))
                        {
                            int i = eld.SelectedIndex;
                            eld.Items.RemoveAt(i);
                            eld.Items.Insert(i, kvpOut);
                            eld.SelectedIndex = i;
                        }
                        else
                        {
                            using MessageDialog md = new(string.Format(Globals.FormatAlreadyExists, k), font: Font);
                            md.ShowDialog();
                        }
                    }
                }
            };
            if (eld.ShowDialog() == DialogResult.OK)
            {
                GPSComboBox.Items.Clear();
                GPSComboBox.Items.AddRange(items.ToArray());
                GPSComboBox.SelectedItem = eld.SelectedItem;

                Options.GeoAreas = GPSComboBox.GetItems<KeyValuePair<string, GeoCoordinates?>>().ToDictionary(x => x.Key, x => x.Value);
                Options.GPSCustomIndex = GPSComboBox.SelectedIndex;

                UpdateGPSControlsA();
                UpdateThumbnailPropsView();
            }
            GeolocationGroupBox.Focus();
        }

        private void DateEditButton_Click(object? sender, EventArgs e)
        {
            List<object> items = new(DateComboBox.Items.Cast<object>());
            using EditListDialog eld = new(ref items, font: Font, editButtons: EditButtons.AddRemoveMoveUpMoveDownSortEdit);
            eld.OnAddItem = () =>
            {
                if (GetDate() is object o)
                {
                    if (!eld.Items.Cast<object>().Contains(o))
                    {
                        eld.SelectedIndex = eld.Items.Add(o);
                    }
                    else
                    {
                        using MessageDialog md = new(string.Format(Globals.FormatAlreadyExists, o), font: Font);
                        md.ShowDialog();
                    }
                }
                return Task.CompletedTask;
            };
            eld.OnEditItem = () =>
            {
                if (GetDate(eld.SelectedItem) is object o)
                {
                    if (!eld.Items.Cast<object>().Contains(o))
                    {
                        int i = eld.SelectedIndex;
                        eld.Items.RemoveAt(i);
                        eld.Items.Insert(i, o);
                        eld.SelectedIndex = i;
                    }
                    else
                    {
                        using MessageDialog md = new(string.Format(Globals.FormatAlreadyExists, o), font: Font);
                        md.ShowDialog();
                    }
                }
                return Task.CompletedTask;
            };
            if (eld.ShowDialog() == DialogResult.OK)
            {
                DateComboBox.Items.Clear();
                DateComboBox.Items.AddRange(items.ToArray());
                DateComboBox.SelectedItem = eld.SelectedItem;

                Options.Dates.Replace(items.Cast<string>());
                Options.DateCustomIndex = DateComboBox.SelectedIndex;

                UpdateDateControlsA();
            }
            DateGroupBox.Focus();
        }

        private void EasyMetadataEditButton_Click(object? sender, EventArgs e)
        {
            List<object> items = new(EasyMetadataComboBox.Items.Cast<object>());
            using EditListDialog eld = new(ref items, font: Font, editButtons: EditButtons.AddRemoveMoveUpMoveDownSortEdit);
            eld.OnAddItem = () =>
            {
                if (GetEasyMetadata() is object o)
                {
                    eld.SelectedIndex = eld.Items.Add(o);
                }
                return Task.CompletedTask;
            };
            eld.OnEditItem = () =>
            {
                if (GetEasyMetadata(eld.SelectedItem) is object o)
                {
                    int i = eld.SelectedIndex;
                    eld.Items.RemoveAt(i);
                    eld.Items.Insert(i, o);
                    eld.SelectedIndex = i;
                }
                return Task.CompletedTask;
            };
            if (eld.ShowDialog() == DialogResult.OK)
            {
                Options.EasyMetadata.Replace(items.Cast<string>().ToArray());
                Options.EasyMetadataIndex = eld.SelectedIndex;

                UpdateEasyMetadataControlsA();
            }
            EasyMetadataGroupBox.Focus();
        }

        private void PreviousFolderButton_Click(object? sender, EventArgs e)
        {
            if (_historyIndex != 0)
            {
                _historyIndex -= 1;
                Update(_history[_historyIndex]);
            }
            PathTableLayoutPanel.Focus();
        }

        private void NextFolderButton_Click(object? sender, EventArgs e)
        {
            if (_historyIndex != _history.Count - 1)
            {
                _historyIndex += 1;
                Update(_history[_historyIndex]);
            }
            PathTableLayoutPanel.Focus();
        }

        private void OneLevelUpButton_Click(object? sender, EventArgs e)
        {
            Update(Folder.ParentPath);
            PathTableLayoutPanel.Focus();
        }

        private void SelectFolderButton_Click(object? sender, EventArgs e)
        {
            OpenFolderDialog.SelectedPath = Folder.Path;
            if (OpenFolderDialog.ShowDialog() == DialogResult.OK)
            {
                Update(OpenFolderDialog.SelectedPath);
            }
            PathTableLayoutPanel.Focus();
        }

        private void SelectBackupFolderButton_Click(object? sender, EventArgs e)
        {
            OpenFolderDialog.SelectedPath = GetValidDirectoryPath(Options.BackupFolderPath);
            if (OpenFolderDialog.ShowDialog() == DialogResult.OK)
            {
                if (OpenFolderDialog.SelectedPath != Options.BackupFolderPath)
                {
                    Options.BackupFolderPath = OpenFolderDialog.SelectedPath;
                    BackupFolderTextBox.Text = OpenFolderDialog.SelectedPath;
                }
            }
            BackupFolderGroupBox.Focus();
        }

        private void SelectDuplicatesFolderButton_Click(object? sender, EventArgs e)
        {
            OpenFolderDialog.SelectedPath = GetValidDirectoryPath(Options.DuplicatesFolderPath);
            if (OpenFolderDialog.ShowDialog() == DialogResult.OK)
            {
                if (OpenFolderDialog.SelectedPath != Options.DuplicatesFolderPath)
                {
                    Options.DuplicatesFolderPath = OpenFolderDialog.SelectedPath;
                    DuplicatesFolderTextBox.Text = OpenFolderDialog.SelectedPath;
                }
            }
            DuplicatesGroupBox.Focus();
        }

        private void DuplicatesCompareEditButton_Click(object? sender, EventArgs e)
        {
            List<object> items = Options.DuplicatesCompareParameters.GetContainingFlags().Select(x => (object)x.GetEasyGlobalStringValue()).ToList();
            using EditListDialog eld = new(ref items, font: Font, editButtons: EditButtons.AddRemove);
            object[] oa = EasyCompareParameter.All.GetContainingFlags().Select(x => x.GetEasyGlobalStringValue()).ToArray();
            eld.OnAddItem = () =>
            {
                using DropDownListDialog ddld = new(oa, 0, font: Font);
                while (ddld.ShowDialog() == DialogResult.OK)
                {
                    if (!eld.Items.Contains(ddld.SelectedItem))
                    {
                        eld.SelectedIndex = eld.Items.Add(ddld.SelectedItem);
                        break;
                    }
                }
                return Task.CompletedTask;
            };
            if (eld.ShowDialog() == DialogResult.OK)
            {
                Options.DuplicatesCompareParameters = EasyCompareParameter.None;
                foreach (string s in items.Cast<string>())
                {
                    Options.DuplicatesCompareParameters |= s.AsEasyEnumFromGlobal<EasyCompareParameter>();
                }
                UpdateDuplicatesControlsB();
            }
            DuplicatesGroupBox.Focus();
        }

        private async void ApplyButton_ClickAsync(object? sender, EventArgs e)
        {
            if (IsApplying)
            {
                _cancellationTokenSource.Cancel();
                return;
            }
            UpdateApplyButton(true);

            List<Action<CancellationToken>> actions = new();
            if (Options.HasCustomizing())
            {
                actions.Add(Customize);
            }
            if (Options.HasRenaming() || Options.HasMoving())
            {
                actions.Add(ApplyFormatting);
            }
            if (Options.HasFinalizing())
            {
                if (Options.DeleteEmptyFolders)
                {
                    actions.Add(DeleteEmptyFolders);
                }
                if (Options.DuplicatesState != CheckState.Unchecked)
                {
                    actions.Add(DeleteOrStoreDuplicates);
                }
            }
            if (Options.LogApplicationEvents)
            {
                Logger.Reset();
            }
            EasyMetadata.Options.Reset();
            await UpdateAsync(actions.ToArray());

            Folders.Replace(Folder.GetDirectoryPaths());
            Files.Replace(Folder.GetFilePaths());

            UpdateApplyButton(false);
            UpdateExplorerView();
            UpdateThumbnailPropsView();

            if (Options.ShutdownUponCompletion)
            {
                static void shutdown() { ShutDown(); Exit(); }
                using DelayedActionDialog dad = new(shutdown, caption: Globals.ShutingDown);
                dad.ShowDialog();
            }
            else if (Options.LogApplicationEvents)
            {
                Logger.Show();
            }
            PathTableLayoutPanel.Focus();
        }

        private async void SelectTopFolderButton_ClickAsync(object? sender, EventArgs e)
        {
            Options.TopFolderPath = GetValidDirectoryPath(Options.TopFolderPath);
            OpenFolderDialog.SelectedPath = Options.TopFolderPath;
            if (OpenFolderDialog.ShowDialog() == DialogResult.OK)
            {
                string p = OpenFolderDialog.SelectedPath;
                if (p != Options.TopFolderPath)
                {
                    Options.TopFolderPath = p;
                    TopFolderTextBox.Text = p;
                    await UpdateFormattingAsync();
                }
            }
            TopFolderGroupBox.Focus();
        }

        private async void SubfoldersEditButton_ClickAsync(object? sender, EventArgs e)
        {
            List<object> items = Options.Subfolders.Select(x => (object)x.GetEasyGlobalStringValue()).ToList();
            using EditListDialog eld = new(ref items, font: Font, editButtons: EditButtons.AddRemoveMoveUpMoveDown);
            object[] oa = Enum.GetValues<EasySubfolder>().Select(x => x.GetEasyGlobalStringValue()).ToArray();
            eld.OnAddItem = () =>
            {
                using DropDownListDialog ddld = new(oa, 0, font: Font);
                if (ddld.ShowDialog() == DialogResult.OK)
                {
                    eld.SelectedIndex = eld.Items.Add(ddld.SelectedItem);
                }
                return Task.CompletedTask;
            };
            if (eld.ShowDialog() == DialogResult.OK)
            {
                Options.Subfolders.Replace(items.Select(x => ((string)x).AsEasyEnumFromGlobal<EasySubfolder>()));
                UpdateFolderControlsB();
                await UpdateFormattingAsync();
            }
            SubfoldersGroupBox.Focus();
        }

        private async void FilterEditButton_ClickAsync(object? sender, EventArgs e)
        {
            List<object> items = Options.TypeFilter.GetContainingFlags().Select(x => (object)x.GetEasyGlobalStringValue()).ToList();
            using EditListDialog eld = new(ref items, font: Font, editButtons: EditButtons.AddRemove);
            object[] oa = EasyTypeFilter.All.GetContainingFlags().Select(x => x.GetEasyGlobalStringValue()).ToArray();
            eld.OnAddItem = () =>
            {
                using DropDownListDialog ddld = new(oa, 0, font: Font);
                while (ddld.ShowDialog() == DialogResult.OK)
                {
                    if (!eld.Items.Contains(ddld.SelectedItem))
                    {
                        eld.SelectedIndex = eld.Items.Add(ddld.SelectedItem);
                        break;
                    }
                }
                return Task.CompletedTask;
            };
            if (eld.ShowDialog() == DialogResult.OK)
            {
                Options.TypeFilter = EasyTypeFilter.None;
                foreach (string s in items.Cast<string>())
                {
                    Options.TypeFilter |= s.AsEasyEnumFromGlobal<EasyTypeFilter>();
                }
                UpdateFilterControlsB();
                await UpdateFormattingAsync();
            }
            FilterGroupBox.Focus();
        }

        #endregion

        #region TextBox

        private void TextBox_MouseDown(object? sender, MouseEventArgs e)
        {
            if (sender is TextBox tb && e.Button == MouseButtons.Right)
            {
                tb.Focus();
            }
        }

        private void TextBox_KeyPress(object? sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == Convert.ToChar(Keys.Enter))
            {
                ActiveControl = null;
                e.Handled = true;
            }
        }

        private void PathTextBox_Leave(object? sender, EventArgs e)
        {
            if (!Update(PathTextBox.Text))
            {
                PathTextBox.Text = Folder.Path;
            }
        }

        private void TitleTextBox_Leave(object? sender, EventArgs e)
        {
            string s = TitleTextBox.Text;
            if (s != Options.Title) { Options.Title = s; }
        }

        private void SubjectTextBox_Leave(object? sender, EventArgs e)
        {
            string s = SubjectTextBox.Text;
            if (s != Options.Subject) { Options.Subject = s; }
        }

        private void CommentTextBox_Leave(object? sender, EventArgs e)
        {
            string s = CommentTextBox.Text;
            if (s != Options.Comment) { Options.Comment = s; }
        }

        private async void PrefixTextBox_LeaveAsync(object? sender, EventArgs e)
        {
            string s = PrefixTextBox.Text;
            if (s != Options.Prefix)
            {
                Options.Prefix = s;
                await UpdateFormattingAsync();
            }
        }

        private async void ReplaceTextBox_LeaveAsync(object? sender, EventArgs e)
        {
            string s = ReplaceTextBox.Text;
            if (s != Options.Replace)
            {
                Options.Replace = s;
                await UpdateFormattingAsync();
            }
        }

        private async void WithTextBox_LeaveAsync(object? sender, EventArgs e)
        {
            string s = WithTextBox.Text;
            if (s != Options.With)
            {
                Options.With = s;
                await UpdateFormattingAsync();
            }
        }

        private async void SuffixTextBox_LeaveAsync(object? sender, EventArgs e)
        {
            string s = SuffixTextBox.Text;
            if (s != Options.Suffix)
            {
                Options.Suffix = s;
                await UpdateFormattingAsync();
            }
        }

        private async void FilterStringTextBox_LeaveAsync(object? sender, EventArgs e)
        {
            string s = FilterStringTextBox.Text;
            if (s != Options.FilterString)
            {
                Options.FilterString = s;
                UpdateFilterControlsC();
                await UpdateFormattingAsync();
            }
        }

        #endregion

        #region CheckBox

        private void CustomizeCheckBox_CheckedChanged(object? sender, EventArgs e)
        {
            Options.CustomizeEnabled = CustomizeCheckBox.Checked;
            UpdateCustomizePanel();
            UpdateEditPanel(customize: true);
        }

        private void TitleCheckBox_CheckedChanged(object? sender, EventArgs e)
        {
            Options.TitleEnabled = TitleCheckBox.Checked;
            TitleTextBox.Enabled = Options.TitleEnabled;
        }

        private void SubjectCheckBox_CheckedChanged(object? sender, EventArgs e)
        {
            Options.SubjectEnabled = SubjectCheckBox.Checked;
            SubjectTextBox.Enabled = Options.SubjectEnabled;
        }

        private void CommentCheckBox_CheckedChanged(object? sender, EventArgs e)
        {
            Options.CommentEnabled = CommentCheckBox.Checked;
            CommentTextBox.Enabled = Options.CommentEnabled;
        }

        private void KeywordsCheckBox_CheckStateChanged(object? sender, EventArgs e)
        {
            Options.KeywordsState = KeywordsCheckBox.CheckState;
            UpdateKeywordsControlsB();
        }

        private void GeolocationCheckBox_CheckedChanged(object? sender, EventArgs e)
        {
            Options.GeolocationEnabled = GeolocationCheckBox.Checked;
            UpdateGPSControlsB();
            UpdateGPSControlsA();
            UpdateGPSControlsA();
        }

        private void WriteGPSCoordsCheckBox_CheckedChanged(object? sender, EventArgs e)
        {
            Options.WriteGPSCoords = WriteGPSCoordsCheckBox.Checked;
            GeolocationGroupBox.Focus();
            UpdateGPSControlsB();
        }

        private void WriteGPSAreaInfoCheckBox_CheckedChanged(object? sender, EventArgs e)
        {
            Options.WriteGPSAreaInfo = WriteGPSAreaInfoCheckBox.Checked;
            GeolocationGroupBox.Focus();
            UpdateGPSControlsB();
        }

        private void WriteGPSEasyMetadataCheckBox_CheckedChanged(object? sender, EventArgs e)
        {
            Options.WriteGPSEasyMetadata = WriteGPSEasyMetadataCheckBox.Checked;
            GeolocationGroupBox.Focus();
            UpdateGPSControlsB();
        }

        private void DateCheckBox_CheckStateChanged(object? sender, EventArgs e)
        {
            Options.DateState = DateCheckBox.CheckState;
            UpdateDateControlsA();
        }

        private void WriteDateModifiedCheckBox_CheckedChanged(object? sender, EventArgs e)
        {
            Options.WriteDateModified = WriteDateModifiedCheckBox.Checked;
            UpdateDateControlsB();
            DateGroupBox.Focus();
        }

        private void WriteDateCreatedCheckBox_CheckedChanged(object? sender, EventArgs e)
        {
            Options.WriteDateCreated = WriteDateCreatedCheckBox.Checked;
            UpdateDateControlsB();
            DateGroupBox.Focus();
        }

        private void WriteDateTakenOrEncodedCheckBox_CheckedChanged(object? sender, EventArgs e)
        {
            Options.WriteDateCamera = WriteDateTakenOrEncodedCheckBox.Checked;
            UpdateDateControlsB();
            DateGroupBox.Focus();
        }

        private void WriteDateEasyMetadataCheckBox_CheckedChanged(object? sender, EventArgs e)
        {
            Options.WriteDateEasyMetadata = WriteDateEasyMetadataCheckBox.Checked;
            UpdateDateControlsB();
            DateGroupBox.Focus();
        }

        private void EasyMetadataCheckBox_CheckedChanged(object? sender, EventArgs e)
        {
            Options.EasyMetadataEnabled = EasyMetadataCheckBox.Checked;
            UpdateEasyMetadataControlsB();
        }

        private void BackupFolderCheckBox_CheckStateChanged(object? sender, EventArgs e)
        {
            Options.BackupFolderState = BackupFolderCheckBox.CheckState;
            UpdateBackupControls();
        }

        private void DeleteEmptyFoldersCheckBox_CheckedChanged(object? sender, EventArgs e)
        {
            Options.DeleteEmptyFolders = DeleteEmptyFoldersCheckBox.Checked;
        }

        private void DuplicatesCheckBox_CheckStateChanged(object? sender, EventArgs e)
        {
            Options.DuplicatesState = DuplicatesCheckBox.CheckState;
            UpdateDuplicatesControlsA();
        }

        private void ShowDuplicatesCompareDialogCheckBox_CheckedChanged(object? sender, EventArgs e)
        {
            Options.ShowDuplicatesCompareDialog = ShowDuplicatesCompareDialogCheckBox.Checked;
            DuplicatesGroupBox.Focus();
        }

        private async void RenameCheckBox_CheckedChangedAsync(object? sender, EventArgs e)
        {
            Options.RenameEnabled = RenameCheckBox.Checked;
            UpdateRenamePanel();
            UpdateEditPanel(rename: true);
            if (!await UpdateFormattingAsync())
            {
                UpdateExplorerView();
            }
        }

        private async void MoveCheckBox_CheckedChangedAsync(object? sender, EventArgs e)
        {
            Options.MoveEnabled = MoveCheckBox.Checked;
            UpdateMovePanel();
            UpdateEditPanel(move: true);
            if (!await UpdateFormattingAsync())
            {
                UpdateExplorerView();
            }
        }

        private async void FinalizeCheckBox_CheckedChangedAsync(object? sender, EventArgs e)
        {
            Options.FinalizeEnabled = FinalizeCheckBox.Checked;
            UpdateFinalizePanel();
            UpdateEditPanel(finalize: true);
            await UpdateFormattingAsync();
        }

        private async void ReplaceWithCheckBox_CheckStateChangedAsync(object? sender, EventArgs e)
        {
            Options.ReplaceWithState = ReplaceWithCheckBox.CheckState;
            UpdateReplaceWithControls();
            await UpdateFormattingAsync();
        }

        private async void DateFormatCheckBox_CheckedChangedAsync(object? sender, EventArgs e)
        {
            Options.DateFormatEnabled = DateFormatCheckBox.Checked;
            DateFormatComboBox.Enabled = Options.DateFormatEnabled;
            await UpdateFormattingAsync();
        }

        private async void TopFolderCheckBox_CheckedChangedAsync(object? sender, EventArgs e)
        {
            Options.TopFolderEnabled = TopFolderCheckBox.Checked;
            UpdateFolderControlsA();
            await UpdateFormattingAsync();
        }

        private async void FilterCheckBox_CheckedChangedAsync(object? sender, EventArgs e)
        {
            Options.FilterEnabled = FilterCheckBox.Checked;
            UpdateFilterControlsA();
            await UpdateFormattingAsync();
        }

        private async void SubfoldersCheckBox_CheckedChangedAsync(object? sender, EventArgs e)
        {
            Options.SubfoldersEnabled = SubfoldersCheckBox.Checked;
            UpdateFolderControlsA();
            await UpdateFormattingAsync();
        }

        #endregion

        #region ComboBox

        private void KeywordsComboBox_SelectedIndexChanged(object? sender, EventArgs e)
        {
            Options.KeywordsIndex = KeywordsComboBox.SelectedIndex;
        }

        private void GPSSourceComboBox_SelectedIndexChanged(object? sender, EventArgs e)
        {
            Options.GPSSourceIndex = GPSSourceComboBox.SelectedIndex;
            UpdateGPSControlsA();
        }

        private void GPSComboBox_SelectedIndexChanged(object? sender, EventArgs e)
        {
            if (Options.GPSSourceIndex == (int)EasyMetadataSource.CustomMetadata)
            {
                Options.GPSCustomIndex = GPSComboBox.SelectedIndex;
            }
            else
            {
                Options.GPSAreaIndex = GPSComboBox.SelectedIndex;
            }
        }

        private void EasyMetadataComboBox_SelectedIndexChanged(object? sender, EventArgs e) => Options.EasyMetadataIndex = EasyMetadataComboBox.SelectedIndex;

        private void DateComboBox_SelectedIndexChanged(object? sender, EventArgs e)
        {
            if (Options.DateState == CheckState.Checked)
            {
                Options.DateCustomIndex = DateComboBox.SelectedIndex;
            }
            else if (Options.DateState == CheckState.Indeterminate)
            {
                Options.DateSourceIndex = DateComboBox.SelectedIndex;
            }
        }

        private async void DateFormatComboBox_SelectedIndexChangedAsync(object? sender, EventArgs e)
        {
            Options.DateFormat = (string)DateFormatComboBox.SelectedItem;
            await UpdateFormattingAsync();
        }

        private async void FilterNameComboBox_SelectedIndexChangedAsync(object? sender, EventArgs e)
        {
            Options.NameFilter = (EasyNameFilter)FilterNameComboBox.SelectedIndex;
            await UpdateFormattingAsync();
        }

        #endregion
    }
}
