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

        private void FolderWatcher_Changed(object? sender, AdaptiveFileSystemWatcherEventArgs e)
        {
            if (_isClosing || _isUpdating || _isApplying)
            {
                return;
            }
            string path = e.Path;
            if (!Directory.Exists(path))
            {
                path = GetParentPath(path);
            }
            Update(path, true);
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
                Update(GetExplorerPath(ed));
            }
            else if (row.Tag is EasyFile ef)
            {
                Process.Start(EXPLORER_EXE, GetExplorerPath(ef));
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
                string[] openExplore = [Globals.Open, Globals.Explore];
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
            string initialDirectory = Options.ImportExportDirectory;
            if (!Directory.Exists(initialDirectory))
            {
                initialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyComputer);
            }
            ExportSettingsDialog.InitialDirectory = initialDirectory;
            if (ExportSettingsDialog.ShowDialog() == DialogResult.OK)
            {
                Options.ImportExportDirectory = ExportSettingsDialog.InitialDirectory;
                string s = Options.ToString();
                Properties.Settings.Default.EasyOptions = s;
                Properties.Settings.Default.Save();
                new FileInfo(ExportSettingsDialog.FileName).WriteText(s);
            }
        }

        private void ImportToolStripMenuItem_Click(object? sender, EventArgs e)
        {
            string initialDirectory = Options.ImportExportDirectory;
            if (!Directory.Exists(initialDirectory))
            {
                initialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyComputer);
            }
            ImportSettingsDialog.InitialDirectory = initialDirectory;
            if (ImportSettingsDialog.ShowDialog() == DialogResult.OK)
            {
                UpdateOptions(ImportSettingsDialog.FileName);
                Options.ImportExportDirectory = ImportSettingsDialog.InitialDirectory;
                SaveOptions();

                WaitUntilTaskCompleted();

                Folders.Clear();
                Files.Clear();
                InitializeLayout();
            }
        }

        private void ExitToolStripMenuItem_Click(object? sender, EventArgs e) => Exit();

        private void ShowOutputToolStripMenuItem_Click(object? sender, EventArgs e)
        {
            bool enabled = !ShowOutputToolStripMenuItem.Checked;
            Options.Show = GetToggledFlags(Options.Show, EasyShow.Output, enabled);
            ShowOutputToolStripMenuItem.Checked = enabled;
            SaveOptions();
            Update(Properties.Settings.Default.StartupDirectory, true);
        }

        private void ShowThumbnailToolStripMenuItem_Click(object? sender, EventArgs e)
        {
            bool enabled = !ShowThumbnailToolStripMenuItem.Checked;
            Options.Show = GetToggledFlags(Options.Show, EasyShow.Thumbnail, enabled);
            ShowThumbnailToolStripMenuItem.Checked = enabled;
            SaveOptions();
            UpdateThumbnailPropsView();
        }

        private void ShowPropertiesToolStripMenuItem_Click(object? sender, EventArgs e)
        {
            bool enabled = !ShowPropertiesToolStripMenuItem.Checked;
            Options.Show = GetToggledFlags(Options.Show, EasyShow.Properties, enabled);
            ShowPropertiesToolStripMenuItem.Checked = enabled;
            SaveOptions();
            UpdateThumbnailPropsView();
        }

        private void ShowMillisecondsToolStripMenuItem_Click(object? sender, EventArgs e)
        {
            bool enabled = !ShowMillisecondsToolStripMenuItem.Checked;
            Options.Show = GetToggledFlags(Options.Show, EasyShow.Milliseconds, enabled);
            ShowMillisecondsToolStripMenuItem.Checked = enabled;
            SaveOptions();
            UpdateThumbnailPropsView();
        }

        private void ShowHiddenItemsToolStripMenuItem_Click(object? sender, EventArgs e)
        {
            bool enabled = !ShowHiddenItemsToolStripMenuItem.Checked;
            Options.Show = GetToggledFlags(Options.Show, EasyShow.HiddenItems, enabled);
            ShowHiddenItemsToolStripMenuItem.Checked = enabled;
            SaveOptions();
            UpdateExplorerView();
        }

        private void ExtractEmbeddedImageToolStripMenuItem_Click(object? sender, EventArgs e)
        {
            bool enabled = !ExtractEmbeddedImageToolStripMenuItem.Checked;
            Options.ExtractEmbeddedMetadata = GetToggledFlags(Options.ExtractEmbeddedMetadata, ExtractEmbeddedMetadata.Image, enabled);
            ExtractEmbeddedImageToolStripMenuItem.Checked = enabled;
            SaveOptions();
            ResetFileDetails();
            UpdateExplorerView();
        }

        private void ExtractEmbeddedVideoToolStripMenuItem_Click(object? sender, EventArgs e)
        {
            bool enabled = !ExtractEmbeddedVideoToolStripMenuItem.Checked;
            Options.ExtractEmbeddedMetadata = GetToggledFlags(Options.ExtractEmbeddedMetadata, ExtractEmbeddedMetadata.Video, enabled);
            ExtractEmbeddedVideoToolStripMenuItem.Checked = enabled;
            SaveOptions();
            ResetFileDetails();
            UpdateExplorerView();
        }

        private void ExtractEmbeddedAudioToolStripMenuItem_Click(object? sender, EventArgs e)
        {
            bool enabled = !ExtractEmbeddedAudioToolStripMenuItem.Checked;
            Options.ExtractEmbeddedMetadata = GetToggledFlags(Options.ExtractEmbeddedMetadata, ExtractEmbeddedMetadata.Audio, enabled);
            ExtractEmbeddedAudioToolStripMenuItem.Checked = enabled;
            SaveOptions();
            ResetFileDetails();
            UpdateExplorerView();
        }

        private void ExtractEmbeddedDocumentToolStripMenuItem_Click(object? sender, EventArgs e)
        {
            bool enabled = !ExtractEmbeddedDocumentToolStripMenuItem.Checked;
            Options.ExtractEmbeddedMetadata = GetToggledFlags(Options.ExtractEmbeddedMetadata, ExtractEmbeddedMetadata.Document, enabled);
            ExtractEmbeddedDocumentToolStripMenuItem.Checked = enabled;
            SaveOptions();
            ResetFileDetails();
            UpdateExplorerView();
        }

        private void ConvertEmbeddedImageToolStripMenuItem_Click(object? sender, EventArgs e) {
            bool enabled = !ConvertEmbeddedImageToolStripMenuItem.Checked;
            Options.ConvertEmbeddedToEasyMetadata = GetToggledFlags(Options.ConvertEmbeddedToEasyMetadata, ConvertEmbeddedToEasyMetadata.Image, enabled);
            ConvertEmbeddedImageToolStripMenuItem.Checked = enabled;
            SaveOptions();
            ResetFileDetails();
            UpdateExplorerView();
        }

        private void ConvertEmbeddedVideoToolStripMenuItem_Click(object? sender, EventArgs e)
        {
            bool enabled = !ConvertEmbeddedVideoToolStripMenuItem.Checked;
            Options.ConvertEmbeddedToEasyMetadata = GetToggledFlags(Options.ConvertEmbeddedToEasyMetadata, ConvertEmbeddedToEasyMetadata.Video, enabled);
            ConvertEmbeddedVideoToolStripMenuItem.Checked = enabled;
            SaveOptions();
            ResetFileDetails();
            UpdateExplorerView();
        }

        private void ConvertEmbeddedAudioToolStripMenuItem_Click(object? sender, EventArgs e)
        {
            bool enabled = !ConvertEmbeddedAudioToolStripMenuItem.Checked;
            Options.ConvertEmbeddedToEasyMetadata = GetToggledFlags(Options.ConvertEmbeddedToEasyMetadata, ConvertEmbeddedToEasyMetadata.Audio, enabled);
            ConvertEmbeddedAudioToolStripMenuItem.Checked = enabled;
            SaveOptions();
            ResetFileDetails();
            UpdateExplorerView();
        }

        private void ConvertEmbeddedDocumentToolStripMenuItem_Click(object? sender, EventArgs e)
        {
            bool enabled = !ConvertEmbeddedDocumentToolStripMenuItem.Checked;
            Options.ConvertEmbeddedToEasyMetadata = GetToggledFlags(Options.ConvertEmbeddedToEasyMetadata, ConvertEmbeddedToEasyMetadata.Document, enabled);
            ConvertEmbeddedDocumentToolStripMenuItem.Checked = enabled;
            SaveOptions();
            ResetFileDetails();
            UpdateExplorerView();
        }

        private void LanguageToolStripMenuItem_Click(object? sender, EventArgs e)
        {

            string[] sa = [.. Enum.GetValues<Language>().Select(x => $"{x.GetGlobalStringValue()} ({x.GetValue<string>()})")];
            using DropDownListDialog ddld = new(sa, Properties.Settings.Default.LanguageIndex);
            if (ddld.ShowDialog() == DialogResult.OK)
            {
                Properties.Settings.Default.LanguageIndex = ddld.SelectedIndex;
                Restart();
            }
        }

        private void PreserveDateCreatedToolStripMenuItem_Click(object? sender, EventArgs e)
        {
            bool b = !PreserveDateCreatedToolStripMenuItem.Checked;
            Options.PreserveDateCreated = b;
            PreserveDateCreatedToolStripMenuItem.Checked = b;
            SaveOptions();
        }

        private void PreserveDateModifiedToolStripMenuItem_Click(object? sender, EventArgs e)
        {
            bool b = !PreserveDateModifiedToolStripMenuItem.Checked;
            Options.PreserveDateModified = b;
            PreserveDateModifiedToolStripMenuItem.Checked = b;
            SaveOptions();
        }

        private void LogApplicationEventsToolStripMenuItem_Click(object? sender, EventArgs e)
        {
            bool b = !LogApplicationEventsToolStripMenuItem.Checked;
            Properties.Settings.Default.LogApplicationEvents = b;
            LogApplicationEventsToolStripMenuItem.Checked = b;
            Properties.Settings.Default.Save();
        }

        private void ShutDownUponCompletionToolStripMenuItem_Click(object? sender, EventArgs e)
        {
            bool b = !ShutDownUponCompletionToolStripMenuItem.Checked;
            Properties.Settings.Default.ShutdownUponCompletion = b;
            ShutDownUponCompletionToolStripMenuItem.Checked = b;
            Properties.Settings.Default.Save();
        }

        private void SplashScreenToolStripMenuItem_Click(object? sender, EventArgs e) => ShowSplash();

        private void AboutEasyFileManagerToolStripMenuItem_Click(object? sender, EventArgs e)
        {
            using MessageDialog md = new(Globals.AboutMessage);
            md.ShowDialog();
        }

        private void RestoreDefaultToolStripMenuItem_Click(object? sender, EventArgs e)
        {
            RestoreToDefault();
            InitializeLayout();
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

        private static void ExplorerToolStripMenuItem_Click(object? sender, EventArgs _, bool open = false)
        {
            if (sender is ToolStripMenuItem tsmi && tsmi.Owner is ContextMenuStrip cms && cms.SourceControl is DataGridView dgv)
            {
                if (dgv.SelectedRows.Count == 1 && dgv.SelectedRows[0].Tag is EasyPath ep)
                {
                    string path = GetExplorerPath(ep);
                    Process.Start(EXPLORER_EXE, open ? path : GetParentPath(path));
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
            List<object> items = [.. KeywordsComboBox.Items.Cast<object>()];
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
                Options.Keywords.Replace([.. items.Cast<EasyList<string>>()]);
                Options.KeywordsIndex = eld.SelectedIndex;

                UpdateKeywordsControlsA();
            }
            KeywordsGroupBox.Focus();
            UpdatePreviewFormatting();
        }

        private void GPSEditButton_Click(object? sender, EventArgs e)
        {
            List<object> items = [.. GPSComboBox.Items.Cast<object>()];
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
                GPSComboBox.Items.AddRange([.. items]);
                GPSComboBox.SelectedItem = eld.SelectedItem;

                Options.GeoAreas = GPSComboBox.GetItems<KeyValuePair<string, GeoCoordinates?>>().ToDictionary(x => x.Key, x => x.Value);
                Options.GPSCustomIndex = GPSComboBox.SelectedIndex;

                UpdateGPSControlsA();
                UpdatePreviewFormatting();
            }
            GeolocationGroupBox.Focus();
        }

        private void DateEditButton_Click(object? sender, EventArgs e)
        {
            List<object> items = [.. DateComboBox.Items.Cast<object>()];
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
                DateComboBox.Items.AddRange([.. items]);
                DateComboBox.SelectedItem = eld.SelectedItem;

                Options.Dates.Replace(items.Cast<string>());
                Options.DateCustomIndex = DateComboBox.SelectedIndex;

                UpdateDateControlsA();
                UpdatePreviewFormatting();
            }
            DateGroupBox.Focus();
        }

        private void EasyMetadataEditButton_Click(object? sender, EventArgs e)
        {
            List<object> items = [.. EasyMetadataComboBox.Items.Cast<object>()];
            using EditListDialog eld = new(ref items, font: Font, editButtons: EditButtons.AddRemoveMoveUpMoveDownSortEdit);
            eld.OnAddItem = async () =>
            {
                if (await GetEasyMetadataAsync() is object o)
                {
                    eld.SelectedIndex = eld.Items.Add(o);
                }
            };
            eld.OnEditItem = async () =>
            {
                if (await GetEasyMetadataAsync(eld.SelectedItem) is object o)
                {
                    int i = eld.SelectedIndex;
                    eld.Items.RemoveAt(i);
                    eld.Items.Insert(i, o);
                    eld.SelectedIndex = i;
                }
            };
            if (eld.ShowDialog() == DialogResult.OK)
            {
                Options.EasyMetadata.Replace([.. items.Cast<string>()]);
                Options.EasyMetadataIndex = eld.SelectedIndex;

                UpdateEasyMetadataControlsA();
                UpdatePreviewFormatting();
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
                    SaveOptions();
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
                    SaveOptions();
                }
            }
            DuplicatesGroupBox.Focus();
        }

        private void CleanUpEditButton_Click(object? sender, EventArgs e)
        {
            List<object> items = [.. Options.CleanUpParameters.GetContainingFlags().Select(x => (object)x.GetEasyGlobalStringValue())];
            using EditListDialog eld = new(ref items, font: Font, editButtons: EditButtons.AddRemove);
            object[] oa = [.. EasyCleanUpParameter.All.GetContainingFlags().Select(x => x.GetEasyGlobalStringValue())];
            eld.OnAddItem = () =>
            {
                using DropDownListDialog ddld = new(oa, 0, font: Font);
                while (ddld.ShowDialog() == DialogResult.OK)
                {
                    if (ddld.SelectedItem is object o && !eld.Items.Contains(o))
                    {
                        eld.SelectedIndex = eld.Items.Add(ddld.SelectedItem);
                        break;
                    }
                }
                return Task.CompletedTask;
            };
            if (eld.ShowDialog() == DialogResult.OK)
            {
                Options.CleanUpParameters = EasyCleanUpParameter.None;
                foreach (string s in items.Cast<string>())
                {
                    Options.CleanUpParameters |= s.AsEasyEnumFromGlobal<EasyCleanUpParameter>();
                }
                SaveOptions();
                UpdateCleanUpControlsB();
            }
            CleanUpGroupBox.Focus();
        }

        private void DuplicatesCompareEditButton_Click(object? sender, EventArgs e)
        {
            List<object> items = [.. Options.DuplicatesCompareParameters.GetContainingFlags().Select(x => (object)x.GetEasyGlobalStringValue())];
            using EditListDialog eld = new(ref items, font: Font, editButtons: EditButtons.AddRemove);
            object[] oa = [.. EasyCompareParameter.All.GetContainingFlags().Select(x => x.GetEasyGlobalStringValue())];
            eld.OnAddItem = () =>
            {
                using DropDownListDialog ddld = new(oa, 0, font: Font);
                while (ddld.ShowDialog() == DialogResult.OK)
                {
                    if (ddld.SelectedItem is object o && !eld.Items.Contains(o))
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
                SaveOptions();
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

            List<Action<CancellationToken>> actions = [];
            if (Options.HasCustomizing())
            {
                actions.Add(Customize);
            }
            if (Options.HasRenaming() || Options.HasCopyingMoving())
            {
                actions.Add(ApplyFormatting);
            }
            if (Options.HasFinalizing())
            {
                if (Options.CleanUpEnabled)
                {
                    actions.Add(CleanUp);
                }
                if (Options.DuplicatesState != CheckState.Unchecked)
                {
                    actions.Add(DeleteOrStoreDuplicates);
                }
            }
            if (Properties.Settings.Default.LogApplicationEvents)
            {
                Logger.Reset();
            }
            EasyMetadata.Options.Reset();
            await UpdateAsync([.. actions]);

            Folders.Replace(Folder.GetDirectoryPaths());
            Files.Replace(Folder.GetFilePaths(), true, false);

            UpdateApplyButton(false);
            UpdateExplorerView();

            if (Properties.Settings.Default.ShutdownUponCompletion)
            {
                static void shutdown() { ShutDown(); Exit(); }
                using DelayedActionDialog dad = new(shutdown, caption: Globals.ShuttingDown);
                dad.ShowDialog();
            }
            else if (Properties.Settings.Default.LogApplicationEvents)
            {
                Logger.Show();
            }
            PathTableLayoutPanel.Focus();
        }

        private void SelectTopFolderButton_Click(object? sender, EventArgs e)
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
                    UpdatePreviewFormatting();
                }
            }
            TopFolderGroupBox.Focus();
        }

        private void SubfoldersEditButton_Click(object? sender, EventArgs e)
        {
            List<object> items = [.. Options.Subfolders.Select(x => (object)x.GetEasyGlobalStringValue())];
            using EditListDialog eld = new(ref items, font: Font, editButtons: EditButtons.AddRemoveMoveUpMoveDown);
            object[] oa = [.. Enum.GetValues<EasySubfolder>().Select(x => x.GetEasyGlobalStringValue())];
            eld.OnAddItem = () =>
            {
                using DropDownListDialog ddld = new(oa, 0, font: Font);
                if (ddld.ShowDialog() == DialogResult.OK && ddld.SelectedItem is object o)
                {
                    eld.SelectedIndex = eld.Items.Add(o);
                }
                return Task.CompletedTask;
            };
            if (eld.ShowDialog() == DialogResult.OK)
            {
                Options.Subfolders.Replace(items.Select(x => ((string)x).AsEasyEnumFromGlobal<EasySubfolder>()));
                UpdateFolderControlsB();
                UpdatePreviewFormatting();
            }
            SubfoldersGroupBox.Focus();
        }

        private void FilterEditButton_Click(object? sender, EventArgs e)
        {
            List<object> items = [.. Options.TypeFilter.GetContainingFlags().Select(x => (object)x.GetEasyGlobalStringValue())];
            using EditListDialog eld = new(ref items, font: Font, editButtons: EditButtons.AddRemove);
            object[] oa = [.. EasyTypeFilter.All.GetContainingFlags().Select(x => x.GetEasyGlobalStringValue())];
            eld.OnAddItem = () =>
            {
                using DropDownListDialog ddld = new(oa, 0, font: Font);
                while (ddld.ShowDialog() == DialogResult.OK)
                {
                    if (ddld.SelectedItem is object o && !eld.Items.Contains(o))
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
                SaveOptions();
                UpdateFilterControlsB();
                UpdatePreviewFormatting();
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
            if (s != Options.Title) { Options.Title = s; SaveOptions(); }
        }

        private void SubjectTextBox_Leave(object? sender, EventArgs e)
        {
            string s = SubjectTextBox.Text;
            if (s != Options.Subject) { Options.Subject = s; SaveOptions(); }
        }

        private void CommentTextBox_Leave(object? sender, EventArgs e)
        {
            string s = CommentTextBox.Text;
            if (s != Options.Comment) { Options.Comment = s; SaveOptions(); }
        }

        private void PrefixTextBox_Leave(object? sender, EventArgs e)
        {
            string s = PrefixTextBox.Text;
            if (s != Options.Prefix)
            {
                Options.Prefix = s;
                SaveOptions();
                UpdatePreviewFormatting();
            }
        }

        private void ReplaceTextBox_Leave(object? sender, EventArgs e)
        {
            string s = ReplaceTextBox.Text;
            if (s != Options.Replace)
            {
                Options.Replace = s;
                SaveOptions();
                UpdatePreviewFormatting();
            }
        }

        private void WithTextBox_Leave(object? sender, EventArgs e)
        {
            string s = WithTextBox.Text;
            if (s != Options.With)
            {
                Options.With = s;
                SaveOptions();
                UpdatePreviewFormatting();
            }
        }

        private void SuffixTextBox_Leave(object? sender, EventArgs e)
        {
            string s = SuffixTextBox.Text;
            if (s != Options.Suffix)
            {
                Options.Suffix = s;
                SaveOptions();
                UpdatePreviewFormatting();
            }
        }

        private void FilterStringTextBox_Leave(object? sender, EventArgs e)
        {
            string s = FilterStringTextBox.Text;
            if (s != Options.FilterString)
            {
                Options.FilterString = s;
                SaveOptions();
                UpdateFilterControlsC();
                UpdatePreviewFormatting();
            }
        }

        #endregion

        #region CheckBox

        private async void PreviewFormattingTimer_TickAsync(object? sender, EventArgs e)
        {
            _previewFormattingTimer.Stop();
            if (!_isClosing)
            {
                await UpdateFormattingAsync();
            }
        }

        private void CustomizeCheckBox_CheckedChanged(object? sender, EventArgs e)
        {
            Options.CustomizeEnabled = CustomizeCheckBox.Checked;
            SaveOptions();
            UpdateCustomizePanel();
            UpdateEditPanel(customize: true);
            UpdatePreviewFormatting();
        }

        private void TitleCheckBox_CheckedChanged(object? sender, EventArgs e)
        {
            Options.TitleEnabled = TitleCheckBox.Checked;
            TitleTextBox.Enabled = Options.TitleEnabled;
            SaveOptions();
        }

        private void SubjectCheckBox_CheckedChanged(object? sender, EventArgs e)
        {
            Options.SubjectEnabled = SubjectCheckBox.Checked;
            SubjectTextBox.Enabled = Options.SubjectEnabled;
            SaveOptions();
        }

        private void CommentCheckBox_CheckedChanged(object? sender, EventArgs e)
        {
            Options.CommentEnabled = CommentCheckBox.Checked;
            CommentTextBox.Enabled = Options.CommentEnabled;
            SaveOptions();
        }

        private void KeywordsCheckBox_CheckStateChanged(object? sender, EventArgs e)
        {
            Options.KeywordsState = KeywordsCheckBox.CheckState;
            UpdateKeywordsControlsB();
            SaveOptions();
        }

        private void GeolocationCheckBox_CheckedChanged(object? sender, EventArgs e)
        {
            Options.GeolocationEnabled = GeolocationCheckBox.Checked;
            SaveOptions();
            UpdateGPSControlsB();
            UpdateGPSControlsA();
            UpdateGPSControlsA();
            UpdatePreviewFormatting();
        }

        private void WriteGPSCoordsCheckBox_CheckedChanged(object? sender, EventArgs e)
        {
            Options.WriteGPSCoords = WriteGPSCoordsCheckBox.Checked;
            SaveOptions();
            GeolocationGroupBox.Focus();
            UpdateGPSControlsB();
            UpdatePreviewFormatting();
        }

        private void WriteGPSAreaInfoCheckBox_CheckedChanged(object? sender, EventArgs e)
        {
            Options.WriteGPSAreaInfo = WriteGPSAreaInfoCheckBox.Checked;
            SaveOptions();
            GeolocationGroupBox.Focus();
            UpdateGPSControlsB();
            UpdatePreviewFormatting();
        }

        private void WriteGPSEasyMetadataCheckBox_CheckedChanged(object? sender, EventArgs e)
        {
            Options.WriteGPSEasyMetadata = WriteGPSEasyMetadataCheckBox.Checked;
            SaveOptions();
            GeolocationGroupBox.Focus();
            UpdateGPSControlsB();
            UpdatePreviewFormatting();
        }

        private void DateCheckBox_CheckStateChanged(object? sender, EventArgs e)
        {
            Options.DateState = DateCheckBox.CheckState;
            SaveOptions();
            UpdateDateControlsA();
            UpdatePreviewFormatting();
        }

        private void WriteDateModifiedCheckBox_CheckedChanged(object? sender, EventArgs e)
        {
            Options.WriteDateModified = WriteDateModifiedCheckBox.Checked;
            SaveOptions();
            UpdateDateControlsB();
            DateGroupBox.Focus();
            UpdatePreviewFormatting();
        }

        private void WriteDateCreatedCheckBox_CheckedChanged(object? sender, EventArgs e)
        {
            Options.WriteDateCreated = WriteDateCreatedCheckBox.Checked;
            SaveOptions();
            UpdateDateControlsB();
            DateGroupBox.Focus();
            UpdatePreviewFormatting();
        }

        private void WriteDateTakenOrEncodedCheckBox_CheckedChanged(object? sender, EventArgs e)
        {
            Options.WriteDateCamera = WriteDateTakenOrEncodedCheckBox.Checked;
            SaveOptions();
            UpdateDateControlsB();
            DateGroupBox.Focus();
            UpdatePreviewFormatting();
        }

        private void WriteDateEasyMetadataCheckBox_CheckedChanged(object? sender, EventArgs e)
        {
            Options.WriteDateEasyMetadata = WriteDateEasyMetadataCheckBox.Checked;
            SaveOptions();
            UpdateDateControlsB();
            DateGroupBox.Focus();
            UpdatePreviewFormatting();
        }

        private void EasyMetadataCheckBox_CheckedChanged(object? sender, EventArgs e)
        {
            Options.EasyMetadataEnabled = EasyMetadataCheckBox.Checked;
            SaveOptions();
            UpdateEasyMetadataControlsB();
            UpdatePreviewFormatting();
        }

        private void BackupFolderCheckBox_CheckedChanged(object? sender, EventArgs e)
        {
            Options.BackupFolderEnabled = BackupFolderCheckBox.Checked;
            UpdateBackupControls();
            SaveOptions();
        }

        private void CleanUpCheckBox_CheckedChanged(object? sender, EventArgs e)
        {
            Options.CleanUpEnabled = CleanUpCheckBox.Checked;
            UpdateCleanUpControlsA();
            SaveOptions();
        }

        private void DuplicatesCheckBox_CheckStateChanged(object? sender, EventArgs e)
        {
            Options.DuplicatesState = DuplicatesCheckBox.CheckState;
            UpdateDuplicatesControlsA();
            SaveOptions();
        }

        private void ShowDuplicatesCompareDialogCheckBox_CheckedChanged(object? sender, EventArgs e)
        {
            Options.ShowDuplicatesCompareDialog = ShowDuplicatesCompareDialogCheckBox.Checked;
            DuplicatesGroupBox.Focus();
            SaveOptions();
        }

        private void RenameCheckBox_CheckedChanged(object? sender, EventArgs e)
        {
            Options.RenameEnabled = RenameCheckBox.Checked;
            SaveOptions();
            UpdateRenamePanel();
            UpdateEditPanel(rename: true);
            UpdatePreviewFormatting();
        }

        private void CopyMoveCheckBox_CheckStateChanged(object? sender, EventArgs e)
        {
            Options.CopyMoveState = CopyMoveCheckBox.CheckState;
            SaveOptions();
            UpdateCopyMoveCheckBox();
            UpdateCopyMovePanel();
            UpdateEditPanel(copyMove: true);
            UpdatePreviewFormatting();
        }

        private void FinalizeCheckBox_CheckedChanged(object? sender, EventArgs e)
        {
            Options.FinalizeEnabled = FinalizeCheckBox.Checked;
            UpdateFinalizePanel();
            UpdateEditPanel(finalize: true);
            UpdateExplorerView();
            SaveOptions();
        }

        private void ReplaceWithCheckBox_CheckStateChanged(object? sender, EventArgs e)
        {
            Options.ReplaceWithState = ReplaceWithCheckBox.CheckState;
            SaveOptions();
            UpdateReplaceWithControls();
            UpdatePreviewFormatting();
        }

        private void DateFormatCheckBox_CheckedChanged(object? sender, EventArgs e)
        {
            Options.DateFormatEnabled = DateFormatCheckBox.Checked;
            SaveOptions();
            DateFormatComboBox.Enabled = Options.DateFormatEnabled;
            UpdatePreviewFormatting();
        }

        private void TopFolderCheckBox_CheckedChanged(object? sender, EventArgs e)
        {
            Options.TopFolderEnabled = TopFolderCheckBox.Checked;
            SaveOptions();
            UpdateFolderControlsA();
            UpdatePreviewFormatting();
        }

        private void FilterCheckBox_CheckedChanged(object? sender, EventArgs e)
        {
            Options.FilterEnabled = FilterCheckBox.Checked;
            SaveOptions();
            UpdateFilterControlsA();
            UpdatePreviewFormatting();
        }

        private void SubfoldersCheckBox_CheckedChanged(object? sender, EventArgs e)
        {
            Options.SubfoldersEnabled = SubfoldersCheckBox.Checked;
            SaveOptions();
            UpdateFolderControlsA();
            UpdatePreviewFormatting();
        }

        #endregion

        #region ComboBox

        private void KeywordsComboBox_SelectedIndexChanged(object? sender, EventArgs e)
        {
            Options.KeywordsIndex = KeywordsComboBox.SelectedIndex;
            SaveOptions();
            UpdatePreviewFormatting();
        }

        private void GPSSourceComboBox_SelectedIndexChanged(object? sender, EventArgs e)
        {
            Options.GPSSourceIndex = GPSSourceComboBox.SelectedIndex;
            SaveOptions();
            UpdateGPSControlsA();
            UpdatePreviewFormatting();
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
            SaveOptions();
            UpdatePreviewFormatting();
        }

        private void EasyMetadataComboBox_SelectedIndexChanged(object? sender, EventArgs e)
        {
            Options.EasyMetadataIndex = EasyMetadataComboBox.SelectedIndex;
            SaveOptions();
            UpdatePreviewFormatting();
        }

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
            SaveOptions();
            UpdatePreviewFormatting();
        }

        private void FilterNameComboBox_SelectedIndexChanged(object? sender, EventArgs e)
        {
            Options.NameFilter = (EasyNameFilter)FilterNameComboBox.SelectedIndex;
            SaveOptions();
            UpdatePreviewFormatting();
        }

        private void DateFormatComboBox_SelectedIndexChanged(object? sender, EventArgs e)
        {
            if (DateFormatComboBox.SelectedItem is object o)
            {
                Options.DateFormat = (string)o;
                SaveOptions();
            }
            UpdatePreviewFormatting();
        }

        #endregion
    }
}
