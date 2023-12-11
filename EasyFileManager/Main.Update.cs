using System.Diagnostics;

using WinFormsLib;

using static WinFormsLib.Chars;
using static WinFormsLib.Forms;
using static WinFormsLib.Buttons;
using static WinFormsLib.Constants;
using static WinFormsLib.GeoLocator;

namespace EasyFileManager
{
    public partial class Main
    {
        public static bool UpdateOptions(string path)
        {
            FileInfo fi = new(path);
            if (!fi.Exists)
            {
                return false;
            }
            string s = fi.ReadText();
            Options = new(s);
            Properties.Settings.Default.EasyOptions = s;
            Properties.Settings.Default.Save();
            return true;
        }

        private bool Update(string path, bool forceRefresh = false)
        {
            if (!Utils.IsValidDirectoryPath(path) || !Directory.Exists(path))
            {
                path = Folder.Path;
                while (!string.IsNullOrEmpty(path) && !Directory.Exists(path))
                {
                    path = Utils.GetDirectoryPath(path);
                }
                if (string.IsNullOrEmpty(path))
                {
                    path = STARTUP_DIRECTORY_DEFAULT;
                }
            }
            EasyMetadata.Options.Reset();
            if (forceRefresh || path != Folder.Path)
            {
                using EasyPath ep = new(path);
                if (!ep.Type.HasFlag(EasyType.System))
                {
                    try
                    {
                        FolderWatcher.EnableRaisingEvents = false;
                        FolderContentsWatcher.EnableRaisingEvents = false;
                        FolderContentsWatcher.Path = path;
                        FolderWatcher.Filter = Utils.GetDirectoryName(path);
                        FolderWatcher.Path = Utils.GetParentPath(path);
                        FolderWatcher.EnableRaisingEvents = true;
                        FolderContentsWatcher.EnableRaisingEvents = true;

                        Properties.Settings.Default.StartupDirectory = path;
                        Properties.Settings.Default.Save();

                        PathTextBox.Text = path;
                        Folder.Initialize(path);

                        Folders.Replace(Folder.GetDirectoryPaths());
                        Files.Replace(Folder.GetFilePaths());

                        UpdateHistory();
                        UpdateOneLevelUpButton();
                        UpdateExplorerView();
                        UpdateThumbnailPropsView();
                        UpdateBackupControls();
                        return true;
                    }
                    catch { }
                }
            }
            return false;
        }

        private void UpdateHistory()
        {
            if (_historyIndex == -1 || _history[_historyIndex] != Folder.Path)
            {
                for (int i = _history.Count - 1; i > _historyIndex; i--)
                {
                    _history.RemoveAt(i);
                }
                _history.Add(Folder.Path);
                _historyIndex = _history.Count - 1;
            }
            PreviousFolderButton.Enabled = _historyIndex != 0;
            NextFolderButton.Enabled = _historyIndex != _history.Count - 1;
        }

        private void UpdateApplyButton(bool isApplying)
        {
            IsApplying = isApplying;
            ApplyButton.Text = isApplying ? Globals.Abort : Globals.Apply;
            NextFolderButton.Enabled = !isApplying;
            PreviousFolderButton.Enabled = !isApplying;
            OneLevelUpButton.Enabled = !isApplying;
            PathTextBox.Enabled = !isApplying;
            SelectFolderButton.Enabled = !isApplying;
            ExplorerThumbnailPropsSplitContainer.Enabled = !isApplying;
            EditTableLayoutPanel.Enabled = !isApplying;
            MainMenuStrip.Enabled = !isApplying;
        }

        private void UpdateOneLevelUpButton() => OneLevelUpButton.Enabled = !string.IsNullOrEmpty(Folder.DirectoryPath);

        private void UpdateEditPanel(bool customize = false, bool rename = false, bool move = false, bool finalize = false)
        {
            SplitContainer esc = ExplorerEditSplitContainer;
            SplitContainer csc = CustomizeSplitContainer;
            SplitContainer rsc = RenameSplitContainer;
            SplitContainer msc = MoveSplitContainer;
            SplitContainer fsc = FinalizeSplitContainer;

            int ch = csc.Panel2Collapsed ? csc.SplitterDistance : _customizeSplitContainerHeight;
            int rh = rsc.Panel2Collapsed ? rsc.SplitterDistance : _renameSplitContainerHeight;
            int mh = msc.Panel2Collapsed ? msc.SplitterDistance : _moveSplitContainerHeight;
            int fh = fsc.Panel2Collapsed ? fsc.SplitterDistance : _finalizeSplitContainerHeight;
            int sw = csc.SplitterWidth + rsc.SplitterWidth + msc.SplitterWidth + fsc.SplitterWidth;
            int mb = csc.Margin.Bottom + rsc.Margin.Bottom + msc.Margin.Bottom + fsc.Margin.Bottom;
            int mt = csc.Margin.Top + rsc.Margin.Top + msc.Margin.Top + fsc.Margin.Top;
            int sd = esc.Height - ch - rh - mh - fh - sw - mb - mt;

            esc.SuspendDrawing();
            if (customize) { csc.Height = ch; }
            if (rename) { rsc.Height = rh; }
            if (move) { msc.Height = mh; }
            if (finalize) { fsc.Height = fh; }
            esc.SplitterDistance = sd;
            esc.ResumeDrawing();
        }

        private void UpdateCustomizePanel() => CustomizeSplitContainer.Panel2Collapsed = !CustomizeCheckBox.Checked;

        private void UpdateRenamePanel() => RenameSplitContainer.Panel2Collapsed = !RenameCheckBox.Checked;

        private void UpdateMovePanel() => MoveSplitContainer.Panel2Collapsed = !MoveCheckBox.Checked;

        private void UpdateFinalizePanel() => FinalizeSplitContainer.Panel2Collapsed = !FinalizeCheckBox.Checked;

        private void UpdateExplorerView()
        {
            IsUpdating = true;

            StoreSelectedPaths();

            foreach (DataGridViewRow row in ExplorerDataGridView.Rows)
            {
                row.Tag = null;
                row.Dispose();
            }
            ExplorerDataGridView.Rows.Clear();
            ExplorerDataGridView.SuspendLayout();
            foreach (EasyFolder ed in Folders)
            {
                if (!Properties.Settings.Default.ShowHiddenItems && ed.IsHidden)
                {
                    continue;
                }
                using DataGridViewRow row = ExplorerDataGridView.AddRow(ed.Name);
                row.Tag = ed;
                if (ed.IsHidden)
                {
                    row.DefaultCellStyle.BackColor = row.DefaultCellStyle.BackColor.Mixed(Color.LightGray, 0.9);
                }
            }
            foreach (EasyFile ef in Files)
            {
                if (!Properties.Settings.Default.ShowHiddenItems && ef.IsHidden)
                {
                    continue;
                }
                using DataGridViewRow row = ExplorerDataGridView.AddRow(ef.Name);
                row.Tag = ef;
                if (ef.IsHidden)
                {
                    row.DefaultCellStyle.BackColor = row.DefaultCellStyle.BackColor.Mixed(Color.LightGray, 0.9);
                }
            }
            ExplorerDataGridView.ResumeLayout(true);

            RestoreSelectedPaths();

            IsUpdating = false;

            UpdateSelection();
        }

        private void UpdateSelection()
        {
            if (IsApplying)
            {
                return;
            }
            UpdateThumbnailPropsView();
            if (IsUpdating)
            {
                return;
            }
            Utils.DelayAsync(UpdateFormattingAsync);
        }

        private void UpdateReplaceWithControls()
        {
            ReplaceTextBox.Enabled = ReplaceWithCheckBox.Checked;
            WithTextBox.Enabled = ReplaceWithCheckBox.Checked;
            ReplaceGroupBox.Text = Options.ReplaceWithState == CheckState.Indeterminate ? Globals.ReplaceRegex : Globals.ReplaceNormal;
        }

        private void UpdateKeywordsControlsA()
        {
            KeywordsComboBox.Items.Clear();
            KeywordsComboBox.Items.AddRange(Options.Keywords.ToArray());
            KeywordsComboBox.SelectedIndex = Options.KeywordsIndex;
            UpdateKeywordsControlsB();
        }

        private void UpdateKeywordsControlsB()
        {
            bool enabled = Options.KeywordsState != CheckState.Unchecked;
            KeywordsComboBox.Enabled = enabled && KeywordsComboBox.Items.Count != 0;
            KeywordsEditButton.Enabled = enabled;
            KeywordsGroupBox.Text = Options.KeywordsState switch
            {
                CheckState.Checked => Globals.KeywordsReplace,
                CheckState.Indeterminate => Globals.KeywordsAdd,
                _ => Globals.KeywordsDefault,
            };
        }

        private void UpdateGPSControlsA()
        {
            bool enabled = Options.GeolocationEnabled;
            WriteGPSCoordsCheckBox.Enabled = enabled;
            WriteGPSAreaInfoCheckBox.Enabled = enabled;
            WriteGPSEasyMetadataCheckBox.Enabled = enabled;

            object[] items;
            GPSComboBox.Items.Clear();
            if (Options.GPSSourceIndex == 0)
            {
                items = Options.GeoAreas.Cast<object>().ToArray();
                GPSComboBox.Items.AddRange(items);
                GPSComboBox.SelectedIndex = Options.GPSCustomIndex;
                GPSEditButton.Visible = true;
            }
            else
            {
                items = Enum.GetValues<GeoAreaDescription>().Select(x => x.GetGlobalStringValue()).ToArray();
                GPSComboBox.Items.AddRange(items);
                GPSComboBox.SelectedIndex = Options.GPSAreaIndex;
                GPSEditButton.Visible = false;
            }
            GPSComboBox.UpdateDropDownList();
            UpdateGPSControlsB();
        }

        private void UpdateGPSControlsB()
        {
            bool enabled = Options.GeolocationEnabled;
            bool writeEnabled = Options.WriteGPSCoords || Options.WriteGPSAreaInfo || Options.WriteGPSEasyMetadata;
            GPSSourceComboBox.Enabled = enabled && writeEnabled;
            GPSComboBox.Enabled = enabled && writeEnabled && GPSComboBox.Items.Count != 0;
            GPSEditButton.Enabled = enabled;
        }

        private void UpdateDateControlsA()
        {
            bool enabled = Options.DateState != CheckState.Unchecked;
            WriteDateCreatedCheckBox.Enabled = enabled;
            WriteDateTakenOrEncodedCheckBox.Enabled = enabled;
            WriteDateModifiedCheckBox.Enabled = enabled;
            WriteDateEasyMetadataCheckBox.Enabled = enabled;

            object[] items;
            DateComboBox.Items.Clear();
            switch (Options.DateState)
            {
                case CheckState.Unchecked:
                default:
                    items = Array.Empty<object>();
                    DateGroupBox.Text = Globals.DateDefault;
                    DateEditButton.Visible = true;
                    break;
                case CheckState.Checked:
                    DateGroupBox.Text = Globals.DateCustomize;
                    items = Options.Dates.ToArray();
                    DateComboBox.Items.AddRange(items);
                    DateComboBox.SelectedIndex = Options.DateCustomIndex;
                    DateEditButton.Visible = true;
                    break;
                case CheckState.Indeterminate:
                    DateGroupBox.Text = Globals.DateMetadata;
                    items = Enum.GetValues<EasyDateSource>().Select(x => x.GetEasyGlobalStringValue()).ToArray();
                    DateComboBox.Items.AddRange(items);
                    DateComboBox.SelectedIndex = Options.DateSourceIndex;
                    DateEditButton.Visible = false;
                    break;
            }
            DateComboBox.UpdateDropDownList();
            UpdateDateControlsB();
        }

        private void UpdateDateControlsB()
        {
            bool enabled = Options.DateState != CheckState.Unchecked;
            bool writeEnabled = Options.WriteDateCreated || Options.WriteDateCamera || Options.WriteDateModified || Options.WriteDateEasyMetadata;
            DateComboBox.Enabled = enabled && writeEnabled && DateComboBox.Items.Count != 0;
            DateEditButton.Enabled = enabled;
        }

        private void UpdateEasyMetadataControlsA()
        {
            EasyMetadataComboBox.Items.Clear();
            EasyMetadataComboBox.Items.AddRange(Options.EasyMetadata.ToArray());
            EasyMetadataComboBox.SelectedIndex = Options.EasyMetadataIndex;
            UpdateEasyMetadataControlsB();
        }

        private void UpdateEasyMetadataControlsB()
        {
            bool enabled = Options.EasyMetadataEnabled;
            EasyMetadataComboBox.Enabled = enabled && EasyMetadataComboBox.Items.Count != 0;
            EasyMetadataEditButton.Enabled = enabled;
        }

        private void UpdateBackupControls()
        {
            BackupFolderGroupBox.Text = Options.BackupFolderState switch
            {
                CheckState.Checked => Globals.BackupFolderMove,
                CheckState.Indeterminate => Globals.BackupFolderCopy,
                _ => Globals.BackupFolderDefault,
            };
            SelectBackupFolderButton.Enabled = Options.BackupFolderState == CheckState.Checked;
            BackupFolderTextBox.BackColor = Options.BackupFolderState != CheckState.Unchecked ? SystemColors.ControlLightLight : SystemColors.Control;
            BackupFolderTextBox.Text = Options.BackupFolderState == CheckState.Indeterminate ? Properties.Settings.Default.StartupDirectory : Options.BackupFolderPath;
        }

        private void UpdateFolderControlsA()
        {
            SelectTopFolderButton.Enabled = Options.TopFolderEnabled;
            TopFolderTextBox.BackColor = Options.TopFolderEnabled ? SystemColors.ControlLightLight : SystemColors.Control;
            SubfoldersTextBox.BackColor = Options.SubfoldersEnabled ? SystemColors.ControlLightLight : SystemColors.Control;
            SubfoldersEditButton.Enabled = Options.SubfoldersEnabled;
            UpdateFolderControlsB();
        }

        private void UpdateFolderControlsB()
        {
            string[] sa = Options.Subfolders.Select(x => $"{LEFT_CURLY_BRACE}{x.GetEasyGlobalStringValue()}{RIGHT_CURLY_BRACE}").ToArray();
            SubfoldersTextBox.Text = string.Join($"{SPACE}{BACKSLASH}{SPACE}", sa);
        }

        private void UpdateFilterControlsA()
        {
            bool enabled = Options.FilterEnabled;
            FilterStringTextBox.Enabled = enabled;
            FilterTypeTextBox.BackColor = enabled ? SystemColors.ControlLightLight : SystemColors.Control;
            FilterEditButton.Enabled = enabled;
            UpdateFilterControlsB();
            UpdateFilterControlsC();
        }

        private void UpdateFilterControlsB()
        {
            string[] sa = Options.TypeFilter.GetContainingFlags().Select(x => x.GetEasyGlobalStringValue()).ToArray();
            FilterTypeTextBox.Text = string.Join($"{SPACE}{VERTICAL_BAR}{SPACE}", sa);
        }

        private void UpdateFilterControlsC() => FilterNameComboBox.Enabled = Options.FilterEnabled && !string.IsNullOrEmpty(FilterStringTextBox.Text);

        private void UpdateDuplicatesControlsA()
        {
            DuplicatesGroupBox.Text = Options.DuplicatesState switch
            {
                CheckState.Checked => Globals.DuplicatesMove,
                CheckState.Indeterminate => Globals.DuplicatesDelete,
                _ => Globals.DuplicatesDefault,
            };
            bool storeDuplicates = Options.DuplicatesState == CheckState.Checked;
            bool compareDuplicates = Options.DuplicatesState != CheckState.Unchecked;
            SelectDuplicatesFolderButton.Enabled = storeDuplicates;
            DuplicatesFolderTextBox.BackColor = storeDuplicates ? SystemColors.ControlLightLight : SystemColors.Control;
            DuplicatesCompareTextBox.BackColor = compareDuplicates ? SystemColors.ControlLightLight : SystemColors.Control;
            ShowDuplicatesCompareDialogCheckBox.Enabled = compareDuplicates;
            DuplicatesCompareEditButton.Enabled = compareDuplicates;
            UpdateDuplicatesControlsB();
        }

        private void UpdateDuplicatesControlsB()
        {
            string[] sa = Options.DuplicatesCompareParameters.GetContainingFlags().Select(x => x.GetEasyGlobalStringValue()).ToArray();
            DuplicatesCompareTextBox.Text = string.Join($"{SPACE}{VERTICAL_BAR}{SPACE}", sa);
        }

        private void UpdateFormatting(bool apply, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();

            int maxValue = 100;
            int numSubSteps = 2;
            int subStepIndex = 0;
            int numValues;

            bool copyInsteadOfMove = false;
            bool filterEnabled = Options.MoveEnabled && Options.FilterEnabled;

            DataGridViewRowCollection rows = ExplorerDataGridView.Rows;

            if (!apply)
            {
                if (!GetSelectedFilePaths().Any())
                {
                    foreach (DataGridViewRow row in rows)
                    {
                        if (row.Tag is EasyFile)
                        {
                            row.Cells[0].Value = EMPTY_STRING;
                        }
                    }
                    return;
                }
            }
            else if (Options.MoveEnabled && Options.BackupFolderState != CheckState.Unchecked)
            {
                string startupFolderPath = Properties.Settings.Default.StartupDirectory;
                string backupFolderPath = Options.BackupFolderState == CheckState.Indeterminate ? startupFolderPath : Options.BackupFolderPath;
                if (!Utils.IsValidDirectoryPath(Options.BackupFolderPath))
                {
                    Options.BackupFolderPath = STARTUP_DIRECTORY_DEFAULT;
                }
                if (backupFolderPath == startupFolderPath)
                {
                    copyInsteadOfMove = true;
                }
                else
                {
                    numSubSteps += 1;

                    string backupName = $"{BACKUP_PREFIX}{DateTime.Now.ToDateSecondString()}";
                    string backupPath = Path.Join(Options.BackupFolderPath, backupName);
                    string[] paths = GetSelectedFilePaths(true);
                    if (filterEnabled)
                    {
                        paths = GetFilteredFilePaths(paths, Options.TypeFilter, Options.NameFilter, Options.FilterString);
                    }
                    numValues = paths.Length;
                    for (int i = 0; i < numValues; i++)
                    {
                        string path = paths[i];
                        EasyPath ep = new(path);
                        string bp = Utils.GetPathDuplicate(Path.Join(backupPath, ep.Name));
                        ep.Copy(bp, Options.PreserveDateCreated, Options.PreserveDateModified);
                        ep.Dispose();

                        Progress.Report(EasyProgress.GetValue(((maxValue * i) / numValues), maxValue, subStepIndex, numSubSteps), $"{path} -> {bp}");
                    }
                    subStepIndex += 1;

                    if (Options.LogApplicationEvents)
                    {
                        Logger.Write($"Backup copied in: {backupFolderPath}");
                    }
                }
            }

            cancellationToken.ThrowIfCancellationRequested();

            string dp = Options.TopFolderPath;
            if (string.IsNullOrEmpty(dp))
            {
                Options.TopFolderPath = Folder.Path;
            }
            OrderedMap<string, string> om = new(Utils.GetAllPaths(Options.TopFolderPath, true).ToDictionary(s => s));

            EasyFiles viewFiles = new();
            foreach (DataGridViewRow row in rows)
            {
                if (row.Tag is EasyFile ef)
                {
                    viewFiles.Add(ef);
                }
            }

            EasyFiles selectedFiles = new();
            foreach (DataGridViewRow row in ExplorerDataGridView.SelectedRows)
            {
                if (row.Tag is EasyFolder ed)
                {
                    string[] paths = ed.GetFilePaths(true);
                    if (filterEnabled)
                    {
                        paths = GetFilteredFilePaths(paths, Options.TypeFilter, Options.NameFilter, Options.FilterString);
                    }
                    selectedFiles.AddRange(paths);
                }
                else if (row.Tag is EasyFile ef)
                {
                    if (!filterEnabled || (NameFilteredPathExists(ef.Path, Options.NameFilter, Options.FilterString) && TypeFilteredPathExists(ef, Options.TypeFilter)))
                    {
                        selectedFiles.Add(ef);
                    }
                }
            }
            EasyFiles selectedViewFiles = new();
            numValues = selectedFiles.Count;
            for (int i = 0; i < numValues; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();

                EasyFile ef = selectedFiles[i];
                string p = ef.Path;
                om.Remove(p);
                ef.UpdateFormatting(Options, om.GetValues());
                string fp = ef.FormattedPath;
                om.Put(p, fp);
                if (apply)
                {
                    if (copyInsteadOfMove)
                    {
                        ef.Copy(fp, Options.PreserveDateCreated, Options.PreserveDateModified);
                        if (Options.LogApplicationEvents)
                        {
                            Logger.Write($"Copied {p} -> {fp}");
                        }
                    }
                    else
                    {
                        ef.Move(fp, Options.PreserveDateCreated, Options.PreserveDateModified);
                        if (Options.LogApplicationEvents)
                        {
                            Logger.Write($"Moved {p} -> {fp}");
                        }
                    }
                }
                else if (Options.LogApplicationEvents)
                {
                    Logger.Write($"Formatted {p} -> {fp}");
                }
                if (viewFiles.Contains(ef))
                {
                    selectedViewFiles.Add(ef);
                }
                else
                {
                    ef.Dispose();
                }

                Progress.Report(EasyProgress.GetValue(((maxValue * i) / numValues), maxValue, subStepIndex, numSubSteps), $"{p} -> {fp}");
            }

            subStepIndex += 1;

            numValues = rows.Count;
            for (int i = 0; i < numValues; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();

                DataGridViewRow row = rows[i];
                if (row.Tag is EasyFile ef)
                {
                    string p = ef.Path;
                    string fp = string.Empty;
                    if (!apply)
                    {
                        if (selectedViewFiles.Contains(ef))
                        {
                            fp = ef.FormattedPath;
                            string pp = $"{Folder.Path}{BACKSLASH}";
                            if (fp.StartsWith(pp))
                            {
                                fp = fp[pp.Length..];
                            }
                        }
                        row.Cells[0].Value = fp;
                    }
                }

                Progress.Report(EasyProgress.GetValue(((maxValue * i) / numValues), maxValue, subStepIndex, numSubSteps), string.Empty);
            }
            Options.TopFolderPath = dp;

            Progress.Report(maxValue);
        }

        private void UpdateFormatting(CancellationToken cancellationToken)
        {
            Debug.WriteLine("Updating formatting...");

            Progress.Reset();

            UpdateFormatting(false, cancellationToken);

            Debug.WriteLine("Formatted.");
        }

        private void ApplyFormatting(CancellationToken cancellationToken)
        {
            Debug.WriteLine("Applying formatting...");

            UpdateFormatting(true, cancellationToken);

            Debug.WriteLine("Formatted.");
        }

        private void Customize(CancellationToken cancellationToken)
        {
            Debug.WriteLine("Customizing...");

            int maxValue = 100;
            int numSubSteps = 2;

            cancellationToken.ThrowIfCancellationRequested();

            Progress.Report(EasyProgress.GetValue(1, maxValue, 0, numSubSteps));

            List<string> l = new();
            foreach (DataGridViewRow row in ExplorerDataGridView.SelectedRows)
            {
                if (row.Tag is EasyFolder ed)
                {
                    l.AddRange(ed.GetFilePaths(true));
                }
                else if (row.Tag is EasyFile ef)
                {
                    l.Add(ef.Path);
                }
            }

            int n = l.Count;
            for (int i = 0; i < n; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();

                string p = l[i];
                using EasyFile ef = new(p);
                ef.CustomizeAsync(Options).Wait(cancellationToken);
                Progress.Report(EasyProgress.GetValue(((maxValue * i) / n), maxValue, 1, numSubSteps), p);
                if (Options.LogApplicationEvents)
                {
                    Logger.Write($"Customized {p}");
                }
            }
            l.Clear();

            Progress.Report(maxValue);

            Debug.WriteLine("Customized.");
        }

        private void DeleteEmptyFolders(CancellationToken cancellationToken)
        {
            Debug.WriteLine("Deleting empty folders, if any...");

            int i = 0;
            foreach (string p in GetSelectedFolderPaths(true))
            {
                using EasyFolder ed = new(p);
                if (!ed.GetAllPaths().Any())
                {
                    if (ed.Delete())
                    {
                        if (Options.LogApplicationEvents)
                        {
                            Logger.Write($"Deleted {ed.Path}");
                        }
                        i += 1;
                    }
                }
                ed.Dispose();
            }
            if (i > 0)
            {
                Utils.Delay(() => { Debug.WriteLine($"Total empty folders deleted: {i}."); });
            }
            else
            {
                Debug.WriteLine($"No empty folders have been deleted.");
            }
        }

        private void DeleteOrStoreDuplicates(CancellationToken cancellationToken)
        {
            bool delete = Options.DuplicatesState == CheckState.Indeterminate;
            string deletOrStor = delete ? "Delet" : "Stor";

            Debug.WriteLine($"{deletOrStor}ing duplicates...");

            int maxValue = 100;
            int numSubSteps = 2;

            cancellationToken.ThrowIfCancellationRequested();

            Progress.Report(EasyProgress.GetValue(0, maxValue, 0, numSubSteps));

            string info;
            if (Utils.IsValidDirectoryPath(Options.DuplicatesFolderPath))
            {
                string[] filePaths = Options.MoveEnabled && Options.TopFolderEnabled && !string.IsNullOrEmpty(Options.TopFolderPath)
                    ? new EasyFolder(Options.TopFolderPath).GetFilePaths(true)
                    : GetSelectedFilePaths(true);

                cancellationToken.ThrowIfCancellationRequested();

                Progress.Report(EasyProgress.GetValue(0, maxValue, 1, numSubSteps));

                Map<string, string[][]> duplicates = EasyComparer.GetDuplicates(filePaths, Options.DuplicatesCompareParameters);

                if (Options.ShowDuplicatesCompareDialog)
                {
                    string deleteOrStore = delete ? Globals.DeleteDuplicates : Globals.StoreDuplicates;
                    string duplicatesName = $"{DUPLICATES_PREFIX}{DateTime.Now.ToDateSecondString()}";
                    string duplicatesPath = Path.Join(Options.DuplicatesFolderPath, duplicatesName);
                    foreach (string[][] nsa in duplicates.Values)
                    {
                        foreach (string[] sa in nsa)
                        {
                            EasyPaths<EasyPath> easyPaths = new(sa.Select(x => new EasyPath(x)));
                            Image[] images = easyPaths.SelectWhereNotNull(x => x.ExtraLargeThumbnail).ToArray();
                            using ImageTableDialog itd = new(images, Screen.PrimaryScreen.Bounds.Size, 0, buttons: DialogResultFlags.YesNo);
                            itd.OnSelectedIndexChanged = () =>
                            {
                                itd.CaptionText = string.Format(deleteOrStore, sa[itd.SelectedIndex]);
                                return Task.CompletedTask;
                            };
                            itd.OnSelectedIndexChanged.Invoke();
                            if (itd.ShowDialog() == DialogResult.Yes)
                            {
                                easyPaths.RemoveAt(itd.SelectedIndex);
                                foreach (EasyPath ep in easyPaths)
                                {
                                    if (delete)
                                    {
                                        ep.Delete();
                                    }
                                    else
                                    {
                                        if (Options.DuplicatesState == CheckState.Checked)
                                        {
                                            string filePath = Utils.GetPathDuplicate(Path.Join(duplicatesPath, ep.Name));
                                            ep.Move(filePath, Options.PreserveDateCreated, Options.PreserveDateModified);
                                        }
                                        ep.Dispose();
                                    }
                                }
                            }
                        }
                    }
                }
                else if (duplicates.Any())
                {
                    string deleteOrStoreAll = delete ? Globals.DeleteAllDuplicates : Globals.StoreAllDuplicates;
                    using MessageDialog md = new(deleteOrStoreAll, buttons: DialogResultFlags.YesNo);
                    if (md.ShowDialog() == DialogResult.Yes)
                    {
                        foreach (string[][] nsa in duplicates.Values)
                        {
                            foreach (string[] sa in nsa)
                            {
                                List<string> l = new(sa);
                                l.RemoveAt(0);
                                foreach (string s in l)
                                {
                                    EasyPath ep = new(s);
                                    if (delete)
                                    {
                                        ep.Delete();
                                    }
                                    else
                                    {
                                        if (Options.DuplicatesState == CheckState.Checked)
                                        {
                                            string p = Path.Join(Options.DuplicatesFolderPath, ep.Name);
                                            ep.Move(p, Options.PreserveDateCreated, Options.PreserveDateModified);
                                        }
                                        ep.Dispose();
                                    }
                                }
                            }
                        }
                    }
                }
                info = $"{deletOrStor}ed duplicates.";
            }
            else
            {
                using MessageDialog md = new(Globals.DuplicatesFolderDoesNotExist);
                md.ShowDialog();

                info = $"{deletOrStor}ing duplicates aborted.";
            }
            Progress.Report(EasyProgress.GetValue(100, maxValue, 1, numSubSteps));

            Debug.WriteLine(info);
            if (Options.LogApplicationEvents)
            {
                Logger.Write(info);
            }
        }

        private void StoreSelectedPaths()
        {
            DataGridViewSelectedRowCollection selectedRows = ExplorerDataGridView.SelectedRows;
            foreach (DataGridViewRow row in selectedRows)
            {
                if (row.Tag is EasyPath ep)
                {
                    _selectedPaths.Add(ep.Path);
                }
            }
        }

        private void RestoreSelectedPaths()
        {
            foreach (DataGridViewRow row in ExplorerDataGridView.Rows)
            {
                row.Selected = row.Tag is EasyPath ep && _selectedPaths.Contains(ep.Path);
            }
            if (ExplorerDataGridView.SelectedRows.Count == 0 && ExplorerDataGridView.Rows.Count != 0)
            {
                ExplorerDataGridView.Rows[0].Selected = true;
            }
            _selectedPaths.Clear();
        }

        private void RotateImage(EasyOrientation orientation)
        {
            DataGridViewSelectedRowCollection rows = ExplorerDataGridView.SelectedRows;
            if (rows.Count == 1 && rows[0].Tag is EasyFile ef)
            {
                ef.Rotate(orientation, Options.PreserveDateModified);

                Debug.WriteLine("Image rotated.");
            }
        }

        private async void UpdateThumbnailPropsView()
        {
            IsUpdating = true;

            WaitUntilTaskCompleted();

            PropsDataGridView.Rows.Clear();
            if (ThumbnailPictureBox.Image is Image image)
            {
                ThumbnailPictureBox.Image = null;
                image.Dispose();
            }
            bool showThumbnail = Properties.Settings.Default.ShowThumbnail;
            bool showProperties = Properties.Settings.Default.ShowProperties;
            DataGridViewSelectedRowCollection rows = ExplorerDataGridView.SelectedRows;
            if (!(showThumbnail || showProperties) || rows.Count != 1 || rows[0].Tag is not EasyPath easyPath || EasyType.Invalid.HasFlag(easyPath.Type))
            {
                ExplorerThumbnailPropsSplitContainer.Panel2Collapsed = true;
                IsUpdating = false;
                return;
            }
            EasyPath ep = (EasyPath)rows[0].Tag!;
            RotateTableLayoutPanel.Visible = showThumbnail && ep.Type == EasyType.Image;
            ExplorerThumbnailPropsSplitContainer.Panel2Collapsed = false;
            ThumbnailPropsSplitContainer.Panel1Collapsed = !showThumbnail;
            ThumbnailPropsSplitContainer.Panel2Collapsed = !showProperties;

            if (showThumbnail && ep.ExtraLargeThumbnail is Image img)
            {
                ThumbnailPictureBox.Image = img.Clone() as Image;
            }
            if (!showProperties)
            {
                IsUpdating = false;
                return;
            }
            PropsDataGridView.SuspendLayout();
            if (!ep.UpdateSize())
            {
                Debug.WriteLine($"Updating size async");

                bool aborted = false;
                Task t = Utils.RunCancellableTask(() =>
                {
                    Progress.Report(99, ep.Path);
                    try
                    {
                        long l = 0;
                        foreach (FileInfo fi in new DirectoryInfo(ep.Path).EnumerateFiles("*", SearchOption.AllDirectories))
                        {
                            _cancellationTokenSource.Token.ThrowIfCancellationRequested();
                            l += fi.Length;
                        }
                        ep.UpdateSize(l);
                    }
                    catch (OperationCanceledException) { aborted = true; }
                    Progress.Report(100, ep.Path);
                }, _cancellationTokenSource, () =>
                {
                    Debug.WriteLine($"Updating size async {(aborted ? "aborted" : "finished")}");
                });
                _task = t;
                await t;
                if (aborted)
                {
                    Progress.Report(100, ep.Path);
                    IsUpdating = false;
                    return;
                }
            }
            double mb = Math.Round(ep.Size / (double)1048576L, 2);
            PropsDataGridView.AddRow(Globals.Name, ep.Name);
            string ext = ep.Extension;
            if (!string.IsNullOrEmpty(ext))
            {
                ext = $" ({ep.Extension.ToUpper()})";
            }
            PropsDataGridView.AddRow(Globals.Type, $"{ep.Type.GetEasyGlobalStringValue()}{ext}");
            PropsDataGridView.AddRow(Globals.Size, $"{mb} MB");
            bool showMillis = Properties.Settings.Default.ShowMilliseconds;
            if (ep.DateModified is DateTime dtm)
            {
                PropsDataGridView.AddRow(Globals.DateModified, showMillis ? dtm.ToDateTimeString() : dtm.ToDateSecondString());
            }
            if (ep.DateCreated is DateTime dtc)
            {
                PropsDataGridView.AddRow(Globals.DateCreated, showMillis ? dtc.ToDateTimeString() : dtc.ToDateSecondString());
            }
            if (ep is EasyFolder ed)
            {
                PropsDataGridView.AddRow(Globals.Contents, ed.ContentText);
            }
            else if (ep is EasyFile ef)
            {
                if (ef.Type == EasyType.Image)
                {
                    if (ef.DateTaken is DateTime dtt)
                    {
                        PropsDataGridView.AddRow(Globals.DateTaken, showMillis ? dtt.ToDateTimeString() : dtt.ToDateSecondString());
                    }
                    if (ef.CameraManufacturer is string man && !string.IsNullOrEmpty(man))
                    {
                        PropsDataGridView.AddRow(Globals.CameraManufacturer, man);
                    }
                    if (ef.CameraModel is string mod && !string.IsNullOrEmpty(mod))
                    {
                        PropsDataGridView.AddRow(Globals.CameraModel, mod);
                    }
                }
                else if (ef.Type == EasyType.Video && ef.DateEncoded is DateTime dte)
                {
                    PropsDataGridView.AddRow(Globals.DateEncoded, showMillis ? dte.ToDateTimeString() : dte.ToDateSecondString());
                }
                PropsDataGridView.AddRow(Globals.Title, ef.Title);
                PropsDataGridView.AddRow(Globals.Subject, ef.Subject);
                PropsDataGridView.AddRow(Globals.Comment, ef.Comment);

                if (Properties.Settings.Default.ShowVideoMetadata)
                {
                    await ef.ExtractExifGeoAreaAsync(Options.UseEasyMetadataWithVideo);
                }

                PropsDataGridView.AddRow(Globals.KeywordsDefault, string.Join(COMMA, ef.Keywords));
                if (ef.GeoCoordinates != null && ef.GeoCoordinates.IsValid)
                {
                    PropsDataGridView.AddRow(Globals.Coordinates, ef.GeoCoordinates.ToString());
                }
                string ai = ef.AreaInfo;
                if (!string.IsNullOrEmpty(ai))
                {
                    PropsDataGridView.AddRow(Globals.AreaInfo, ai);
                }
                if (ef.GeoObject is GeoObject go && go.address is GeoAddress gad)
                {
                    PropsDataGridView.AddRow(Globals.Address, gad.ToString());
                }
                if (ef.EasyMetadata is EasyMetadata emd && emd.CustomDict is Dictionary<string, object> cd)
                {
                    PropsDataGridView.AddRow(Globals.CustomDict, cd.ToJson());
                }
            }
            foreach (DataGridViewRow row in PropsDataGridView.SelectedRows) { row.Selected = false; }

            PropsDataGridView.ResumeLayout(true);

            IsUpdating = false;
        }

        private async Task<bool> UpdateAsync(Action<CancellationToken> action) => await UpdateAsync(new Action<CancellationToken>[] { action });

        private async Task<bool> UpdateAsync(Action<CancellationToken>[] actions)
        {
            IsUpdating = true;

            WaitUntilTaskCompleted();

            using Task task = Task.Run(() =>
            {
                if (_isClosing)
                {
                    Exit();
                }
                else
                {
                    try
                    {
                        int n = actions.Length;
                        if (n > 0)
                        {
                            Progress.NumSteps = n;

                            for (int i = 0; i < n; i++)
                            {
                                Progress.StepIndex = i;

                                actions[i](_cancellationTokenSource.Token);
                            }
                        }
                    }
                    catch (OperationCanceledException)
                    {
                        Progress.Report(0);
                        _cancellationTokenSource.Dispose();

                        Debug.WriteLine("Refreshed progress.");
                    }
                }
            }, _cancellationTokenSource.Token);

            try
            {
                _task = task;
                await _task;
                if (_task.IsCompletedSuccessfully)
                {
                    IsUpdating = false;
                    return true;
                }
            }
            catch (TaskCanceledException) { }

            IsUpdating = false;
            return false;
        }

        private async Task<bool> UpdateFormattingAsync() => (Options.RenameEnabled || Options.MoveEnabled) && await UpdateAsync(UpdateFormatting);
    }
}
