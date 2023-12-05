using System.Diagnostics;
using System.Reflection;
using Microsoft.Win32;

using WinFormsLib;

using static WinFormsLib.Forms;
using static WinFormsLib.Constants;
using static WinFormsLib.GeoLocator;

namespace EasyFileManager
{
    public partial class Main
    {
        private EasyList<string>? GetKeywords(object? editItem = null)
        {
            editItem ??= new EasyList<string>();
            List<object> items = new((EasyList<string>)editItem);
            using EditListDialog eld = new(ref items, font: Font, editButtons: Buttons.EditButtons.AddRemoveRenameMoveUpMoveDownSort);
            return eld.ShowDialog() == DialogResult.OK ? new EasyList<string>(items.Cast<string>()) : null;
        }

        private string[] GetSelectedFilePaths(bool recursive = false)
        {
            List<string> l = new();
            foreach (DataGridViewRow row in ExplorerDataGridView.SelectedRows)
            {
                if (row.Tag is EasyFile ef)
                {
                    l.Add(ef.Path);
                }
                else if (recursive && row.Tag is EasyFolder ed)
                {
                    l.AddRange(ed.GetFilePaths(true));
                }
            }
            return l.ToArray();
        }

        private string[] GetSelectedFolderPaths(bool recursive = false)
        {
            List<string> l = new();
            foreach (DataGridViewRow row in ExplorerDataGridView.SelectedRows)
            {
                if (row.Tag is EasyFolder ed)
                {
                    if (recursive)
                    {
                        l.AddRange(ed.GetDirectoryPaths(true));
                    }
                    l.Add(ed.Path);
                }
            }
            return l.ToArray();
        }

        private object? GetDate(object? editItem = null)
        {
            Map<string, object?> m = new() { { string.Empty, editItem } };
            using InputDialog id = new(ref m, font: Font);
            id.OnUpdate = () =>
            {
                string? s = (string?)m[string.Empty];
                m[string.Empty] = string.IsNullOrEmpty(s) ? s : s.AsDateTime() is DateTime dt ? dt.ToDateTimeString() : null;
                return Task.CompletedTask;
            };
            return id.ShowDialog() == DialogResult.OK ? m[string.Empty] : null;
        }

        private object? GetEasyMetadata(object? editItem = null)
        {
            EasyMetadata emd = editItem is not string s || string.IsNullOrEmpty(s) ? new() : new(s);
            Map<string, object?> m = emd.ToMap();
            using InputDialog id = new(ref m, font: Font);
            id.OnUpdate = () =>
            {
                foreach (KeyValuePair<string, string?> kvp in m.ToDictionary(kvp => kvp.Key, kvp => (string?)kvp.Value))
                {
                    string k = kvp.Key;
                    string? s = kvp.Value;
                    object? o = null;
                    if (!string.IsNullOrEmpty(s))
                    {
                        switch (k)
                        {
                            case nameof(EasyMetadata.Date):
                                if (s.AsDateTime() is DateTime dtfs)
                                {
                                    s = dtfs.ToDateTimeString();
                                }
                                break;
                            case nameof(EasyMetadata.GeoArea):
                                try { o = new GeoArea(s); } catch { }
                                break;
                            case nameof(EasyMetadata.Tags):
                                try { o = new EasyList<string>(s); } catch { }
                                break;
                            case nameof(EasyMetadata.CustomDict):
                                try { o = new Map<string, object?>(s); } catch { }
                                break;
                        }
                    }
                    m[k] = o;
                }
                return Task.CompletedTask;
            };
            return id.ShowDialog() == DialogResult.OK ? $"{new EasyMetadata(m)}" : null;
        }

        private async Task<KeyValuePair<string, GeoCoordinates?>?> GetGPSAsync(object? editItem = null)
        {
            editItem ??= new KeyValuePair<string, GeoCoordinates?>(string.Empty, null);
            KeyValuePair<string, GeoCoordinates?> kvp = (KeyValuePair<string, GeoCoordinates?>)editItem;
            string query = kvp.Value != null ? kvp.Value.ToShortString() : string.Empty;
            Map<string, object?> m = new() { { Globals.AreaInfo, kvp.Key }, { Globals.Geocoords, query } };
            using InputDialog id = new(ref m, font: Font);
            if (id.ShowDialog() == DialogResult.Cancel)
            {
                return null;
            }
            string n = (string)m[Globals.AreaInfo]!;
            string q = (string)m[Globals.Geocoords]!;
            if (string.IsNullOrEmpty(q) && !string.IsNullOrEmpty(n))
            {
                q = n;
            }
            GeoCoordinates? gc = null;
            if (await GetGeoObjectAsync(q) is GeoObject go)
            {
                gc = go.GetGeoCoordinates();
                if (gc == null)
                {
                    gc = kvp.Value;
                }
                else if (string.IsNullOrEmpty(n) && go.GetGeoAddress() is GeoAddress ga)
                {
                    n = ga.GetArea();
                }
            }
            return new(n, gc);
        }

        private static void CloseForms(bool includeMainForm = true, CloseReason closeReason = CloseReason.None)
        {
            Form[] forms = Application.OpenForms.Cast<Form>().ToArray();
            for (int i = forms.Length - 1; i >= Convert.ToSingle(!includeMainForm); i--)
            {
                Form form = forms[i];
                string n = form.Name;
                string d = string.Empty;
                if (form.Tag is CloseReason cr)
                {
                    d = $" ({cr}) ";
                    Debug.WriteLine(d);
                }
                else if (closeReason != CloseReason.None)
                {
                    form.Tag = closeReason;
                    d = $" ({closeReason}) ";
                }
                Debug.WriteLine($"Closing{d}'{n}'...");
                form.Close();
            }
        }

        private static void WaitUntilTaskCompleted()
        {
            try { _cancellationTokenSource.Cancel(); } catch { }
            while (!_task.IsCompleted)
            {
                try
                {
                    Debug.WriteLine("Waiting 10 milliseconds until previous task is completed.");
                    _task.Wait(10);
                }
                catch { }
            }
            try { _task.Dispose(); _cancellationTokenSource.Dispose(); _cancellationTokenSource = new(); } catch { }
        }

        private static bool FilteredPathExists(string path, EasyTypeFilter typeFilter = default, EasyNameFilter nameFilter = default, string filterString = EMPTY_STRING)
        {
            bool result = false;
            string n = Utils.GetFileNameWithoutExtension(path);
            if (string.IsNullOrEmpty(filterString) || (nameFilter switch
            {
                EasyNameFilter.StartsWith => n.StartsWith(filterString),
                EasyNameFilter.EndsWith => n.EndsWith(filterString),
                EasyNameFilter.Contains => n.Contains(filterString),
                _ => n.StartsWith(filterString),
            }))
            {
                EasyType et = Options.TypeFilter.GetContainingFlags().Select(x => x.GetValue<EasyType>()).Aggregate((x, y) => x |= y);
                if (et == EasyType.None)
                {
                    result = true;
                }
                else
                {
                    EasyPath ep = new(path);
                    if (et.HasFlag(ep.Type))
                    {
                        result = true;
                    }
                    ep.Dispose();
                }
            }
            return result;
        }

        private static string[] GetFilteredFilePaths(string[] paths, EasyTypeFilter typeFilter, EasyNameFilter nameFilter = default, string filterString = EMPTY_STRING)
        {
            return paths.Where(path => FilteredPathExists(path, typeFilter, nameFilter, filterString)).ToArray();
        }

        private static bool SetAddRemoveProgramsIcon()
        {
            using RegistryKey? uninstallRegKey = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall");
            bool result = false;
            if (uninstallRegKey != null)
            {
                string appName = GetAssemblyName();
                string iconPath = Path.Join(Application.StartupPath, @"Resources\Icons\ApplicationIcon.ico");

                string[] subKeyNames = uninstallRegKey.GetSubKeyNames();
                for (int i = 0; i < subKeyNames.Length; i++)
                {
                    using RegistryKey? regKey = uninstallRegKey.OpenSubKey(subKeyNames[i], true);
                    if (regKey != null && regKey.GetValue("DisplayName") as string == appName)
                    {
                        if (regKey.GetValue("DisplayIcon") as string != iconPath)
                        {
                            regKey.SetValue("DisplayIcon", iconPath);
                            result = true;
                        }
                        break;
                    }
                }
            }
            return result;
        }

        public static Main? GetMainForm() => Application.OpenForms.Count == 0 ? null : Application.OpenForms[0] as Main;

        public static Assembly GetAssembly() => Assembly.GetExecutingAssembly();

        public static string GetAssemblyName() => GetAssembly().GetSimpleName();

        public static string GetAppDataDirectoryPath() => Path.Join(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), GetAssemblyName());

        public static string GetAppDataConfigPath() => Path.Join(GetAppDataDirectoryPath(), $"{GetAssemblyName()}.cfg");
    }
}
