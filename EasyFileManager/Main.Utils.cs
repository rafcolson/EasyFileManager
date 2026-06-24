using System.Diagnostics;
using System.Reflection;

using WinFormsLib;

using static WinFormsLib.Forms;
using static WinFormsLib.Constants;
using static WinFormsLib.GeoLocator;

namespace EasyFileManager
{
    public partial class Main
    {
        private static T GetToggledFlags<T>(T flags, T flag, bool enabled) where T : struct, Enum
        {
            long flagsValue = Convert.ToInt64(flags);
            long flagValue = Convert.ToInt64(flag);
            long value = enabled ? flagsValue | flagValue : flagsValue & ~flagValue;
            return (T)Enum.ToObject(typeof(T), value);
        }

        private EasyList<string>? GetKeywords(object? editItem = null)
        {
            editItem ??= new EasyList<string>();
            List<object> items = [.. (EasyList<string>)editItem];
            using EditListDialog eld = new(ref items, font: Font, editButtons: Buttons.EditButtons.AddRemoveRenameMoveUpMoveDownSort);
            return eld.ShowDialog() == DialogResult.OK ? [.. items.Cast<string>()] : null;
        }

        private string[] GetSelectedFilePaths(bool recursive = false)
        {
            List<string> l = [];
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
            return [.. l];
        }

        private string[] GetSelectedFolderPaths(bool recursive = false)
        {
            List<string> l = [];
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
            return [.. l];
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

        private static string GetExplorerPath(EasyPath easyPath) => easyPath.TargetPath ?? easyPath.Path;

        private static bool HasGeoAreaData(GeoArea geoArea)
        {
            return !string.IsNullOrEmpty(geoArea.AreaInfo) || geoArea.GeoCoords is GeoCoordinates gc && gc.IsValid;
        }

        private async Task<string?> GetEasyMetadataAsync(object? editItem = null)
        {
            EasyMetadata emd = editItem is not string s || string.IsNullOrEmpty(s) ? new() : new(s);
            Map<string, object?> m = emd.ToEditableMap();
            using InputDialog id = new(ref m, font: Font);
            id.OnUpdate = () =>
            {
                foreach (KeyValuePair<string, string?> kvp in m.ToDictionary(kvp => kvp.Key, kvp => kvp.Value?.ToString()))
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
                                    o = dtfs.ToDateTimeString();
                                }
                                break;
                            case nameof(EasyMetadata.GeoArea):
                                try
                                {
                                    GeoArea ga = new(s);
                                    o = HasGeoAreaData(ga) ? ga : s;
                                }
                                catch { o = s; }
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
            if (id.ShowDialog() != DialogResult.OK)
            {
                return null;
            }
            if (m.TryGetValue(nameof(EasyMetadata.GeoArea), out object? value) && value is string geoAreaInfo && !string.IsNullOrEmpty(geoAreaInfo))
            {
                KeyValuePair<string, GeoCoordinates?>? kvp = await GetGPSAsync(new KeyValuePair<string, GeoCoordinates?>(geoAreaInfo, null));
                if (kvp is KeyValuePair<string, GeoCoordinates?> geoArea)
                {
                    m[nameof(EasyMetadata.GeoArea)] = new GeoArea()
                    {
                        AreaInfo = geoArea.Key,
                        GeoCoords = geoArea.Value
                    };
                }
                else
                {
                    m[nameof(EasyMetadata.GeoArea)] = new GeoArea()
                    {
                        AreaInfo = geoAreaInfo
                    };
                }
            }
            return $"{new EasyMetadata(m)}";
        }

        private async Task<KeyValuePair<string, GeoCoordinates?>?> GetGPSAsync(object? editItem = null)
        {
            editItem ??= new KeyValuePair<string, GeoCoordinates?>(string.Empty, null);
            KeyValuePair<string, GeoCoordinates?> kvp = (KeyValuePair<string, GeoCoordinates?>)editItem;
            string query = kvp.Value != null ? kvp.Value.ToShortString() : string.Empty;
            Map<string, object?> m = new() { { Globals.AreaInfo, kvp.Key }, { Globals.GPSCoordinates, query } };
            using InputDialog id = new(ref m, font: Font);
            if (id.ShowDialog() == DialogResult.Cancel)
            {
                return null;
            }
            string n = (string)m[Globals.AreaInfo]!;
            string q = (string)m[Globals.GPSCoordinates]!;
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
            Form[] forms = [.. Application.OpenForms.Cast<Form>()];
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

        private static bool NameFilteredPathExists(string path, EasyNameFilter nameFilter, string filterString)
        {
            string n = Utils.GetFileNameWithoutExtension(path);
            return string.IsNullOrEmpty(filterString) || (nameFilter switch
            {
                EasyNameFilter.StartsWith => n.StartsWith(filterString),
                EasyNameFilter.EndsWith => n.EndsWith(filterString),
                EasyNameFilter.Contains => n.Contains(filterString),
                EasyNameFilter.DoesNotStartWith => !n.StartsWith(filterString),
                EasyNameFilter.DoesNotEndWith => !n.EndsWith(filterString),
                EasyNameFilter.DoesNotContain => !n.Contains(filterString),
                _ => n.StartsWith(filterString),
            });
        }

        private static bool TypeFilteredPathExists(string path, EasyTypeFilter typeFilter)
        {
            using EasyPath ep = new(path);
            return TypeFilteredPathExists(ep, typeFilter);
        }

        private static bool TypeFilteredPathExists(EasyPath easyPath, EasyTypeFilter typeFilter)
        {
            EasyType et = typeFilter.GetContainingFlags().Select(x => x.GetValue<EasyType>()).Aggregate((x, y) => x |= y);
            return et == EasyType.None || et.HasFlag(easyPath.Type);
        }

        private static string[] GetFilteredFilePaths(string[] paths, EasyTypeFilter typeFilter, EasyNameFilter nameFilter = default, string filterString = EMPTY_STRING)
        {
            return [.. paths.Where(path => NameFilteredPathExists(path, nameFilter, filterString) && TypeFilteredPathExists(path, typeFilter))];
        }

        //private static bool SetAddRemoveProgramsIcon()
        //{
        //    if (Debugger.IsAttached)
        //    {
        //        return false;
        //    }
        //    using RegistryKey? uninstallRegKey = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall");
        //    bool result = false;
        //    if (uninstallRegKey != null)
        //    {
        //        string appName = GetAssemblyName();
        //        string iconPath = Path.Join(Application.StartupPath, @"Resources\Icons\ApplicationIcon.ico");

        //        string[] subKeyNames = uninstallRegKey.GetSubKeyNames();
        //        for (int i = 0; i < subKeyNames.Length; i++)
        //        {
        //            using RegistryKey? regKey = uninstallRegKey.OpenSubKey(subKeyNames[i], true);
        //            if (regKey != null && regKey.GetValue("DisplayName") as string == appName)
        //            {
        //                if (regKey.GetValue("DisplayIcon") as string != iconPath)
        //                {
        //                    regKey.SetValue("DisplayIcon", iconPath);
        //                    result = true;
        //                }
        //                break;
        //            }
        //        }
        //    }
        //    return result;
        //}

        private static Main? GetMainForm() => Application.OpenForms.Count == 0 ? null : Application.OpenForms[0] as Main;

        private static Assembly GetAssembly() => Assembly.GetExecutingAssembly();

        private static string GetAssemblyName() => GetAssembly().GetSimpleName();

        private static string GetAppDataDirectoryPath() => Path.Join(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), GetAssemblyName());

        private static string GetAppDataConfigPath() => Path.Join(GetAppDataDirectoryPath(), $"{GetAssemblyName()}.cfg");
    }
}
