using Microsoft.WindowsAPICodePack.Shell;

using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using System.Text.Encodings.Web;
using System.Diagnostics;
using System.Reflection;
using System.Text.Json;

using WinFormsLib;

using static WinFormsLib.Chars;
using static WinFormsLib.Forms;
using static WinFormsLib.Utils;
using static WinFormsLib.Constants;
using static WinFormsLib.GeoLocator;

namespace EasyFileManager
{
    public static class EasyExtensions
    {
        public static string GetEasyGlobalStringValue(this Enum super)
        {
            Type type = super.GetType();
            return Enum.GetName(type, super) is string s
                   && type.GetField(s) is FieldInfo fi
                   && fi.GetCustomAttributes(typeof(EasyGlobalStringValueAttribute), false) is EasyGlobalStringValueAttribute[] attribs
                   && attribs.Length > 0
                ? attribs[0].Value
                : string.Empty;
        }

        public static T? AsEasyEnumFromGlobal<T>(this string super) where T : Enum
        {
            foreach (T t in Enum.GetValues(typeof(T)))
            {
                if (super == t.GetEasyGlobalStringValue())
                {
                    return t;
                }
            }
            return default;
        }
    }

    public static class EasyComparer
    {
        private class Lib : Map<object, EasyFiles>
        {
            public void Put(object o, EasyFile ef)
            {
                if (!ContainsKey(o))
                {
                    this[o] = new();
                }
                this[o].Add(ef);
            }

            public void Filter(bool dispose = false)
            {
                List<object> l = new();
                foreach (KeyValuePair<object, EasyFiles> kvp in this)
                {
                    if (kvp.Value.Count == 1)
                    {
                        l.Add(kvp.Key);
                    }
                }
                foreach (object o in l)
                {
                    Remove(o, dispose);
                }
            }

            public bool Remove(object o, bool dispose = false)
            {
                if (dispose)
                {
                    this[o].Clear();
                }
                return base.Remove(o);
            }
            
            public void Clear(bool dispose = false)
            {
                if (dispose)
                {
                    foreach (EasyFiles easyFiles in Values)
                    {
                        easyFiles.Clear();
                    }
                }
                base.Clear();
            }

            public void Replace(IDictionary<object, EasyFiles> id, bool dispose = false)
            {
                Clear(dispose);
                Union(id);
            }

            public string[][] GetFilePaths() => Values.Select(x => x.GetPaths()).ToArray();

            public override string ToString() => GetFilePaths().ToJson();
        }

        public static Map<string, string[][]> GetDuplicates(string[] filePaths, EasyCompareParameter parameters)
        {
            Array.Sort(filePaths);
            Map<string, Lib> m = new();
            foreach (string p in filePaths)
            {
                EasyFile ef = new(p);
                string ext = ef.Extension.ToUpper();
                if (!m.ContainsKey(ext))
                {
                    m[ext] = new();
                }
                m[ext].Put(string.Empty, ef);
            }
            Lib lib;
            foreach (KeyValuePair<string, Lib> kvp in m)
            {
                foreach (EasyCompareParameter ecp in parameters.GetContainingFlags())
                {
                    foreach (object o in kvp.Value.GetKeys())
                    {
                        List<EasyFile> l = kvp.Value[o].Cast<EasyFile>().ToList();
                        lib = new();
                        switch (ecp)
                        {
                            case EasyCompareParameter.Name:
                                foreach (EasyFile ef in l) { lib.Put(ef.Name, ef); }
                                break;
                            case EasyCompareParameter.RawName:
                                foreach (EasyFile ef in l) { lib.Put(ef.RawName, ef); }
                                break;
                            case EasyCompareParameter.Size:
                                foreach (EasyFile ef in l) { lib.Put(ef.Size, ef); }
                                break;
                            case EasyCompareParameter.DateModified:
                                foreach (EasyFile ef in l) { if (ef.DateModified is DateTime dt) { lib.Put(dt.ToDateTimeString(), ef); } }
                                break;
                            case EasyCompareParameter.DateCreated:
                                foreach (EasyFile ef in l) { if (ef.DateCreated is DateTime dt) { lib.Put(dt.ToDateTimeString(), ef); } }
                                break;
                            case EasyCompareParameter.DateTakenOrEncoded:
                                foreach (EasyFile ef in l) { if (ef.DateTakenOrEncoded is DateTime dt) { lib.Put(dt.ToDateTimeString(), ef); } }
                                break;
                            case EasyCompareParameter.SmallThumbnail:
                                foreach (EasyFile ef in l)
                                {
                                    if ((EasyType.Image | EasyType.Video).HasFlag(ef.Type) && ef.SmallThumbnail is Image img)
                                    {
                                        lib.Put(img.ToJson(), ef);
                                        img.Dispose();
                                    }
                                }
                                break;
                            case EasyCompareParameter.MediumThumbnail:
                                foreach (EasyFile ef in l)
                                {
                                    if ((EasyType.Image | EasyType.Video).HasFlag(ef.Type) && ef.MediumThumbnail is Image img)
                                    {
                                        lib.Put(img.ToJson(), ef);
                                        img.Dispose();
                                    }
                                }
                                break;
                            case EasyCompareParameter.LargeThumbnail:
                                foreach (EasyFile ef in l)
                                {
                                    if ((EasyType.Image | EasyType.Video).HasFlag(ef.Type) && ef.LargeThumbnail is Image img)
                                    {
                                        lib.Put(img.ToJson(), ef);
                                        img.Dispose();
                                    }
                                }
                                break;
                            case EasyCompareParameter.ExtraLargeThumbnail:
                                foreach (EasyFile ef in l)
                                {
                                    if ((EasyType.Image | EasyType.Video).HasFlag(ef.Type) && ef.ExtraLargeThumbnail is Image img)
                                    {
                                        lib.Put(img.ToJson(), ef);
                                        img.Dispose();
                                    }
                                }
                                break;
                            case EasyCompareParameter.Contents:
                                // TODO
                                break;
                        }
                        lib.Filter(true);
                        l.Clear();
                        kvp.Value.Remove(o);
                        kvp.Value.Union(lib);
                    }
                }
            }
            Map<string, string[][]> duplicates = new();
            foreach (KeyValuePair<string, Lib> kvp in m)
            {
                if (kvp.Value.Any())
                {
                    duplicates[kvp.Key] = kvp.Value.GetFilePaths();
                    kvp.Value.Clear(true);
                }
            }
            return duplicates;
        }
    }

    public class EasyProgress : Progress<int>
    {
        private int _maxValue;
        private int _stepIndex;
        private int _numSteps;

        public string Info { get; private set; } = EMPTY_STRING;
        public int Value { get; private set; } = 0;
        public int MaxValue { get => _maxValue; set { if (value < 0) { throw new ArgumentOutOfRangeException(nameof(MaxValue)); } _maxValue = value; } }
        public int StepIndex { get => _stepIndex; set { if (value < 0) { throw new ArgumentOutOfRangeException(nameof(StepIndex)); } _stepIndex = value; } }
        public int NumSteps { get => _numSteps; set { if (value < 1) { throw new ArgumentOutOfRangeException(nameof(NumSteps)); } _numSteps = value; } }
        public EasyProgress(Action<int> handler) : base(handler) { MaxValue = 100; StepIndex = 0; NumSteps = 1; }
        public static int GetValue(int value, int maxValue, int stepIndex, int numSteps)
        {
            double f = ((double)stepIndex) / numSteps;
            return (int)(f * maxValue) + (value / numSteps);
        }
        public void ReportValue(int value, int stepIndex, int numSteps, bool update = true)
        {
            if (update)
            {
                StepIndex = stepIndex;
                NumSteps = numSteps;
            }
            Value = GetValue(value, MaxValue, stepIndex, numSteps);
            ((IProgress<int>)this).Report(Value);
        }
        public void Report(int value, int stepIndex, int numSteps, bool update = true, string info = EMPTY_STRING)
        {
            Info = info;
            ReportValue(value, stepIndex, numSteps, update);
        }
        public void Report(int value, string info = EMPTY_STRING) => Report(value, StepIndex, NumSteps, false, info);
        public void Reset(bool report = true) { Value = 0; StepIndex = 0; NumSteps = 1; if (report) { Report(0); } }
    }

    [AttributeUsage(AttributeTargets.Field)]
    public class EasyGlobalStringValueAttribute : Attribute
    {
        public string Value { get; protected set; }
        public EasyGlobalStringValueAttribute(string name) => Value = Globals.ResourceManager.GetString(name) is string s ? s : string.Empty;
    }

    [Flags]
    public enum EasyType
    {
        None = 0,
        [EasyGlobalStringValue("Custom")]
        Custom = 1 << 0,
        [EasyGlobalStringValue("Unspecified")]
        Unspecified = 1 << 1,
        [EasyGlobalStringValue("Folder")]
        Folder = 1 << 2,
        [EasyGlobalStringValue("Unknown")]
        Unknown = 1 << 3,
        [EasyGlobalStringValue("Text")]
        Text = 1 << 4,
        [EasyGlobalStringValue("Image")]
        Image = 1 << 5,
        [EasyGlobalStringValue("Audio")]
        Audio = 1 << 6,
        [EasyGlobalStringValue("Video")]
        Video = 1 << 7,
        [EasyGlobalStringValue("Compressed")]
        Compressed = 1 << 8,
        [EasyGlobalStringValue("Document")]
        Document = 1 << 9,
        [EasyGlobalStringValue("System")]
        System = 1 << 10,
        [EasyGlobalStringValue("Application")]
        Application = 1 << 11,
        [EasyGlobalStringValue("GameMedia")]
        GameMedia = 1 << 12,
        [EasyGlobalStringValue("Contacts")]
        Contacts = 1 << 13,
        [EasyGlobalStringValue("SymbolicLink")]
        SymbolicLink = 1 << 14,
        [EasyGlobalStringValue("Invalid")]
        Invalid = 1 << 15,
        File = Text | Image | Audio | Video | Compressed | Document | Application | GameMedia | Contacts
    }

    public enum EasyMetadataSource
    {
        [EasyGlobalStringValue("CustomMetadata")]
        CustomMetadata,
        [EasyGlobalStringValue("EasyMetadata")]
        EasyMetadata,
        [EasyGlobalStringValue("VideoMetadata")]
        VideoMetadata,
        [EasyGlobalStringValue("ShellProperties")]
        ShellProperties,
        [EasyGlobalStringValue("EasyShellOrVideoMetadata")]
        EasyShellOrVideoMetadata
    }

    public enum EasyDateSource
    {
        [EasyGlobalStringValue("DateModified")]
        DateModified,
        [EasyGlobalStringValue("DateCreated")]
        DateCreated,
        [EasyGlobalStringValue("DateTakenOrEncoded")]
        DateTakenOrEncoded,
        [EasyGlobalStringValue("DateFromEasyMetadata")]
        DateFromEasyMetadata,
        [EasyGlobalStringValue("DateFromFolderName")]
        DateFromFolderName,
        [EasyGlobalStringValue("DateEarliest")]
        DateEarliest
    }

    public enum EasySubfolder
    {
        [EasyGlobalStringValue("FileType")]
        FileType,
        [EasyGlobalStringValue("FileTypeText")]
        FileTypeText,
        [EasyGlobalStringValue("FileExtension")]
        FileExtension,
        [EasyGlobalStringValue("Subject")]
        Subject,
        [EasyGlobalStringValue("AreaInfo")]
        AreaInfo,
        [EasyGlobalStringValue("AreaInfoFromEasyMetadata")]
        AreaInfoFromEasyMetadata,
        [EasyGlobalStringValue("YearCreated")]
        YearCreated,
        [EasyGlobalStringValue("YearTakenOrEncoded")]
        YearTakenOrEncoded,
        [EasyGlobalStringValue("YearFromEasyMetadata")]
        YearFromEasyMetadata,
        [EasyGlobalStringValue("YearEarliest")]
        YearEarliest,
        [EasyGlobalStringValue("MonthCreated")]
        MonthCreated,
        [EasyGlobalStringValue("MonthTakenOrEncoded")]
        MonthTakenOrEncoded,
        [EasyGlobalStringValue("MonthFromEasyMetadata")]
        MonthFromEasyMetadata,
        [EasyGlobalStringValue("MonthEarliest")]
        MonthEarliest,
        [EasyGlobalStringValue("DayCreated")]
        DayCreated,
        [EasyGlobalStringValue("DayTakenOrEncoded")]
        DayTakenOrEncoded,
        [EasyGlobalStringValue("DayFromEasyMetadata")]
        DayFromEasyMetadata,
        [EasyGlobalStringValue("DayEarliest")]
        DayEarliest,
        [EasyGlobalStringValue("CameraManufacturer")]
        CameraManufacturer,
        [EasyGlobalStringValue("CameraModel")]
        CameraModel
    }

    [Flags]
    public enum EasyTypeFilter
    {
        None = 0,
        [EasyGlobalStringValue("Text")]
        [Value(EasyType.Text)]
        Text = 1 << 0,
        [EasyGlobalStringValue("Document")]
        [Value(EasyType.Document)]
        Document = 1 << 1,
        [EasyGlobalStringValue("Image")]
        [Value(EasyType.Image)]
        Image = 1 << 2,
        [EasyGlobalStringValue("Video")]
        [Value(EasyType.Video)]
        Video = 1 << 3,
        [EasyGlobalStringValue("Audio")]
        [Value(EasyType.Audio)]
        Audio = 1 << 4,
        [EasyGlobalStringValue("Compressed")]
        [Value(EasyType.Compressed)]
        Compressed = 1 << 5,
        Default = Image | Video,
        All = Text | Document | Image | Video | Audio | Compressed
    }

    public enum EasyNameFilter
    {
        [EasyGlobalStringValue("StartsWith")]
        StartsWith,
        [EasyGlobalStringValue("EndsWith")]
        EndsWith,
        [EasyGlobalStringValue("Contains")]
        Contains,
        [EasyGlobalStringValue("DoesNotStartWith")]
        DoesNotStartWith,
        [EasyGlobalStringValue("DoesNotEndWith")]
        DoesNotEndWith,
        [EasyGlobalStringValue("DoesNotContain")]
        DoesNotContain
    }

    [Flags]
    public enum EasyFileProperty
    {
        None = 0,
        [EasyGlobalStringValue("Title")]
        Title = 1 << 0,
        [EasyGlobalStringValue("Subject")]
        Subject = 1 << 1,
        [EasyGlobalStringValue("Comment")]
        Comment = 1 << 2,
        [EasyGlobalStringValue("Keywords")]
        Keywords = 1 << 3,
        [EasyGlobalStringValue("GPSCoordinates")]
        GPSCoordinates = 1 << 4,
        [EasyGlobalStringValue("AreaInfo")]
        AreaInfo = 1 << 5,
        [EasyGlobalStringValue("EasyMetadata")]
        EasyMetadata = 1 << 6,
        All = Title | Subject | Comment | Keywords | GPSCoordinates | AreaInfo | EasyMetadata
    }

    [Flags]
    public enum EasyCleanUpParameter
    {
        None = 0,
        [EasyGlobalStringValue("Title")]
        [Value(EasyFileProperty.Title)]
        Title = 1 << 0,
        [EasyGlobalStringValue("Subject")]
        [Value(EasyFileProperty.Subject)]
        Subject = 1 << 1,
        [EasyGlobalStringValue("Comment")]
        [Value(EasyFileProperty.Comment)]
        Comment = 1 << 2,
        [EasyGlobalStringValue("Keywords")]
        [Value(EasyFileProperty.Keywords)]
        Keywords = 1 << 3,
        [EasyGlobalStringValue("GPSCoordinates")]
        [Value(EasyFileProperty.GPSCoordinates)]
        GPSCoordinates = 1 << 4,
        [EasyGlobalStringValue("AreaInfo")]
        [Value(EasyFileProperty.AreaInfo)]
        AreaInfo = 1 << 5,
        [EasyGlobalStringValue("EasyMetadata")]
        [Value(EasyFileProperty.EasyMetadata)]
        EasyMetadata = 1 << 6,
        [EasyGlobalStringValue("EmptyFolders")]
        [Value(EasyFileProperty.None)]
        EmptyFolders = 1 << 7,
        Default = EmptyFolders,
        All = EasyFileProperty.All | EmptyFolders
    }

    [Flags]
    public enum EasyCompareParameter
    {
        None = 0,
        [EasyGlobalStringValue("Name")]
        Name = 1 << 0,
        [EasyGlobalStringValue("RawName")]
        RawName = 1 << 1,
        [EasyGlobalStringValue("Size")]
        Size = 1 << 2,
        [EasyGlobalStringValue("DateModified")]
        DateModified = 1 << 3,
        [EasyGlobalStringValue("DateCreated")]
        DateCreated = 1 << 4,
        [EasyGlobalStringValue("DateTakenOrEncoded")]
        DateTakenOrEncoded = 1 << 5,
        [EasyGlobalStringValue("SmallThumbnail")]
        SmallThumbnail = 1 << 6,
        [EasyGlobalStringValue("MediumThumbnail")]
        MediumThumbnail = 1 << 7,
        [EasyGlobalStringValue("LargeThumbnail")]
        LargeThumbnail = 1 << 8,
        [EasyGlobalStringValue("ExtraLargeThumbnail")]
        ExtraLargeThumbnail = 1 << 9,
        [EasyGlobalStringValue("Contents")]
        Contents = 1 << 10,
        Default = RawName | Size | DateTakenOrEncoded | ExtraLargeThumbnail,
        All = Name | RawName | Size | DateModified | DateCreated | DateTakenOrEncoded | SmallThumbnail | MediumThumbnail | LargeThumbnail | ExtraLargeThumbnail | Contents
    }

    public enum EasyOrientation
    {
        [Value(8)]
        Rotate270,
        [Value(3)]
        Rotate180,
        [Value(6)]
        Rotate90
    }

    public class EasyLogger
    {
        private const string FILENAME = "log.txt";

        private readonly FileInfo _fileInfo;

        public EasyLogger(string directoryPath = EMPTY_STRING)
        {
            if (string.IsNullOrEmpty(directoryPath))
            {
                directoryPath = Path.Join(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), Assembly.GetExecutingAssembly().GetSimpleName());
            }
            _fileInfo = new(Path.Join(directoryPath, FILENAME));
        }
        public void Write(string info) => _fileInfo.WriteText($"{DateTime.Now.ToDateTimeString()}: {info}", true);
        public void Reset() => _fileInfo.WriteText(string.Empty);
        public void Show() => Process.Start("notepad.exe", _fileInfo.FullName);
    }

    public class EasyOptions
    {
        private static JsonSerializerOptions JsonSerializerOptions => new()
        {
            Converters = { EasyFileManager.EasyMetadata.JsonConverter },
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
            WriteIndented = true,
            IgnoreReadOnlyFields = true,
            IgnoreReadOnlyProperties = true
        };

        public string Title { get; set; } = EMPTY_STRING;
        public string Subject { get; set; } = EMPTY_STRING;
        public string Comment { get; set; } = EMPTY_STRING;
        public string Prefix { get; set; } = EMPTY_STRING;
        public string Replace { get; set; } = EMPTY_STRING;
        public string With { get; set; } = EMPTY_STRING;
        public string Suffix { get; set; } = EMPTY_STRING;
        public string FilterString { get; set; } = EMPTY_STRING;
        public string DateFormat { get; set; } = DateTimeExtensions.FORMAT_DATE_SECOND;
        public string BackupFolderPath { get; set; } = STARTUP_DIRECTORY_DEFAULT;
        public string TopFolderPath { get; set; } = STARTUP_DIRECTORY_DEFAULT;
        public string DuplicatesFolderPath { get; set; } = STARTUP_DIRECTORY_DEFAULT;

        public bool CustomizeEnabled { get; set; }
        public bool RenameEnabled { get; set; }
        public bool FinalizeEnabled { get; set; }
        public bool TitleEnabled { get; set; }
        public bool SubjectEnabled { get; set; }
        public bool CommentEnabled { get; set; }
        public bool DateFormatEnabled { get; set; }
        public bool GeolocationEnabled { get; set; }
        public bool EasyMetadataEnabled { get; set; }
        public bool WriteGPSCoords { get; set; }
        public bool WriteGPSAreaInfo { get; set; }
        public bool WriteGPSEasyMetadata { get; set; }
        public bool WriteDateCreated { get; set; }
        public bool WriteDateCamera { get; set; }
        public bool WriteDateModified { get; set; }
        public bool WriteDateEasyMetadata { get; set; }
        public bool BackupFolderEnabled { get; set; }
        public bool TopFolderEnabled { get; set; }
        public bool SubfoldersEnabled { get; set; }
        public bool FilterEnabled { get; set; }
        public bool UseEasyMetadataWithVideo { get; set; }
        public bool PreserveDateModified { get; set; }
        public bool PreserveDateCreated { get; set; }
        public bool LogApplicationEvents { get; set; }
        public bool ShutdownUponCompletion { get; set; }
        public bool CleanUpEnabled { get; set; }
        public bool ShowDuplicatesCompareDialog { get; set; }

        public int KeywordsIndex { get; set; } = -1;
        public int GPSSourceIndex { get; set; } = 4;
        public int GPSCustomIndex { get; set; } = -1;
        public int GPSAreaIndex { get; set; } = 6;
        public int EasyMetadataIndex { get; set; } = -1;
        public int DateSourceIndex { get; set; } = 4;
        public int DateCustomIndex { get; set; } = -1;

        public CheckState CopyMoveState { get; set; } = CheckState.Unchecked;
        public CheckState KeywordsState { get; set; } = CheckState.Unchecked;
        public CheckState DateState { get; set; } = CheckState.Unchecked;
        public CheckState ReplaceWithState { get; set; } = CheckState.Unchecked;
        public CheckState DuplicatesState { get; set; } = CheckState.Unchecked;

        public EasyTypeFilter TypeFilter { get; set; } = EasyTypeFilter.Default;
        public EasyNameFilter NameFilter { get; set; } = EasyNameFilter.Contains;
        public EasyCleanUpParameter CleanUpParameters { get; set; } = EasyCleanUpParameter.Default;
        public EasyCompareParameter DuplicatesCompareParameters { get; set; } = EasyCompareParameter.Default;

        public List<EasyList<string>> Keywords { get; set; } = new();
        public Dictionary<string, GeoCoordinates?> GeoAreas { get; set; } = new();
        public EasyList<string> EasyMetadata { get; set; } = new();
        public EasyList<string> Dates { get; set; } = new();
        public EasyList<EasySubfolder> Subfolders { get; set; } = new();

        public EasyOptions() { }

        public EasyOptions(string json)
        {
            if (!string.IsNullOrEmpty(json) && JsonSerializer.Deserialize<EasyOptions>(json, JsonSerializerOptions) is EasyOptions eo)
            {
                Title = eo.Title;
                Subject = eo.Subject;
                Comment = eo.Comment;
                Prefix = eo.Prefix;
                Replace = eo.Replace;
                With = eo.With;
                Suffix = eo.Suffix;
                DateFormat = eo.DateFormat;
                FilterString = eo.FilterString;
                BackupFolderPath = GetValidDirectoryPath(eo.BackupFolderPath);
                TopFolderPath = GetValidDirectoryPath(eo.TopFolderPath);
                DuplicatesFolderPath = GetValidDirectoryPath(eo.DuplicatesFolderPath);

                CustomizeEnabled = eo.CustomizeEnabled;
                RenameEnabled = eo.RenameEnabled;
                FinalizeEnabled = eo.FinalizeEnabled;
                TitleEnabled = eo.TitleEnabled;
                SubjectEnabled = eo.SubjectEnabled;
                CommentEnabled = eo.CommentEnabled;
                DateFormatEnabled = eo.DateFormatEnabled;
                GeolocationEnabled = eo.GeolocationEnabled;
                EasyMetadataEnabled = eo.EasyMetadataEnabled;
                WriteGPSCoords = eo.WriteGPSCoords;
                WriteGPSAreaInfo = eo.WriteGPSAreaInfo;
                WriteGPSEasyMetadata = eo.WriteGPSEasyMetadata;
                WriteDateCreated = eo.WriteDateCreated;
                WriteDateCamera = eo.WriteDateCamera;
                WriteDateModified = eo.WriteDateModified;
                BackupFolderEnabled = eo.BackupFolderEnabled;
                TopFolderEnabled = eo.TopFolderEnabled;
                SubfoldersEnabled = eo.SubfoldersEnabled;
                FilterEnabled = eo.FilterEnabled;
                WriteDateEasyMetadata = eo.WriteDateEasyMetadata;
                UseEasyMetadataWithVideo = eo.UseEasyMetadataWithVideo;
                PreserveDateModified = eo.PreserveDateModified;
                PreserveDateCreated = eo.PreserveDateCreated;
                LogApplicationEvents = eo.LogApplicationEvents;
                ShutdownUponCompletion = eo.ShutdownUponCompletion;
                CleanUpEnabled = eo.CleanUpEnabled;
                ShowDuplicatesCompareDialog = eo.ShowDuplicatesCompareDialog;

                KeywordsIndex = eo.KeywordsIndex;
                GPSSourceIndex = eo.GPSSourceIndex;
                GPSCustomIndex = eo.GPSCustomIndex;
                GPSAreaIndex = eo.GPSAreaIndex;
                DateSourceIndex = eo.DateSourceIndex;
                DateCustomIndex = eo.DateCustomIndex;
                EasyMetadataIndex = eo.EasyMetadataIndex;

                CopyMoveState = eo.CopyMoveState;
                KeywordsState = eo.KeywordsState;
                DateState = eo.DateState;
                ReplaceWithState = eo.ReplaceWithState;
                DuplicatesState = eo.DuplicatesState;

                TypeFilter = eo.TypeFilter;
                NameFilter = eo.NameFilter;
                CleanUpParameters = eo.CleanUpParameters;
                DuplicatesCompareParameters = eo.DuplicatesCompareParameters;

                Keywords = eo.Keywords;
                GeoAreas = eo.GeoAreas;
                EasyMetadata = eo.EasyMetadata;
                Dates = eo.Dates;
                Subfolders = eo.Subfolders;
            }
        }

        public bool HasCustomizing() => CustomizeEnabled && (TitleEnabled || SubjectEnabled || CommentEnabled || KeywordsState != CheckState.Unchecked || GeolocationEnabled || DateState != CheckState.Unchecked || EasyMetadataEnabled);

        public bool HasRenaming() => RenameEnabled && (!string.IsNullOrEmpty(Prefix + Suffix) || (ReplaceWithState != CheckState.Unchecked && !string.IsNullOrEmpty(Replace)) || DateFormatEnabled);

        public bool HasCopyingMoving(bool apply = true) => CopyMoveState != CheckState.Unchecked && ((apply && BackupFolderEnabled) || TopFolderEnabled || SubfoldersEnabled);

        public bool HasFinalizing()
        {
            return FinalizeEnabled && ((CleanUpEnabled && CleanUpParameters != EasyCleanUpParameter.None) || (DuplicatesState != CheckState.Unchecked && DuplicatesCompareParameters != EasyCompareParameter.None));
        }

        public override string ToString() => JsonSerializer.Serialize(this, JsonSerializerOptions);
    }

    public class EasyMetadata
    {
        public static class Options
        {
            public static bool? Delete { get; set; }
            public static bool? Overwrite { get; set; }
            public static bool? VideoMetadata { get; set; }
            public static void Reset()
            {
                Delete = null;
                Overwrite = null;
                VideoMetadata = null;
            }
            static Options() => Reset();
        }

        private class EasyMetadataJsonConverter : JsonConverter<EasyMetadata>
        {
            public override EasyMetadata Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
            {
                return reader.GetString() is string json ? new(json) : new();
            }
            public override void Write(Utf8JsonWriter writer, EasyMetadata emd, JsonSerializerOptions options)
            {
                writer.WriteStringValue(emd.ToString());
            }
        }

        public static JsonConverter JsonConverter => new EasyMetadataJsonConverter();

        private static JsonSerializerOptions JsonSerializerOptions => new()
        {
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingDefault,
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
            IgnoreReadOnlyFields = true,
            IgnoreReadOnlyProperties = true
        };

        public string? Date { get; set; } = null;
        public GeoArea? GeoArea { get; set; } = null;
        public EasyList<string>? Tags { get; set; } = null;
        public Map<string, object?>? CustomDict { get; set; } = null;

        public bool IsValid => Date != null || GeoArea != null || Tags != null || CustomDict != null;

        public EasyMetadata() { }

        public EasyMetadata(string json)
        {
            if (!string.IsNullOrEmpty(json) && JsonSerializer.Deserialize<EasyMetadata>(json, JsonSerializerOptions) is EasyMetadata emd)
            {
                Date = emd.Date;
                GeoArea = emd.GeoArea;
                Tags = emd.Tags;
                CustomDict = emd.CustomDict;
            }
        }

        public EasyMetadata(Map<string, object?> map) : this(map.ToString()) { }

        public Map<string, object?> ToMap() => new(ToJson());

        public string ToJson() => JsonSerializer.Serialize(this, new JsonSerializerOptions(JsonSerializerOptions) { DefaultIgnoreCondition = JsonIgnoreCondition.Never });

        public override string ToString() => JsonSerializer.Serialize(this, JsonSerializerOptions);
    }

    public class EasyList<T> : List<T>
    {
        public EasyList(){}

        public EasyList(string json)
        {
            if (string.IsNullOrEmpty(json))
            {
                return;
            }
            if (!json.IsArray())
            {
                json = $"{LEFT_SQUARE_BRACKET}{json}{RIGHT_SQUARE_BRACKET}";
            }
            if (JsonSerializer.Deserialize<IEnumerable<T>>(json) is IEnumerable<T> ie)
            {
                AddRange(ie);
            }
        }

        public EasyList(IEnumerable<T> ie) => AddRange(ie);

        public void Replace(IEnumerable<T> ie)
        {
            Clear();
            AddRange(ie);
        }

        public override string ToString() => string.Join($"{COMMA}{SPACE}", this);
    }

    public class EasyPaths<T> : EasyList<EasyPath> where T : EasyPath
    {
        public EasyPaths() { }

        public EasyPaths(IEnumerable<string> paths)
        {
            foreach (string path in paths)
            {
                Add(path);
            }
        }

        public EasyPaths(IEnumerable<EasyPath> easyPaths) => AddRange(easyPaths);

        public bool Contains(string path)
        {
            foreach (EasyPath ep in this)
            {
                if (ep.Path == path)
                {
                    return true;
                }
            }
            return false;
        }

        public T? Get(string path)
        {
            foreach (EasyPath ep in this)
            {
                if (ep.Path == path)
                {
                    return (T)ep;
                }
            }
            return null;
        }

        public virtual EasyPath? Add(string path, bool readShell = true)
        {
            if (!Contains(path))
            {
                EasyPath ep = new(path, readShell);
                Add(ep);
                return ep;
            }
            return null;
        }

        public void AddRange(IEnumerable<string> paths, bool readShell = true)
        {
            foreach (string path in paths)
            {
                Add(path, readShell);
            }
        }

        public bool Remove(string path)
        {
            foreach (EasyPath ep in ToArray())
            {
                if (ep.Path == path)
                {
                    Remove(ep);
                    return true;
                }
            }
            return false;
        }

        public new void Remove(EasyPath easyPath)
        {
            easyPath.Dispose();
            base.Remove(easyPath);
        }

        public new int RemoveAll(Predicate<EasyPath> predicate)
        {
            EasyPaths<T> el = new(this);
            int result = base.RemoveAll(predicate);
            el.ExceptWith(this);
            el.Clear();
            return result;
        }

        public new void RemoveAt(int index)
        {
            this[index].Dispose();
            base.RemoveAt(index);
        }

        public new void RemoveRange(int index, int count)
        {
            EasyPaths<T> el = new(this);
            base.RemoveRange(index, count);
            el.ExceptWith(this);
            el.Clear();
        }

        public new void Clear()
        {
            foreach (EasyPath ep in this)
            {
                ep.Dispose();
            }
            base.Clear();
        }

        public void Replace(IEnumerable<string> paths, bool readShell = true)
        {
            Clear();
            AddRange(paths, readShell);
        }

        public string[] GetPaths() => ToArray().Select(x => x.Path).ToArray();

        public string[] GetRawPaths() => ToArray().Select(x => x.RawPath).ToHashSet().ToArray();

        public override string ToString() => ToArray().ToJson(EasyPath.JsonConverter);
    }

    public class EasyFolders : EasyPaths<EasyFolder>
    {
        public new EasyFolder this[int index] { get => (EasyFolder)base[index]; set => base[index] = value; }

        public EasyFolders() { }

        public EasyFolders(IEnumerable<string> paths) => AddRange(paths);

        public EasyFolders(IEnumerable<EasyFolder> easyFolders) => AddRange(easyFolders);

        public new IEnumerator<EasyFolder> GetEnumerator()
        {
            using IEnumerator<EasyPath> ie = base.GetEnumerator();
            while (ie.MoveNext())
            {
                yield return (EasyFolder)ie.Current;
            }
        }

        public override EasyFolder? Add(string path, bool readShell = true)
        {
            if (!Contains(path))
            {
                EasyFolder ed;
                try
                {
                    ed = new(path, readShell);
                    Add(ed);
                    return ed;
                }
                catch (ArgumentException) { }
            }
            return null;
        }
    }

    public class EasyFiles : EasyPaths<EasyFile>
    {
        public new EasyFile this[int index] { get => (EasyFile)base[index]; set => base[index] = value; }

        public EasyFiles() { }

        public EasyFiles(IEnumerable<string> paths) => AddRange(paths);

        public EasyFiles(IEnumerable<EasyFile> easyFiles) => AddRange(easyFiles);

        public new IEnumerator<EasyFile> GetEnumerator()
        {
            using IEnumerator<EasyPath> ie = base.GetEnumerator();
            while (ie.MoveNext())
            {
                yield return (EasyFile)ie.Current;
            }
        }

        public override EasyFile? Add(string path, bool readShell = true)
        {
            if (!Contains(path))
            {
                EasyFile ef;
                try
                {
                    ef = new(path, readShell);
                    Add(ef);
                    return ef;
                }
                catch (ArgumentException) { }
            }
            return null;
        }
    }

    public class EasyPath : IDisposable, IEquatable<EasyPath>, IComparable<EasyPath>
    {
        private class EasyPathJsonConverter : JsonConverter<EasyPath>
        {
            public override EasyPath Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
            {
                if (reader.GetString() is string s)
                {
                    return new(s);
                }
                return new();
            }
            public override void Write(Utf8JsonWriter writer, EasyPath easyPath, JsonSerializerOptions options)
            {
                writer.WriteStringValue(easyPath.ToString());
            }
        }

        private bool _isDisposed;
        private FileAttributes _fileAttributes;

        private long? _size;

        protected string _path = EMPTY_STRING;
        protected DateTime? _dateCreated;
        protected DateTime? _dateModified;
        protected ShellObject? _shellObject;

        public static JsonConverter JsonConverter => new EasyPathJsonConverter();

        public long Size => UpdateSize() && _size is long l ? l : 0;
        public string Path => _path;
        public DateTime? DateCreated => _dateCreated;
        public DateTime? DateModified => _dateModified;

        public bool Exists => IsFolder ? Directory.Exists(_path) : File.Exists(_path);
        public bool IsInvalid => Type == EasyType.Invalid;
        public bool IsFolder => _fileAttributes.HasFlag(FileAttributes.Directory);
        public bool IsFile => !IsFolder;
        public bool IsSystem => _fileAttributes.HasFlag(FileAttributes.System);
        public bool IsHidden => _fileAttributes.HasFlag(FileAttributes.Hidden) || _path.StartsWith(PERIOD);
        public bool IsReadOnly => _fileAttributes.HasFlag(FileAttributes.ReadOnly);
        public bool IsSymbolicLink => _fileAttributes.HasFlag(FileAttributes.ReparsePoint);
        public EasyType Type
        {
            get
            {
                EasyType et = EasyType.None;
                if (_shellObject is ShellObject so)
                {
                    if (IsSymbolicLink)
                    {
                        et |= EasyType.SymbolicLink;
                    }
                    if (IsSystem)
                    {
                        et |= EasyType.System;
                    }
                    if (IsFolder)
                    {
                        et |= EasyType.Folder;
                    }
                    else if (so.Properties.System.PerceivedType.Value is int i)
                    {
                        et |= i switch
                        {
                            -3 => EasyType.Custom,
                            -2 => EasyType.Unspecified,
                            -1 => EasyType.Folder,
                            0 => EasyType.Unknown,
                            1 => EasyType.Text,
                            2 => EasyType.Image,
                            3 => EasyType.Audio,
                            4 => EasyType.Video,
                            5 => EasyType.Compressed,
                            6 => EasyType.Document,
                            7 => EasyType.System,
                            8 => EasyType.Application,
                            9 => EasyType.GameMedia,
                            10 => EasyType.Contacts,
                            _ => EasyType.Invalid,
                        };
                    }
                }
                else
                {
                    et |= EasyType.Invalid;
                }
                return et;
            }
        }
        public string TypeText
        {
            get
            {
                if (_shellObject is ShellObject so)
                {
                    string t = $"{so.Properties.System.ItemTypeText.Value}";
                    if (so.Properties.System.ItemType.Value is string it)
                    {
                        t += $" ({it})";
                    }
                    return t;
                }
                return EMPTY_STRING;
            }
        }
        public string Name => GetFileName(_path);
        public string NameWithoutExtension => GetFileNameWithoutExtension(_path);
        public string PathWithoutExtension => GetPathWithoutExtension(_path);
        public string Extension => GetExtension(_path);
        public string DirectoryPath => GetDirectoryPath(_path);
        public string DirectoryName => GetFileName(DirectoryPath);
        public string ParentPath => GetParentPath(_path);
        public string RawName => GetRawName(_path);
        public string RawPath => GetRawPath(_path);
        public FileInfo FileInfo => new(_path);
        public Image? SmallThumbnail => GetThumbnailImage(_shellObject?.Thumbnail.SmallBitmap);
        public Image? MediumThumbnail => GetThumbnailImage(_shellObject?.Thumbnail.MediumBitmap);
        public Image? LargeThumbnail => GetThumbnailImage(_shellObject?.Thumbnail.LargeBitmap);
        public Image? ExtraLargeThumbnail => GetThumbnailImage(_shellObject?.Thumbnail.ExtraLargeBitmap);

        private bool MoveOrCopy(bool copy, string path, bool preserveDateCreated = true, bool preserveDateModified = true)
        {
            bool result = false;
            if (!IsInvalid && (copy || !IsReadOnly))
            {
                DateTime? dtc = preserveDateCreated ? DateCreated : null;
                DateTime? dtm = preserveDateModified ? DateModified : null;
                result = copy ? IsFolder ? CopyDirectory(_path, path) : CopyFile(_path, path) : IsFolder ? MoveDirectory(_path, path) : MoveFile(_path, path);
                if (result)
                {
                    if (preserveDateCreated || preserveDateModified) { Initialize(path); }
                    if (preserveDateCreated) { _dateCreated = dtc; WriteDateCreated(); }
                    if (preserveDateModified) { _dateModified = dtm; WriteDateModified(); }
                }
            }
            return result;
        }
        private bool ReadShellObject()
        {
            try
            {
                _shellObject = ShellObject.FromParsingName(_path);
                return true;
            }
            catch { return false; }
        }
        private bool ReadFileAttributes()
        {
            try
            {
                if (_shellObject is ShellObject so)
                {
                    try
                    {
                        uint i = (uint)so.Properties.DefaultPropertyCollection["System.FileAttributes"].ValueAsObject;
                        _fileAttributes = (FileAttributes)Enum.ToObject(typeof(FileAttributes), i);
                    }
                    catch (IndexOutOfRangeException) { }
                }
                _fileAttributes = File.GetAttributes(_path);
                return true;
            }
            catch { return false; }
        }
        private bool ReadDateModified()
        {
            if (_shellObject is ShellObject so)
            {
                _dateModified = so.Properties.System.DateModified.Value is DateTime dt ? dt : File.GetLastWriteTime(_path);
                return true;
            }
            return false;
        }
        private bool ReadDateCreated()
        {
            if (_shellObject is ShellObject so)
            {
                _dateCreated = so.Properties.System.DateCreated.Value is DateTime dt ? dt : File.GetCreationTime(_path);
                return true;
            }
            return false;
        }

        protected bool WriteDateModified()
        {
            if (DateModified is DateTime dt)
            {
                File.SetLastWriteTime(_path, dt);
                return true;
            }
            return false;
        }
        protected bool WriteDateCreated()
        {
            if (DateCreated is DateTime dt)
            {
                File.SetCreationTime(_path, dt);
                return true;
            }
            return false;
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                _shellObject?.Dispose();
                _shellObject = null;
            }
        }

        public EasyPath(string path = EMPTY_STRING, bool readShell = true) => Initialize(path, readShell);

        public virtual void Initialize(string path = EMPTY_STRING, bool readShell = true)
        {
            if (string.IsNullOrEmpty(path))
            {
                path = _path;
            }
            else
            {
                _path = path;
            }
            _shellObject?.Dispose();
            if (string.IsNullOrEmpty(path))
            {
                Debug.WriteLine($"EasyPath is empty.");
                return;
            }
            if (path == GetParentPath(path))
            {
                Debug.WriteLine($"EasyPath is root.");
                return;
            }
            try
            {
                ReadFileAttributes();
                if (readShell)
                {
                    ReadShellObject();
                    ReadDateModified();
                    ReadDateCreated();
                }
                if (IsSymbolicLink)
                {
                    if (IsFolder)
                    {
                        if (GetTargetDirectory(path) is string td)
                        {
                            Initialize(td);
                        }
                    }
                    else if (GetTargetFile(path) is string tf)
                    {
                        Initialize(tf);
                    }
                }
            }
            catch (Exception e)
            {
                Debug.WriteLine($"EasyPath '{path}' invalid: {e.Message}");
            }
        }

        public void Dispose()
        {
            if (!_isDisposed)
            {
                Dispose(true);
                GC.SuppressFinalize(this);
                _isDisposed = true;
            }
        }
        public bool UpdateSize(long? size = null)
        {
            if (size is long l)
            {
                _size = l;
            }
            else if (_size == null && !IsSymbolicLink)
            {
                if (IsFolder)
                {
                    CancellationTokenSource cts = new();
                    Task t = RunCancellableTask(() =>
                    {
                        try
                        {
                            long l = 0;
                            foreach (FileInfo fi in new DirectoryInfo(_path).EnumerateFiles("*", SearchOption.AllDirectories))
                            {
                                cts.Token.ThrowIfCancellationRequested();
                                l += fi.Length;
                            }
                            UpdateSize(l);
                        }
                        catch { }
                    }, cts);
                    if (!t.Wait(1000))
                    {
                        Debug.WriteLine($"Cancelling updating size");
                        cts.Cancel();
                        return false;
                    }
                }
                else if (IsFile && _shellObject is ShellObject so && so.Properties.System.FileAllocationSize.Value is ulong ul)
                {
                    _size = Convert.ToInt64(ul);
                }
                else if (!IsInvalid)
                {
                    try { _size = FileInfo.Length; } catch { }
                }
            }
            return true;
        }
        public bool Move(string path, bool preserveDateCreated = true, bool preserveDateModified = true)
        {
            return MoveOrCopy(false, path, preserveDateCreated, preserveDateModified);
        }
        public bool Copy(string path, bool preserveDateCreated = true, bool preserveDateModified = true)
        {
            return MoveOrCopy(true, path, preserveDateCreated, preserveDateModified);
        }
        public bool Delete()
        {
            if (IsInvalid || IsReadOnly)
            {
                return false;
            }
            bool result = IsFolder ? DeleteDirectory(_path) : DeleteFile(_path);
            if (result) { Dispose(); }
            return result;
        }
        public int CompareTo(EasyPath? other) => other != null ? _path.CompareTo(other._path) : throw new ArgumentNullException(nameof(other));
        public bool Equals(EasyPath? other) => _shellObject is ShellObject so && other is EasyPath ep && ep._shellObject is ShellObject so2 && so.Equals(so2);

        public override bool Equals(object? obj) => Equals(obj as EasyPath);
        public override int GetHashCode() => HashCode.Combine(_path);
        public override string ToString() => _path;

        public static bool operator ==(EasyPath? left, EasyPath? right) => EqualityComparer<EasyPath>.Default.Equals(left, right);
        public static bool operator !=(EasyPath? left, EasyPath? right) => !(left == right);
        public static bool operator <(EasyPath left, EasyPath right) => left.CompareTo(right) < 0;
        public static bool operator <=(EasyPath left, EasyPath right) => left.CompareTo(right) <= 0;
        public static bool operator >(EasyPath left, EasyPath right) => left.CompareTo(right) > 0;
        public static bool operator >=(EasyPath left, EasyPath right) => left.CompareTo(right) >= 0;
    }

    public class EasyFolder : EasyPath
    {
        public string ContentText
        {
            get
            {
                int filesCount = GetFilePaths().Length;
                int foldersCount = GetDirectoryPaths().Length;
                string filesText = filesCount == 1 ? Globals.File : Globals.Files;
                string foldersText = foldersCount == 1 ? Globals.Folder : Globals.Folders;
                return $"{filesCount} {filesText}, {foldersCount} {foldersText}";
            }
        }

        public EasyFolder() : base() { }

        public EasyFolder(string path, bool readShell = true) : base(path, readShell)
        {
            if (IsFolder) { return; }
            Dispose(); throw new ArgumentException("Path is not a folder.");
        }

        public string[] GetFilePaths(bool recursive = false) => Utils.GetFilePaths(_path, recursive);

        public string[] GetDirectoryPaths(bool recursive = false) => Utils.GetDirectoryPaths(_path, recursive);

        public string[] GetAllPaths(bool recursive = false) => Utils.GetAllPaths(_path, recursive);

        public string[] GetRawPaths(bool recursive = false) => Utils.GetRawPaths(GetAllPaths(recursive));
    }

    public class EasyFile : EasyPath
    {
        private string _formattedPath = EMPTY_STRING;
        private string _areaInfo = EMPTY_STRING;
        private DateTime? _dateTaken;
        private DateTime? _dateEncoded;
        private EasyMetadata? _easyMetadata;
        private GeoCoordinates? _geoCoordinates;
        private GeoObject? _geoObject;

        public string FormattedPath => _formattedPath;
        public string AreaInfo => _areaInfo;
        public DateTime? DateTaken => _dateTaken;
        public DateTime? DateEncoded => _dateEncoded;
        public EasyMetadata? EasyMetadata => _easyMetadata;
        public GeoCoordinates? GeoCoordinates => _geoCoordinates;
        public GeoObject? GeoObject => _geoObject;

        public DateTime? DateTakenOrEncoded => Type == EasyType.Image ? DateTaken : Type == EasyType.Video ? DateEncoded : null;
        public DateTime? DateFromEasyMetadata
        {
            get => _easyMetadata is EasyMetadata emd && emd.Date is string s ? s.AsDateTime() : null;
            set
            {
                if (_easyMetadata == null)
                {
                    if (value == null)
                    {
                        return;
                    }
                    _easyMetadata = new();
                }
                _easyMetadata.Date = value?.ToDateTimeString();
            }
        }
        public DateTime? DateEarliest
        {
            get
            {
                List<DateTime> l = new();
                if (DateFromEasyMetadata is DateTime dfem)
                {
                    l.Add(dfem);
                }
                if (DateCreated is DateTime dc)
                {
                    l.Add(dc);
                }
                if (DateTaken is DateTime dt)
                {
                    l.Add(dt);
                }
                else if (DateEncoded is DateTime de)
                {
                    l.Add(de);
                }
                return l.Any() ? l.Min() : null;
            }
        }

        public string FormattedName => GetFileName(FormattedPath);
        public string FormattedParentPath => GetParentPath(FormattedPath);
        public string FormattedNameWithoutExtension => GetFileNameWithoutExtension(FormattedPath);

        public string Title
        {
            get => _shellObject is ShellObject so ? so.Properties.System.Title.Value : EMPTY_STRING;
            set { if (_shellObject is ShellObject so) { so.Properties.System.Title.Value = value; } }
        }
        public string Subject
        {
            get => _shellObject is ShellObject so ? so.Properties.System.Subject.Value : EMPTY_STRING;
            set { if (_shellObject is ShellObject so) { so.Properties.System.Subject.Value = value; } }
        }
        public string Comment
        {
            get => _shellObject is ShellObject so ? so.Properties.System.Comment.Value : EMPTY_STRING;
            set { if (_shellObject is ShellObject so) { so.Properties.System.Comment.Value = value; } }
        }
        public string[] Keywords
        {
            get => _shellObject is ShellObject so && so.Properties.System.Keywords.Value is string[] kw ? kw : Array.Empty<string>();
            set { if (_shellObject is ShellObject so) { so.Properties.System.Keywords.Value = value; } }
        }
        public string? AreaInfoFromEasyMetadata
        {
            get => _easyMetadata is EasyMetadata emd && emd.GeoArea is GeoArea ga && !string.IsNullOrEmpty(ga.AreaInfo) ? ga.AreaInfo : null;
            set
            {
                if (_easyMetadata == null) { if (value == null) { return; } _easyMetadata = new(); }
                if (_easyMetadata.GeoArea == null) { if (value == null) { return; } _easyMetadata.GeoArea = new(); }
                _easyMetadata.GeoArea.AreaInfo = value;
            }
        }
        public string? CameraManufacturer
        {
            get => _shellObject is ShellObject so && Type == EasyType.Image ? so.Properties.System.Photo.CameraManufacturer.Value : null;
            set { if (_shellObject is ShellObject so && value is string s && Type == EasyType.Image) { so.Properties.System.Photo.CameraManufacturer.Value = s; } }
        }
        public string? CameraModel
        {
            get => _shellObject is ShellObject so && Type == EasyType.Image ? so.Properties.System.Photo.CameraModel.Value : null;
            set { if (_shellObject is ShellObject so && value is string s && Type == EasyType.Image) { so.Properties.System.Photo.CameraModel.Value = s; } }
        }

        private bool ReadDateTaken()
        {
            if (_shellObject is ShellObject so && Type == EasyType.Image) { _dateTaken = so.Properties.System.Photo.DateTaken.Value; return true; }
            return false;
        }
        private bool ReadDateEncoded()
        {
            if (_shellObject is ShellObject so && Type == EasyType.Video) { _dateEncoded = so.Properties.System.Media.DateEncoded.Value; return true; }
            return false;
        }
        private bool ReadAreaInfo()
        {
            if (_shellObject is ShellObject so) { _areaInfo = so.Properties.System.GPS.AreaInformation.Value; return true; }
            return false;
        }
        private bool ReadGeoCoordinates()
        {
            if (_shellObject is ShellObject so)
            {
                double[] lat = so.Properties.System.GPS.Latitude.Value;
                double[] lon = so.Properties.System.GPS.Longitude.Value;
                string latRef = so.Properties.System.GPS.LatitudeRef.Value;
                string lonRef = so.Properties.System.GPS.LongitudeRef.Value;
                if (lat != null && lon != null && !string.IsNullOrEmpty(latRef) & !string.IsNullOrEmpty(lonRef))
                {
                    _geoCoordinates = new(lat, lon, latRef, lonRef);
                    return true;
                }
            }
            return false;
        }
        private bool ReadEasyMetadata()
        {
            string delimiter = $"{nameof(EasyMetadata)}{EQUALS_SIGN}";
            foreach (string kw in Keywords)
            {
                if (kw.SplitLast(delimiter) is string[] sa)
                {
                    _easyMetadata = new(sa.Last());
                    return true;
                }
            }
            return false;
        }
        private bool WriteDateTaken()
        {
            if (_shellObject is ShellObject so && Type == EasyType.Image) { so.Properties.System.Photo.DateTaken.Value = DateTaken; return true; }
            return false;
        }
        private bool WriteDateEncoded()
        {
            if (_shellObject is ShellObject so && Type == EasyType.Video) { so.Properties.System.Media.DateEncoded.Value = DateEncoded; return true; }
            return false;
        }
        private bool WriteAreaInfo(string? areaInfo = null)
        {
            areaInfo ??= AreaInfo;
            if (_shellObject is ShellObject so && (EasyType.Image | EasyType.Video).HasFlag(Type))
            {
                so.Properties.System.GPS.AreaInformation.Value = areaInfo;
                return true;
            }
            return false;
        }
        private bool WriteGeoCoordinates(double[] latCoords, double[] lonCoords, string latRef, string lonRef)
        {
            if (_shellObject is ShellObject so && (EasyType.Image | EasyType.Video).HasFlag(Type))
            {
                uint denominator = 10000;
                so.Properties.System.GPS.LatitudeRef.Value = latRef;
                so.Properties.System.GPS.LongitudeRef.Value = lonRef;
                so.Properties.System.GPS.LatitudeNumerator.Value = new uint[]
                {
                    (uint)latCoords[0],
                    (uint)latCoords[1],
                    (uint)(latCoords[2] * denominator)
                };
                so.Properties.System.GPS.LongitudeNumerator.Value = new uint[]
                {
                    (uint)lonCoords[0],
                    (uint)lonCoords[1],
                    (uint)(lonCoords[2] * denominator)
                };
                so.Properties.System.GPS.LatitudeDenominator.Value = new uint[] { 1, 1, denominator };
                so.Properties.System.GPS.LongitudeDenominator.Value = new uint[] { 1, 1, denominator };
                return true;
            }
            return false;
        }
        private bool WriteGeoCoordinates(GeoCoordinates? geoCoords = null)
        {
            geoCoords ??= GeoCoordinates ?? new();
            return WriteGeoCoordinates(geoCoords.LatitudeCoordinates, geoCoords.LongitudeCoordinates, geoCoords.LatitudeWindDirection, geoCoords.LongitudeWindDirection);
        }
        private bool WriteEasyMetaData(EasyMetadata? easyMetadata = null)
        {
            easyMetadata ??= EasyMetadata;
            string delimiter = $"{nameof(EasyMetadata)}{EQUALS_SIGN}";
            string? keyword = null;
            foreach (string kw in Keywords)
            {
                if (kw.SplitLast(delimiter) == null)
                {
                    continue;
                }
                keyword = kw;
                break;
            }
            bool delete = EasyMetadata == null || !EasyMetadata.IsValid;
            bool? deleteOrOverwrite = delete ? EasyMetadata.Options.Delete : EasyMetadata.Options.Overwrite;
            if (keyword == null)
            {
                if (easyMetadata == null)
                {
                    return false;
                }
                return AddKeyword($"{delimiter}{EasyMetadata}");
            }
            string caption = easyMetadata == null ? Globals.DeleteEasyMetadata : Globals.OverwriteEasyMetadata;
            if (deleteOrOverwrite == null)
            {
                using MessageCheckDialog md = new(Globals.MessageCheckDialogText, caption);
                deleteOrOverwrite = md.ShowDialog() == DialogResult.Yes;
                if (md.Checked)
                {
                    if (delete)
                    {
                        EasyMetadata.Options.Delete = deleteOrOverwrite;
                    }
                    else
                    {
                        EasyMetadata.Options.Overwrite = deleteOrOverwrite;
                    }
                }
            }
            if (deleteOrOverwrite is bool b && b)
            {
                if (keyword != null)
                {
                    RemoveKeyword(keyword);
                }
                if (!delete)
                {
                    return AddKeyword($"{delimiter}{easyMetadata}");
                }
            }
            return false;
        }

        public EasyFile() : base() { }

        public EasyFile(string path, bool readShell = true) : base(path, readShell)
        {
            if (IsFile) { return; }
            Dispose(); throw new ArgumentException("Path is not a file.");
        }

        public override void Initialize(string path = EMPTY_STRING, bool readShell = true)
        {
            base.Initialize(path, readShell);
            _formattedPath = path;
            _easyMetadata = null;
            _geoCoordinates = null;
            _geoObject = null;

            if (!IsInvalid || readShell)
            {
                ReadDateTaken();
                ReadDateEncoded();
                ReadAreaInfo();
                ReadGeoCoordinates();
                ReadEasyMetadata();
            }
        }

        public bool AddKeyword(string keyword)
        {
            List<string> l = Keywords.ToList();
            if (!l.Contains(keyword))
            {
                l.Add(keyword);
                Keywords = l.ToArray();
                return true;
            }
            return false;
        }

        public bool RemoveKeyword(string keyword)
        {
            List<string> l = Keywords.ToList();
            if (l.Remove(keyword))
            {
                Keywords = l.ToArray();
                return true;
            }
            return false;
        }

        public bool RemoveProperty(EasyFileProperty property)
        {
            bool result = false;
            switch (property)
            {
                case EasyFileProperty.Title: if (!string.IsNullOrEmpty(Title)) { Title = string.Empty; result = true; }; break;
                case EasyFileProperty.Comment: if (!string.IsNullOrEmpty(Comment)) { Comment = string.Empty; result = true; }; break;
                case EasyFileProperty.Keywords: if (Keywords.Any()) { Keywords = Array.Empty<string>(); result = true; }; break;
                case EasyFileProperty.GPSCoordinates: if (_geoCoordinates != null || ReadGeoCoordinates()) { _geoCoordinates = null; WriteGeoCoordinates(); result = true; } break;
                case EasyFileProperty.AreaInfo: if (!string.IsNullOrEmpty(AreaInfo)) { _areaInfo = string.Empty; result = WriteAreaInfo(); }; break;
                case EasyFileProperty.EasyMetadata: if (EasyMetadata != null || ReadEasyMetadata()) { _easyMetadata = null; WriteEasyMetaData(); result = true; } break;
            }
            return result;
        }

        public bool UpdateFormatting(EasyOptions? options = null, string[]? excludedPaths = null)
        {
            if (IsInvalid || IsReadOnly)
            {
                return false;
            }
            Initialize();
            if (options != null && options.CustomizeEnabled && options.CopyMoveState != CheckState.Unchecked && options.SubfoldersEnabled)
            {
                foreach (EasySubfolder esf in options.Subfolders)
                {
                    switch (esf)
                    {
                        case EasySubfolder.AreaInfo:
                        case EasySubfolder.AreaInfoFromEasyMetadata:
                        case EasySubfolder.YearCreated:
                        case EasySubfolder.YearTakenOrEncoded:
                        case EasySubfolder.YearFromEasyMetadata:
                        case EasySubfolder.YearEarliest:
                        case EasySubfolder.MonthCreated:
                        case EasySubfolder.MonthTakenOrEncoded:
                        case EasySubfolder.MonthFromEasyMetadata:
                        case EasySubfolder.MonthEarliest:
                        case EasySubfolder.DayCreated:
                        case EasySubfolder.DayTakenOrEncoded:
                        case EasySubfolder.DayFromEasyMetadata:
                        case EasySubfolder.DayEarliest:
                            CustomizeAsync(options, false).Wait();
                            break;
                    }
                }
            }
            string n = NameWithoutExtension;
            string pp = ParentPath;
            string ext = Extension;
            DateTime? dte = DateEarliest;
            if (options != null && (options.RenameEnabled is bool r | (options.CopyMoveState != CheckState.Unchecked) is bool cm))
            {
                if (cm)
                {
                    if (options.TopFolderEnabled && !string.IsNullOrEmpty(options.TopFolderPath))
                    {
                        pp = options.TopFolderPath;
                    }
                    if (options.SubfoldersEnabled)
                    {
                        foreach (EasySubfolder esf in options.Subfolders)
                        {
                            string sf = string.Empty;
                            switch (esf)
                            {
                                case EasySubfolder.FileType:
                                    if (Enum.GetName(Type) is string s) { sf = s; }
                                    break;
                                case EasySubfolder.FileTypeText:
                                    sf = TypeText;
                                    break;
                                case EasySubfolder.FileExtension:
                                    sf = ext.Replace($"{PERIOD}", EMPTY_STRING).ToUpper();
                                    break;
                                case EasySubfolder.Subject:
                                    sf = Subject;
                                    break;
                                case EasySubfolder.AreaInfo:
                                    sf = AreaInfo;
                                    break;
                                case EasySubfolder.AreaInfoFromEasyMetadata:
                                    if (AreaInfoFromEasyMetadata is string aifem) { sf = aifem; }
                                    break;
                                case EasySubfolder.YearCreated:
                                    if (DateCreated is DateTime yc) { sf = yc.ToDateYearString(); }
                                    break;
                                case EasySubfolder.YearTakenOrEncoded:
                                    if (DateTakenOrEncoded is DateTime ytoe) { sf = ytoe.ToDateYearString(); }
                                    break;
                                case EasySubfolder.YearFromEasyMetadata:
                                    if (DateFromEasyMetadata is DateTime yfem) { sf = yfem.ToDateYearString(); }
                                    break;
                                case EasySubfolder.YearEarliest:
                                    if (DateEarliest is DateTime ye) { sf = ye.ToDateYearString(); }
                                    break;
                                case EasySubfolder.MonthCreated:
                                    if (DateCreated is DateTime mc) { sf = mc.ToDateMonthString(); }
                                    break;
                                case EasySubfolder.MonthTakenOrEncoded:
                                    if (DateTakenOrEncoded is DateTime mtoe) { sf = mtoe.ToDateMonthString(); }
                                    break;
                                case EasySubfolder.MonthFromEasyMetadata:
                                    if (DateFromEasyMetadata is DateTime mfem) { sf = mfem.ToDateMonthString(); }
                                    break;
                                case EasySubfolder.MonthEarliest:
                                    if (DateEarliest is DateTime me) { sf = me.ToDateMonthString(); }
                                    break;
                                case EasySubfolder.DayCreated:
                                    if (DateCreated is DateTime dc) { sf = dc.ToDateDayString(); }
                                    break;
                                case EasySubfolder.DayTakenOrEncoded:
                                    if (DateTakenOrEncoded is DateTime dtoe) { sf = dtoe.ToDateDayString(); }
                                    break;
                                case EasySubfolder.DayFromEasyMetadata:
                                    if (DateFromEasyMetadata is DateTime dfem) { sf = dfem.ToDateDayString(); }
                                    break;
                                case EasySubfolder.DayEarliest:
                                    if (DateEarliest is DateTime de) { sf = de.ToDateDayString(); }
                                    break;
                                case EasySubfolder.CameraManufacturer:
                                    if (CameraManufacturer is string man) { sf = man; }
                                    break;
                                case EasySubfolder.CameraModel:
                                    if (CameraModel is string mod) { sf = mod; }
                                    break;
                            }
                            if (!string.IsNullOrEmpty(sf))
                            {
                                pp = System.IO.Path.Join(pp, sf);
                            }
                        }
                    }
                }
                if (r)
                {
                    if (options.ReplaceWithState != CheckState.Unchecked && !string.IsNullOrEmpty(options.Replace))
                    {
                        if (options.ReplaceWithState == CheckState.Indeterminate)
                        {
                            try
                            {
                                n = Regex.Replace(n, options.Replace, options.With);
                                goto DatePattern;
                            }
                            catch (RegexParseException) { }
                        }
                        n = n.Replace(options.Replace, options.With);
                        goto DatePattern;
                    }
                DatePattern:
                    if (options.DateFormatEnabled && dte is DateTime dt)
                    {
                        n += dt.ToString(options.DateFormat);
                    }
                    n = $"{options.Prefix}{n}{options.Suffix}";
                }
            }
            if (string.IsNullOrEmpty(n))
            {
                n = WinFormsLib.Globals.FileNameDefault;
            }
            string p = System.IO.Path.Join(pp, $"{n}{ext}");
            _formattedPath = GetPathDuplicate(p, excludedPaths ?? GetAllPaths(pp));
            return p != _path;
        }

        public bool Rotate(EasyOrientation orientation, bool preserveDateModified = false)
        {
            if (_shellObject is ShellObject so && Type == EasyType.Image)
            {
                DateTime? dtm = preserveDateModified ? DateModified : null;

                int i = orientation.GetValue<int>();
                if (so.Properties.System.Photo.Orientation.Value is ushort us)
                {
                    switch (us)
                    {
                        case 6:
                            switch (i) { case 6: i = 3; break; case 3: i = 8; break; case 8: i = 1; break; }
                            break;
                        case 3:
                            switch (i) { case 6: i = 8; break; case 3: i = 1; break; case 8: i = 6; break; }
                            break;
                        case 8:
                            switch (i) { case 6: i = 1; break; case 3: i = 6; break; case 8: i = 3; break; }
                            break;
                    }
                }
                so.Properties.System.Photo.Orientation.Value = (ushort?)i;

                if (preserveDateModified && dtm is DateTime dt) { _dateModified = dt.Similar(); WriteDateModified(); Initialize(_path); }

                return true;
            }
            return false;
        }

        public void ExtractExifGeoArea(bool useEasyMetadataWithVideo = true)
        {
            if (Type != EasyType.Video
                || (_geoCoordinates is GeoCoordinates gc && gc.IsValid)
                || (useEasyMetadataWithVideo && EasyMetadata is EasyMetadata emd && emd.GeoArea != null)
                || ExtractExifGPSCoords(_path) is not GeoCoordinates gc0 || !gc0.IsValid)
            {
                return;
            }
            _geoCoordinates = gc0;
            if (useEasyMetadataWithVideo)
            {
                bool result = true;
                if (EasyMetadata.Options.VideoMetadata is not bool ov)
                {
                    using MessageCheckDialog md = new(Globals.MessageCheckDialogText, Globals.WriteEasyMetadata);
                    result = md.ShowDialog() == DialogResult.Yes;
                    if (md.Checked)
                    {
                        EasyMetadata.Options.VideoMetadata = result;
                    }
                }
                else
                {
                    result = ov;
                }
                if (result)
                {
                    _easyMetadata ??= new();
                    _easyMetadata.GeoArea ??= new() { AreaInfo = AreaInfo };
                    if (GeoCoordinates is GeoCoordinates gc1)
                    {
                        _easyMetadata.GeoArea.GeoCoords = gc1;
                    }
                    WriteEasyMetaData();
                }
            }
        }

        public async Task CustomizeAsync(EasyOptions? options = null, bool apply = true)
        {
            if (options != null && options.CustomizeEnabled)
            {
                DateTime? dtm = options.PreserveDateModified ? DateModified : null;
                if (options.TitleEnabled)
                {
                    string s = Subject;
                    Title = options.Title;
                    Subject = s;
                }
                if (options.SubjectEnabled) { Subject = options.Subject; }
                if (options.CommentEnabled) { Comment = options.Comment; }
                if (options.KeywordsState != CheckState.Unchecked)
                {
                    List<string> l = options.KeywordsState == CheckState.Indeterminate ? Keywords.ToList() : new();
                    if (options.KeywordsIndex != -1)
                    {
                        l.AddRange(options.Keywords[options.KeywordsIndex]);
                    }
                    Keywords = l.ToArray();
                }
                if (options.GeolocationEnabled && (options.WriteGPSEasyMetadata || options.WriteGPSCoords || options.WriteGPSAreaInfo))
                {
                    string? areaInfo = null;
                    GeoCoordinates? geoCoords = null;

                    EasyMetadataSource egds = Enum.GetValues<EasyMetadataSource>()[options.GPSSourceIndex];
                    if (egds != EasyMetadataSource.CustomMetadata)
                    {
                        GeoObject? geoObject = null;
                        bool allTypes = egds == EasyMetadataSource.EasyShellOrVideoMetadata;
                        if ((allTypes || egds == EasyMetadataSource.EasyMetadata) && EasyMetadata is EasyMetadata emd && emd.GeoArea is GeoArea ga)
                        {
                            areaInfo = ga.AreaInfo;
                            geoCoords = ga.GeoCoords;
                        }
                        if (allTypes || egds == EasyMetadataSource.ShellProperties)
                        {
                            if (string.IsNullOrEmpty(areaInfo))
                            {
                                areaInfo = AreaInfo;
                            }
                            if (geoCoords == null || !geoCoords.IsValid)
                            {
                                geoCoords = GeoCoordinates;
                            }
                            geoObject = GeoObject;
                        }
                        if ((allTypes || egds == EasyMetadataSource.VideoMetadata) && Type == EasyType.Video)
                        {
                            geoCoords = ExtractExifGPSCoords(_path);
                        }

                        if (geoObject == null) // TODO: add option to prioritize area info over gps coordinates
                        {
                            if (geoCoords != null && geoCoords.IsValid)
                            {
                                geoObject = await GetGeoObjectAsync(geoCoords.Latitude, geoCoords.Longitude);
                                if (geoObject != null && geoObject.GetGeoAddress() is GeoAddress geoAddress)
                                {
                                    areaInfo = geoAddress.GetArea();
                                }
                            }
                            else if (!string.IsNullOrEmpty(areaInfo))
                            {
                                geoObject = await GetGeoObjectAsync(areaInfo);
                            }
                        }

                        if (geoObject == null)
                        {
                            Debug.WriteLine($"File {Name} has no available geolocation data.");
                        }
                        else
                        {
                            GeoAreaDescription gad = Enum.GetValues<GeoAreaDescription>()[options.GPSAreaIndex];
                            if (gad != GeoAreaDescription.Address)
                            {
                                if (geoObject.GetGeoAddress() is GeoAddress geoAddress)
                                {
                                    areaInfo = geoAddress.GetArea(gad);
                                    if (!string.IsNullOrEmpty(areaInfo))
                                    {
                                        geoObject = await GetGeoObjectAsync(areaInfo);
                                    }
                                }
                            }
                            if (geoObject != null)
                            {
                                geoCoords ??= geoObject.GetGeoCoordinates();
                                _geoObject = geoObject;
                                _geoCoordinates = geoCoords;
                            }
                        }
                    }
                    else if (options.GPSCustomIndex != -1)
                    {
                        KeyValuePair<string, GeoCoordinates?> kvp = options.GeoAreas.ToArray()[options.GPSCustomIndex];
                        areaInfo = kvp.Key;
                        geoCoords = kvp.Value;
                    }

                    if (options.WriteGPSEasyMetadata)
                    {
                        GeoArea? geoArea = null;
                        if (!string.IsNullOrEmpty(areaInfo))
                        {
                            geoArea = new() { AreaInfo = areaInfo };
                        }
                        if (geoCoords is GeoCoordinates gc && gc.IsValid)
                        {
                            geoArea ??= new();
                            geoArea.GeoCoords = gc;
                        }
                        _easyMetadata ??= new();
                        if (geoArea != null)
                        {
                            _easyMetadata.GeoArea = geoArea;
                            if (apply)
                            {
                                WriteEasyMetaData();
                            }
                        }
                        else if (_easyMetadata.GeoArea != null)
                        {
                            _easyMetadata.GeoArea = null;
                            if (apply)
                            {
                                WriteEasyMetaData();
                            }
                        }
                    }
                    if (options.WriteGPSAreaInfo)
                    {
                        _areaInfo = areaInfo ?? string.Empty;
                        if (apply)
                        {
                            WriteAreaInfo();
                        }
                    }
                    if (options.WriteGPSCoords)
                    {
                        _geoCoordinates = geoCoords;
                        if (apply)
                        {
                            WriteGeoCoordinates();
                        }
                    }
                }
                if (options.EasyMetadataEnabled)
                {
                    if (options.EasyMetadataIndex != -1)
                    {
                        _easyMetadata = new(options.EasyMetadata[options.EasyMetadataIndex]);
                        if (apply)
                        {
                            WriteEasyMetaData();
                        }
                    }
                }
                if (options.DateState != CheckState.Unchecked && (options.WriteDateModified || options.WriteDateCreated || options.WriteDateCamera || options.WriteDateEasyMetadata))
                {
                    DateTime? dt = null;
                    if (options.DateState == CheckState.Indeterminate)
                    {
                        dt = Enum.GetValues<EasyDateSource>()[options.DateSourceIndex] switch
                        {
                            EasyDateSource.DateModified => DateModified,
                            EasyDateSource.DateCreated => DateCreated,
                            EasyDateSource.DateTakenOrEncoded => DateTakenOrEncoded,
                            EasyDateSource.DateFromEasyMetadata => DateFromEasyMetadata,
                            EasyDateSource.DateEarliest => DateEarliest,
                            EasyDateSource.DateFromFolderName => DirectoryName.AsDateTime(),
                            _ => null,
                        };
                    }
                    else if (options.DateCustomIndex != -1)
                    {
                        string s = options.Dates[options.DateCustomIndex];
                        dt = string.IsNullOrEmpty(s) ? DateTime.MaxValue : s.AsDateTime();
                    }
                    if (dt != null)
                    {
                        if (dt == DateTime.MaxValue)
                        {
                            dt = null;
                        }
                        if (options.WriteDateEasyMetadata)
                        {
                            _easyMetadata ??= new();
                            _easyMetadata.Date = dt?.ToDateTimeString();
                            if (apply)
                            {
                                WriteEasyMetaData();
                            }
                        }
                        if (options.WriteDateCamera)
                        {
                            if (Type == EasyType.Image)
                            {
                                _dateTaken = dt;
                                if (apply)
                                {
                                    WriteDateTaken();
                                }
                            }
                            else if (Type == EasyType.Video)
                            {
                                _dateEncoded = dt;
                                if (apply)
                                {
                                    WriteDateEncoded();
                                }
                            }
                        }
                        if (options.WriteDateCreated)
                        {
                            _dateCreated = dt;
                            if (apply)
                            {
                                WriteDateCreated();
                            }
                        }
                        if (options.WriteDateModified)
                        {
                            _dateModified = dt;
                            if (apply)
                            {
                                WriteDateModified();
                            }
                            dtm = null;
                        }
                    }
                }
                if (dtm != null)
                {
                    _dateModified = dtm;
                    if (apply)
                    {
                        WriteDateModified();
                    }
                }
            }
        }
    }
}
