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
                        EasyFiles l = kvp.Value[o];
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
            return (int)((1 - f) * value + f * maxValue);
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
    public enum EasyFilter
    {
        None = 0,
        [EasyGlobalStringValue("Text")]
        [Utils.Value(EasyType.Text)]
        Text = 1 << 0,
        [EasyGlobalStringValue("Document")]
        [Utils.Value(EasyType.Document)]
        Document = 1 << 1,
        [EasyGlobalStringValue("Image")]
        [Utils.Value(EasyType.Image)]
        Image = 1 << 2,
        [EasyGlobalStringValue("Video")]
        [Utils.Value(EasyType.Video)]
        Video = 1 << 3,
        [EasyGlobalStringValue("Audio")]
        [Utils.Value(EasyType.Audio)]
        Audio = 1 << 4,
        [EasyGlobalStringValue("Compressed")]
        [Utils.Value(EasyType.Compressed)]
        Compressed = 1 << 5,
        Default = Image | Video,
        All = Text | Document | Image | Video | Audio | Compressed
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
        [Utils.Value(8)]
        Rotate270,
        [Utils.Value(3)]
        Rotate180,
        [Utils.Value(6)]
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
        public string DateFormat { get; set; } = DateTimeExtensions.FORMAT_DATE_SECOND;
        public string BackupFolderPath { get; set; } = STARTUP_DIRECTORY_DEFAULT;
        public string TopFolderPath { get; set; } = STARTUP_DIRECTORY_DEFAULT;
        public string DuplicatesFolderPath { get; set; } = STARTUP_DIRECTORY_DEFAULT;

        public bool CustomizeEnabled { get; set; }
        public bool RenameEnabled { get; set; }
        public bool MoveEnabled { get; set; }
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
        public bool TopFolderEnabled { get; set; }
        public bool SubfoldersEnabled { get; set; }
        public bool FilterEnabled { get; set; }
        public bool UseEasyMetadataWithVideo { get; set; }
        public bool PreserveDateModified { get; set; }
        public bool PreserveDateCreated { get; set; }
        public bool LogApplicationEvents { get; set; }
        public bool ShutdownUponCompletion { get; set; }
        public bool DeleteEmptyFolders { get; set; }
        public bool ShowDuplicatesCompareDialog { get; set; }

        public int KeywordsIndex { get; set; } = -1;
        public int GPSSourceIndex { get; set; } = 4;
        public int GPSCustomIndex { get; set; } = -1;
        public int GPSAreaIndex { get; set; } = 6;
        public int EasyMetadataIndex { get; set; } = -1;
        public int DateSourceIndex { get; set; } = 4;
        public int DateCustomIndex { get; set; } = -1;

        public CheckState KeywordsState { get; set; } = CheckState.Unchecked;
        public CheckState DateState { get; set; } = CheckState.Unchecked;
        public CheckState ReplaceWithState { get; set; } = CheckState.Unchecked;
        public CheckState BackupFolderState { get; set; } = CheckState.Unchecked;
        public CheckState DuplicatesState { get; set; } = CheckState.Unchecked;

        public EasyFilter Filter { get; set; } = EasyFilter.Default;
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
                BackupFolderPath = Utils.IsValidDirectoryPath(eo.BackupFolderPath) ? eo.BackupFolderPath : STARTUP_DIRECTORY_DEFAULT;
                TopFolderPath = Utils.IsValidDirectoryPath(eo.TopFolderPath) ? eo.TopFolderPath : STARTUP_DIRECTORY_DEFAULT;
                DuplicatesFolderPath = Utils.IsValidDirectoryPath(eo.DuplicatesFolderPath) ? eo.DuplicatesFolderPath : STARTUP_DIRECTORY_DEFAULT;

                CustomizeEnabled = eo.CustomizeEnabled;
                RenameEnabled = eo.RenameEnabled;
                MoveEnabled = eo.MoveEnabled;
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
                TopFolderEnabled = eo.TopFolderEnabled;
                SubfoldersEnabled = eo.SubfoldersEnabled;
                FilterEnabled = eo.FilterEnabled;
                WriteDateEasyMetadata = eo.WriteDateEasyMetadata;
                UseEasyMetadataWithVideo = eo.UseEasyMetadataWithVideo;
                PreserveDateModified = eo.PreserveDateModified;
                PreserveDateCreated = eo.PreserveDateCreated;
                LogApplicationEvents = eo.LogApplicationEvents;
                ShutdownUponCompletion = eo.ShutdownUponCompletion;
                DeleteEmptyFolders = eo.DeleteEmptyFolders;
                ShowDuplicatesCompareDialog = eo.ShowDuplicatesCompareDialog;

                KeywordsIndex = eo.KeywordsIndex;
                GPSSourceIndex = eo.GPSSourceIndex;
                GPSCustomIndex = eo.GPSCustomIndex;
                GPSAreaIndex = eo.GPSAreaIndex;
                DateSourceIndex = eo.DateSourceIndex;
                DateCustomIndex = eo.DateCustomIndex;
                EasyMetadataIndex = eo.EasyMetadataIndex;

                KeywordsState = eo.KeywordsState;
                DateState = eo.DateState;
                ReplaceWithState = eo.ReplaceWithState;
                BackupFolderState = eo.BackupFolderState;
                DuplicatesState = eo.DuplicatesState;
                DuplicatesCompareParameters = eo.DuplicatesCompareParameters;

                Keywords = eo.Keywords;
                GeoAreas = eo.GeoAreas;
                EasyMetadata = eo.EasyMetadata;
                Dates = eo.Dates;
                Subfolders = eo.Subfolders;
                Filter = eo.Filter;
            }
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

        public virtual EasyPath? Add(string path)
        {
            if (!Contains(path))
            {
                EasyPath ep = new(path);
                if (!ep.IsInvalid)
                {
                    Add(ep);
                    return ep;
                }
                ep.Dispose();
            }
            return null;
        }

        public void AddRange(IEnumerable<string> paths)
        {
            foreach (string path in paths)
            {
                Add(path);
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
            base.Remove(easyPath);
            easyPath.Dispose();
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
            EasyPath ep = this[index];
            base.RemoveAt(index);
            ep.Dispose();
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

        public void Replace(IEnumerable<string> paths)
        {
            Clear();
            AddRange(paths);
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

        public override EasyFolder? Add(string path)
        {
            if (!Contains(path))
            {
                EasyFolder ef = new(path);
                if (!ef.IsInvalid)
                {
                    Add(ef);
                    return ef;
                }
                ef.Dispose();
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

        public override EasyFile? Add(string path)
        {
            if (!Contains(path))
            {
                EasyFile ef = new(path);
                if (!ef.IsInvalid)
                {
                    Add(ef);
                    return ef;
                }
                ef.Dispose();
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
        public static JsonConverter JsonConverter => new EasyPathJsonConverter();

        private long? _size = null;

        public FileAttributes FileAttributes { get; private set; }
        public ShellObject? ShellObject { get; private set; } = null;
        public string Path { get; private set; } = EMPTY_STRING;
        public bool IsDisposed { get; private set; }

        public string Name => Utils.GetFileName(Path);
        public string NameWithoutExtension => Utils.GetFileNameWithoutExtension(Path);
        public string PathWithoutExtension => Utils.GetPathWithoutExtension(Path);
        public string Extension => Utils.GetExtension(Path);
        public string DirectoryPath => Utils.GetDirectoryPath(Path);
        public string DirectoryName => Utils.GetFileName(DirectoryPath);
        public string ParentPath => Utils.GetParentPath(Path);
        public string RawName => Utils.GetRawName(Path);
        public string RawPath => Utils.GetRawPath(Path);
        public bool Exists => IsFolder ? Directory.Exists(Path) : File.Exists(Path);
        public bool IsInvalid => Type == EasyType.Invalid;
        public bool IsFile => EasyType.File.HasFlag(Type);
        public bool IsFolder => FileAttributes.HasFlag(FileAttributes.Directory);
        public bool IsSystem => FileAttributes.HasFlag(FileAttributes.System);
        public bool IsHidden => FileAttributes.HasFlag(FileAttributes.Hidden) || Path.StartsWith(PERIOD);
        public bool IsReadOnly => FileAttributes.HasFlag(FileAttributes.ReadOnly);
        public bool IsSymbolicLink => FileAttributes.HasFlag(FileAttributes.ReparsePoint);
        public long Size
        {
            get
            {
                UpdateSize();
                return _size is long l ? l : 0;
            }
        }
        public Image? SmallThumbnail => Utils.GetThumbnailImage(ShellObject?.Thumbnail.SmallBitmap);
        public Image? MediumThumbnail => Utils.GetThumbnailImage(ShellObject?.Thumbnail.MediumBitmap);
        public Image? LargeThumbnail => Utils.GetThumbnailImage(ShellObject?.Thumbnail.LargeBitmap);
        public Image? ExtraLargeThumbnail => Utils.GetThumbnailImage(ShellObject?.Thumbnail.ExtraLargeBitmap);
        public FileInfo FileInfo => new(Path);

        public DateTime? DateModified
        {
            get => ShellObject != null && ShellObject.Properties.System.DateModified.Value is DateTime dt ? dt : File.GetLastWriteTime(Path);
            set { if (value is DateTime dt) { File.SetLastWriteTime(Path, dt); } }
        }
        public DateTime? DateCreated
        {
            get => ShellObject != null && ShellObject.Properties.System.DateCreated.Value is DateTime dt ? dt : File.GetCreationTime(Path);
            set { if (value is DateTime dt) { File.SetCreationTime(Path, dt); } }
        }
        public EasyType Type
        {
            get
            {
                EasyType et = EasyType.None;
                if (ShellObject is ShellObject so)
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
                if (ShellObject != null)
                {
                    string t = $"{ShellObject.Properties.System.ItemTypeText.Value}";
                    if (ShellObject.Properties.System.ItemType.Value is string it)
                    {
                        t += $" ({it})";
                    }
                    return t;
                }
                return EMPTY_STRING;
            }
        }

        public EasyPath(string path = EMPTY_STRING) => Initialize(path);

        public virtual void Initialize(string path)
        {
            Path = path;
            ShellObject?.Dispose();
            if (string.IsNullOrEmpty(path))
            {
                Debug.WriteLine($"EasyPath is empty.");
                return;
            }
            if (path == Utils.GetParentPath(path))
            {
                Debug.WriteLine($"EasyPath is root.");
                return;
            }
            try
            {
                ShellObject = ShellObject.FromParsingName(path);
                FileAttributes = GetFileAttributes();
                if (IsSymbolicLink)
                {
                    if (IsFolder)
                    {
                        if (Utils.GetTargetDirectory(path) is string td)
                        {
                            Initialize(td);
                        }
                    }
                    else if (Utils.GetTargetFile(path) is string tf)
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

        private FileAttributes GetFileAttributes()
        {
            if (ShellObject != null)
            {
                try
                {
                    uint i = (uint)ShellObject.Properties.DefaultPropertyCollection["System.FileAttributes"].ValueAsObject;
                    return (FileAttributes)Enum.ToObject(typeof(FileAttributes), i);
                }
                catch (IndexOutOfRangeException) { }
            }
            return File.GetAttributes(Path);
        }

        private bool MoveOrCopy(bool copy, string path, bool preserveDateCreated = true, bool preserveDateModified = true)
        {
            bool result = false;
            if (!IsInvalid && (copy || !IsReadOnly))
            {
                DateTime? dtc = preserveDateCreated ? DateCreated : null;
                DateTime? dtm = preserveDateModified ? DateModified : null;
                result = copy
                    ? IsFolder ? Utils.CopyDirectory(Path, path) : Utils.CopyFile(Path, path)
                    : IsFolder ? Utils.MoveDirectory(Path, path) : Utils.MoveFile(Path, path);
                if (result)
                {
                    if (preserveDateCreated || preserveDateModified) { Initialize(path); }
                    if (preserveDateCreated) { DateCreated = dtc; }
                    if (preserveDateModified) { DateModified = dtm; }
                }
            }
            return result;
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                ShellObject?.Dispose();
                ShellObject = null;
            }
        }

        public void Dispose()
        {
            if (!IsDisposed)
            {
                Dispose(true);
                GC.SuppressFinalize(this);
                IsDisposed = true;
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
                    Task t = Utils.RunCancellableTask(() =>
                    {
                        try
                        {
                            long l = 0;
                            foreach (FileInfo fi in new DirectoryInfo(Path).EnumerateFiles("*", SearchOption.AllDirectories))
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
                else if (IsFile && ShellObject != null && ShellObject.Properties.System.FileAllocationSize.Value is ulong ul)
                {
                    _size = Convert.ToInt64(ul);
                }
                else if (!IsInvalid)
                {
                    try { _size = new FileInfo(Path).Length; } catch { }
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
            bool result = IsFolder ? Utils.DeleteDirectory(Path) : Utils.DeleteFile(Path);
            if (result) { Dispose(); }
            return result;
        }

        public int CompareTo(EasyPath? other) => other != null ? Path.CompareTo(other.Path) : throw new ArgumentNullException(nameof(other));

        public bool Equals(EasyPath? other)
        {
            return other != null && ((ShellObject == null
                && other.ShellObject == null) || (ShellObject != null && ShellObject.Equals(other.ShellObject)));
        }

        public override bool Equals(object? obj) => Equals(obj as EasyPath);

        public override int GetHashCode() => HashCode.Combine(Path);

        public override string ToString() => Path;

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

        public EasyFolder(string path = EMPTY_STRING) : base(path) { }

        public string[] GetFilePaths(bool recursive = false) => Utils.GetFilePaths(Path, recursive);

        public string[] GetDirectoryPaths(bool recursive = false) => Utils.GetDirectoryPaths(Path, recursive);

        public string[] GetAllPaths(bool recursive = false) => Utils.GetAllPaths(Path, recursive);

        public string[] GetRawPaths(bool recursive = false) => Utils.GetRawPaths(GetAllPaths(recursive));
    }

    public class EasyFile : EasyPath
    {
        public string FormattedPath { get; private set; } = EMPTY_STRING;
        public EasyMetadata? EasyMetadata { get; private set; } = null;
        public GeoCoordinates? GeoCoordinates { get; private set; } = null;
        public GeoObject? GeoObject { get; private set; } = null;

        public DateTime? DateTaken
        {
            get => Type == EasyType.Image && ShellObject is ShellObject so && so.Properties.System.Photo.DateTaken.Value is DateTime dt ? dt : null;
            set { if (Type == EasyType.Image && ShellObject is ShellObject so) { so.Properties.System.Photo.DateTaken.Value = value; } }
        }
        public DateTime? DateEncoded
        {
            get => Type == EasyType.Video && ShellObject is ShellObject so && so.Properties.System.Media.DateEncoded.Value is DateTime dt ? dt : null;
            set { if (Type == EasyType.Video && ShellObject is ShellObject so) { so.Properties.System.Media.DateEncoded.Value = value; } }
        }
        public DateTime? DateTakenOrEncoded => Type == EasyType.Image ? DateTaken : Type == EasyType.Video ? DateEncoded : null;
        public DateTime? DateFromEasyMetadata
        {
            get => EasyMetadata is EasyMetadata emd && emd.Date is string s ? s.AsDateTime() : null;
            set
            {
                if (EasyMetadata == null)
                {
                    if (value == null)
                    {
                        return;
                    }
                    EasyMetadata = new();
                }
                EasyMetadata.Date = value?.ToDateTimeString();
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

        public string FormattedName => Utils.GetFileName(FormattedPath);
        public string FormattedParentPath => Utils.GetParentPath(FormattedPath);
        public string FormattedNameWithoutExtension => Utils.GetFileNameWithoutExtension(FormattedPath);

        public string Title
        {
            get => ShellObject != null ? ShellObject.Properties.System.Title.Value : EMPTY_STRING;
            set { if (ShellObject != null) { ShellObject.Properties.System.Title.Value = value; } }
        }
        public string Subject
        {
            get => ShellObject != null ? ShellObject.Properties.System.Subject.Value : EMPTY_STRING;
            set { if (ShellObject != null) { ShellObject.Properties.System.Subject.Value = value; } }
        }
        public string Comment
        {
            get => ShellObject != null ? ShellObject.Properties.System.Comment.Value : EMPTY_STRING;
            set { if (ShellObject != null) { ShellObject.Properties.System.Comment.Value = value; } }
        }
        public string[] Keywords
        {
            get => ShellObject != null && ShellObject.Properties.System.Keywords.Value is string[] kw ? kw : Array.Empty<string>();
            set { if (ShellObject != null) { ShellObject.Properties.System.Keywords.Value = value; } }
        }
        public string AreaInfo
        {
            get => ShellObject != null ? ShellObject.Properties.System.GPS.AreaInformation.Value : EMPTY_STRING;
            set
            {
                if (ShellObject is ShellObject so && (EasyType.Image | EasyType.Video).HasFlag(Type))
                {
                    so.Properties.System.GPS.AreaInformation.Value = value;
                }
            }
        }
        public string? AreaInfoFromEasyMetadata
        {
            get => EasyMetadata is EasyMetadata emd && emd.GeoArea is GeoArea ga && !string.IsNullOrEmpty(ga.AreaInfo) ? ga.AreaInfo : null;
            set
            {
                if (EasyMetadata == null) { if (value == null) { return; } EasyMetadata = new(); }
                if (EasyMetadata.GeoArea == null) { if (value == null) { return; } EasyMetadata.GeoArea = new(); }
                EasyMetadata.GeoArea.AreaInfo = value;
            }
        }
        public string? CameraManufacturer
        {
            get => ShellObject is ShellObject so && Type == EasyType.Image ? so.Properties.System.Photo.CameraManufacturer.Value : null;
            set
            {
                if (ShellObject is ShellObject so && value is string s && Type == EasyType.Image)
                {
                    so.Properties.System.Photo.CameraManufacturer.Value = s;
                }
            }
        }
        public string? CameraModel
        {
            get => ShellObject is ShellObject so && Type == EasyType.Image ? so.Properties.System.Photo.CameraModel.Value : null;
            set
            {
                if (ShellObject is ShellObject so && value is string s && Type == EasyType.Image)
                {
                    so.Properties.System.Photo.CameraModel.Value = s;
                }
            }
        }

        public EasyFile(string path) : base(path) { }

        public override void Initialize(string path)
        {
            base.Initialize(path);
            FormattedPath = path;

            if (!IsInvalid)
            {
                ReadGeoCoordinates();
                ReadEasyMetadata();
            }
        }

        private void ReadGeoCoordinates()
        {
            if (ShellObject != null)
            {
                double[] lat = ShellObject.Properties.System.GPS.Latitude.Value;
                double[] lon = ShellObject.Properties.System.GPS.Longitude.Value;
                string latRef = ShellObject.Properties.System.GPS.LatitudeRef.Value;
                string lonRef = ShellObject.Properties.System.GPS.LongitudeRef.Value;
                if (lat != null && lon != null && !string.IsNullOrEmpty(latRef) & !string.IsNullOrEmpty(lonRef))
                {
                    GeoCoordinates = new(lat, lon, latRef, lonRef);
                }
            }
        }

        private bool WriteGeoCoordinates(double[] latCoords, double[] lonCoords, string latRef, string lonRef)
        {
            if (ShellObject != null && (EasyType.Image | EasyType.Video).HasFlag(Type))
            {
                uint denominator = 10000;
                ShellObject.Properties.System.GPS.LatitudeRef.Value = latRef;
                ShellObject.Properties.System.GPS.LongitudeRef.Value = lonRef;
                ShellObject.Properties.System.GPS.LatitudeNumerator.Value = new uint[]
                {
                    (uint)latCoords[0],
                    (uint)latCoords[1],
                    (uint)(latCoords[2] * denominator)
                };
                ShellObject.Properties.System.GPS.LongitudeNumerator.Value = new uint[]
                {
                    (uint)lonCoords[0],
                    (uint)lonCoords[1],
                    (uint)(lonCoords[2] * denominator)
                };
                ShellObject.Properties.System.GPS.LatitudeDenominator.Value = new uint[] { 1, 1, denominator };
                ShellObject.Properties.System.GPS.LongitudeDenominator.Value = new uint[] { 1, 1, denominator };
                return true;
            }
            return false;
        }

        private bool WriteGeoCoordinates(GeoCoordinates geoCoordinates) => WriteGeoCoordinates
        (
            geoCoordinates.LatitudeCoordinates,
            geoCoordinates.LongitudeCoordinates,
            geoCoordinates.LatitudeWindDirection,
            geoCoordinates.LongitudeWindDirection
        );

        private void ReadEasyMetadata()
        {
            string delimiter = $"{nameof(EasyMetadata)}{EQUALS_SIGN}";
            foreach (string kw in Keywords)
            {
                if (kw.SplitLast(delimiter) is string[] sa)
                {
                    EasyMetadata = new(sa.Last());
                    break;
                }
            }
        }

        private void DeleteOrOverwriteEasyMetaData(bool overwrite = false)
        {
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
            bool? deleteOrOverwrite = overwrite ? EasyMetadata.Options.Overwrite : EasyMetadata.Options.Delete;
            if (keyword == null)
            {
                AddKeyword($"{delimiter}{EasyMetadata}");
                return;
            }
            string caption = overwrite ? Globals.OverwriteEasyMetadata : Globals.DeleteEasyMetadata;
            if (deleteOrOverwrite == null)
            {
                using MessageCheckDialog md = new(Globals.MessageCheckDialogText, caption);
                deleteOrOverwrite = md.ShowDialog() == DialogResult.Yes;
                if (md.Checked)
                {
                    if (overwrite)
                    {
                        EasyMetadata.Options.Overwrite = deleteOrOverwrite;
                    }
                    else
                    {
                        EasyMetadata.Options.Delete = deleteOrOverwrite;
                    }
                }
            }
            if (deleteOrOverwrite is bool b && b)
            {
                if (!string.IsNullOrEmpty(keyword))
                {
                    RemoveKeyword(keyword);
                }
                if (overwrite)
                {
                    AddKeyword($"{delimiter}{EasyMetadata}");
                }
            }
        }

        private void WriteEasyMetaData() => DeleteOrOverwriteEasyMetaData(!(EasyMetadata == null || !EasyMetadata.IsValid));

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

        public bool UpdateFormatting(EasyOptions? options = null, string[]? excludedPaths = null)
        {
            if (IsInvalid || IsReadOnly)
            {
                return false;
            }
            string n = NameWithoutExtension;
            string pp = ParentPath;
            string ext = Extension;
            DateTime? dte = DateEarliest;
            if (options != null && (options.RenameEnabled is bool r | options.MoveEnabled is bool m))
            {
                if (m)
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
            FormattedPath = Utils.GetPathDuplicate(p, excludedPaths ?? Utils.GetAllPaths(pp));
            return p != Path;
        }

        public bool Rotate(EasyOrientation orientation, bool preserveDateModified = false)
        {
            if (ShellObject != null && Type == EasyType.Image)
            {
                DateTime? dtm = preserveDateModified ? DateModified : null;

                int i = orientation.GetValue<int>();
                if (ShellObject.Properties.System.Photo.Orientation.Value is ushort us)
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
                ShellObject.Properties.System.Photo.Orientation.Value = (ushort?)i;

                if (preserveDateModified && dtm is DateTime dt) { DateModified = dt.Similar(); Initialize(Path); }

                return true;
            }
            return false;
        }

        public async Task ExtractExifGeoAreaAsync(bool writeEasyMetadata = true)
        {
            if ((EasyMetadata is EasyMetadata emd && emd.GeoArea != null)
                || Type != EasyType.Video
                || GeoCoordinates is GeoCoordinates gc && gc.IsValid
                || await ExtractExifGPSCoords(Path) is not GeoCoordinates gc0)
            {
                return;
            }
            GeoCoordinates = gc0;
            if (writeEasyMetadata)
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
                    EasyMetadata ??= new();
                    EasyMetadata.GeoArea ??= new() { AreaInfo = AreaInfo };
                    if (GeoCoordinates is GeoCoordinates gc1)
                    {
                        EasyMetadata.GeoArea.GeoCoords = gc1;
                    }
                    WriteEasyMetaData();
                }
            }
        }

        public async Task CustomizeAsync(EasyOptions? options = null)
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
                            geoCoords = await ExtractExifGPSCoords(Path);
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
                                GeoObject = geoObject;
                                GeoCoordinates = geoCoords;
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
                        EasyMetadata ??= new();
                        if (geoArea != null)
                        {
                            EasyMetadata.GeoArea = geoArea;
                            WriteEasyMetaData();
                        }
                        else if (EasyMetadata.GeoArea != null)
                        {
                            EasyMetadata.GeoArea = null;
                            WriteEasyMetaData();
                        }
                    }
                    if (options.WriteGPSAreaInfo)
                    {
                        AreaInfo = areaInfo ?? string.Empty;
                    }
                    if (options.WriteGPSCoords && !WriteGeoCoordinates(geoCoords ?? new()))
                    {
                        Debug.WriteLine("Failed writing GPS coordinates.");
                    }
                }
                if (options.EasyMetadataEnabled)
                {
                    if (options.EasyMetadataIndex != -1)
                    {
                        EasyMetadata = new(options.EasyMetadata[options.EasyMetadataIndex]);
                        WriteEasyMetaData();
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
                            EasyMetadata ??= new();
                            EasyMetadata.Date = dt?.ToDateTimeString();
                            WriteEasyMetaData();
                        }
                        if (options.WriteDateCamera)
                        {
                            if (Type == EasyType.Image)
                            {
                                DateTaken = dt;
                            }
                            else if (Type == EasyType.Video)
                            {
                                DateEncoded = dt;
                            }
                        }
                        if (options.WriteDateCreated)
                        {
                            DateCreated = dt;
                        }
                        if (options.WriteDateModified)
                        {
                            DateModified = dt;
                            dtm = null;
                        }
                    }
                }
                if (dtm != null) { DateModified = dtm; }
            }
        }
    }
}
