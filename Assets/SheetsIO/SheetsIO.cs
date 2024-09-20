using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Numerics;
using System.Reflection;
using System.Threading.Tasks;
using Google.Apis.Http;
using Google.Apis.Requests;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util;

namespace SheetsIO
{
    public sealed class SheetsIO : IDisposable
    {
        internal const string FirstCell = "B2";
        internal const int MaxArrayElements = 100;
        
        internal delegate bool ReadObjectDelegate(IOPointer p, out object o);

        readonly SheetsService service;
        readonly IValueSerializer serializer;

		/// <summary>Creates a new ready-to-use instance of the serializer.</summary>
		/// <param name="initializer"> An initializer for Google Sheets service. </param>
		/// <param name="valueSerializer"> Replaces the default serialization strategy. </param>
		public SheetsIO(BaseClientService.Initializer initializer, IValueSerializer valueSerializer = null)
		{
			// If you got an issue with missing System.IO.Compression lib, set GZipEnabled = FALSE in your initializer.
			service = new SheetsService(initializer ?? throw new ArgumentException("SheetsService can't be null"));
			serializer = valueSerializer ?? new DefaultValueSerializer();
		}

		/// <summary>Read the data from a spreadsheet and deserialize it to object of type T.</summary>
		/// <param name="spreadsheet">Spreadsheet ID.</param>
		/// <param name="sheet">Specify the name of the sheet, if the spreadsheet contains multiple sheets of a type.</param>
		/// <typeparam name="T">Type of object to read from the spreadsheet data.</typeparam>
		/// <returns>Object of type T.</returns>
		public async Task<T> ReadAsync<T>(string spreadsheet, string sheet = "") {
            var spreadsheets = await GetSpreadsheetAsync(spreadsheet);
            var context = new ReadContext(GetSheetsList(spreadsheets), serializer);
            var meta = typeof(T).GetIOAttribute();
            if (!context.TryReadType(meta, meta.AppendNamePart(sheet), out var result))
                throw new Exception("Can't parse the requested object. Some required sheets are missing in the provided spreadsheet");

            var valueRanges = await GetValueRanges(spreadsheet, context.Ranges);
            if (!valueRanges.All(context.TryApplyRange))
                throw new Exception("Failed to assemble the object.");
            return (T)result;
        }

        /// <summary>Write a serialized object of type T to the spreadsheet.</summary>
        /// <param name="obj">Object to serialize.</param>
        /// <param name="spreadsheet">Spreadsheet ID.</param>
        /// <param name="sheet">Specify the name of the sheet, if the spreadsheet contains multiple sheets of a type.</param>
        /// <typeparam name="T">Type of serialized object.</typeparam>
        public Task<bool> WriteAsync<T>(T obj, string spreadsheet, string sheet = "") {
            var meta = typeof(T).GetIOAttribute();
            var ranges = new WriteContext(meta, meta.AppendNamePart(sheet), obj, serializer).ValueRanges;
            return WriteRangesAsync(spreadsheet, ranges);
        }

        public void Dispose() => service.Dispose();

        async Task<bool> WriteRangesAsync(string spreadsheet, IList<ValueRange> values) {
            var hashset = new HashSet<string>(values.Select(range => range.Range.GetSheetName()));
            bool hasRequiredSheets = await CreateSheetsAsync(spreadsheet, hashset);
            if (!hasRequiredSheets)
                return false;

            var result = await ExecuteWithBackOffHandler(service.Spreadsheets.Values.BatchUpdate(UpdateRequest(values), spreadsheet));
            return result.TotalUpdatedCells > 0;
        }

        async Task<Spreadsheet> GetSpreadsheetAsync(string spreadsheetID) => await ExecuteWithBackOffHandler(service.Spreadsheets.Get(spreadsheetID));

        async Task<IList<ValueRange>> GetValueRanges(string spreadsheet, IEnumerable<string> ranges) {
            var request = service.Spreadsheets.Values.BatchGet(spreadsheet);
            request.Ranges         = ranges.ToArray();
            request.MajorDimension = SpreadsheetsResource.ValuesResource.BatchGetRequest.MajorDimensionEnum.COLUMNS;
            var result = await ExecuteWithBackOffHandler(request);
            return result.ValueRanges;
        }

        async Task<bool> CreateSheetsAsync(string spreadsheet, IEnumerable<string> requiredSheets) {
            var spreadsheets = await GetSpreadsheetAsync(spreadsheet);
            string[] sheetsToCreate = requiredSheets.Except(GetSheetsList(spreadsheets)).ToArray();
            if (sheetsToCreate.Length == 0) return true;

            var result = await ExecuteWithBackOffHandler(service.Spreadsheets.BatchUpdate(AddSheet(sheetsToCreate), spreadsheet));
            return result.Replies.All(reply => reply.AddSheet.Properties != null);
        }

        IEnumerable<string> GetSheetsList(Spreadsheet spreadsheet) => spreadsheet.Sheets.Select(sheet => sheet.Properties.Title).ToArray();
        static BatchUpdateSpreadsheetRequest AddSheet(IEnumerable<string> list) => new BatchUpdateSpreadsheetRequest {Requests = list.Select(AddSheet).ToList()};
        static Request AddSheet(string name) => new Request {AddSheet = new AddSheetRequest {Properties = new SheetProperties {Title = name}}};
        static BatchUpdateValuesRequest UpdateRequest(IList<ValueRange> data) => new BatchUpdateValuesRequest {Data = data, ValueInputOption = "USER_ENTERED"};

        static Task<T> ExecuteWithBackOffHandler<T>(ClientServiceRequest<T> request, BackOffHandler handler = null) {
            request.AddExceptionHandler(handler ?? new BackOffHandler(new ExponentialBackOff(TimeSpan.FromMilliseconds(250), 5)));
            return request.ExecuteAsync();
        }
        
        readonly struct ReadContext
        {
            readonly IDictionary<string, (IOPointer[] pointers, object obj)> dictionary;
            readonly HashSet<string> sheets;
            readonly IValueSerializer serializer;
            public IEnumerable<string> Ranges => dictionary.Select(x => x.Key);

            public ReadContext(IEnumerable<string> sheets, IValueSerializer serializer) {
                this.sheets     = new HashSet<string>(sheets);
                this.serializer = serializer;
                dictionary      = new Dictionary<string, (IOPointer[] pointers, object obj)>();
            }

            public bool TryReadType(IOMetaAttribute type, string name, out object result) {
                result = Activator.CreateInstance(type.Type);
                var sPointers = type.GetSheetPointers(name).ToArray();
                if (!sPointers.TryGetChildren(TryCreateSheet, out var children)) // todo: probably can be merged with p.TryCreateFromChildren()
                    return false;
                result.SetFields(sPointers.Select(p => p.Field.FieldInfo), children);
                if (type.Regions.Count == 0) return true;
            
                if (!sheets.Contains(name)) return false;
                dictionary.Add(type.GetA1Range(name, SheetsIO.FirstCell), (type.GetPointers(V2Int.Zero).ToArray(), result));
                return true;
            }

            bool TryCreateSheet(IOPointer p, out object result) => (p.Rank == p.Field.Rank
                                                                         ? TryReadType(p.Field.Meta, p.Name, out result)
                                                                         : p.TryCreateFromChildren(TryCreateSheet, IOPointer.GetSheetPointers, out result))
                                                                 || p.Optional;

            public bool TryApplyRange(ValueRange range) { // todo: probably can be merged with p.TryCreateFromChildren()
                var (pointers, obj) = dictionary.First(pair => StringComparer.Ordinal.Equals(range.Range.GetSheetName(), pair.Key.GetSheetName())).Value;
                if (!pointers.TryGetChildren(new ReadRangeContext(range, serializer).Delegate, out var children))
                    return false;
                obj.SetFields(pointers.Select(x => x.Field.FieldInfo), children);
                return true;
            }
        }
        
        readonly struct ReadRangeContext
        {
            readonly IList<IList<object>> values;
            readonly IValueSerializer serializer;
            public SheetsIO.ReadObjectDelegate Delegate => TryCreate;

            public ReadRangeContext(ValueRange range, IValueSerializer s) {
                values     = range.Values;
                serializer = s;
            }

            bool TryCreate(IOPointer p, out object result) => (p.IsValue
                                                                   ? TryReadRegion(p, out result)
                                                                   : p.TryCreateFromChildren(TryCreate, IOPointer.GetRegionPointers, out result))
                                                           || p.Optional;

            bool TryReadRegion(IOPointer p, out object result) => (result = values.TryGetElement(p.Pos.X, out var column) && column.TryGetElement(p.Pos.Y, out var cell)
                                                                          ? serializer.Deserialize(p.Field.Types[p.Rank], cell)
                                                                          : null) != null;
        }
        
        readonly struct WriteContext
        {
            public readonly IList<ValueRange> ValueRanges;
            readonly IValueSerializer serializer;

            public WriteContext(IOMetaAttribute type, string name, object obj, IValueSerializer serializer) {
                ValueRanges     = new List<ValueRange>();
                this.serializer = serializer;
                WriteType(type, name, obj);
            }

            void WriteType(IOMetaAttribute type, string name, object obj) {
                obj.ForEachChild(type.GetSheetPointers(name), WriteSheet);
                if (type.Regions.Count == 0) return;

                var sheet = new WriteSheetContext(serializer);
                obj.ForEachChild(type.GetPointers(V2Int.Zero), sheet.WriteRegion);
                ValueRanges.Add(new ValueRange {Values = sheet.Values, MajorDimension = "COLUMNS", Range = type.GetA1Range(name, SheetsIO.FirstCell)});
            }

            void WriteSheet(IOPointer pointer, object obj) {
                if (pointer.Rank == pointer.Field.Rank)
                    WriteType(pointer.Field.Meta, pointer.Name, obj);
                else
                    obj.ForEachChild(IOPointer.GetSheetPointers(pointer), WriteSheet);
            }
        }

        readonly struct WriteSheetContext
        {
            public readonly IList<IList<object>> Values;
            readonly IValueSerializer serializer;

            public WriteSheetContext(IValueSerializer s) {
                Values     = new List<IList<object>>();
                serializer = s;
            }

            public void WriteRegion(IOPointer pointer, object obj) {
                if (pointer.IsValue)
                    WriteValue(obj, pointer.Pos);
                else
                    obj.ForEachChild(IOPointer.GetRegionPointers(pointer), WriteRegion);
            } // todo: highlight elements of free array with color (chessboard-ish order)    with   ((index.X + index.Y) & 1) ? color1 : color2

            void WriteValue(object value, V2Int pos) {
                for (int i = Values.Count; i <= pos.X; i++) 
                    Values.Add(new List<object>());
                for (int i = Values[pos.X].Count; i < pos.Y; i++) 
                    Values[pos.X].Add(null);
                Values[pos.X].Add(serializer.Serialize(value));
            }
        }
    }
    
    static class ExtensionMethods
    {
        static readonly MethodInfo addMethodInfo = typeof(ICollection<>).GetMethod("Add", BindingFlags.Instance | BindingFlags.Public);

        public static IOFieldAttribute GetIOAttribute(this FieldInfo field) {
            var attribute = (IOFieldAttribute) Attribute.GetCustomAttribute(field, typeof(IOFieldAttribute));
            if (attribute != null && !attribute.Initialized)
                attribute.CacheMeta(field);
            return attribute;
        }

        public static IOMetaAttribute GetIOAttribute(this Type type) {
            var attribute = (IOMetaAttribute) Attribute.GetCustomAttribute(type, typeof(IOMetaAttribute));
            if (attribute != null && !attribute.Initialized)
                attribute.CacheMeta(type);
            return attribute;
        }

        public static V2Int GetSize(this IOMetaAttribute meta) => meta?.Size ?? new V2Int(1, 1);

        public static bool TryGetElement<T>(this IList<T> target, int index, out T result) {
            bool exists = (target?.Count ?? 0) > index;
            result = exists ? target[index] : default; // safe, index >= 0
            return exists;
        }

        public static IEnumerable<T> Produce<T>(this T start, int max, Func<T, int, T> func) {
            int i = max;
            var value = start;
            do yield return value; // kind of Enumerable.Aggregate(), but after each step a current value is returned
            while (--i >= 0 && (value = func(value, i)) != null);
        }
        
        public static void ForEachChild(this object parent, IEnumerable<IOPointer> pointers, Action<IOPointer, object> action) {
            using var e = pointers.GetEnumerator();
            while (e.MoveNext() && TryGetChild(parent, e.Current, out var child))
                action.Invoke(e.Current, child);
        }
        
        public static bool TryGetChildren(this IEnumerable<IOPointer> p, SheetsIO.ReadObjectDelegate create, out ArrayList list) {
            list = new ArrayList();
            foreach (var child in p)
                if (create(child, out var childObj) || child.Optional)
                    list.Add(childObj);
                else
                    return false;
            return true;
        }

        public static bool TryCreateFromChildren(this IOPointer p, SheetsIO.ReadObjectDelegate create, Func<IOPointer, IEnumerable<IOPointer>> func, out object result) =>
            (result = func(p).TryGetChildren(create, out var childrenList) || p.IsValidContent(childrenList)
                          ? MakeObject(p, childrenList)
                          : null) != null;

        public static void SetFields(this object parent, IEnumerable<FieldInfo> fields, ArrayList children) {
            foreach (var (f, child) in fields.Zip(children.Cast<object>(), (f, child) => (f, child)))
                f.SetValue(parent, child);
        }

        static object MakeObject(IOPointer p, ArrayList children) {
            if (p.Field.Types[p.Rank].IsArray) 
                return MakeArray(p, children);
            
            var result = Activator.CreateInstance(p.TargetType);
            AddChildrenToObject(p, children, result);
            return result;
        }

        static void AddChildrenToObject(IOPointer p, ArrayList children, object parent) {
            if (p.Rank == p.Field.Rank)
                parent.SetFields(p.Field.Meta.Regions.Select(x => x.FieldInfo), children);
            else if (parent is IList list)
                foreach (var child in children)
                    list.Add(child);
            else SetValues(p, children, parent);
        }

        static void SetValues(IOPointer p, IEnumerable children, object parent) {
            var method = addMethodInfo.MakeGenericMethod(p.Field.Types[p.Rank]);
            foreach (var child in children)
                method.Invoke(parent, new[]{child});
        }

        static object MakeArray(IOPointer p, IList children) {
            var result = (IList)Array.CreateInstance(p.Field.Types[p.Rank + 1], children.Count);
            for (int i = 0; i < children.Count; i++)
                result[i] = children[i];
            return result;
        }

        static bool TryGetChild(object parent, IOPointer p, out object child) {
            if (parent != null && p.Rank == 0) {
                child = p.Field.FieldInfo.GetValue(parent);
                return true;
            }
            if (parent is IList list && list.Count > p.Index) {
                child = list[p.Index];
                return true;
            }
            child = null;
            return !p.IsFreeSize;
        }
    }
    
    static class A1Notation
	{
		const int A1LettersCount = 26;
        const int BigNumber = 999;

		public static string GetA1Range(this IOMetaAttribute type, string sheet, string a2First) =>
            $"'{sheet.Trim()}'!{a2First}:{WriteA1(type.Size + ReadA1(a2First) + new V2Int(-1, -1))}";

        public static string GetSheetName(this string range) => range.Split('!')[0].Replace("''", "'").Trim('\'', ' ');

        static V2Int ReadA1(string a1) => new V2Int(Evaluate(a1.Where(char.IsLetter).Select(char.ToUpperInvariant), '@', A1LettersCount),
                                                    Evaluate(a1.Where(char.IsDigit), '0', 10));

        static string WriteA1(V2Int a1) => (a1.X >= BigNumber ? string.Empty : new string(ToLetters(a1.X)))
                                         + (a1.Y >= BigNumber ? string.Empty : (a1.Y + 1).ToString());

        static char[] ToLetters(int number) => number < A1LettersCount
                                                              ? new[] {(char) ('A' + number)}
                                                              : ToLetters(number / A1LettersCount - 1).Append((char) ('A' + number % A1LettersCount)).ToArray();

        static int Evaluate(IEnumerable<char> digits, char zero, int @base) {
            int result = (int) digits.Reverse().Select((c, i) => (c - zero) * Math.Pow(@base, i)).Sum();
            return result-- > 0 ? result : BigNumber; // In Google Sheets notation, upper boundary of the range may be missing - it means "up to a big number"
        }
    }
    
    readonly struct V2Int
    {
        public static V2Int Zero => new(0, 0);
        public readonly int X, Y;

        public V2Int(int x, int y) {
            X = x;
            Y = y;
        }

        public static V2Int operator +(V2Int a, V2Int b) => new(a.X + b.X, a.Y + b.Y);
        public static V2Int operator *(V2Int a, V2Int b) => new(a.X * b.X, a.Y * b.Y);
		public static V2Int operator *(int a, V2Int b) => new(a * b.X, a * b.Y);
		public static V2Int operator *(V2Int a, int b) => new(a.X * b, a.Y * b);

	}

    readonly struct IOPointer
    {
        public readonly IOFieldAttribute Field;
        public readonly int Rank, Index;
        public readonly V2Int Pos;
        public readonly string Name;

        public bool IsValue => Rank == Field.Rank && Field.Meta is null;
        public bool IsFreeSize => Field.Rank > 0 && Field.ElementsCount.Count == 0;
        public bool Optional => Field.IsOptional || Rank > 0 && Field.ElementsCount.Count > 0;
        public Type TargetType => Field.Types[Rank];
        public IEnumerable<int> ChildIndices => Enumerable.Range(0, Field.MaxCount(Rank));

        public IOPointer(IOFieldAttribute field, int rank, int index, V2Int pos, string name) {
            Field = field;
            Rank  = rank;
            Index = index;
            Pos   = pos;
            Name  = name;
        }

        public bool IsValidContent(ArrayList children) => Rank < Field.Rank 
                                                       && (children.Count > 0 || Rank == 0);

        public static IEnumerable<IOPointer> GetRegionPointers(IOPointer p) => p.Rank == p.Field.Rank
                ? p.Field.Meta.GetPointers(p.Pos)
                : p.ChildIndices.Select(i => new IOPointer(p.Field, p.Rank + 1, i, p.Pos + p.Field.Offsets[p.Rank + 1] * i, ""));
        public static IEnumerable<IOPointer> GetSheetPointers(IOPointer p) => p.Rank == p.Field.Rank
                ? p.Field.Meta.GetSheetPointers(p.Name)
                : p.ChildIndices.Select(i => new IOPointer(p.Field, p.Rank + 1, i, V2Int.Zero, $"{p.Name} {i + 1}"));
    }
}
