using System;
using System.Collections.Generic;
using System.Linq;
using Google.Apis.Sheets.v4.Data;

namespace SheetsIO
{
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

        public bool ReadType(IOMetaAttribute type, string name, out object result) {
            result = Activator.CreateInstance(type.Type);
            var sPointers = type.GetSheetPointers(name).ToArray();
            if (!sPointers.TryGetChildren(ReadSheetObject, out var children))
                return false;
            result.SetFields(sPointers.Select(x => x.Field.FieldInfo), children);
            if (type.Regions.Count == 0) return true;
            
            if (!sheets.Contains(name)) return false;
            dictionary.Add(type.GetA1Range(name, SheetsIO.FirstCell), (type.GetPointers(V2Int.Zero).ToArray(), result));
            return true;
        }

        bool ReadSheetObject(IOPointer p, out object obj) {
            return (p.Rank == p.Field.Rank
                        ? ReadType(p.Field.Meta, p.Name, out obj)
                        : (obj = IOPointer.GetChildrenSheets(p).TryGetChildren(ReadSheetObject, out var l) ? p.MakeObject(l) : null) != null)
                || p.Optional;
        }

        public bool TryApplyRange(ValueRange range) {
            var (pointers, obj) = dictionary.First(pair => StringComparer.Ordinal.Equals(range.Range.GetSheetName(), pair.Key.GetSheetName())).Value;
            bool result = pointers.TryGetChildren(new ReadRangeContext(range, serializer).Delegate, out var children);
            if (result)
                obj.SetFields(pointers.Select(x => x.Field.FieldInfo), children);
            return result;
        }
    }
}
