using System.Collections.Generic;
using Google.Apis.Sheets.v4.Data;

namespace SheetsIO
{
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
            obj.ForEachChild(type.GetSheetPointers(name), WriteSheetObject);
            if (type.Regions.Count == 0) return;

            var sheet = new WriteSheetContext(serializer);
            obj.ForEachChild(type.GetPointers(V2Int.Zero), sheet.WriteObject);
            ValueRanges.Add(new ValueRange {Values = sheet.Values, MajorDimension = "COLUMNS", Range = type.GetA1Range(name, SheetsIO.FirstCell)});
        }

        void WriteSheetObject(IOPointer pointer, object obj) {
            if (pointer.Rank == pointer.Field.Rank)
                WriteType(pointer.Field.Meta, pointer.Name, obj);
            else
                obj.ForEachChild(IOPointer.GetChildrenSheets(pointer), WriteSheetObject);
        }
    }
}
