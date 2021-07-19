using System.Collections.Generic;
using Google.Apis.Sheets.v4.Data;

namespace SheetsIO
{
    readonly struct ReadRangeContext
    {
        readonly IList<IList<object>> values;
        readonly IValueSerializer serializer;
        public SheetsIO.ReadObjectDelegate Delegate => TryCreate;

        public ReadRangeContext(ValueRange range, IValueSerializer s) {
            values     = range.Values;
            serializer = s;
        }

        bool TryCreate(IOPointer p, out object result) =>
            (p.IsValue
                 ? TryRead(p, out result)
                 : (result = IOPointer.GetChildren(p).TryGetChildren(Delegate, out var l) || p.Rank < p.Field.Rank && (l.Count > 0 || p.Rank == 0)
                                 ? p.MakeObject(l)
                                 : null) != null) || p.Optional;

        bool TryRead(IOPointer p, out object result) => (result = values.TryGetElement(p.Pos.X, out var column) && column.TryGetElement(p.Pos.Y, out var cell)
                                                                      ? serializer.Deserialize(p.Field.Types[p.Rank], cell)
                                                                      : null) != null;
        
    }
}
