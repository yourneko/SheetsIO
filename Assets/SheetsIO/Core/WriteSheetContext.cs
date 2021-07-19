using System.Collections.Generic;

namespace SheetsIO
{
    readonly struct WriteSheetContext
    {
        public readonly IList<IList<object>> Values;
        readonly IValueSerializer serializer;

        public WriteSheetContext(IValueSerializer s) {
            Values     = new List<IList<object>>();
            serializer = s;
        }

        public void WriteObject(IOPointer pointer, object obj) {
            if (pointer.IsValue)
                WriteValue(obj, pointer.Pos);
            else
                obj.ForEachChild(IOPointer.GetChildren(pointer), WriteObject);
        } // todo: highlight elements of free array with color (chessboard-ish order)    with   ((index.X + index.Y) & 1) ? color1 : color2

        void WriteValue(object value, V2Int pos) {
            for (int i = Values.Count; i <= pos.X; i++) Values.Add(new List<object>());
            for (int i = Values[pos.X].Count; i < pos.Y; i++) Values[pos.X].Add(null);
            Values[pos.X].Add(serializer.Serialize(value));
        }
    }
}
