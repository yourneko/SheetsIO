using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace SheetsIO
{
    /// <summary>Contains metadata of classes.</summary>
    [AttributeUsage(AttributeTargets.Class)]
    public sealed class IOMetaAttribute : Attribute
    {
        internal readonly string SheetName;

        internal bool Initialized { get; private set; }
        internal Type Type;
        internal V2Int Size { get; private set; }
        internal IReadOnlyList<IOFieldAttribute> Regions { get; private set; }
        IReadOnlyList<IOFieldAttribute> Sheets { get; set; }

        /// <summary>Map this class to Google Spreadsheets.</summary>
        /// <param name="sheetName">Types with a sheet name always occupy the whole sheet.</param>
        /// <remarks>Avoid loops in hierarchy of types. It will cause the stack overflow.</remarks>
        public IOMetaAttribute(string sheetName = null) {
            SheetName = sheetName;
        }

        internal void CacheMeta(Type type) {
            Initialized = true;
            Type        = type;
            var allFields = type.GetFields(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
                                .Select(field => field.GetIOAttribute()).Where(x => x != null).ToArray();
            if (allFields.Length == 0) 
                throw new Exception($"Class {type.Name} has no IOField's!");
            Sheets  = allFields.Where(x => !string.IsNullOrEmpty(x.Meta?.SheetName ?? string.Empty)).ToArray();
            Regions = allFields.Where(x => x.Meta is null || string.IsNullOrEmpty(x.Meta.SheetName)).OrderBy(x => x.SortOrder).ToArray();
            Size    = V2Int.Zero;
            foreach (var f in Regions) {
                f.PosInType = new V2Int(Size.X, 0);
                Size        = new V2Int(Size.X + f.Sizes[0].X, Math.Max(Size.Y, f.Sizes[0].Y));
            }
        }

        internal string AppendNamePart(string parentName) => $"{parentName} {SheetName}".Trim();
        internal IEnumerable<IOPointer> GetSheetPointers(string name) => Sheets.Select((f, i) => new IOPointer(f, 0, i, V2Int.Zero, f.Meta.AppendNamePart(name)));
        internal IEnumerable<IOPointer> GetPointers(V2Int pos) => Regions.Select((f, i) => new IOPointer(f, 0, i, pos.Add(f.PosInType), ""));
    }
}
