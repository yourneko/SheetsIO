using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace SheetsIO
{
    /// <summary>Contains metadata of fields.</summary>
    [AttributeUsage(AttributeTargets.Field)]
    public sealed class IOFieldAttribute : Attribute
    {
        /// <summary>NULL is a valid value for this field.</summary>
        public bool IsOptional;

        internal readonly IReadOnlyList<int> ElementsCount;
        internal V2Int PosInType;

        internal bool Initialized { get; private set; }
        internal IOMetaAttribute Meta { get; private set; }
        internal FieldInfo FieldInfo { get; private set; }
        internal int Rank { get; private set; }
        internal IReadOnlyList<Type> Types { get; private set; }
        internal IReadOnlyList<V2Int> Sizes { get; private set; }

        /// <summary>Map this field to Google Spreadsheets.</summary>
        /// <param name="elementsCount">Fixed number of elements for each rank of array.</param>
        public IOFieldAttribute(params int[] elementsCount) {
            ElementsCount = elementsCount;
        }

        internal void CacheMeta(FieldInfo field) {
            Initialized = true;
            FieldInfo   = field;
            Types       = field.FieldType.RepeatAggregated(Math.Max(ElementsCount.Count, 2), NextType).ToArray();
            Rank        = Types.Count - 1;
            Meta        = Types[Rank].GetIOAttribute();
            Sizes       = (Meta?.Size ?? new V2Int(1, 1)).RepeatAggregated(Rank, NextRankSize).Reverse().ToArray();
            if (Rank == 0) return;
            
            if (!string.IsNullOrEmpty(Meta?.SheetName) && FieldInfo.GetCustomAttribute<IOPlacementAttribute>() != null)
                throw new Exception($"Remove MapPlacement attribute from a field {FieldInfo.Name}. " +
                                    $"Instances of type {Types[Rank].Name} can't be arranged, because they are placed on separate sheets.");
        }

        internal int MaxCount(int rank) => rank < ElementsCount.Count ? ElementsCount[rank] : SheetsIO.MaxArrayElements;
        static Type NextType(Type type, int rank) => type.GetTypeInfo().GetInterfaces().FirstOrDefault(IsCollection)?.GetGenericArguments()[0];
        static bool IsCollection(Type type) => type != typeof(string) && type.IsGenericType && type.GetGenericTypeDefinition() == typeof(ICollection<>);
        V2Int NextRankSize(V2Int v2, int rank) => v2.Scale((rank & 1) == 0 ? MaxCount(rank - 1) : 1,
                                                           (rank & 1) == 0 ? 1 : MaxCount(rank - 1));
    }
}
