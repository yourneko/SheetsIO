using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace SheetsIO
{
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
    }
}
