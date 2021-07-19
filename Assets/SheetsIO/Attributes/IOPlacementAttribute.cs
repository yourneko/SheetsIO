using System;

namespace SheetsIO
{
    /// <summary>Adjusts the placement of target field content.</summary>
    /// <remarks>This attribute shouldn't be placed on fields of whole-sheet-sized type.</remarks>
    [AttributeUsage(AttributeTargets.Field)]
    public sealed class IOPlacementAttribute : Attribute
    {
        /// <summary>Fields are processed in ascending order. Default value of SortOrder is 1000.</summary>
        public int SortOrder;
    }
}
