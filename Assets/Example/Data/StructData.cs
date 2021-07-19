using System;
using SheetsIO;

namespace Example
{
    [Serializable, IOMeta] public class StructData
    {
        [IOField, IOPlacement (SortOrder = 0)] public string key;
        [IOField] public string value;
    }
}
