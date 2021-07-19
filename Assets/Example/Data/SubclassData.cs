using System;
using SheetsIO;
using UnityEngine;

namespace Example
{
    [Serializable, IOMeta] public class SubclassData
    {
        [IOField, SerializeField] string name;
        [IOField, SerializeField] int subvalue;
        [IOField(2), SerializeField] StructData[] keyValueArray;
    }
}
