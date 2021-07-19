using System;
using SheetsIO;
using UnityEngine;

namespace Example
{
    [Serializable, IOMeta ("Group")]
    public class SuperclassData
    {
        [SerializeField, IOField(2)] ExampleData[] data;
        [SerializeField, IOField, IOPlacement(SortOrder = 0)] string title;
    }
}
