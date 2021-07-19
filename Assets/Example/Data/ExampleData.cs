using System;
using System.Collections.Generic;
using SheetsIO;
using UnityEngine;

namespace Example
{
    [Serializable, IOMeta ("ExampleSheet")]
    public class ExampleData
    {
        [IOField, IOPlacement (SortOrder = 5), SerializeField] string[] ss1;
        [IOField, IOPlacement (SortOrder = 3), SerializeField] List<string> ss2; 
        [IOField, IOPlacement (SortOrder = 2), SerializeField] StructData keyValueArray;
        [IOField, IOPlacement (SortOrder = 0), SerializeField] string s;
        [IOField(IsOptional = true), IOPlacement (SortOrder = 4), SerializeField] int i;
        [IOField, IOPlacement (SortOrder = 6), SerializeField] List<SubclassData> subList;

        [IOField(3, 3), IOPlacement (SortOrder = 1)] int[][] grid = {new[] {1, 2, 3}, new[] {4, 5, 6,}, new[] {7, 8, 9}};
    }
}
