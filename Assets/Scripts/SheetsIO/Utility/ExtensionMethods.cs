using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using UnityEngine;

namespace SheetsIO
{
    static class ExtensionMethods
    {
        static readonly MethodInfo addMethodInfo = typeof(ICollection<>).GetMethod("Add", BindingFlags.Instance | BindingFlags.Public);

        public static IOFieldAttribute GetIOAttribute(this FieldInfo field) {
            var attribute = (IOFieldAttribute) Attribute.GetCustomAttribute(field, typeof(IOFieldAttribute));
            if (attribute != null && !attribute.Initialized)
                attribute.CacheMeta(field);
            return attribute;
        }

        public static IOMetaAttribute GetIOAttribute(this Type type) {
            var attribute = (IOMetaAttribute) Attribute.GetCustomAttribute(type, typeof(IOMetaAttribute));
            if (attribute != null && !attribute.Initialized)
                attribute.CacheMeta(type);
            return attribute;
        }

        public static bool TryGetElement<T>(this IList<T> target, int index, out T result) {
            bool exists = (target?.Count ?? 0) > index;
            result = exists ? target[index] : default; // safe, index >= 0
            return exists;
        }

        public static IEnumerable<T> RepeatAggregated<T>(this T start, int max, Func<T, int, T> func) {
            int i = max;
            var value = start;
            do yield return value; // kind of Enumerable.Aggregate(), but after each step returns a current value
            while (--i >= 0 && (value = func(value, i)) != null);
        }
        
        public static void ForEachChild(this object parent, IEnumerable<IOPointer> pointers, Action<IOPointer, object> action) {
            using var e = pointers.GetEnumerator();
            while (e.MoveNext() && parent.TryGetChild(e.Current, out var child))
                action.Invoke(e.Current, child);
        }

        static bool TryGetChild(this object parent, IOPointer p, out object child) {
            Debug.Log(p);
            if (parent != null && p.Rank == 0) {
                child = p.Field.FieldInfo.GetValue(parent);
                return true;
            }
            if (parent is IList list && list.Count > p.Index) {
                child = list[p.Index];
                return true;
            }
            child = null;
            return !p.IsFreeSize;
        }
        
        public static bool TryGetChildren(this IEnumerable<IOPointer> p, SheetsIO.ReadObjectDelegate create, out ArrayList list) {
            list = new ArrayList();
            foreach (var child in p)
                if (create(child, out var childObj) || child.Optional)
                    list.Add(childObj);
                else
                    return false;
            return true;
        }
        
        public static object MakeObject(this IOPointer p, ArrayList children) {
            if (p.Field.Types[p.Rank].IsArray) return children.ToArray();
            
            var result = Activator.CreateInstance(p.TargetType);
            if (p.Rank == p.Field.Rank)
                result.SetFields(p.Field.Meta.Regions.Select(x => x.FieldInfo), children);
            else if (result is IList list)
                foreach (var child in children)
                    list.Add(child);
            else SetFieldValues(p, children, result);
            return result;
        }

        public static void SetFields(this object parent, IEnumerable<FieldInfo> fields, ArrayList children) {
            foreach (var (f, child) in fields.Zip(children.Cast<object>(), (f, child) => (f, child)))
                f.SetValue(parent, child);
        }

        static void SetFieldValues(IOPointer p, IEnumerable children, object parent) {
            var method = addMethodInfo.MakeGenericMethod(p.Field.Types[p.Rank]);
            foreach (var child in children)
                method.Invoke(parent, new[]{child});
        }
    }
}
