using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

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

        public static V2Int GetSize(this IOMetaAttribute meta) => meta?.Size ?? new V2Int(1, 1);

        public static bool TryGetElement<T>(this IList<T> target, int index, out T result) {
            bool exists = (target?.Count ?? 0) > index;
            result = exists ? target[index] : default; // safe, index >= 0
            return exists;
        }

        public static IEnumerable<T> RepeatAggregated<T>(this T start, int max, Func<T, int, T> func) {
            int i = max;
            var value = start;
            do yield return value; // kind of Enumerable.Aggregate(), but after each step a current value is returned
            while (--i >= 0 && (value = func(value, i)) != null);
        }
        
        public static void ForEachChild(this object parent, IEnumerable<IOPointer> pointers, Action<IOPointer, object> action) {
            using var e = pointers.GetEnumerator();
            while (e.MoveNext() && TryGetChild(parent, e.Current, out var child))
                action.Invoke(e.Current, child);
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

        public static bool TryCreateFromChildren(this IOPointer p, SheetsIO.ReadObjectDelegate create, Func<IOPointer, IEnumerable<IOPointer>> func, out object result) =>
            (result = func(p).TryGetChildren(create, out var childrenList) || p.IsValidContent(childrenList)
                          ? MakeObject(p, childrenList)
                          : null) != null;

        public static void SetFields(this object parent, IEnumerable<FieldInfo> fields, ArrayList children) {
            foreach (var (f, child) in fields.Zip(children.Cast<object>(), (f, child) => (f, child)))
                f.SetValue(parent, child);
        }

        static object MakeObject(IOPointer p, ArrayList children) {
            if (p.Field.Types[p.Rank].IsArray) 
                return MakeArray(p, children);
            
            var result = Activator.CreateInstance(p.TargetType);
            AddChildrenToObject(p, children, result);
            return result;
        }

        static void AddChildrenToObject(IOPointer p, ArrayList children, object parent) {
            if (p.Rank == p.Field.Rank)
                parent.SetFields(p.Field.Meta.Regions.Select(x => x.FieldInfo), children);
            else if (parent is IList list)
                foreach (var child in children)
                    list.Add(child);
            else SetValues(p, children, parent);
        }

        static void SetValues(IOPointer p, IEnumerable children, object parent) {
            var method = addMethodInfo.MakeGenericMethod(p.Field.Types[p.Rank]);
            foreach (var child in children)
                method.Invoke(parent, new[]{child});
        }

        static object MakeArray(IOPointer p, IList children) {
            var result = (IList)Array.CreateInstance(p.Field.Types[p.Rank + 1], children.Count);
            for (int i = 0; i < children.Count; i++)
                result[i] = children[i];
            return result;
        }

        static bool TryGetChild(object parent, IOPointer p, out object child) {
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
    }
}
