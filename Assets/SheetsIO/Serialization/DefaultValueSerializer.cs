using System;
using System.Collections.Generic;
using System.Globalization;

namespace SheetsIO
{
    class DefaultValueSerializer : IValueSerializer
    {
        static readonly IReadOnlyDictionary<Type, Func<string, object>> @switch = new Dictionary<Type, Func<string, object>> {
             {typeof(string), s => s},
             {typeof(int), value => int.TryParse(value, out int i) ? i : 0}, 
             {typeof(float), value => float.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out float f) ? f : 0f}, 
             {typeof(double), value => double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double d) ? d : 0},
             {typeof(bool), value => StringComparer.OrdinalIgnoreCase.Equals(value, "true")}, 
             {typeof(DateTime), s => DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, out var a) ? a : DateTime.MinValue}
         };

        public string Serialize(object target) => target switch {
                                                      int i      => i.ToString(),
                                                      bool b     => b ? "true" : "false",
                                                      string s   => s,
                                                      float f    => f.ToString(CultureInfo.InvariantCulture),
                                                      DateTime a => a.ToString(CultureInfo.InvariantCulture),
                                                      null       => string.Empty,
                                                      _          => target.ToString()
                                                  };

        public object Deserialize(Type type, object value) => @switch.TryGetValue(type, out var func)
                                                                  ? func.Invoke((string) value ?? string.Empty)
                                                                  : throw new Exception($"Value type {type.Name} is not supported by default serializer.");
    }
}
