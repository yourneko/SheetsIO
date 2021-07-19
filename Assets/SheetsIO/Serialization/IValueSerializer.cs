using System;

namespace SheetsIO
{
    /// <summary>Provides a strategy of value serialization.</summary>
    public interface IValueSerializer
    {
        /// <summary>Represents object as a string value.</summary>
        /// <param name="value">User's object</param>
        /// <returns>Culture-invariant formatting for simple types.</returns>
        string Serialize(object value);

        /// <summary>Creates an object of the given type out of the string.</summary>
        /// <param name="type">Type of object.</param>
        /// <param name="text">A source string.</param>
        /// <returns>Result object.</returns>
        /// <exception cref="NotSupportedException">Serialization of a given type is not implemented.</exception>
        object Deserialize(Type type, object text);
    }
}
