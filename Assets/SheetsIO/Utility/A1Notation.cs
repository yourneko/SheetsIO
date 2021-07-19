using System;
using System.Collections.Generic;
using System.Linq;

namespace SheetsIO
{
    static class A1Notation
    {
        public static string GetA1Range(this IOMetaAttribute type, string sheet, string a2First) =>
            $"'{sheet.Trim()}'!{a2First}:{WriteA1(type.Size.Add(ReadA1(a2First)).Add(new V2Int(-1, -1)))}";

        public static string GetSheetName(this string range) => range.Split('!')[0].Replace("''", "'").Trim('\'', ' ');

        static V2Int ReadA1(string a1) => new V2Int(Evaluate(a1.Where(char.IsLetter).Select(char.ToUpperInvariant), '@', SheetsIO.A1LettersCount),
                                                    Evaluate(a1.Where(char.IsDigit), '0', 10));

        static string WriteA1(V2Int a1) => (a1.X >= 999 ? string.Empty : new string(ToLetters(a1.X).ToArray()))
                                         + (a1.Y >= 999 ? string.Empty : (a1.Y + 1).ToString());

        static IEnumerable<char> ToLetters(int number) => number < SheetsIO.A1LettersCount
                                                              ? new[] {(char) ('A' + number)}
                                                              : ToLetters(number / SheetsIO.A1LettersCount - 1).Append((char) ('A' + number % SheetsIO.A1LettersCount));

        static int Evaluate(IEnumerable<char> digits, char zero, int @base) {
            int result = (int) digits.Reverse().Select((c, i) => (c - zero) * Math.Pow(@base, i)).Sum();
            return result-- > 0 ? result : 999; // In Google Sheets notation, upper boundary of the range may be missing - it means "up to a big number"
        }
    }
}
