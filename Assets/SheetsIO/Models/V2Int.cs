using System.Numerics;

namespace SheetsIO
{
    readonly struct V2Int
    {
        public static V2Int Zero => new(0, 0);
        public readonly int X, Y;

        public V2Int(int x, int y) {
            X = x;
            Y = y;
        }

        public static V2Int operator +(V2Int a, V2Int b) => new(a.X + b.X, a.Y + b.Y);
        public static V2Int operator *(V2Int a, V2Int b) => new(a.X * b.X, a.Y * b.Y);
		public static V2Int operator *(int a, V2Int b) => new(a * b.X, a * b.Y);
		public static V2Int operator *(V2Int a, int b) => new(a.X * b, a.Y * b);

	}
}
