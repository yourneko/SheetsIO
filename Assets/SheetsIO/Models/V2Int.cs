namespace SheetsIO
{
    readonly struct V2Int
    {
        public static V2Int Zero => new V2Int(0, 0);
        public readonly int X, Y;

        public V2Int(int x, int y) {
            X = x;
            Y = y;
        }

        public V2Int Add(V2Int other) => new V2Int(X + other.X, Y + other.Y);
        public V2Int Scale(int x, int y) => new V2Int(X * x, Y * y);
        public V2Int Scale(int n) => Scale(n, n);
    }
}
