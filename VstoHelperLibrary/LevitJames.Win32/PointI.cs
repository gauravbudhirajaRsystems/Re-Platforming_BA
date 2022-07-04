// © Copyright 2020 Levit & James, Inc.

using System.Runtime.InteropServices;

namespace LevitJames.Win32
{

    [StructLayout(LayoutKind.Sequential)]
    public struct PointI
    {

        public PointI(int x, int y)
        {
            X = x;
            Y = y;
        }
 
        public int X { get; set; }
        public int Y { get; set; }

        public bool Equals(PointI other)
        {
            return X == other.X && Y == other.Y;
        }

        public override bool Equals(object obj)
        {
            if ((obj is PointI pt))
                return Equals(pt);
            return false;
        }

        /// <inheritdoc />
        public override int GetHashCode()
        {
            unchecked
            {
                // ReSharper disable NonReadonlyMemberInGetHashCode
                return (X * 397) ^ Y;
                // ReSharper restore NonReadonlyMemberInGetHashCode
            }
        }

        public static bool operator ==(PointI left, PointI right) => left.Equals(right);
 
        public static bool operator !=(PointI left, PointI right) => !left.Equals(right);

        public override string ToString()
        {
            return $"X={X}, Y={Y}";
        }
    }
}