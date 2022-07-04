// © Copyright 2020 Levit & James, Inc.

using System.Runtime.InteropServices;

namespace LevitJames.Win32
{
	[StructLayout(LayoutKind.Sequential)]
    public struct SizeI
    {

        public SizeI(int width, int height)
        {
            Width = width;
            Height = height;
        }


        public int Width { get; set; }
        public int Height { get; set; }

        public bool Equals(SizeI other)
        {
            return Width == other.Width && Height == other.Height;
        }

        public override bool Equals(object obj)
        {
            if (obj is SizeI sz)
                return Equals(sz);

            return false;
            
        }
        public static bool operator ==(SizeI left, SizeI right) => left.Equals(right);
 
        public static bool operator !=(SizeI left, SizeI right) => !left.Equals(right);

        public override int GetHashCode()
        {
            unchecked
            {
                // ReSharper disable NonReadonlyMemberInGetHashCode
                return (Width * 397) ^ Height;
                // ReSharper restore NonReadonlyMemberInGetHashCode
            }
        }
        public override string ToString()
        {
            return $"Width={Width}, Height={Height}";
        }
    }
}