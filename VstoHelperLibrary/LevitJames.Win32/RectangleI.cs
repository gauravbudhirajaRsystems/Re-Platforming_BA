using System;
using System.Runtime.InteropServices;

namespace LevitJames.Win32
{
    [StructLayout(LayoutKind.Sequential)]
    public struct RectangleI
    {

        public int Left { get; set; }
        public int Top { get; set; }
        public int Right { get; set; }
        public int Bottom { get; set; }

        // ReSharper disable once InconsistentNaming
        public static RectangleI FromXYWH(int x, int y, int width, int height)
        {
            return new RectangleI {Left = x, Top = y, Width = width, Height = height};
        }

        public RectangleI(int x, int y, int right, int bottom)
        {
            Left = x;
            Top = y;

            Right = right;
            Bottom = bottom;
        }

        public void Move(int left, int top)
        {
            var length = Width;
            Left = left;
            Width = length;
            length = Height;
            Top = top;
            Height = length;
        }

        public bool IsEmpty => Left == 0 && Top == 0 && Right == 0 && Bottom == 0;

        /// <summary>
        /// Gets sets the X position of the rectangle
        /// The Width and right position stay unchanged. 
        /// </summary>
        public int X
        {
            get => Left;
            set
            {
                var width = Width;
                Left = value;
                Width = width;
            }
        }

        /// <summary>
        /// Gets sets the X position of the rectangle
        /// The Height and Bottom position stay unchanged. 
        /// </summary>
        public int Y
        {
            get => Top;
            set
            {
                var height = Height;
                Top = value;
                Height = height;
            }
        }

        public int Width
        {
            get => Right - Left;
            set => Right = Left + value;
        }

        public int Height
        {
            get => Bottom - Top;
            set => Bottom = Top + value;
        }

        public SizeI Size
        {
            get => new SizeI(Right - Left, Bottom - Top);
            set
            {
                Right = Left + value.Width;
                Bottom = Top + value.Height;
            }
        }

        public PointI Location
        {
            get => new PointI(Left, Top);
            set
            {
                Left = value.X;
                Top = value.Y;
            }
        }

        public void Offset(int x, int y)
        {
            if (x != 0)
            {
                Left += x;
                Right += x;
            }

            if (y != 0)
            {
                Top += y;
                Bottom += y;
            }
        }

        public void Scale(double scaleFactor)
        {
            if (scaleFactor == 1)
                return;

            var length = Width;
            Left = (int)Math.Round(Left * scaleFactor);
            Width = (int)Math.Round(length * scaleFactor);
            
            length = Height;
            Top = (int)Math.Round(Top * scaleFactor);
            Height = (int)Math.Round(length * scaleFactor);
        }

        public void Inflate(int x, int y)
        {
            Left -= x;
            Top -= x;
            Right += x;
            Bottom += y;
        }

        public RectangleI Center(int width, int height)
        {
            return RectangleI.FromXYWH(X + (Width - width) / 2, Y + (Height - height) / 2, width, height);
        }

        public RectangleI FlipAxis(bool vertical)
        {
            return vertical ? new RectangleI(Top, Left, Bottom, Right) : this;
        }

        public static RectangleI Intersect(RectangleI rect1, RectangleI rect2)
        {
            //Get the largest x location
            var x = Math.Max(rect1.X, rect2.X);
            //Get the largest y location
            var y = Math.Max(rect1.Y, rect2.Y);

            var minRightPos = Math.Min(rect1.Right, rect2.Right);
            var minBottomPos = Math.Min(rect1.Bottom, rect2.Bottom);

            return (((minRightPos < x) || (minBottomPos < y)) ? new RectangleI() : new RectangleI(x, y, minRightPos - x, minBottomPos - y));

        }

        public static RectangleI Union(RectangleI rect1, RectangleI rect2)
        {
            //Get the largest x location
            var x = Math.Min(rect1.X, rect2.X);
            //Get the largest y location
            var y = Math.Min(rect1.Y, rect2.Y);

            return new RectangleI(x, y, Math.Max(rect1.Right, rect2.Right), Math.Max(rect1.Bottom, rect2.Bottom));

        }

        public bool Contains(int x, int y) => Left <= x && x <= Right && Top <= y && y <= Bottom;
        public bool Contains(PointI pt) => Contains(pt.X, pt.Y);

        public bool Contains(RectangleI rect) => Left <= rect.Left && rect.Right <= Right && Top <= rect.Top && rect.Bottom <= Bottom;

        public bool Equals(RectangleI rect) => rect.Left == Left && rect.Top == Top && rect.Right == Right && Bottom == rect.Bottom;

        public override bool Equals(object obj)
        {
            if (obj is RectangleI rc)
                return Equals(rc);

            return false;

        }

        public static bool operator ==(RectangleI left, RectangleI right) => left.Equals(right);


        public static bool operator !=(RectangleI left, RectangleI right) => !left.Equals(right);

        // ReSharper disable NonReadonlyMemberInGetHashCode
        public override int GetHashCode() => Left ^ ((Top << 13) | (Top >> 0x13)) ^ ((Right << 0x1a) | (Right >> 6)) ^ ((Bottom << 7) | (Bottom >> 0x19));
        // ReSharper restore NonReadonlyMemberInGetHashCode

        /// <inheritdoc />
        public override string ToString()
        {
            return $"X={Left}, Y={Top}, Width={Width}, Height={Height}, Right={Right}, Bottom={Bottom}";
        }
    }

}
