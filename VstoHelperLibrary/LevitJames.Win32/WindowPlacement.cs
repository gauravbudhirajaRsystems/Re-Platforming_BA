using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
 
namespace LevitJames.Win32
{

	[TypeConverter(typeof(WindowPlacementConverter))]
    [Serializable]
    [StructLayout(LayoutKind.Sequential)]
	public sealed class WindowPlacement 
	{
		private int _length = Marshal.SizeOf(typeof(WindowPlacement));
		public int _flags;
		public int _showCmd;
		private PointI _minPosition;
		private PointI _maxPosition;
		private RectangleI _normalPosition;

		public int Flags { get => _flags; set => _flags = value; }
		public int ShowCmd { get => _showCmd; set => _showCmd = value; }

		
		public PointI MinPosition { get => _minPosition; set => _minPosition = value; }
		public PointI MaxPosition { get => _maxPosition; set => _maxPosition = value; }
		public RectangleI NormalPosition { get => _normalPosition; set => _normalPosition = value; }

 
    }
}