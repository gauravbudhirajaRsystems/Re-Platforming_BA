using System;
using Microsoft.Office.Interop.Word;

namespace LevitJames.MSOffice.MSWord
{
	public class WordSelectionEventArgs : EventArgs
	{
		internal WordSelectionEventArgs(bool isNativeSelection, int lastSelectionChangeTick, Range range)
		{
			IsNativeSelection = isNativeSelection;
			LastChangeTickCount = lastSelectionChangeTick;
			Range = range;
		}

		public bool IsNativeSelection { get; }

		public int LastChangeTickCount { get; }


		public Range Range { get; }
	}
}