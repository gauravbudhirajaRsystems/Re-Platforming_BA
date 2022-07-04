using System;

namespace LevitJames.MSOffice.MSWord
{
	public interface IWordSelectionChange
	{

		event EventHandler<WordSelectionEventArgs> SelectionChanged;

		int Resolution { get; set; }
		SelectionChangeOption Options { get; set; }

		/// <summary>
		///     Causes the selection change to update immediately and raise it's selection change event if required, instead of
		///     waiting for idle time.
		/// </summary>
		/// <param name="raiseSelectionChangeEvent">true (default) raises teh selection change if the selection has changed.</param>
		void Update(bool raiseSelectionChangeEvent);
	}
}