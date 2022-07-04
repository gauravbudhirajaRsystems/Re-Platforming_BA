using System;

namespace LevitJames.Core
{
	/// <summary>
	/// A class for providing a Text/Value pair
	/// </summary>
	/// <typeparam name="T"></typeparam>
	public class TextValuePair<T>
	{

		/// <summary>
		/// The value of the item
		/// </summary>
		public T Value { get; set; }
		
		/// <summary>
		/// A string describing the value
		/// </summary>
		public string Text { get; set; }

        /// <summary>
        /// Creates a new instance of the TextValuePair class
        /// </summary>
        /// <param name="text"></param>
        /// <param name="value"></param>
        public TextValuePair(string text, T value)
		{
			Value = value;
			Text = text;
		}
 

	}
}