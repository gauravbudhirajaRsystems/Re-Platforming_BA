// © Copyright 2018 Levit & James, Inc.

namespace LevitJames.MSOffice.MSWord
{
    /// <summary>
    ///     Enum used to control how and when <see cref="WordExtensions.SelectionChanged">SelectionChange</see> events should
    ///     be raised
    /// </summary>
    
    public enum SelectionChangeOption
    {
        /// <summary>
        ///     Stops the raising of <see cref="WordExtensions.SelectionChanged">SelectionChange</see> Events.
        /// </summary>
        
        None = 0x0,

        /// <summary>
        ///     Use Words built-in Selection event handler.
        /// </summary>
        NativeSelection = 1,

        /// <summary>
        ///     Raises the <see cref="WordExtensions.SelectionChanged">SelectionChange</see> event when the cursor is moved by the
        ///     cursor keys or the mouse in the Word Document or any character is typed in the document.
        /// </summary>
        
        CustomSelection = 2
    }
}