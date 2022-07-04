// © Copyright 2018 Levit & James, Inc.

using System;

namespace LevitJames.MSOffice.MSWord
{
    /// <summary>
    ///     The event arguments used by the WordExtensions.ShortcutKeyPressed Event.
    /// </summary>
    
    public class OfficeShortcutKeyPressedEventArgs : EventArgs
    {
        /// <summary>
        ///     Creates a new instance of the OfficeShortcutKeyPressedEventArgs type.
        /// </summary>
        /// <param name="shortcutKey"></param>
        /// <param name="isKeyUp"></param>
        /// <param name="keyRepeatCount"></param>
        public OfficeShortcutKeyPressedEventArgs(OfficeShortcutKey shortcutKey, bool isKeyUp, int keyRepeatCount)
        {
            Shortcut = shortcutKey;
            IsKeyUp = isKeyUp;
            RepeatCount = keyRepeatCount;
        }

        public bool Handled { get; set; }

        /// <summary>
        ///     Returns true if this is the first time this shortcut has been pressed since the last key up
        /// </summary>
        
        
        /// <remarks>This can be used so that a Shortcut key is not repeatedly executed when the key combination is held down.</remarks>
        public bool IsFirst => !Shortcut.IsRepeat;

        /// <summary>
        ///     The number of times the keystroke is repeated as a result of the user holding down the key
        /// </summary>
        
        
        
        public int RepeatCount { get; }

        /// <summary>
        ///     Returns false if the key is currently down and true if it is being released.
        /// </summary>
        
        
        
        public bool IsKeyUp { get; }

        /// <summary>
        ///     Returns the OfficeShortcutKey instance associated with the key combination currently pressed.
        /// </summary>
        
        
        
        public OfficeShortcutKey Shortcut { get; }
    }
}