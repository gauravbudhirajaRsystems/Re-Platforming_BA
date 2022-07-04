// © Copyright 2018 Levit & James, Inc.

using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using LevitJames.Core;
using LevitJames.Core.Diagnostics;
using LevitJames.Libraries.Hooks;


namespace LevitJames.MSOffice.MSWord
{
    /// <summary>
    ///     A collection of OfficeShortcutKey instances. This class is exposed via the WordExtensions class.
    /// </summary>
    
    public sealed class OfficeShortcutKeyCollection : KeyedCollection<int, OfficeShortcutKey>
    {
        internal OfficeShortcutKeyCollection() { }

        /// <summary>
        ///     Suspends the detection of keyboard shortcuts
        /// </summary>
        public bool SuspendShortcuts { get; set; }


        protected override int GetKeyForItem(OfficeShortcutKey item)
        {
            return item.Keys;
        }


        public int RegisterShortcut(OfficeShortcutKey item)
        {
            Check.NotNull(item, "item");
            base.Add(item);
            return item.RegisteredCount;
        }


        public int UnregisterShortcut(int key)
        {
            return UnregisterShortcut(key, permanently: false);
        }

        public int UnregisterShortcut(int key, bool permanently)
        {
            return UnregisterShortcut(base[key], permanently);
        }

        private int UnregisterShortcut(OfficeShortcutKey item, bool permanently)
        {
            permanently = permanently || item.RegisteredCount == 1;
            if (permanently)
            {
                base.Remove(item.Keys);
                item.RegisteredCount = 0;
            }
            else
            {
                item.RegisteredCount--;
            }

            return item.RegisteredCount;
        }


        //Shadowed members so they are hidden, (but still work just in case they are called)


        [EditorBrowsable(EditorBrowsableState.Never)]
        public new void Add(OfficeShortcutKey item)
        {
            RegisterShortcut(item);
        }


        [EditorBrowsable(EditorBrowsableState.Never)]
        public new void Remove(int key)
        {
            UnregisterShortcut(key);
        }


        [EditorBrowsable(EditorBrowsableState.Never)]
        public new void RemoveAt(int index)
        {
            var osk = Items[index];
            UnregisterShortcut(osk.Keys);
        }


        protected override void InsertItem(int index, OfficeShortcutKey item)
        {
            Check.NotNull(item, nameof(item));

            OfficeShortcutKey existingItem = null;

            TryGetValue(item.Keys, ref existingItem);
            if (existingItem != null)
            {
                if (string.Equals(existingItem.Name, item.Name, StringComparison.OrdinalIgnoreCase))
                {
                    //If an item already exists with the same name then  we bump up the registered count.
                    // Then when the item is removed we remove the registered count by one, and only remove when it equals 0
                    // this way we can register keyboard shortcuts from many places.
                    existingItem.RegisteredCount++;
                    return;
                }

                throw new LJException("Shortcut already exists");
            }

            item.RegisteredCount = 1;
            base.InsertItem(index, item);
            StartStopHook();
        }


        protected override void ClearItems()
        {
            base.ClearItems();
            StartStopHook();
        }


        protected override void RemoveItem(int index)
        {
            var itm = this[index];
            base.RemoveItem(index);
            itm.RegisteredCount = 0;
            StartStopHook();
        }


        [SuppressMessage("Microsoft.Design", "CA1045:DoNotPassTypesByReference", MessageId = "1#")]
        public bool TryGetValue(int key, ref OfficeShortcutKey value)
        {
            if (Dictionary == null)
            {
                return false;
            }

            return Dictionary.TryGetValue(key, out value);
        }


        private void AppDiagnosticsChangedHandler(object sender, AppDiagnosticOptionsChangedEventArgs e)
        {
            //If AppDiagnostics.HasOption(AppDiagnosticOptions.NoIdleHandlers) = False Then
            //AddHandler AppDiagnostics.OptionChanged, AddressOf AppDiagnosticsChangedHandler
            if (Count > 0 && e.HasItemChanged(AppDiagnosticOptions.NoHooks))
            {
                StartStopHook(start: false);
            }
        }


        private void StartStopHook(bool start)
        {
            if (start)
            {
                if (AppDiagnostics.GetOption(AppDiagnosticOptions.NoHooks) == false)
                {
                    WindowsHook.Add(WindowsHookType.WH_KEYBOARD, KeyboardProc);
                }
            }
            else if (Count > 0)
            {
                WindowsHook.Remove(WindowsHookType.WH_KEYBOARD, KeyboardProc);
            }
        }

        private void StartStopHook()
        {
            switch (Count)
            {
            case 0:
                AppDiagnostics.OptionChanged -= AppDiagnosticsChangedHandler;
                StartStopHook(start: false);
                break;
            case 1:
                AppDiagnostics.OptionChanged += AppDiagnosticsChangedHandler;
                StartStopHook(start: true);
                break;
            }
        }


        private void KeyboardProc(ref HookMessage m)
        {
            if (m.Code != (int) HC.ACTION)
            {
                return;
            }

            if (SuspendShortcuts)
            {
                return;
            }

            // ReSharper disable once InconsistentNaming
            const long KBH_STATE = unchecked((int) 0x80000000); //bit 31 0=down 1=up

            int keys = 0;
            if (KeyboardHelper.IsShiftKeyPressed())
	            keys |= OfficeShortcutKey.ShiftModifierKey;

            if (KeyboardHelper.IsControlKeyPressed())
	            keys |= OfficeShortcutKey.ControlModifierKey;

            if (KeyboardHelper.IsAltKeyPressed())
	            keys |= OfficeShortcutKey.AltModifierKey;

            if (keys == 0)
                return;

            keys |= (int) m.WParam;

            OfficeShortcutKey osk;

            Dictionary.TryGetValue(keys, out osk);
            if (osk != null)
            {
                int repeatCount = Libraries.NativeMethods.LoWord(m.LParam.ToInt32());
                var isKeyUp = (m.LParam.ToInt32() & KBH_STATE) == KBH_STATE;

                if (isKeyUp)
                {
                    osk.IsRepeat = true;
                }

                var e = new OfficeShortcutKeyPressedEventArgs(osk, isKeyUp, repeatCount);
                WordExtensions.RaiseShortcutKeyPressed(e);

                if (e.Handled)
                {
                    m.Result = new IntPtr(value: 1);
                }

                if (isKeyUp)
                {
                    osk.IsRepeat = false;
                }
                else
                {
                    osk.IsRepeat = true;
                }

                osk.IsRepeat = isKeyUp == false;
            }
        }
    }
}