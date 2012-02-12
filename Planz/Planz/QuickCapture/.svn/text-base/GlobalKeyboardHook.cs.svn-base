/*************************************************************************
*
* DO NOT ALTER OR REMOVE COPYRIGHT NOTICES OR THIS FILE HEADER.
 * 
* Copyright (c) 2010 by the Keeping Found Things Found group, 
*                       the Information School, University of Washington
*
* Planz - Bring it together: Capture, Connect, … Complete!
*
* This file is part of Planz
*
* Planz is free software: you can redistribute it and/or modify
* it under the terms of the GNU General Public License version 3
* only, as published by the Free Software Foundation.
*
* Planz is distributed in the hope that it will be useful,
* but WITHOUT ANY WARRANTY; without even the implied warranty of
* MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
* GNU General Public License version 3 for more details
* (a copy is included in the LICENSE file that accompanied this code).
*
* You should have received a copy of the GNU General Public License
* version 3 along with Planz.  If not, see
* <http://www.gnu.org/licenses/gpl.html>
* for a copy of the GPLv3 License.
*
************************************************************************/

using System;
using System.Runtime.InteropServices;
using System.Windows.Interop;
using System.Windows.Media;
using System.Threading;
using System.Windows;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using QuickCapture;
using Planz;

namespace GlobalKeyboardHook
{
    /// Represents a system hot key.
    public class HotKey
    {
        /// Initializes a new instance of the <see cref="HotKey"/> by specific name, modifiers, and key.
        /// <param name="name">A <see cref="string"/> to identify the system hot key.</param>
        /// <param name="modifiers">Modify keys (Shift, Ctrl, Alt, and/or Windows Logo key).</param>
        /// <param name="key">Key code.</param>
        public HotKey(string name, KeyModifiers modifiers, Keys key)
        {
            if (name == null)
            {
                throw new ArgumentNullException("name");
            }

            if (key == Keys.None)
            {
                throw new ArgumentException("key cannot be Keys.None. Must specify a key for the hot key.", "key");
            }

            _name = name;
            _modifiers = modifiers;
            _key = key;
        }

        private int _id = 0;
        /// Gets or sets the ID of the hot key.
        internal int Id
        {
            get { return _id; }
            set { _id = value; }
        }

        private string _name = null;
        /// A <see cref="string"/> used to identify the hot key.
        public string Name
        {
            get { return _name; }
        }

        private KeyModifiers _modifiers = KeyModifiers.None;
        /// Modifier keys of this hot key.
        public KeyModifiers Modifiers
        {
            get { return _modifiers; }
        }

        private Keys _key = Keys.None;
        /// Key of this hot key.
        public Keys Key
        {
            get { return _key; }
        }
    }

    /// Modifier keys.
    [Flags]
    public enum KeyModifiers
    {
        /// No modifier key was pressed.
        None = 0,
        /// Alt key was pressed.
        Alt = 1,
        /// Ctrl key was pressed.
        Control = 2,
        /// Shift key was pressed.
        Shift = 4,
        /// Windows Logo key was pressed.
        Windows = 8
    }

    public class KeyboardHandler : IDisposable
    {

        public const int WM_HOTKEY = 0x0312;

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vlc);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        private readonly Window _mainWindow;
        private WindowInteropHelper _host;
        //private HotKey _hotkey;

        public KeyboardHandler(Window mainWindow)
        {
            _mainWindow = mainWindow;
            _host = new WindowInteropHelper(_mainWindow);

            SetupHotKey(_host.Handle);
            ComponentDispatcher.ThreadPreprocessMessage += ComponentDispatcher_ThreadPreprocessMessage;
        }

        void ComponentDispatcher_ThreadPreprocessMessage(ref MSG msg, ref bool handled)
        {
            if (msg.message == WM_HOTKEY)
            {
                //Handle hot key kere
                ((QuickCaptureWindow)_mainWindow).On_Activate();  
            }
        }

        private void SetupHotKey(IntPtr handle)
        {
            //RegisterHotKey(handle, GetType().GetHashCode(), (int)(KeyModifiers.Control | KeyModifiers.Shift), (int)Keys.C);
            RegisterHotKey(handle, GetType().GetHashCode(), (int)(KeyModifiers.Windows), (int)Keys.C);
        }

        public void Dispose()
        {
            UnregisterHotKey(_host.Handle, GetType().GetHashCode());
        }
    }
}