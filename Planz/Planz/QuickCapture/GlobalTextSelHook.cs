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
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;

using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Windows.Interop;
using System.Windows.Threading;
using System.Diagnostics;
using System.Threading;
using System.Text.RegularExpressions;
using Accessibility;
using System.Windows.Automation;
using MS_Outlook = Microsoft.Office.Interop.Outlook;


namespace GlobalTextSelHook
{
	public class TextSelHandler
	{
        private IntPtr _hWnd = IntPtr.Zero;

        [DllImport("User32.dll")] private static extern bool

        SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll", CharSet=CharSet.Auto)]

        static public extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]

        static extern void keybd_event(byte bVk, byte bScan, uint dwFlags,

        uint dwExtraInfo);

        private readonly Window _activeWindow;

        public TextSelHandler(Window activeWindow)
        {
            _activeWindow = activeWindow;
        }
        //forced Ctrl C to copy the clippboard
        public void SendCtrlC(IntPtr hWnd){

            uint KEYEVENTF_KEYUP = 2;
            byte VK_CONTROL = 0x11;

            SetForegroundWindow(hWnd);
            keybd_event(VK_CONTROL,0,0,0);
            keybd_event (0x43, 0, 0, 0 ); //Send the C key (43 is "C")

            keybd_event (0x43, 0, KEYEVENTF_KEYUP, 0);

            keybd_event (VK_CONTROL, 0, KEYEVENTF_KEYUP, 0);// 'Left Control Up
        }
	}
}
