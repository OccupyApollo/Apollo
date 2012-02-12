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
using System.IO;
using System.Windows;
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

namespace Planz
{
    public class InfoItem
    {
        public string Title
        {
            get;
            set;
        }

        public string Uri
        {
            get;
            set;
        }

        public InfoItemType Type
        {
            get;
            set;
        }
    }

    public class ActiveWindow
    {
        private IntPtr last_handle = IntPtr.Zero;
        private string last_windowtitle = null;
        private InfoItem activeItem = null;

        public const int WM_GETTEXTLENGTH = 0x000E;
        public const int WM_GETTEXT = 0x000D; 

        public ActiveWindow()
        {

        }

        public IntPtr LastHandle
        {
            get
            {
                return this.last_handle;
            }
            set
            {
                this.last_handle = value;
            }
        }

        public string LastWindowTitle
        {
            get
            {
                return this.last_windowtitle;
            }
            set
            {
                this.last_windowtitle = value;
            }
        }

        public IntPtr GetActiveWindowHandle()
        {
            try
            {
                return GetForegroundWindow();
            }
            catch (Exception)
            {
                return IntPtr.Zero;
            }
        }

        public string GetActiveWindowText(IntPtr handle)
        {
            try
            {
                const int nChars = 256;
                StringBuilder Buff = new StringBuilder(nChars);

                if (GetWindowText(handle.ToInt32(), Buff, nChars) > 0)
                {
                    return Buff.ToString();
                }

                return null;
            }
            catch (Exception)
            {
                return null;
            }
        }

        public List<InfoItem> GetActiveItemList()
        {
            List<InfoItem> itemList = new List<InfoItem>();

            Process[] procs = Process.GetProcesses();
            IntPtr hWnd;
            foreach (Process proc in procs)
            {
                if ((hWnd = proc.MainWindowHandle) != IntPtr.Zero)
                {
                    InfoItem item = GetActiveItemFromHandle(proc.ProcessName, hWnd);
                    if (item != null)
                    {
                        itemList.Add(item);
                    }
                }
            }      

            itemList.AddRange(GetActiveItemFromROT().ToList());

            return itemList;
        }

        public InfoItem GetActiveItem()
        {
            IntPtr hActiveWin = GetActiveWindowHandle();
            String activeTitle = GetActiveWindowText(hActiveWin);

            // Get the active item if it is a Web page
            Process[] procs = Process.GetProcesses();

            foreach (Process proc in procs)
            {
                if ((proc.MainWindowHandle) == hActiveWin)
                {
                    const int nChars = 256;
                    StringBuilder Buff = new StringBuilder(nChars);

                    GetWindowText((int)hActiveWin, Buff, nChars);

                    activeItem = GetActiveItemFromHandle(proc.ProcessName, hActiveWin);
                    if (activeItem != null)
                        return activeItem;
                }
            }

            // Get the active item if it is a File with type of pdf, word, excel, ppt
            activeItem = GetActiveItemFromROT(hActiveWin);
            if (activeItem != null)
                return activeItem;

            return null;
        }

        private InfoItem GetActiveItemFromHandle(string processName, IntPtr handle)
        {
            try
            {
                if (processName == ProcessList.iexplore.ToString())
                {
                    #region IE

                    IEAccessible tabs = new IEAccessible();

                    string pageTitle = tabs.GetActiveTabCaption(handle);
                    string pageURL = tabs.GetActiveTabUrl(handle);

                    return new InfoItem
                    {
                        Title = pageTitle,
                        Uri = pageURL,
                        Type = InfoItemType.Web,
                    };

                    #endregion
                }
                else if (processName == ProcessList.firefox.ToString())
                {
                    #region FireFox

                    AutomationElement ff = AutomationElement.FromHandle(handle);

                    System.Windows.Automation.Condition condition1 = new PropertyCondition(AutomationElement.IsContentElementProperty, true);
                    System.Windows.Automation.Condition condition2 = new PropertyCondition(AutomationElement.ClassNameProperty, "MozillaContentWindowClass");
                    TreeWalker walker = new TreeWalker(new AndCondition(condition1, condition2));
                    AutomationElement elementNode = walker.GetFirstChild(ff);
                    ValuePattern valuePattern;

                    if (elementNode != null)
                    {
                        AutomationPattern[] pattern = elementNode.GetSupportedPatterns();

                        foreach (AutomationPattern autop in pattern)
                        {
                            if (autop.ProgrammaticName.Equals("ValuePatternIdentifiers.Pattern"))
                            {
                                valuePattern = elementNode.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;

                                string pageURL = valuePattern.Current.Value.ToString();
                                string pageTitle = GetActiveWindowText(handle);
                                pageTitle = pageTitle.Remove(pageTitle.IndexOf(" - Mozilla Firefox"));

                                return new InfoItem
                                {
                                    Title = pageTitle,
                                    Uri = pageURL,
                                    Type = InfoItemType.Web,
                                };
                            }
                        }
                    }
                    
                    return null;

                    #endregion
                }
                else if (processName == "chrome")
                {
                    String pageURL = string.Empty;
                    String pageTitle = string.Empty;

                    IntPtr urlHandle = FindWindowEx(handle, IntPtr.Zero, "Chrome_AutocompleteEditView", null);
                    //IntPtr titleHandle = FindWindowEx(handle, IntPtr.Zero, "Chrome_WindowImpl_0", null);
                    IntPtr titleHandle = FindWindowEx(handle, IntPtr.Zero, "Chrome_WidgetWin_0", null);

                    const int nChars = 256;
                    StringBuilder Buff = new StringBuilder(nChars);

                    int length = SendMessage(urlHandle, WM_GETTEXTLENGTH, 0, 0);
                    if (length > 0)
                    {
                        //Get URL from chrome tab
                        SendMessage(urlHandle, WM_GETTEXT, nChars, Buff);
                        pageURL = Buff.ToString();

                        //Get the title
                        GetWindowText((int)titleHandle, Buff, nChars);
                        pageTitle = Buff.ToString();

                        return new InfoItem
                        {
                            Title = pageTitle,
                            Uri = pageURL,
                            Type = InfoItemType.Web,
                        };
                    }
                    else
                        return null;

                }
                else if (processName == ProcessList.OUTLOOK.ToString())
                {
                    #region Outlook

                    MS_Outlook.Application outlookApp;
                    MS_Outlook.MAPIFolder currFolder;
                    MS_Outlook.Inspector inspector;
                    MS_Outlook.MailItem mailItem;

                    outlookApp = (MS_Outlook.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Outlook.Application");
                    currFolder = outlookApp.ActiveExplorer().CurrentFolder;
                    inspector = outlookApp.ActiveInspector();

                    if (inspector == null)
                    {
                        mailItem = (MS_Outlook.MailItem)outlookApp.ActiveExplorer().Selection[1];
                    }
                    else
                    {
                        Regex regex = new Regex(@"\w* - Microsoft Outlook");
                        if (regex.IsMatch(GetActiveWindowText(handle)))
                        {
                            mailItem = (MS_Outlook.MailItem)outlookApp.ActiveExplorer().Selection[1];
                        }
                        else
                        {
                            mailItem = (MS_Outlook.MailItem)inspector.CurrentItem;
                        }
                    }

                    string subject = mailItem.Subject;
                    string entryID = mailItem.EntryID;
                    
                    return new InfoItem
                    {
                        Title = subject,
                        Uri = entryID,
                        Type = InfoItemType.Email,
                    };

                    #endregion
                }
                else if (processName == ProcessList.WINWORD.ToString() ||
                    processName == ProcessList.EXCEL.ToString() ||
                    processName == ProcessList.POWERPNT.ToString() ||
                    processName == ProcessList.AcroRd32.ToString())
                {
                    #region Word, Excel, PPT, Adobe Reader

                    return null;

                    #endregion         
                }
                else
                {
                    return null;
                }
            }
            catch (Exception)
            {
                return null;
            }
        }

        private InfoItem GetActiveItemFromROT(IntPtr handle)
        {
            activeItem = new InfoItem();
            String activeTitle = GetActiveWindowText(handle);
            String title;

            IntPtr numFetched = IntPtr.Zero;
            IRunningObjectTable runningObjectTable;
            IEnumMoniker monikerEnumerator;
            IMoniker[] monikers = new IMoniker[1];

            GetRunningObjectTable(0, out runningObjectTable);
            runningObjectTable.EnumRunning(out monikerEnumerator);
            monikerEnumerator.Reset();

            while (monikerEnumerator.Next(1, monikers, numFetched) == 0)
            {
                IBindCtx ctx;
                CreateBindCtx(0, out ctx);

                string runningObjectName;
                monikers[0].GetDisplayName(ctx, null, out runningObjectName);

                object runningObjectVal;
                runningObjectTable.GetObject(monikers[0], out runningObjectVal);

                try
                {
                    string ext = Path.GetExtension(runningObjectName).ToLower();
                    if (ext == ".pdf" ||
                        ext == ".docx" || ext == ".doc" ||
                        ext == ".xlsx" || ext == ".xls" ||
                        ext == ".pptx" || ext == ".ppt")
                    {
                        title = Path.GetFileNameWithoutExtension(runningObjectName);
                        if (activeTitle.IndexOf(title) != -1)
                        {
                            activeItem.Title = Path.GetFileNameWithoutExtension(runningObjectName);
                            activeItem.Uri = runningObjectName;
                            activeItem.Type = InfoItemType.File;

                            return activeItem;
                        }

                    }
                }
                catch (Exception)
                {
                    continue;
                }
            }

            return null;
        }

        private List<InfoItem> GetActiveItemFromROT()
        {
            List<InfoItem> rotItem = new List<InfoItem>();

            IntPtr numFetched = IntPtr.Zero;
            IRunningObjectTable runningObjectTable;
            IEnumMoniker monikerEnumerator;
            IMoniker[] monikers = new IMoniker[1];

            GetRunningObjectTable(0, out runningObjectTable);
            runningObjectTable.EnumRunning(out monikerEnumerator);
            monikerEnumerator.Reset();

            while (monikerEnumerator.Next(1, monikers, numFetched) == 0)
            {
                IBindCtx ctx;
                CreateBindCtx(0, out ctx);

                string runningObjectName;
                monikers[0].GetDisplayName(ctx, null, out runningObjectName);

                object runningObjectVal;
                runningObjectTable.GetObject(monikers[0], out runningObjectVal);

                try
                {
                    string ext = Path.GetExtension(runningObjectName).ToLower();
                    if (ext == ".pdf" ||
                        ext == ".docx" || ext == ".doc" ||
                        ext == ".xlsx" || ext == ".xls" ||
                        ext == ".pptx" || ext == ".ppt")
                    {
                        rotItem.Add(new InfoItem{
                            Title = Path.GetFileNameWithoutExtension(runningObjectName),
                            Uri = runningObjectName,
                            Type = InfoItemType.File,
                        });

                        runningObjectTable.Register(0, runningObjectVal, monikers[0]);
                    }
                }
                catch (Exception)
                {
                    continue;
                }
            }

            return rotItem;
        }

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern int GetWindowText(int hWnd, StringBuilder text, int count);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("ole32.dll")]
        private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable prot);

        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(int reserved, out IBindCtx ppbc);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr FindWindowEx(IntPtr parentHandle, IntPtr childAfter, string className, string windowTitle);

        [DllImport("user32.dll")]
        private static extern int SendMessage(IntPtr hWnd, uint Msg, int wParam, int lParam);

        [DllImport("user32.dll")]
        private static extern int SendMessage(IntPtr hWnd, uint Msg, int wParam, StringBuilder lParam);
    }

    public class IEAccessible
    {
        private enum OBJID : uint
        {
            OBJID_WINDOW = 0x00000000,
        }

        private const int IE_ACTIVE_TAB = 2097154;
        private const int CHILDID_SELF = 0;

        private IAccessible accessible;

        private IEAccessible[] Children
        {
            get
            {
                int num = 0;

                object[] res = GetAccessibleChildren(accessible, out num);

                if (res == null) return new IEAccessible[0];

                List<IEAccessible> list = new List<IEAccessible>(res.Length);

                foreach (object obj in res)
                {
                    IAccessible acc = obj as IAccessible;

                    if (acc != null) list.Add(new IEAccessible(acc));
                }

                return list.ToArray();
            }

        }

        private string Name
        {
            get
            {
                string ret = accessible.get_accName(CHILDID_SELF);
                return ret;
            }
        }

        private int ChildCount
        {
            get
            {
                int ret = accessible.accChildCount;
                return ret;
            }
        }

        public IEAccessible()
        {

        }

        public IEAccessible(IntPtr ieHandle, string tabCaptionToActivate)
        {
            AccessibleObjectFromWindow(GetDirectUIHWND(ieHandle), OBJID.OBJID_WINDOW, ref accessible);

            if (accessible == null) throw new Exception();

            var ieDirectUIHWND = new IEAccessible(ieHandle);

            foreach (IEAccessible accessor in ieDirectUIHWND.Children)
            {
                foreach (var child in accessor.Children)
                {
                    foreach (var tab in child.Children)
                    {
                        if (tab.Name == tabCaptionToActivate)
                        {
                            tab.Activate();
                            return;
                        }
                    }
                }
            }
        }

        public IEAccessible(IntPtr ieHandle, int tabIndexToActivate)
        {
            AccessibleObjectFromWindow(GetDirectUIHWND(ieHandle), OBJID.OBJID_WINDOW, ref accessible);

            if (accessible == null) throw new Exception();

            var index = 0;

            var ieDirectUIHWND = new IEAccessible(ieHandle);

            foreach (IEAccessible accessor in ieDirectUIHWND.Children)
            {
                foreach (var child in accessor.Children)
                {
                    foreach (var tab in child.Children)
                    {
                        if (tabIndexToActivate >= child.ChildCount - 1) return;
                        if (index == tabIndexToActivate)
                        {
                            tab.Activate();
                            return;
                        }

                        index++;
                    }
                }
            }
        }

        private IEAccessible(IntPtr ieHandle)
        {
            AccessibleObjectFromWindow(GetDirectUIHWND(ieHandle), OBJID.OBJID_WINDOW, ref accessible);

            if (accessible == null) throw new Exception();
        }

        public string GetActiveTabUrl(IntPtr ieHandle)
        {
            AccessibleObjectFromWindow(GetDirectUIHWND(ieHandle), OBJID.OBJID_WINDOW, ref accessible);

            if (accessible == null) throw new Exception();

            var ieDirectUIHWND = new IEAccessible(ieHandle);

            foreach (IEAccessible accessor in ieDirectUIHWND.Children)
            {
                foreach (var child in accessor.Children)
                {
                    foreach (var tab in child.Children)
                    {
                        object tabIndex = tab.accessible.get_accState(CHILDID_SELF);

                        if ((int)tabIndex == IE_ACTIVE_TAB)
                        {
                            var description = tab.accessible.get_accDescription(CHILDID_SELF);

                            if (!string.IsNullOrEmpty(description))
                            {
                                if (description.Contains(Environment.NewLine))
                                {
                                    var url = description.Substring(description.IndexOf(Environment.NewLine)).Trim();
                                    return url;
                                }
                            }
                        }
                    }
                }
            }

            return String.Empty;
        }

        public int GetActiveTabIndex(IntPtr ieHandle)
        {
            AccessibleObjectFromWindow(GetDirectUIHWND(ieHandle), OBJID.OBJID_WINDOW, ref accessible);

            if (accessible == null) throw new Exception();

            var index = 0;
            var ieDirectUIHWND = new IEAccessible(ieHandle);

            foreach (IEAccessible accessor in ieDirectUIHWND.Children)
            {
                foreach (var child in accessor.Children)
                {
                    foreach (var tab in child.Children)
                    {
                        object tabIndex = tab.accessible.get_accState(0);

                        if ((int)tabIndex == IE_ACTIVE_TAB) return index;

                        index++;
                    }
                }
            }

            return -1;
        }

        public string GetActiveTabCaption(IntPtr ieHandle)
        {
            AccessibleObjectFromWindow(GetDirectUIHWND(ieHandle), OBJID.OBJID_WINDOW, ref accessible);

            if (accessible == null) throw new Exception();

            var ieDirectUIHWND = new IEAccessible(ieHandle);

            foreach (IEAccessible accessor in ieDirectUIHWND.Children)
            {
                foreach (var child in accessor.Children)
                {
                    foreach (var tab in child.Children)
                    {
                        object tabIndex = tab.accessible.get_accState(0);

                        if ((int)tabIndex == IE_ACTIVE_TAB) return tab.Name;
                    }
                }
            }

            return String.Empty;
        }

        public List<string> GetTabCaptions(IntPtr ieHandle)
        {
            AccessibleObjectFromWindow(GetDirectUIHWND(ieHandle), OBJID.OBJID_WINDOW, ref accessible);

            if (accessible == null) throw new Exception();

            var ieDirectUIHWND = new IEAccessible(ieHandle);
            var captionList = new List<string>();

            foreach (IEAccessible accessor in ieDirectUIHWND.Children)
            {
                foreach (var child in accessor.Children)
                    foreach (var tab in child.Children)
                        captionList.Add(tab.Name);
            }

            if (captionList.Count > 0) captionList.RemoveAt(captionList.Count - 1);

            return captionList;
        }

        public int GetTabCount(IntPtr ieHandle)
        {
            AccessibleObjectFromWindow(GetDirectUIHWND(ieHandle), OBJID.OBJID_WINDOW, ref accessible);

            if (accessible == null) throw new Exception();

            var ieDirectUIHWND = new IEAccessible(ieHandle);

            foreach (IEAccessible accessor in ieDirectUIHWND.Children)
            {
                foreach (var child in accessor.Children)
                {
                    foreach (var tab in child.Children)
                        return child.ChildCount - 1;
                }
            }

            return 0;
        }

        private IntPtr GetDirectUIHWND(IntPtr ieFrame)
        {
            var directUI = FindWindowEx(ieFrame, IntPtr.Zero, "CommandBarClass", null);
            directUI = FindWindowEx(directUI, IntPtr.Zero, "ReBarWindow32", null);
            directUI = FindWindowEx(directUI, IntPtr.Zero, "TabBandClass", null);
            directUI = FindWindowEx(directUI, IntPtr.Zero, "DirectUIHWND", null);

            return directUI;
        }

        private IEAccessible(IAccessible acc)
        {
            if (acc == null) throw new Exception();

            accessible = acc;
        }

        private void Activate()
        {
            accessible.accDoDefaultAction(CHILDID_SELF);
        }

        private static object[] GetAccessibleChildren(IAccessible ao, out int childs)
        {
            childs = 0;

            object[] ret = null;

            int count = ao.accChildCount;

            if (count > 0)
            {
                ret = new object[count];

                AccessibleChildren(ao, 0, count, ret, out childs);
            }

            return ret;
        }

        #region Interop

        [DllImport("user32.dll", SetLastError = true)]

        private static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass,

        string lpszWindow);

        private static int AccessibleObjectFromWindow(IntPtr hwnd, OBJID idObject, ref IAccessible acc)
        {
            Guid guid = new Guid("{618736e0-3c3d-11cf-810c-00aa00389b71}"); // IAccessible

            object obj = null;

            int num = AccessibleObjectFromWindow(hwnd, (uint)idObject, ref guid, ref obj);

            acc = (IAccessible)obj;

            return num;
        }

        [DllImport("oleacc.dll")]

        private static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint id, ref Guid iid, [In, Out, MarshalAs(UnmanagedType.IUnknown)] ref object ppvObject);

        [DllImport("oleacc.dll")]

        private static extern int AccessibleChildren(IAccessible paccContainer, int iChildStart, int cChildren, [In, Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 2)] object[] rgvarChildren, out int pcObtained);

        #endregion

    }

}