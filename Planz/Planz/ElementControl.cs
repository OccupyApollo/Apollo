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
using System.Net;
using System.Windows;
using System.Collections.Generic;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;
using System.Text.RegularExpressions;
using System.IO;
using System.Collections;
using System.ComponentModel;
using System.Windows.Threading;
using System.Reflection;
using System.Xml;
using IWshRuntimeLibrary;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Outlook = Microsoft.Office.Interop.Outlook;
using OneNote = Microsoft.Office.Interop.OneNote;
using Microsoft.VisualBasic.FileIO;

namespace Planz
{
    public delegate void FSSyncReporter(object sender, FileSystemSyncEventArgs e);
    public delegate void HasOpenFileDelegate(object sender, EventArgs e);
    public delegate void GeneralOperationErrorDelegate(object sender, OperationErrorEventArgs e);

    public class OperationErrorEventArgs : EventArgs
    {
        public string Message
        {
            get;
            set;
        }

        public OperationErrorEventArgs()
        {

        }
    }

    public class FileSystemSyncEventArgs : EventArgs
    {
        public FileSystemSyncMessage Message
        {
            get;
            set;
        }

        public FileSystemSyncEventArgs()
        {

        }
    }

    public class ElementControl
    {
        private DatabaseControl dbControl;
        private SearchFunction sf = new SearchFunction();

        private Stack pendingHeadingHighP = new Stack();
        private Stack pendingHeadingLowP = new Stack();

        public delegate void AddElementMethodInvoker();
        private System.Threading.Thread mainThread = System.Threading.Thread.CurrentThread;

        private Element root;
        private Element currentElement;
        private Element previousElement;
        private Element hoverElement;
        private List<Element> selectedElements = new List<Element>();
        private List<Element> updateStatusList = new List<Element>();
        private Stack<Element> tobeDeletedElements = new Stack<Element>();

        private NavigationItem ni_root;
        private bool isNavigationExpansion = false;
        private int MAX_NEWASSOCIATION_VISIBLE = 30;

        public event FSSyncReporter fsSyncReporter;
        public event HasOpenFileDelegate hasOpenFileDelegate;
        public event GeneralOperationErrorDelegate generalOperationErrorDelegate;


        // Note: param "rootElementPath" should be ending with System.IO.Path.DirectorySeparatorChar
        public ElementControl(string rootElementPath)
        {
            root = new Element
            {
                ParentElement = null,
                HeadImageSource = String.Empty,
                TailImageSource = String.Empty,
                NoteText = String.Empty,
                IsExpanded = true,
                Path = rootElementPath,
                Type = ElementType.Heading,
            };

            ni_root = new NavigationItem
            {
                Name = String.Empty,
                Path = rootElementPath,
                Parent = null,
            };

            dbControl = new DatabaseControl(rootElementPath);
            dbControl.newXooMLCreate += new NewXooMLCreateDelegate(dbControl_newXooMLCreate);
            dbControl.OpenConnection();
            dbControl.newDBControlHighP += new NewDatabaseControlHandler(dbControl_newDBControlHighP);
            dbControl.newDBControlLowP += new NewDatabaseControlHandler(dbControl_newDBControlLowP);
            dbControl.elementUpdate += new ElementUpdateDelegate(dbControl_elementUpdate);
            dbControl.elementStatusChangedDelegate += new ElementStatusChangedDelegate(dbControl_elementStatusChanged);
            dbControl.elementDelete += new ElementDeleteDelegate(dbControl_elementDelete);

            root.ShowAssociationMarkedDone = dbControl.GetFragmentElementFromXML().ShowAssociationMarkedDone;
            root.ShowAssociationMarkedDefer = dbControl.GetFragmentElementFromXML().ShowAssociationMarkedDefer;

            foreach (Element element in dbControl.GetAllElementFromXML())
            {
                if (element.IsVisible == Visibility.Collapsed)
                {
                    continue;
                }

                AddElement(element, root);

                if (element.IsHeading)
                {
                    NavigationItem ni = new NavigationItem
                    {
                        Name = element.NoteText,
                        Path = element.Path,
                        Parent = ni_root,
                    };
                    ni.Items.Add(new NavigationItem());
                    ni_root.Items.Add(ni);
                }
            }

            currentElement = root;

            RunXMLBackgroundWorker();
        }

        public void UpdateRootPath(string newRootPath)
        {
            root.Path = newRootPath;
            root.Elements.Clear();

            ni_root.Path = newRootPath;
            ni_root.Items.Clear();
            ni_root.Items.Clear();
            selectedElements.Clear();
            updateStatusList.Clear();
            tobeDeletedElements.Clear();

            GC.Collect();

            dbControl = new DatabaseControl(newRootPath);
            dbControl.newXooMLCreate += new NewXooMLCreateDelegate(dbControl_newXooMLCreate);
            dbControl.OpenConnection();
            dbControl.newDBControlHighP += new NewDatabaseControlHandler(dbControl_newDBControlHighP);
            dbControl.newDBControlLowP += new NewDatabaseControlHandler(dbControl_newDBControlLowP);
            dbControl.elementUpdate += new ElementUpdateDelegate(dbControl_elementUpdate);
            dbControl.elementStatusChangedDelegate += new ElementStatusChangedDelegate(dbControl_elementStatusChanged);
            dbControl.elementDelete += new ElementDeleteDelegate(dbControl_elementDelete);

            foreach (Element element in dbControl.GetAllElementFromXML())
            {
                AddElement(element, root);

                if (element.IsHeading)
                {
                    NavigationItem ni = new NavigationItem
                    {
                        Name = element.NoteText,
                        Path = element.Path,
                        Parent = ni_root,
                    };
                    ni.Items.Add(new NavigationItem());
                    ni_root.Items.Add(ni);
                }
            }

            currentElement = root;

            RunXMLBackgroundWorker();
        }
        
        public Element Root
        {
            get { return root; }
        }

        public Element CurrentElement
        {
            get 
            { 
                return currentElement; 
            }
            set 
            { 
                currentElement = value;
                Properties.Settings.Default.LastFocusedElementGuid = currentElement.ID.ToString();
                Properties.Settings.Default.Save();
            }
        }

        public Element PreviousElement
        {
            get { return previousElement; }
            set { previousElement = value; }
        }

        public Element HoverElement
        {
            get { return hoverElement; }
            set { hoverElement = value; }
        }

        public List<Element> SelectedElements
        {
            get { return selectedElements; }
            set { selectedElements = value; }
        }

        public NavigationItem NI_Root
        {
            get { return ni_root; }
        }

        #region FS Sync

        private void dbControl_newDBControlHighP(object sender, EventArgs e)
        {
            Element element = sender as Element;
            if (pendingHeadingHighP.Contains(element) == false)
            {
                pendingHeadingHighP.Push(element);
            }
        }

        private void dbControl_newDBControlLowP(object sender, EventArgs e)
        {
            Element element = sender as Element;
            if (pendingHeadingLowP.Contains(element) == false)
            {
                pendingHeadingLowP.Push(element);
            }
        }

        private void dbControl_newXooMLCreate(object sender, EventArgs e)
        {
            DatabaseControl temp_dbControl = sender as DatabaseControl;
            List<Element> eleList = temp_dbControl.GetAllElementFromXML();

            if (eleList.Count < MAX_NEWASSOCIATION_VISIBLE) 
                return;

            for (int i = MAX_NEWASSOCIATION_VISIBLE; i < eleList.Count; i++)
            {
                Element ele = eleList[i];
                ele.IsVisible = Visibility.Collapsed;
                temp_dbControl.UpdateElementIntoXML(ele);
            }
        }

        private void dbControl_elementUpdate(object sender, EventArgs e)
        {
            Element element = sender as Element;
            UpdateElement(element);
        }

        private void dbControl_elementStatusChanged(object sender, EventArgs e)
        {
            Element element = sender as Element;
            updateStatusList.Add(element);
        }

        private void dbControl_elementDelete(object sender, EventArgs e)
        {
            Element element = sender as Element;
            tobeDeletedElements.Push(element);
        }

        private void RunXMLBackgroundWorker()
        {
            BackgroundWorker xmlReader = new BackgroundWorker();
            xmlReader.DoWork += new DoWorkEventHandler(xmlReader_DoWork);
            xmlReader.RunWorkerCompleted += new RunWorkerCompletedEventHandler(xmlReader_RunWorkerCompleted);
            xmlReader.RunWorkerAsync();
        }

        public void UpdateStatusChangedElementList()
        {
            foreach (Element element in updateStatusList)
            {
                if (element.Status == ElementStatus.New ||
                    element.Status == ElementStatus.MissingAssociation)
                {
                    if (element.IsHeading)
                    {
                        element.FontColor = ElementColor.DarkBlue.ToString();
                    }
                    else
                    {
                        element.FontColor = ElementColor.Black.ToString();
                    }

                    if (element.Status == ElementStatus.New)
                        element.Status = ElementStatus.Normal;

                    UpdateElement(element);
                }
            }
        }

        private void xmlReader_DoWork(object sender, DoWorkEventArgs e)
        {
            while (pendingHeadingHighP.Count > 0)
            {
                try
                {
                    Element element = pendingHeadingHighP.Pop() as Element;
                    DatabaseControl temp_dbControl;
                    if (element.IsRemoteHeading)
                    {
                        string path = String.Empty;
                        IWshRuntimeLibrary.IWshShortcut shortcut = GetShortcut(element);
                        if (shortcut != null)
                        {
                            path = shortcut.TargetPath + System.IO.Path.DirectorySeparatorChar;
                        }
                        temp_dbControl = new DatabaseControl(path);
                    }
                    else
                    {
                        temp_dbControl = new DatabaseControl(element.Path);
                    }
                    temp_dbControl.newXooMLCreate += new NewXooMLCreateDelegate(dbControl_newXooMLCreate);
                    temp_dbControl.OpenConnection();
                    temp_dbControl.newDBControlHighP += new NewDatabaseControlHandler(dbControl_newDBControlHighP);
                    temp_dbControl.newDBControlLowP += new NewDatabaseControlHandler(dbControl_newDBControlLowP);

                    temp_dbControl.elementStatusChangedDelegate += new ElementStatusChangedDelegate(dbControl_elementStatusChanged);
                    temp_dbControl.elementDelete += new ElementDeleteDelegate(dbControl_elementDelete);

                    List<Element> eleList = temp_dbControl.GetAllElementFromXML();

                    int hiddenElements = 0;
                    foreach (Element ele in eleList)
                    {
                        Dispatcher.FromThread(mainThread).Invoke((AddElementMethodInvoker)delegate
                        {
                            if (tobeDeletedElements.Contains(ele))
                            {

                            }
                            else
                            {
                                Element temp_ele = temp_dbControl.GetFragmentElementFromXML();
                                element.ShowAssociationMarkedDefer = temp_ele.ShowAssociationMarkedDefer;
                                element.ShowAssociationMarkedDone = temp_ele.ShowAssociationMarkedDone;
                                element.StartDate = temp_ele.StartDate;
                                element.DueDate = temp_ele.DueDate;

                                if (ele.IsVisible == Visibility.Visible)
                                {
                                    if (ele.IsHeading)
                                    {
                                        ele.ParentElement = element;
                                        if (IsCircularHeading(ele))
                                        {
                                            if (ele.IsExpanded)
                                            {
                                                Stack<Element> s = new Stack<Element>();
                                                while (pendingHeadingHighP.Count != 0)
                                                {
                                                    Element _ele = pendingHeadingHighP.Pop() as Element;
                                                    if (_ele != ele)
                                                    {
                                                        s.Push(_ele);
                                                    }
                                                    else
                                                    {
                                                        while (s.Count != 0)
                                                        {
                                                            pendingHeadingHighP.Push(s.Pop());
                                                        }
                                                        break;
                                                    }
                                                }

                                                ele.IsExpanded = false;
                                            }
                                        }
                                    }
                                    AddElement(ele, element);
                                }
                                else
                                {
                                    hiddenElements++;
                                }
                            }

                        });
                    }

                    if (eleList.Count != 0 && hiddenElements == eleList.Count)
                    {
                        Dispatcher.FromThread(mainThread).Invoke((AddElementMethodInvoker)delegate
                        {
                            InsertCommandNote("Click the icon on the right to show hidden associations", ElementCommand.ShowHiddenAssociations, element);
                        });
                    }

                    if (element.CanImportCalendar)
                    {
                        Dispatcher.FromThread(mainThread).Invoke((AddElementMethodInvoker)delegate
                        {
                            InsertCommandNote("Click the icon on the right to import Meetings and Appointments from Outlook", ElementCommand.ImportMeetingsAndAppointmentsFromOutlook, element);
                        });
                    }

                    while (tobeDeletedElements.Count > 0)
                    {
                        Element tobedeleted = tobeDeletedElements.Pop();
                        tobedeleted.ParentElement = element;
                        DeleteElement(tobedeleted);
                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }
        }

        private void xmlReader_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            ((BackgroundWorker)sender).Dispose();

            OnFSSyncFinished(null, new FileSystemSyncEventArgs{ Message = FileSystemSyncMessage.HighPriorityWorkFinished });

            if (pendingHeadingLowP.Count != 0)
            {
                while (pendingHeadingLowP.Count > 0)
                {
                    pendingHeadingHighP.Push(pendingHeadingLowP.Pop());
                }

                ((BackgroundWorker)sender).RunWorkerAsync();
            }
            else
            {
                if (isNavigationExpansion)
                {
                    isNavigationExpansion = false;
                    OnFSSyncFinished(null, new FileSystemSyncEventArgs { Message = FileSystemSyncMessage.NavigationTreeExpanded });
                }
                else
                {
                    OnFSSyncFinished(null, new FileSystemSyncEventArgs { Message = FileSystemSyncMessage.LowPriorityWorkFinished });
                }
            }
        }

        protected virtual void OnFSSyncFinished(object sender, FileSystemSyncEventArgs e)
        {
            if (fsSyncReporter != null)
                fsSyncReporter(sender, e);
        }

        public void SyncHeadingElement(Element element)
        {
            element.Elements.Clear();
            GC.Collect();

            pendingHeadingHighP.Push(element);

            RunXMLBackgroundWorker();
        }

        public void ShowOrHidePowerDElement(Element parent, PowerDStatus status, bool show)
        {
            try
            {
                DatabaseControl temp_dbControl;
                List<Element> allElements = new List<Element>();

                if (parent.Path != root.Path)
                {
                    temp_dbControl = new DatabaseControl(parent.Path);
                    temp_dbControl.OpenConnection();
                    temp_dbControl.DO_NOT_UPDATE_POWERDREGION = true;
                    allElements = temp_dbControl.GetAllElementFromXML();

                }
                else
                {
                    temp_dbControl = dbControl;
                    temp_dbControl.DO_NOT_UPDATE_POWERDREGION = true;
                    allElements = temp_dbControl.GetAllElementFromXML();
                    temp_dbControl.DO_NOT_UPDATE_POWERDREGION = false;
                }

                for (int i = allElements.Count - 1; i >= 0; i--)
                {
                    if (allElements[i].PowerDStatus == status)
                    {
                        //allElements[i].ParentElement = parent;
                        if (show)
                        {
                            allElements[i].IsVisible = Visibility.Visible;
                        }
                        else
                        {
                            allElements[i].IsVisible = Visibility.Collapsed;
                            for (int j = parent.Elements.Count - 1; j >= 0 ; j--)
                            {
                                if (parent.Elements[j].ID == allElements[i].ID)
                                {
                                    parent.Elements.RemoveAt(j);
                                    break;
                                }
                            }
                        }
                        UpdateElement(allElements[i]);
                    }
                    else
                    {
                        continue;
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }

        public void SyncNavigationItem(NavigationItem ni_root)
        {
            List<NavigationItem> ni_list;

            if (ni_root.Path != root.Path)
            {
                DatabaseControl temp_dbControl = new DatabaseControl(ni_root.Path);
                temp_dbControl.OpenConnection();
                ni_list = temp_dbControl.GetAllHeadingElementFromXML();
                temp_dbControl.CloseConnection();
            }
            else
            {
                ni_list = dbControl.GetAllHeadingElementFromXML();
            }

            ni_root.Items.Clear();
            foreach (NavigationItem ni in ni_list)
            {
                ni.Parent = ni_root;
                ni.Items.Add(new NavigationItem());
                ni_root.Items.Add(ni);
            }
        }

        #endregion

        #region Element Ops and Outline Ops

        public Element CreateNewElement(ElementType type, string noteText)
        {
            Element newElement = new Element
            {
                NoteText = noteText,
                ParentElement = root,
            };
            switch (type)
            {
                case ElementType.Heading:
                    newElement.Type = type;
                    newElement.FlagImageSource = String.Format("pack://application:,,,/{0};component/{1}", "Planz", "Images/normal.gif");
                    newElement.IsExpanded = false;
                    newElement.Path = newElement.ParentElement.Path + HeadingNameConverter.ConvertFromHeadingNameToFolderName(newElement) + System.IO.Path.DirectorySeparatorChar;
                    // Remove while spaces
                    newElement.NoteText = newElement.NoteText.Trim();
                    if (!CreateFolder(newElement))
                    {
                        newElement.Type = ElementType.Note;
                        MessageBox.Show("The heading name is too long, please shorten the name and try again.");
                    }
                    else
                    {
                        if (Properties.Settings.Default.ShowOutline)
                        {
                            newElement.ShowExpander = Visibility.Visible;
                        }
                    }
                    break;
                case ElementType.Note:
                    newElement.Type = type;
                    newElement.FlagImageSource = String.Format("pack://application:,,,/{0};component/{1}", "Planz", "Images/normal.gif");
                    break;
            };

            return newElement;
        }

        public Element CreateNewElement(Element element)
        {
            Element newElement = CreateNewElement(element.Type, element.NoteText);

            newElement.AssociationType = element.AssociationType;
            newElement.AssociationURI = element.AssociationURI;
            newElement.FlagStatus = element.FlagStatus;
            newElement.FontColor = element.FontColor;
            newElement.IsExpanded = element.IsExpanded;
            newElement.IsVisible = element.IsVisible;
            newElement.LevelOfSynchronization = element.LevelOfSynchronization;
            newElement.ModifiedOn = element.ModifiedOn;
            newElement.PowerDStatus = element.PowerDStatus;
            newElement.PowerDTimeStamp = element.PowerDTimeStamp;
            newElement.ShowAssociationMarkedDefer = element.ShowAssociationMarkedDefer;
            newElement.ShowAssociationMarkedDone = element.ShowAssociationMarkedDone;
            newElement.StartDate = element.StartDate;
            newElement.Status = element.Status;
            newElement.TempData = element.TempData;

            return newElement;
        }

        public bool ChangeElementType(ElementType type, ref Element element)
        {
            switch (type)
            {
                case ElementType.Heading:
                    element.Type = type;
                    element.IsExpanded = false;
                    element.LevelOfSynchronization = 0;
                    string path = element.Path;
                    string associatedURI = element.AssociationURI;
                    if (element.AssociationType != ElementAssociationType.FolderShortcut)
                    {
                        element.Path = element.ParentElement.Path + HeadingNameConverter.ConvertFromHeadingNameToFolderName(element) + System.IO.Path.DirectorySeparatorChar;
                        element.NoteText = element.NoteText.Trim();

                        if (!CreateFolder(element))
                        {
                            element.Type = ElementType.Note;
                            element.AssociationURI = associatedURI;
                            element.Path = path;
                            //MessageBox.Show("The heading name is too long, please shorten the name and try again.");
                            return false;
                        }
                        else
                        {
                            element.AssociationURI = System.IO.Path.GetFileName(System.IO.Path.GetDirectoryName(element.Path));
                            if (Properties.Settings.Default.ShowOutline)
                            {
                                element.ShowExpander = Visibility.Visible;
                            }
                            return true;
                        }
                    }
                    else
                    {
                        IWshRuntimeLibrary.WshShellClass wshShell = new IWshRuntimeLibrary.WshShellClass();
                        IWshRuntimeLibrary.IWshShortcut shortcut = (IWshRuntimeLibrary.IWshShortcut)wshShell.CreateShortcut(element.ParentElement.Path + element.AssociationURI);
                        if (shortcut != null)
                        {
                            element.Path = shortcut.TargetPath + System.IO.Path.DirectorySeparatorChar;
                        }
                        if (Properties.Settings.Default.ShowOutline)
                        {
                            element.ShowExpander = Visibility.Visible;
                        }
                    }
                    
                    break;
                case ElementType.Note:
                    if (!element.IsRemoteHeading)
                    {
                        RemoveFolder(element);
                        element.AssociationType = ElementAssociationType.None;
                        element.AssociationURI = String.Empty;
                    }
                    else
                    {
                        element.AssociationType = ElementAssociationType.FolderShortcut;
                    }
                    element.Type = type;
                    element.IsExpanded = false;
                    element.LevelOfSynchronization = 0;
                    break;
            };
            return true;
        }

        public void AddElement(Element newElement, Element parentElement)
        {
            if (FindElementByGuid(parentElement, newElement.ID) == null)
            {
                newElement.ParentElement = parentElement;
                parentElement.Elements.Add(newElement);
            }
        }

        // Includes UI change and XML IO
        public void InsertElement(Element newElement, Element parentElement, int index)
        {
            newElement.ParentElement = parentElement;
            parentElement.Elements.Insert(index, newElement);

            if (newElement.IsCommandNote)
            {
                return;
            }

            if (newElement.ParentElement.Path != root.Path)
            {
                DatabaseControl temp_dbControl = new DatabaseControl(newElement.ParentElement.Path);
                temp_dbControl.OpenConnection();
                temp_dbControl.InsertElementIntoXML(newElement);
                temp_dbControl.CloseConnection();
            }
            else
            {
                dbControl.InsertElementIntoXML(newElement);
            }
        }

        // Includes UI change and XML IO
        public void RemoveElement(Element element, Element parentElement)
        {
            string previousText = element.NoteText;
            parentElement.Elements.Remove(element);
            element.NoteText = previousText;

            if (element.ParentElement.Path != root.Path)
            {
                DatabaseControl temp_dbControl = new DatabaseControl(element.ParentElement.Path);
                temp_dbControl.OpenConnection();
                temp_dbControl.RemoveElementFromXML(element);
                temp_dbControl.CloseConnection();
            }
            else
            {
                dbControl.RemoveElementFromXML(element);
            }
        }

        public void Promote(Element element)
        {
            Element tempElement = element;

            // Promote from note to remote heading
            if (element.Type == ElementType.Note && element.AssociationType == ElementAssociationType.FolderShortcut)
            {
                IWshRuntimeLibrary.IWshShortcut shortcut = GetShortcut(element);
                string target = shortcut.TargetPath;

                if (IsSubfolder(target, element.ParentElement.Path) == true)
                {
                    const string message = "You cannot promote a folder within itself or its subfolder.";
                    OnOperationError(element, new OperationErrorEventArgs { Message = message });
                    return;
                }

                ChangeElementType(ElementType.Heading, ref element);
                UpdateElement(element);

                return;

            }

            if (element.HasAssociation && element.IsHeading == false)
            {
                return;
            }

            Element newElement = null;
            bool bLongName = false;

            if (element.ParentElement == root)
            {
                switch (element.Type)
                {
                    case ElementType.Heading:
                        break;
                    case ElementType.Note:

                        string noteText = element.NoteText;
                        bLongName = noteText.Length > StartProcess.MAX_EXTRACTNAME_LENGTH;

                        if (bLongName == true)
                        {
                            newElement = CreateNewElement(ElementType.Note, noteText);
                            element.NoteText = noteText.Substring(0, StartProcess.MAX_EXTRACTNAME_LENGTH);
                        }

                        if (ChangeElementType(ElementType.Heading, ref element))
                            UpdateElement(element);
                        else
                        {
                            element.NoteText = noteText;
                            return;
                        }

                        if (bLongName == true)
                            InsertElement(newElement, element, 0);
                        break;
                };
            }
            else
            {
                switch (element.Type)
                {
                    case ElementType.Heading:

                        UpdateElement(element);
                        if (element.IsLocalHeading && (CheckOpenFiles(element) == true))
                        {
                            return;
                        }

                        int index = element.ParentElement.ParentElement.Elements.IndexOf(element.ParentElement);
                        if (element == element.ParentElement.FirstChild && element.ParentElement.Elements.Count == 1)
                        {
                            element.ParentElement.IsExpanded = false;
                            UpdateElement(element.ParentElement);
                        }
                        if (element.IsLocalHeading)
                        {

                            string previousPath = element.Path;
                            element.Path = element.ParentElement.ParentElement.Path + HeadingNameConverter.ConvertFromHeadingNameToFolderName(element) + System.IO.Path.DirectorySeparatorChar;

                            RemoveElement(element, element.ParentElement);
                            InsertElement(element, element.ParentElement.ParentElement, index + 1);
                            UpdateElement(element);
                            MoveFolder(element, previousPath);
                        }
                        else
                        {
                            MoveElement(element, element.ParentElement.ParentElement, index + 1);
                        }
                        break;
                    case ElementType.Note:
                        string noteText = element.NoteText;
                        bLongName = noteText.Length > StartProcess.MAX_EXTRACTNAME_LENGTH;

                        if (bLongName == true)
                        {
                            newElement = CreateNewElement(ElementType.Note, noteText);
                            element.NoteText = noteText.Substring(0, StartProcess.MAX_EXTRACTNAME_LENGTH);
                        }

                        if (ChangeElementType(ElementType.Heading, ref element))
                            UpdateElement(element);
                        else
                        {
                            element.NoteText = noteText;
                            return;
                        }

                        if (bLongName == true)
                            InsertElement(newElement, element, 0);

                        break;
                };
            }
        }

        public void Demote(Element element)
        {
            if (element.ParentElement.FirstChild == element)
            {
                switch (element.Type)
                {
                    case ElementType.Heading:
                        if (element.IsRemoteHeading)
                        {
                            ChangeElementType(ElementType.Note, ref element);
                            element.Elements.Clear();
                            GC.Collect();
                            UpdateElement(element);
                        }
                        else
                        {
                            if (HasChildOrContent(element) == false)
                            {
                                ChangeElementType(ElementType.Note, ref element);
                                UpdateElement(element);
                            }
                        }
                        break;
                    case ElementType.Note:
                        break;
                };
            }
            else
            {
                switch (element.Type)
                {
                    case ElementType.Heading:
                        if ((HasChildOrContent(element) == false) || (!element.IsLocalHeading && element.IsCollapsed))
                        {
                            ChangeElementType(ElementType.Note, ref element);
                            UpdateElement(element);
                        }
                        else
                        {
                            Element elementAbove = element.ElementAboveUnderSameParent;
                            if (elementAbove != null)
                            {
                                if (elementAbove.Type == ElementType.Heading)
                                {
                                    if (element.IsLocalHeading && (CheckOpenFiles(element) == true))
                                        return;

                                    elementAbove.IsExpanded = true;

                                    UpdateElement(elementAbove);
                                    if (element.IsLocalHeading)
                                    {
                                        string previousPath = element.Path;
                                        element.Path = elementAbove.Path + HeadingNameConverter.ConvertFromHeadingNameToFolderName(element) + System.IO.Path.DirectorySeparatorChar;

                                        RemoveElement(element, element.ParentElement);
                                        InsertElement(element, elementAbove, elementAbove.Elements.Count);
                                        MoveFolder(element, previousPath);
                                        UpdateElement(element);
                                    }
                                    else
                                    {
                                        MoveElement(element, elementAbove, elementAbove.Elements.Count);
                                    }
                                }
                                else
                                {

                                    return;
                                }
                            }
                        }
                        break;

                    case ElementType.Note:
                        Element elementAboveUnderSameParent = element.ElementAboveUnderSameParent;
                        if (elementAboveUnderSameParent.Type == ElementType.Heading)
                        {
                            elementAboveUnderSameParent.IsExpanded = true;

                            UpdateElement(elementAboveUnderSameParent);
                            MoveElement(element, elementAboveUnderSameParent, elementAboveUnderSameParent.Elements.Count);
                        }
                        break;
                };
            }
        }

        public void MoveUp(Element element)
        {
            if (element.IsCommandNote)
            {
                return;
            }

            switch (element.Type)
            {
                case ElementType.Heading:
                case ElementType.Note:
                    // If it is the first child of its parent
                    if (element == element.ParentElement.FirstChild)
                    {
                        if (element.ParentElement == root)
                        {
                            // If its parent is Root
                            // It cant move up
                            return;
                        }
                        else
                        {
                            // If its parent is not Root
                            // Move to new parent
                            Element newParent = element.ParentElement.ParentElement;
                            MoveElement(element, newParent, newParent.Elements.IndexOf(element.ParentElement));
                        }
                    }
                    // If it is not the first child of its parent
                    else
                    {
                        Element elementAbove = element.ElementAboveUnderSameParent;
                        switch (elementAbove.Type)
                        {
                            // If its above Element under same parent is heading
                            case ElementType.Heading:
                                if (elementAbove.IsExpanded)
                                {
                                    // If the heading is expanded
                                    // Move to this heading
                                    MoveElement(element, elementAbove, elementAbove.Elements.Count);
                                }
                                else
                                {
                                    // If the heading is not expanded
                                    // Swap with this heading
                                    SwapPosition(element, elementAbove);
                                }
                                break;
                            // If its above Element under same parent is note
                            case ElementType.Note:
                                // Swap with above Element
                                SwapPosition(element, elementAbove);
                                break;
                        }; 
                    }
                    break;
            };
        }

        public void MoveDown(Element element)
        {
            if (element.IsCommandNote)
            {
                return;
            }

            switch (element.Type)
            {
                case ElementType.Heading:
                case ElementType.Note:
                    // If it is the last child of its parent
                    if (element == element.ParentElement.LastChild)
                    {
                        if (element.ElementBelow == null)
                        {
                            // If it is the last Element
                            // It cant move down
                            return;
                        }
                        else
                        {
                            // If its parent Element is Root
                            // It cant move down
                            if (element.ParentElement == root)
                            {                      
                                return;
                            }

                            // If it is not the last Element
                            // Move to new parent
                            Element newParent = element.ParentElement.ParentElement;
                            MoveElement(element, newParent, newParent.Elements.IndexOf(element.ParentElement) + 1);
                        }
                    }
                    // If it is not the last child of its parent
                    else
                    {
                        Element elementBelow = element.ElementBelowUnderSameParent;
                        switch (elementBelow.Type)
                        {
                            // If its following Element under same parent is heading
                            case ElementType.Heading:
                                if (elementBelow.IsExpanded)
                                {
                                    // If the heading is expanded
                                    // Move to this heading
                                    MoveElement(element, elementBelow, 0);
                                }
                                else
                                {
                                    // If the heading is not expanded
                                    // Swap with this heading
                                    SwapPosition(element, elementBelow);
                                }
                                break;
                            // If its above Element under same parent is note
                            case ElementType.Note:
                                // Swap with following Element
                                SwapPosition(element, elementBelow);
                                break;
                        }; 
                    }
                    break;
            };
        }

        private void SwapPosition(Element element1, Element element2)
        {
            int pos1 = element1.Position;
            int pos2 = element2.Position;

            Element parentElement = element1.ParentElement;
            parentElement.Elements.Remove(element1);
            parentElement.Elements.Insert(pos2, element1);
            parentElement.Elements.Remove(element2);
            parentElement.Elements.Insert(pos1, element2);

            element1.Position = pos2;
            element2.Position = pos1;

            if ((element1.PowerDStatus == PowerDStatus.Done || element1.PowerDStatus == PowerDStatus.None) && element2.PowerDStatus == PowerDStatus.Deferred)
            {
                element1.PowerDStatus = PowerDStatus.Deferred;
                element1.FontColor = ElementColor.Purple.ToString();
                element1.PowerDTimeStamp = element2.PowerDTimeStamp;
            }
            else if ((element1.PowerDStatus == PowerDStatus.Deferred || element1.PowerDStatus == PowerDStatus.None) && element2.PowerDStatus == PowerDStatus.Done)
            {
                element1.PowerDStatus = PowerDStatus.Done;
                element1.FontColor = ElementColor.SeaGreen.ToString();
                element1.PowerDTimeStamp = element2.PowerDTimeStamp;
            }
            else if ((element1.PowerDStatus == PowerDStatus.Done || element1.PowerDStatus == PowerDStatus.Deferred) && element2.PowerDStatus == PowerDStatus.None)
            {
                element1.PowerDStatus = PowerDStatus.None;
                element1.FontColor = ElementColor.Black.ToString();
            }

            UpdateElement(element1);
            UpdateElement(element2);
        }

        private void MoveElement(Element element, Element newParentElement, int index)
        {
            // If move across heading
            // Check open files
            Element previousParentElement = element.ParentElement;
            if (previousParentElement != newParentElement)
            {
                if (CheckOpenFiles(element) == true)
                {
                    return;
                }
            }
            
            // Associated file change
            string newPath = newParentElement.Path;
            MoveFile(element, newPath);

            // XML and UI change
            RemoveElement(element, previousParentElement);
            InsertElement(element, newParentElement, index);
            
            UpdateElement(element);
            if (previousParentElement != root)
            {
                UpdateElement(previousParentElement);
            }
        }

        // Includes XML IO
        public void UpdateElement(Element element)
        {
            if (element == root)
            {
                dbControl.UpdateFragmentElementIntoXML(element);
                return;
            }

            if (element.IsCommandNote)
            {
                return;
            }

            if (element.ParentElement == null)
            {
                return;
            }

            if (element.IsLocalHeading)
            {
                string previousName = System.IO.Directory.GetParent(element.Path).Name;
                string currentName = HeadingNameConverter.ConvertFromHeadingNameToFolderName(element);
                if (previousName != currentName)
                {
                    if (CheckOpenFiles(element) == true)
                    {
                        return;
                    }
                }
            }

            if (element.ParentElement.Path != root.Path)
            {
                DatabaseControl temp_dbControl = new DatabaseControl(element.ParentElement.Path);
                temp_dbControl.newXooMLCreate += new NewXooMLCreateDelegate(dbControl_newXooMLCreate);
                temp_dbControl.OpenConnection();
                temp_dbControl.elementStatusChangedDelegate += new ElementStatusChangedDelegate(dbControl_elementStatusChanged);
                temp_dbControl.UpdateElementIntoXML(element);
                temp_dbControl.CloseConnection();
            }
            else
            {
                dbControl.UpdateElementIntoXML(element);
            }

            switch (element.Type)
            {
                case ElementType.Heading:
                    if (element.IsLocalHeading)
                    {
                        string previousPath = element.Path;
                        element.Path = element.ParentElement.Path + HeadingNameConverter.ConvertFromHeadingNameToFolderName(element) + System.IO.Path.DirectorySeparatorChar;
                        try
                        {
                            RenameFolder(element, previousPath);
                        }
                        catch (Exception)
                        {
                            element.Path = previousPath;
                            element.NoteText = System.IO.Directory.GetParent(element.Path).Name;
                            MessageBox.Show("The heading name is too long, please shorten the name and try again.");
                            return;
                        }
                    }
                    DatabaseControl temp_dbControl = new DatabaseControl(element.Path);
                    temp_dbControl.newXooMLCreate +=new NewXooMLCreateDelegate(dbControl_newXooMLCreate);
                    temp_dbControl.OpenConnection();
                    temp_dbControl.UpdateFragmentElementIntoXML(element);
                    temp_dbControl.CloseConnection();
                    break;
                case ElementType.Note:
                    break;
            };
        }

        public void HideElement(Element element)
        {
            if (element.IsCommandNote)
            {
                return;
            }

            if (element.ElementAboveUnderSameParent != null &&
                element.ElementAboveUnderSameParent.IsCommandNote &&
                element.ElementAboveUnderSameParent.Command == ElementCommand.DisplayMoreAssociations)
            {
                int hiddenAssociationCount = Int32.Parse(element.ElementAboveUnderSameParent.NoteText.Split(' ')[10]);
                element.ElementAboveUnderSameParent.NoteText = String.Format("Click the icon on the right to show the next {0} association(s)", ++hiddenAssociationCount);
            }
            else if (element.ElementBelowUnderSameParent != null &&
                element.ElementBelowUnderSameParent.IsCommandNote &&
                element.ElementBelowUnderSameParent.Command == ElementCommand.DisplayMoreAssociations)
            {
                int hiddenAssociationCount = Int32.Parse(element.ElementBelowUnderSameParent.NoteText.Split(' ')[10]);
                element.ElementBelowUnderSameParent.NoteText = String.Format("Click the icon on the right to show the next {0} association(s)", ++hiddenAssociationCount);
            }
            else 
            {
                int hiddenAssociationCount = 0;
                Element commandNote = new Element();
                commandNote.Type = ElementType.Note;
                commandNote.Command = ElementCommand.DisplayMoreAssociations;
                commandNote.Status = ElementStatus.Special;
                commandNote.TailImageSource = String.Format("pack://application:,,,/{0};component/{1}", "Planz", "Images/command.png");
                commandNote.HasAssociation = true;
                commandNote.CanOpen = false;
                commandNote.CanExplore = false;
                commandNote.CanRename = false;
                commandNote.CanDelete = false;
                commandNote.NoteText = String.Format("Click the icon on the right to show the next {0} association(s)", ++hiddenAssociationCount);

                InsertElement(commandNote, element.ParentElement, element.Position);
            }

            element.IsVisible = Visibility.Collapsed;

            DatabaseControl temp_dbControl = new DatabaseControl(element.ParentElement.Path);
            temp_dbControl.OpenConnection();
            temp_dbControl.UpdateElementIntoXML(element);
            temp_dbControl.CloseConnection();
        }

        public void DeleteElement(Element element)
        {
            try
            {
                // Note with association AND remote Heading
                if (element.Status != ElementStatus.MissingAssociation &&
                    (element.IsRemoteHeading || 
                    (element.IsNote && element.HasAssociation)))
                {
                    MoveAssociationToRecycleBin(element.AssociationURIFullPath);
                }
                
                if (element.IsLocalHeading)
                {
                    MoveFolderToRecycleBin(element.Path);
                }
                RemoveElement(element, element.ParentElement);
            }
            catch (Exception ex)
            {
                //string message = "DeleteElement\n" + ex.Message;
                //OnOperationError(element, new OperationErrorEventArgs{ Message = message });
            }
            finally
            {
                
            }
        }

        private void MoveAssociationToRecycleBin(string fileFullName)
        {
            FileSystem.DeleteFile(fileFullName, UIOption.OnlyErrorDialogs, RecycleOption.SendToRecycleBin);
        }

        private void MoveFolderToRecycleBin(string fileFullName)
        {
            FileSystem.DeleteDirectory(fileFullName, UIOption.OnlyErrorDialogs, RecycleOption.SendToRecycleBin);
        }

        private void MoveFile(Element element, string newPath)
        {
            switch (element.Type)
            {
                case ElementType.Heading:
                    if (element.IsRemoteHeading)
                    {
                        string srcPath = element.AssociationURIFullPath;
                        string destPath = newPath + element.AssociationURI;
                        System.IO.File.Move(srcPath, destPath);
                    }
                    else
                    {
                        System.IO.Directory.Move(element.Path, newPath + HeadingNameConverter.ConvertFromHeadingNameToFolderName(element) + System.IO.Path.DirectorySeparatorChar);
                    }
                    break;
                case ElementType.Note:
                    if (element.HasAssociation)
                    {
                        string srcPath = element.AssociationURIFullPath;
                        string destPath = newPath + element.AssociationURI;
                        System.IO.File.Move(srcPath, destPath);
                    }
                    break;
            };
        }

        #endregion

        #region ICC and D&L

        public void AddAssociation(Element element, string fileFullName, ElementAssociationType type, string text)
        {
            string noteText = String.Empty;
            string shortcutName = String.Empty;
            string folderPath = element.Path;
            string title = string.Empty;

            List<Element> emailAttachmentElementList = new List<Element>();

            if (text != null)
            {
                noteText = text;
                if (noteText.Length > StartProcess.MAX_EXTRACTTEXT_LENGTH)
                {
                    noteText = noteText.Substring(0, StartProcess.MAX_EXTRACTTEXT_LENGTH) + "...";
                }
                title = noteText.Replace("/", "").Replace("\\", "").Replace("*", "").Replace("?", "").Replace("\"", "").Replace("<", "").Replace(">", "").Replace("|", "").Replace(":", "");
                if (title.Length > StartProcess.MAX_EXTRACTNAME_LENGTH)
                    title = title.Substring(0, StartProcess.MAX_EXTRACTNAME_LENGTH);
                while (title.EndsWith("."))
                    title.TrimEnd('.');
            }

            if (type == ElementAssociationType.File)
            {

            }
            else if (type == ElementAssociationType.FileShortcut)
            {
                if (text == null)
                {
                    noteText = noteText = System.IO.Path.GetFileNameWithoutExtension(fileFullName);
                    title = noteText;
                }
                
                string fs_fileName = title + System.IO.Path.GetExtension(fileFullName);
                shortcutName = ShortcutNameConverter.GenerateShortcutNameFromFileName(fs_fileName, folderPath);
            }
            else if (type == ElementAssociationType.FolderShortcut)
            {
                if (fileFullName.EndsWith(System.IO.Path.DirectorySeparatorChar.ToString()))
                {
                    fileFullName = fileFullName.Substring(0, fileFullName.Length - 1);
                }
                noteText = System.IO.Path.GetFileName(fileFullName);

                string fs_fileName = noteText;
                shortcutName = ShortcutNameConverter.GenerateShortcutNameFromFileName(fs_fileName, folderPath);
            }
            else if (type == ElementAssociationType.Web)
            {
                if (text == null)
                {
                    ActiveWindow activeWindow = new ActiveWindow();
                    title = activeWindow.GetActiveWindowText(activeWindow.GetActiveWindowHandle());
                    if(title.Contains(" - Windows Internet Explorer"))
                    {
                        // IE 
                        int labelIndex1 = title.LastIndexOf(" - Windows Internet Explorer");
                        if (labelIndex1 != -1)
                        {
                            title = title.Remove(labelIndex1);
                            noteText = title;
                        }
                    }else if(title.Contains(" - Mozilla Firefox"))
                    {
                        // Firefox
                        int labelIndex2 = title.LastIndexOf(" - Mozilla Firefox");
                        if (labelIndex2 != -1)
                        {
                            title = title.Remove(labelIndex2);
                            noteText = title;
                        }
                    }else
                    {
                        noteText = fileFullName;
                        title = string.Empty;
                    }

                    if (noteText.Length > StartProcess.MAX_EXTRACTTEXT_LENGTH)
                    {
                        noteText = noteText.Substring(0, StartProcess.MAX_EXTRACTTEXT_LENGTH) + "...";
                    }
                }
                shortcutName = ShortcutNameConverter.GenerateShortcutNameFromWebTitle(title, folderPath);
         
            }
            else if (type == ElementAssociationType.Email)
            {
                Outlook.Application outlookApp = new Outlook.Application();
                Outlook.MailItem mailItem = null;
                if (outlookApp.ActiveExplorer().Selection.Count > 0)
                {
                    mailItem = outlookApp.ActiveExplorer().Selection[1] as Outlook.MailItem;

                    if (mailItem == null)
                        return;

                    if (mailItem != null)
                    {
                        Element associatedElement = element;
                        if (element.IsHeading && element.IsCollapsed)
                        {
                            associatedElement = element.ParentElement;
                        }

                        Regex byproduct = new Regex("att[0-9]+[.txt|.c]");

                        foreach (Outlook.Attachment attachment in mailItem.Attachments)
                        {
                            string attachmentFileName = attachment.FileName;

                            if (!byproduct.IsMatch(attachmentFileName.ToLower()))
                            {
                                string fileNameWithoutExt = System.IO.Path.GetFileNameWithoutExtension(attachmentFileName);
                                string fileNameExt = System.IO.Path.GetExtension(attachmentFileName);
                                string copyPath = associatedElement.Path + attachmentFileName;
                                int index = 2;
                                while (FileNameChecker.Exist(copyPath))
                                {
                                    copyPath = associatedElement.Path + fileNameWithoutExt + " (" + index.ToString() + ")" + fileNameExt;
                                    index++;
                                }
                                attachment.SaveAsFile(copyPath);

                                Element attachmentElement = CreateNewElement(ElementType.Note, " --- " + fileNameWithoutExt);
                                attachmentElement.AssociationType = ElementAssociationType.File;
                                attachmentElement.AssociationURI = System.IO.Path.GetFileName(copyPath);
                                attachmentElement.TailImageSource = FileTypeHandler.GetIcon((ElementAssociationType)attachmentElement.AssociationType, copyPath);
                                emailAttachmentElementList.Add(attachmentElement);
                            }
                        }
                    }
                }
                if ((fileFullName == null) && (mailItem != null))
                {
                    noteText = mailItem.Subject;
                    if (noteText.Length > StartProcess.MAX_EXTRACTTEXT_LENGTH)
                    {
                        noteText = noteText.Substring(0, StartProcess.MAX_EXTRACTTEXT_LENGTH) + "...";
                    }
                    fileFullName = mailItem.EntryID;
                    shortcutName = ShortcutNameConverter.GenerateShortcutNameFromEmailSubject(mailItem.Subject, folderPath);
                }
                else
                {
                    shortcutName = ShortcutNameConverter.GenerateShortcutNameFromEmailSubject(title, folderPath);
                }
            }

            switch (element.Type)
            {
                case ElementType.Heading:
                    if (element.IsExpanded)
                    {
                        Element firstElement = CreateNewElement(ElementType.Note, noteText);
                        InsertElement(firstElement, element, 0);
                        AssignAssociationInfo(firstElement, fileFullName, shortcutName, type);
                        currentElement = firstElement;
                    }
                    else
                    {
                        Element newElement = CreateNewElement(ElementType.Note, noteText);
                        InsertElement(newElement, element.ParentElement, element.Position + 1);
                        AssignAssociationInfo(newElement, fileFullName, shortcutName, type);
                        currentElement = newElement;
                    }
                    break;
                case ElementType.Note:
                    if (element.HasAssociation == false)
                    {
                        if (element.NoteText.Trim() == String.Empty)
                        {
                            element.NoteText = noteText;
                        }
                        AssignAssociationInfo(element, fileFullName, shortcutName, type);
                        currentElement = element;
                    }
                    else
                    {
                        Element newElement = CreateNewElement(ElementType.Note, noteText);
                        InsertElement(newElement, element.ParentElement, element.Position + 1);
                        AssignAssociationInfo(newElement, fileFullName, shortcutName, type);
                        currentElement = newElement;
                    }
                    break;
            };

            int curr_index = currentElement.Position;
            foreach (Element emailAttachmentElement in emailAttachmentElementList)
            {
                InsertElement(emailAttachmentElement, currentElement.ParentElement, curr_index);
                curr_index++;
            }
        }

        public void AddAppointments(Element element, DateTime dt)
        {
            Outlook.Application outlookApp = new Outlook.Application();

            // Get the NameSpace and Logon information.
            Outlook.NameSpace outlookNS = outlookApp.GetNamespace("mapi");

            //Log on by using a dialog box to choose the profile.
            outlookNS.Logon(Missing.Value, Missing.Value, true, true);

            // Lets Get The Calendar folder.
            Outlook.MAPIFolder outlookCal = outlookNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);

            // Get the Items (Appointments) collection from the Calendar folder.
            Outlook.Items appointItems = outlookCal.Items;

            // Set start value
            DateTime startDt = 
                new DateTime(dt.Year, dt.Month, dt.Day, 0, 0, 0);
            // Set end value
            DateTime endDt =
                new DateTime(dt.Year, dt.Month, dt.Day, 23, 59, 59);
            // Initial restriction is Jet query for date range
            string sCriteria = "[Start] >= '" +
                startDt.ToString("g")
                + "' AND [End] <= '" +
                endDt.ToString("g") + "'";

            //Use the Restrict method to reduce the number of items to process.
            Outlook.Items restrictedItems = appointItems.Restrict(sCriteria);

            restrictedItems.Sort("[Start]", Type.Missing);
            restrictedItems.IncludeRecurrences = true;

            //Get each item until item null.
            Outlook.AppointmentItem appointItem;
            //Preferably the first Item
            appointItem = (Outlook.AppointmentItem) restrictedItems.GetFirst();

            if (appointItem == null)
            {
                MessageBox.Show("There are no meetings or appointments in the Outlook calendar to import for this date.");
            }

            int count =0;
            foreach (Element ele in element.Elements)
            {
                if (ele.IsCommandNote == false)
                    break;
                count++;
            }

            while (appointItem != null)
            {
                //create heading for each appoints
                string newNoteText = appointItem.Start.ToShortTimeString() + " " + appointItem.Subject.ToString();

                if(newNoteText.Length > StartProcess.MAX_EXTRACTNAME_LENGTH)
                    newNoteText = newNoteText.Substring(0, StartProcess.MAX_EXTRACTNAME_LENGTH);                   
                
                bool bDuplicate = false;

                Element dupEle = element.FirstChild;

                bool bOverwrite = false;
                int index = 2;
                                
                //Check whether there is duplicate appointments with same heading
                foreach (Element ele in FindAllHeadingElements(element))
                {
                    if (newNoteText == ele.NoteText)
                    {
                        //System.Windows.MessageBox.Show("Duplicate appointments found!\n");
                        dupEle = ele;
                        bDuplicate = true;
                        break;
                    }
                }
                System.Windows.MessageBoxButton buttons = System.Windows.MessageBoxButton.YesNo;
                string message = "The appointment has already existed,do you want to overwrite it?";
                string caption = newNoteText;

                //if duplicate appointment exists and user don't want to merge them, skip it
                if ((bDuplicate == true) && (System.Windows.MessageBox.Show(message, caption, buttons, System.Windows.MessageBoxImage.Warning, System.Windows.MessageBoxResult.Yes) == System.Windows.MessageBoxResult.No))
                {
                    appointItem = (Outlook.AppointmentItem)restrictedItems.GetNext();
                    continue;
                }
                Element appointEle = dupEle;
                if( bDuplicate== true)
                {
                    if (dupEle.IsExpanded)
                    {
                        //overwrite the current appointments
                        int remain = 0;
                        while (dupEle.Elements.Count > remain)
                        {
                            if (dupEle.FirstChild.AssociationType == ElementAssociationType.File)
                            {
                                remain++;
                                dupEle.Elements.Move(0, dupEle.Elements.Count - 1);
                                continue;
                            }
                            DeleteElement(dupEle.FirstChild);
                        }
                    }
                    else
                    {
                        DatabaseControl temp_dbControl = new DatabaseControl(dupEle.Path);
                        temp_dbControl.OpenConnection();
                        List<Element> eleList = temp_dbControl.GetAllElementFromXML();
                        foreach (Element ele in eleList)
                        {

                            if (ele.AssociationType != ElementAssociationType.File)
                            {
                                ele.ParentElement = dupEle;
                                dupEle.Elements.Add(ele);
                                DeleteElement(ele);
                            }
                        }
                        temp_dbControl.CloseConnection();
                    }

                    UpdateElement(dupEle);

                    appointEle = dupEle;
                    appointEle.Status = ElementStatus.New;
                    appointEle.FontColor = ElementColor.Blue.ToString();
                    appointEle.Position = count++;
                    UpdateElement(appointEle);
                }else
                {
                    //create new appointment
                    appointEle = CreateNewElement(ElementType.Heading, newNoteText);
                    appointEle.Status = ElementStatus.New;
                    appointEle.FontColor = ElementColor.Blue.ToString();
                    InsertElement(appointEle, element, count++);
                    UpdateElement(element);
                }

                //add subject
                Element subEle = CreateNewElement(ElementType.Note, string.Empty);
                subEle.FontColor = ElementColor.Blue.ToString();
                subEle.Status = ElementStatus.New;
                subEle.NoteText = "<no subject>";
                if (appointItem.Subject != null)
                    subEle.NoteText = "Subject: " + appointItem.Subject.ToString();

                if (subEle.NoteText.Length > StartProcess.MAX_EXTRACTTEXT_LENGTH)
                    subEle.NoteText = subEle.NoteText.Substring(0, StartProcess.MAX_EXTRACTTEXT_LENGTH) + "...";

                InsertElement(subEle, appointEle, appointEle.Elements.Count);


                //add location
                Element locEle = CreateNewElement(ElementType.Note, string.Empty);
                locEle.Status = ElementStatus.New;
                locEle.FontColor = ElementColor.Blue.ToString();

                locEle.NoteText = "<no location>";
                if (appointItem.Location != null)
                    locEle.NoteText = "Location: " + appointItem.Location.Trim().ToString();
                InsertElement(locEle, appointEle, appointEle.Elements.Count);

                //add start time
                Element startEle = CreateNewElement(ElementType.Note, string.Empty);
                startEle.FontColor = ElementColor.Blue.ToString();
                startEle.Status = ElementStatus.New;
                startEle.NoteText = "<no start time>";
                if (appointItem.Start != null)
                    startEle.NoteText = "Start time: " + appointItem.Start.ToShortTimeString();
                InsertElement(startEle, appointEle, appointEle.Elements.Count);

                //add end time
                Element endEle = CreateNewElement(ElementType.Note, string.Empty);
                endEle.Status = ElementStatus.New;
                endEle.FontColor = ElementColor.Blue.ToString();
                endEle.NoteText = "<no end time>";
                if (appointItem.End != null)
                    endEle.NoteText = "End Time: " + appointItem.End.ToShortTimeString();
                InsertElement(endEle, appointEle, appointEle.Elements.Count);

                //add attachments
                foreach (Outlook.Attachment attachment in appointItem.Attachments)
                {
                    string attachmentFileName = attachment.FileName;

                    Regex byproduct = new Regex("att[0-9]+[.txt|.c]");

                    if (!byproduct.IsMatch(attachmentFileName.ToLower()))
                    {
                        string fileNameWithoutExt = System.IO.Path.GetFileNameWithoutExtension(attachmentFileName);
                        string fileNameExt = System.IO.Path.GetExtension(attachmentFileName);
                        string copyPath = appointEle.Path + attachmentFileName;

                        message = "The attachment has already existed,do you want to overwriet it?";

                        bOverwrite = true;
                        index = 2;
                        while (FileNameChecker.Exist(copyPath))
                        {
                            caption = attachmentFileName;
                            if (System.Windows.MessageBox.Show(message, caption, buttons, System.Windows.MessageBoxImage.Warning, System.Windows.MessageBoxResult.Yes) == System.Windows.MessageBoxResult.Yes)
                            {
                                bOverwrite = true;
                                break;
                            }
                            copyPath = appointEle.Path + fileNameWithoutExt + " (" + index.ToString() + ")" + fileNameExt;
                            attachmentFileName = fileNameWithoutExt + " (" + index.ToString() + ")" + fileNameExt;
                            index++;
                        }

                        attachment.SaveAsFile(copyPath);

                        if (bOverwrite == false)
                        {
                            Element attachmentElement = CreateNewElement(ElementType.Note, " --- " + fileNameWithoutExt);
                            attachmentElement.FontColor = ElementColor.Blue.ToString();
                            attachmentElement.Status = ElementStatus.New;
                            attachmentElement.AssociationType = ElementAssociationType.File;
                            attachmentElement.AssociationURI = System.IO.Path.GetFileName(copyPath);
                            attachmentElement.TailImageSource = FileTypeHandler.GetIcon((ElementAssociationType)attachmentElement.AssociationType, copyPath);
                            InsertElement(attachmentElement, appointEle, appointEle.Elements.Count);
                        }
                        else
                            appointEle.Elements.Move(0, appointEle.Elements.Count - 1);
                    }
                }

                //add body text

                Element bodyEle = CreateNewElement(ElementType.Note, string.Empty);
                bodyEle.Status = ElementStatus.New;
                bodyEle.FontColor = ElementColor.Blue.ToString();

                bodyEle.NoteText = "<no message>";
                if ((appointItem.Body != null) && (appointItem.Body.Trim() != string.Empty))
                {
                    bodyEle.NoteText = appointItem.Body.Replace("\r\n", " ").Replace("\t"," ").Trim();
                    if(bodyEle.NoteText.Length>StartProcess.MAX_EXTRACTTEXT_LENGTH)
                        bodyEle.NoteText = bodyEle.NoteText.Substring(0, StartProcess.MAX_EXTRACTTEXT_LENGTH) + "...";
                }
                //add the uri to appoint item in outlook
                InsertElement(bodyEle, appointEle, appointEle.Elements.Count);

                string fileFullName = appointItem.EntryID;

                
                ElementAssociationType type = ElementAssociationType.Appointment;
                string folderPath = appointEle.Path;

               
                string fileName = "<no subject>";
                if (appointItem.Subject != null)
                {
                    fileName = appointItem.Subject.ToString();
                    if (appointItem.Subject.Length > StartProcess.MAX_EXTRACTNAME_LENGTH)
                        fileName = appointItem.Subject.Substring(0, StartProcess.MAX_EXTRACTNAME_LENGTH);
                }

                string shortcutName = ShortcutNameConverter.GenerateShortcutNameFromAppointmentSubject(fileName, folderPath);
                AssignAssociationInfo(bodyEle, fileFullName, shortcutName, type);

                //appointEle.IsExpanded = true;
                UpdateElement(appointEle);

                appointItem = (Outlook.AppointmentItem)restrictedItems.GetNext();
            }

            // Log off
            outlookNS.Logoff();

            // Clean up
            appointItem = null;
            restrictedItems = null;
            appointItems = null;
            outlookCal = null;
            outlookNS = null;
            outlookApp = null;
        }

        public void RemoveAssociation(Element element)
        {
            try
            {
                MoveAssociationToRecycleBin(element.AssociationURIFullPath);
            }
            catch (Exception)
            {

            }
            finally
            {
                if (element.Status == ElementStatus.MissingAssociation)
                {
                    element.Status = ElementStatus.Normal;
                    if (element.FontColor == ElementColor.Gray.ToString())
                    {
                        element.FontColor = ElementColor.Black.ToString();
                    }
                }

                element.AssociationType = ElementAssociationType.None;
                element.AssociationURI = String.Empty;

                UpdateElement(element);
            }
        }

        private void AssignAssociationInfo(Element element, string fileFullName, string shortcutName, ElementAssociationType type)
        {
            string element_path = element.Path;
            string shortcut_path = element_path + shortcutName;

            element.AssociationType = type;
            element.AssociationURI = shortcutName;

            CreateShortcut(element, fileFullName, shortcut_path);

            element.TailImageSource = FileTypeHandler.GetIcon((ElementAssociationType)element.AssociationType, element.AssociationURIFullPath);

            UpdateElement(element);
        }

        private void CreateShortcut(Element element, string fileFullName, string shortcutPath)
        {
            try
            {
                WshShellClass wshShell = new WshShellClass();
                IWshRuntimeLibrary.IWshShortcut shortcut = (IWshRuntimeLibrary.IWshShortcut)wshShell.CreateShortcut(shortcutPath);

                switch (element.AssociationType)
                {
                    case ElementAssociationType.FileShortcut:
                    case ElementAssociationType.FolderShortcut:
                    case ElementAssociationType.Web:
                    case ElementAssociationType.Tweet:
                        shortcut.TargetPath = fileFullName;
                        shortcut.Description = fileFullName;
                        break;
                    case ElementAssociationType.Email:
                    case ElementAssociationType.Appointment:
                        string targetPath = OutlookSupportFunction.GenerateShortcutTargetPath();
                        shortcut.TargetPath = targetPath;
                        shortcut.Arguments = "/select outlook:" + fileFullName;
                        shortcut.Description = targetPath;
                        break;
                    default:
                        break;
                };

                shortcut.Save();
            }
            catch (Exception ex)
            {
                string message = "Oops!\nSomething unexpected happened, please close this Planz window and reopen.";

                //string message = "CreateShortcut\n" + ex.Message;
                OnOperationError(element, new OperationErrorEventArgs { Message = message });
            }
            finally
            {

            }
        }

        public IWshRuntimeLibrary.IWshShortcut GetShortcut(Element element)
        {
            try
            {
                if (System.IO.File.Exists(element.AssociationURIFullPath))
                {
                    WshShellClass wshShell = new WshShellClass();
                    IWshRuntimeLibrary.IWshShortcut shortcut = (IWshRuntimeLibrary.IWshShortcut)wshShell.CreateShortcut(element.AssociationURIFullPath);
                    return shortcut;
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

        public FileInfo GetFileInfo(Element element)
        {
            try
            {
                if (System.IO.File.Exists(element.AssociationURIFullPath))
                {
                    return new FileInfo(element.AssociationURIFullPath);
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

        public string GetFlagDescription(Element element)
        {
            string description = String.Empty;
            
            if (element.FlagStatus == FlagStatus.Flag)
            {
                if (element.StartDate != new DateTime())
                {
                    description += "Start Date: " + element.StartDate.ToShortDateString() + " " + element.StartDate.ToShortTimeString() + "\r\n";
                }
                if (element.DueDate != new DateTime())
                {
                    description += "Due Date: " + element.DueDate.ToShortDateString() + " " + element.DueDate.ToShortTimeString() + "\r\n"; 
                }
                description += "Check as Completed";
            }
            else if (element.FlagStatus == FlagStatus.Check)
            {
                description += "Reset";
            }
            else
            {
                description += "Flag as To-do";
            }

            return description;
        }

        public string GetAssociationDescription(Element element)
        {
            try
            {
                string description = String.Empty;
                string missingFileMessage = "File no longer exist";

                switch (element.AssociationType)
                {
                    case ElementAssociationType.File:
                        FileInfo fi = GetFileInfo(element);
                        if (fi != null)
                        {
                            description = String.Format("{0, -15}{1, 0}\n", "Name: ", fi.Name) +
                                String.Format("{0, -15}{1, 0}\n", "Location: ", fi.Directory.FullName) +
                                String.Format("{0, -15}{1, 0}\n", "Created: ", fi.CreationTime.ToShortDateString() + " " + fi.CreationTime.ToShortTimeString()) +
                                String.Format("{0, -15}{1, 0}\n", "Modified: ", fi.LastWriteTime.ToShortDateString() + " " + fi.LastWriteTime.ToShortTimeString()) +
                                String.Format("{0, -15}{1, 0}\n", "Accessed: ", fi.LastAccessTime.ToShortDateString() + " " + fi.LastAccessTime.ToShortTimeString()) +
                                String.Format("{0, -15}{1, 0}", "Size: ", (fi.Length / 1024).ToString() + "KB");
                        }
                        else
                        {
                            throw new Exception(missingFileMessage);
                        }
                        break;
                    case ElementAssociationType.Folder:
                        DirectoryInfo di = new DirectoryInfo(element.AssociationURIFullPath);
                        description = String.Format("{0, -15}{1, 0}\n", "Location: ", di.FullName) +
                               String.Format("{0, -15}{1, 0}\n", "Created: ", di.CreationTime.ToShortDateString() + " " + di.CreationTime.ToShortTimeString()) +
                               String.Format("{0, -15}{1, 0}\n", "Modified: ", di.LastWriteTime.ToShortDateString() + " " + di.LastWriteTime.ToShortTimeString()) +
                               String.Format("{0, -15}{1, 0}", "Accessed: ", di.LastAccessTime.ToShortDateString() + " " + di.LastAccessTime.ToShortTimeString());
                        break;
                    case ElementAssociationType.FileShortcut:
                    case ElementAssociationType.FolderShortcut:
                    case ElementAssociationType.Web:
                    case ElementAssociationType.Tweet:
                        IWshRuntimeLibrary.IWshShortcut shortcut = GetShortcut(element);
                        FileInfo fis = new FileInfo(shortcut.FullName);
                        if (shortcut != null)
                        {
                            string location;
                            if (shortcut.TargetPath != String.Empty)
                            {
                                location = shortcut.TargetPath;
                            }
                            else
                            {
                                location = shortcut.Description;
                            }
                            description = String.Format("{0, -15}{1, 0}\n", "Location: ", location) +
                                String.Format("{0, -15}{1, 0}\n", "Created: ", fis.CreationTime.ToShortDateString() + " " + fis.CreationTime.ToShortTimeString()) +
                                String.Format("{0, -15}{1, 0}\n", "Modified: ", fis.LastWriteTime.ToShortDateString() + " " + fis.LastWriteTime.ToShortTimeString()) +
                                String.Format("{0, -15}{1, 0}", "Accessed: ", fis.LastAccessTime.ToShortDateString() + " " + fis.LastAccessTime.ToShortTimeString());
                        }
                        else
                        {
                            throw new Exception(missingFileMessage);
                        }
                        break;
                    case ElementAssociationType.Email:
                        description = "Email";
                        break;
                    case ElementAssociationType.Appointment:
                        description = "Appointment";
                        break;
                };

                return description;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        public void ICC(Element element, Object obj, bool insertBelow, ICCAssociationType type)
        {
            Element targetElement = null;

            try
            {
                string fileName = String.Empty;
                if (obj is Tweet)
                {
                    fileName = ((Tweet)obj).Message;
                }
                else if (obj is String)
                {
                    fileName = (string)obj;
                }

                string fileFullName = element.Path + ICCFileNameHandler.GenerateFileName(fileName, type);

                switch (element.Type)
                {
                    case ElementType.Heading:
                        if (element.IsExpanded)
                        {
                            targetElement = CreateNewElement(ElementType.Note, fileName);
                            InsertElement(targetElement, element, 0);
                        }
                        else
                        {
                            fileFullName = element.ParentElement.Path + ICCFileNameHandler.GenerateFileName(fileName, type);
                            if (insertBelow)
                            {
                                targetElement = CreateNewElement(ElementType.Note, fileName);
                                InsertElement(targetElement, element.ParentElement, element.Position + 1);
                            }
                            else
                            {
                                targetElement = CreateNewElement(ElementType.Note, fileName);
                                InsertElement(targetElement, element.ParentElement, element.Position);
                            }
                        }
                        break;
                    case ElementType.Note:
                        if (element.HasAssociation == false)
                        {
                            if (insertBelow)
                            {
                                if (element.NoteText.Trim() == String.Empty)
                                {
                                    element.NoteText = fileName;
                                }
                                targetElement = element;
                            }
                            else
                            {
                                targetElement = CreateNewElement(ElementType.Note, fileName);
                                InsertElement(targetElement, element.ParentElement, element.Position);
                            }
                        }
                        else
                        {
                            if (insertBelow)
                            {
                                targetElement = CreateNewElement(ElementType.Note, fileName);
                                InsertElement(targetElement, element.ParentElement, element.Position + 1);
                            }
                            else
                            {
                                targetElement = CreateNewElement(ElementType.Note, fileName);
                                InsertElement(targetElement, element.ParentElement, element.Position);
                            }
                        }
                        break;
                };

                switch (type)
                {
                    case ICCAssociationType.OutlookEmailMessage:
                        targetElement.AssociationType = ElementAssociationType.Email;
                        if (element.AssociationType == ElementAssociationType.File)
                        {
                            targetElement.Status = ElementStatus.TempStatus_HasEmailAttachment;
                            targetElement.TempData = element.AssociationURIFullPath;
                        }

                        string folderPath = System.IO.Directory.GetParent(fileFullName).FullName + System.IO.Path.DirectorySeparatorChar;
                        string fileNameWithoutExt = System.IO.Path.GetFileNameWithoutExtension(fileFullName);
                        string ext = ShortcutNameConverter.mailLinkExt;
                        string newShortcutName = ShortcutNameConverter.RenameShortcutName(fileNameWithoutExt, ext, folderPath);
                        fileFullName = folderPath + newShortcutName;
                        if (fileNameWithoutExt + ext != newShortcutName)
                        {
                            targetElement.AssociationURI = newShortcutName;
                        }

                        break;
                    case ICCAssociationType.TwitterUpdate:
                        targetElement.AssociationType = ElementAssociationType.Tweet;
                        break;
                    default:
                        targetElement.AssociationType = ElementAssociationType.File;
                        break;
                };

                if (type == ICCAssociationType.TwitterUpdate)
                {
                    CreateTwitterUpdate(targetElement, fileFullName, obj as Tweet);
                }
                else
                {
                    CreateAndOpenICCFile(targetElement, fileFullName, type);
                }

                AssignICCInfo(targetElement, fileFullName);

                currentElement = targetElement;
            }
            catch (Exception ex)
            {
                targetElement.AssociationType = ElementAssociationType.None;
                targetElement.AssociationURI = String.Empty;
                targetElement.TailImageSource = String.Empty;

                UpdateElement(targetElement);

                string message = "ICC\n" + ex.Message;
                OnOperationError(targetElement, new OperationErrorEventArgs { Message = message });
            }
        }

        private void CreateAndOpenICCFile(Element element, string fileFullName, ICCAssociationType type)
        {
            switch (type)
            {
                case ICCAssociationType.NotepadTextDocument:
                    /*TextWriter tw = new StreamWriter(fileFullName);
                    tw.Close();
                    System.Diagnostics.ProcessStartInfo psi = new System.Diagnostics.ProcessStartInfo(fileFullName);
                    psi.ErrorDialog = false;
                    psi.UseShellExecute = true;
                    System.Diagnostics.Process.Start(psi);*/
                    System.IO.File.Copy(FileTypeHandler.GetTemplate(ICCAssociationType.NotepadTextDocument), fileFullName);
                    System.Diagnostics.Process.Start(fileFullName);
                    break;
                case ICCAssociationType.WordDocument:
                    /*Word.Application wordApp = new Word.Application();
                    wordApp.Visible = true;
                    object oMissing = System.Reflection.Missing.Value;
                    Word.Document wordDoc = wordApp.Documents.Add(ref oMissing, ref oMissing, ref oMissing, ref oMissing);
                    object oSavePath = fileFullName;
                    wordDoc.SaveAs(ref oSavePath,
                        ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
                    wordApp = null;*/
                    System.IO.File.Copy(FileTypeHandler.GetTemplate(ICCAssociationType.WordDocument), fileFullName);
                    System.Diagnostics.Process.Start(fileFullName);
                    break;
                case ICCAssociationType.ExcelSpreadsheet:
                    /*Excel.Application excelApp = new Excel.Application();
                    excelApp.Visible = true;
                    Excel.Workbooks workbooks = excelApp.Workbooks;
                    Excel._Workbook workbook = workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                    Excel._Worksheet worksheet = (Excel._Worksheet)workbook.ActiveSheet;
                    worksheet.SaveAs(fileFullName, 
                        Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                    excelApp = null;*/
                    System.IO.File.Copy(FileTypeHandler.GetTemplate(ICCAssociationType.ExcelSpreadsheet), fileFullName);
                    System.Diagnostics.Process.Start(fileFullName);
                    break;
                case ICCAssociationType.PowerPointPresentation:
                    /*PowerPoint.Application pptApp = new PowerPoint.Application();
                    pptApp.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                    PowerPoint.Presentations presentations = pptApp.Presentations;
                    PowerPoint._Presentation presentation = presentations.Add(Microsoft.Office.Core.MsoTriState.msoTrue);
                    presentation.SaveAs(fileFullName, 
                        Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsDefault, Microsoft.Office.Core.MsoTriState.msoTrue);
                    pptApp = null;*/
                    System.IO.File.Copy(FileTypeHandler.GetTemplate(ICCAssociationType.PowerPointPresentation), fileFullName);
                    System.Diagnostics.Process.Start(fileFullName);
                    break;
                case ICCAssociationType.OutlookEmailMessage:
                    Outlook.Application olApp = new Outlook.Application();
                    Outlook.MailItem mail = olApp.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem) as Outlook.MailItem;
                    mail.Subject = System.IO.Path.GetFileNameWithoutExtension(fileFullName);
                    /*mail.Body += "\n\n\n----------\n" +
                        "This email was sent from Planz(TM).\n" +
                        "To find out more, see a video or download, visit: http://kftf.ischool.washington.edu/planner_index.htm";*/
                    if (element.Status == ElementStatus.TempStatus_HasEmailAttachment)
                    {
                        mail.Attachments.Add(element.TempData, Missing.Value, Missing.Value, Missing.Value);
                        element.Status = ElementStatus.Normal;
                        element.TempData = String.Empty;
                    }
                    Outlook.UserProperties props = mail.UserProperties;
                    Outlook.UserProperty prop = props.Add("PlanzID", Microsoft.Office.Interop.Outlook.OlUserPropertyType.olText, false, 0);
                    prop.Value = element.ID.ToString();
                    mail.Display(false);
                    mail.Save();
                    WshShellClass wshShell = new WshShellClass();
                    IWshRuntimeLibrary.IWshShortcut shortcut = (IWshRuntimeLibrary.IWshShortcut)wshShell.CreateShortcut(fileFullName);
                    string targetPath = OutlookSupportFunction.GenerateShortcutTargetPath();
                    shortcut.TargetPath = targetPath;
                    shortcut.Arguments = "/select outlook:" + mail.EntryID;
                    shortcut.Description = "PlanzICC";
                    shortcut.Save();
                    break;
                case ICCAssociationType.OneNoteSection:
                    /*OneNote.ApplicationClass onApp = new OneNote.ApplicationClass();
                    string sectionID;
                    string pageID;
                    //string outputXML;
                    onApp.OpenHierarchy(fileFullName, String.Empty, out sectionID, Microsoft.Office.Interop.OneNote.CreateFileType.cftSection);
                    //onApp.GetHierarchy(sectionID, Microsoft.Office.Interop.OneNote.HierarchyScope.hsSections, out outputXML);
                    onApp.CreateNewPage(sectionID, out pageID, Microsoft.Office.Interop.OneNote.NewPageStyle.npsBlankPageWithTitle);
                    onApp.NavigateTo(sectionID, pageID, true);*/
                    System.IO.File.Copy(FileTypeHandler.GetTemplate(ICCAssociationType.OneNoteSection), fileFullName);
                    System.Diagnostics.Process.Start(fileFullName);
                    break;
                default:
                    break;
            };
        }

        private void AssignICCInfo(Element element, string fileFullName)
        {
            string element_path = element.Path;

            element.AssociationURI = System.IO.Path.GetFileName(fileFullName);
            element.TailImageSource = FileTypeHandler.GetIcon((ElementAssociationType)element.AssociationType, element.AssociationURIFullPath);

            UpdateElement(element);
        }

        private void CreateTwitterUpdate(Element element, string fileFullName, Tweet tweet)
        {
            TwitterControl tc = new TwitterControl(tweet);
            if (tc.Update())
            {
                CreateShortcut(element, tc.GetTweetURL(), fileFullName);
            }
            else
            {
                throw new Exception("Failed to create Twitter update.");
            }
        }

        public void ReplyEmail(Element element)
        {
            // NOTICE! Only works for Inbox under Personal Folders of Outlook!
            string entryID = OutlookSupportFunction.GetOutlookEntryIDFromShortcut(element.AssociationURIFullPath);

            Element newElement = CreateNewElement(ElementType.Note, String.Empty);

            Outlook.Application olApp = new Outlook.Application();
            Outlook.NameSpace ns = olApp.GetNamespace("MAPI");
            ns.Logon(Missing.Value, Missing.Value, true, false);
            Outlook.MAPIFolder inboxFolder = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            Outlook.MailItem mail = (Outlook.MailItem)ns.GetItemFromID(entryID, inboxFolder.StoreID);
            Outlook.MailItem replyMail = mail.ReplyAll();
            Outlook.UserProperties props = replyMail.UserProperties;
            Outlook.UserProperty prop = props.Add("PlanzID", Microsoft.Office.Interop.Outlook.OlUserPropertyType.olText, false, 0);
            prop.Value = newElement.ID.ToString();
            replyMail.Display(false);
            replyMail.Save();

            string subject = replyMail.Subject;
            string fileFullName = element.Path + ShortcutNameConverter.GenerateShortcutNameFromEmailSubject(subject, element.Path);

            WshShellClass wshShell = new WshShellClass();
            IWshRuntimeLibrary.IWshShortcut shortcut = (IWshRuntimeLibrary.IWshShortcut)wshShell.CreateShortcut(fileFullName);
            string targetPath = OutlookSupportFunction.GenerateShortcutTargetPath();
            shortcut.TargetPath = targetPath;
            shortcut.Arguments = "/select outlook:" + replyMail.EntryID;
            shortcut.Description = "PlanzICC";
            shortcut.Save();

            newElement.NoteText = " --- " + replyMail.Subject;
            newElement.AssociationType = ElementAssociationType.Email;
            InsertElement(newElement, element.ParentElement, element.Position);

            AssignICCInfo(newElement, fileFullName);
        }

        public void RelatedMessages(Element element)
        {
            //System.Windows.MessageBox.Show("UI for conversation in outlook");
        }

        #endregion

        #region Folder Ops

        public bool CreateFolder(Element element)
        {
            bool isFolderCreated = false;
            if (element.ParentElement.Path != root.Path)
            {
                DatabaseControl temp_dbControl = new DatabaseControl(element.ParentElement.Path);
                temp_dbControl.OpenConnection();
                isFolderCreated = temp_dbControl.CreateFolder(element);
                temp_dbControl.CloseConnection();
            }
            else
            {
                isFolderCreated = dbControl.CreateFolder(element);
            }
            return isFolderCreated;
        }

        public void RemoveFolder(Element element)
        {
            // Remove folder permanently
            dbControl.RemoveFolder(element);
        }

        public void RenameFolder(Element element, string previousPath)
        {
            // If duplex folder name

            string currentFolderName = System.IO.Path.GetFileName(System.IO.Path.GetDirectoryName(element.Path));
            string previousFolderName = System.IO.Path.GetFileName(System.IO.Path.GetDirectoryName(previousPath));

            if (HeadingNameConverter.Exist(element.Path) && 
                previousPath.ToLower() != element.Path.ToLower() &&
                previousFolderName != currentFolderName)
            {
                int index = 2;
                string folderName = element.NoteText;
                while (Directory.Exists(element.Path))
                {
                    folderName = element.NoteText + " (" + index.ToString() + ")";
                    element.NoteText = folderName;
                    element.Path = element.ParentElement.Path + HeadingNameConverter.ConvertFromHeadingNameToFolderName(element) + System.IO.Path.DirectorySeparatorChar;
                    currentFolderName = System.IO.Path.GetFileName(System.IO.Path.GetDirectoryName(element.Path));
                    index++;
                }
                element.NoteText = folderName;
                element.AssociationURI = folderName;
                Directory.Move(previousPath, element.Path);
                UpdateElement(element);
            }

            dbControl.RenameFolder(element, previousPath);

            if (currentFolderName != previousFolderName)
            {
                LogControl.Write(
                    element,
                    LogEventAccess.NoteTextBox,
                    LogEventType.RenameHeading,
                    LogEventStatus.NULL,
                    LogEventInfo.PreviousName + LogControl.COMMA + previousFolderName + LogControl.DELIMITER +
                    LogEventInfo.NewName + LogControl.COMMA + currentFolderName);
            }
        }

        public void MoveFolder(Element element, string previousPath)
        {
            // If duplex folder name
            if (HeadingNameConverter.Exist(element.Path) && previousPath != element.Path)
            {
                int index = 2;
                string folderName = element.NoteText;
                while (Directory.Exists(element.Path))
                {
                    folderName = element.NoteText + " (" + index.ToString() + ")";
                    element.NoteText = folderName;
                    element.Path = element.ParentElement.Path + HeadingNameConverter.ConvertFromHeadingNameToFolderName(element) + System.IO.Path.DirectorySeparatorChar;
                    index++;
                }
                element.NoteText = folderName;
                Directory.Move(previousPath, element.Path);
                UpdateElement(element);
            }

            dbControl.MoveFolder(element, previousPath);
        }

        #endregion

        #region Save As and Export

        public void SaveAsTXT(string fileFullName)
        {
            StreamWriter sw = new StreamWriter(fileFullName);

            foreach (Element element in root.Elements)
            {
                ProcessSaveAsTXT(element, sw);
            }

            sw.Close();

            System.Diagnostics.Process.Start(fileFullName);
        }

        private void ProcessSaveAsTXT(Element element, StreamWriter sw)
        {
            sw.Write(element.NoteText);
            sw.Write(sw.NewLine);
            foreach (Element ele in element.Elements)
            {
                for (int i = 0; i < ele.Level; i++)
                {
                    sw.Write("  ");
                }
                

                if (ele.IsHeading && ele.IsExpanded)
                {
                    ProcessSaveAsTXT(ele, sw);
                }
                else
                {
                    sw.Write(ele.NoteText);
                    sw.Write(sw.NewLine);
                }
                
            }
        }

        public void SaveAsHTML(string fileFullName)
        {
            StreamWriter sw = new StreamWriter(fileFullName);

            sw.Write("<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.01 Transitional//EN\" \"http://www.w3.org/TR/html4/loose.dtd\">\n");
            sw.Write("<html xmlns=\"http://www.w3.org/1999/xhtml\">\n");
            sw.Write("<head>\n<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\" />\n");
            sw.Write("<title>" + System.IO.Path.GetFileName(root.Path) + "</title>\n</head>\n<body>\n");

            foreach (Element element in root.Elements)
            {
                ProcessSaveAsHTML(element, sw);
            }

            sw.Write("</body>\n</html>");

            sw.Close();

            System.Diagnostics.Process.Start(fileFullName);
        }

        private void ProcessSaveAsHTML(Element element, StreamWriter sw)
        {
            try
            {
                ProcessSaveAsHTMLElement(element, sw);

                sw.Write("<div style=\"margin-left: 40px;\">\n");
                foreach (Element ele in element.Elements)
                {
                    /*for (int i = 0; i < ele.Level; i++)
                    {
                        sw.Write("&nbsp;&nbsp;");
                    }*/

                    if (ele.IsHeading && ele.IsExpanded)
                    {
                        ProcessSaveAsHTML(ele, sw);
                    }
                    else
                    {
                        ProcessSaveAsHTMLElement(ele, sw);
                    }
                }
                sw.Write("</div>\n");
            }
            catch (Exception)
            {
                return;
            }
        }

        private void ProcessSaveAsHTMLElement(Element element, StreamWriter sw)
        {
            try
            {
                if (element.IsHeading)
                {

                    int headingLevel = element.Level;
                    string headingTagStart = "<font color=\"" + element.FontColor + "\" " +
                        "face=\"" + element.FontFamily + "\">" + "<h" + headingLevel + "><a href=\"file:///" +
                        element.Path + "\">";
                    string headngTagEnd = "</a></h" + headingLevel + "></font>\n";

                    sw.Write(headingTagStart + element.NoteText + headngTagEnd);
                }
                else if (element.HasAssociation)
                {
                    string url = String.Empty;
                    string association = String.Empty;
                    try
                    {
                        switch (element.AssociationType)
                        {
                            case ElementAssociationType.Email:
                                association = "Email";
                                url = element.AssociationURIFullPath;
                                break;
                            case ElementAssociationType.File:
                                association = "File";
                                url = element.AssociationURIFullPath;
                                break;
                            case ElementAssociationType.FileShortcut:
                                association = "File";
                                url = GetShortcut(element).TargetPath;
                                break;
                            case ElementAssociationType.FolderShortcut:
                                association = "Folder";
                                url = GetShortcut(element).TargetPath;
                                break;
                            case ElementAssociationType.Tweet:
                                association = "Tweet";
                                url = GetShortcut(element).Description;
                                break;
                            case ElementAssociationType.Web:
                                association = "Web";
                                url = GetShortcut(element).Description;
                                break;
                        }
                    }
                    catch (Exception)
                    {
                        url = String.Empty;
                        association = String.Empty;
                    }

                    sw.Write("<font color=\"" + element.FontColor + "\" " +
                        "face=\"" + element.FontFamily + "\">" +
                        element.NoteText.Replace("\r\n", "<br>\n").Replace(" ", "&nbsp;") + "&nbsp;&nbsp;" +
                        "<a href=\"" + url + "\">" +
                        association + "</a><br><br>" + "</font>\n");
                }
                else
                {
                    sw.Write("<font color=\"" + element.FontColor + "\" " +
                        "face=\"" + element.FontFamily + "\">" + element.NoteText.Replace("\r\n", "<br>\n").Replace(" ", "&nbsp;") + "<br><br></font>\n");
                }

            }
            catch (Exception)
            {

            }
        }

        public void ExportStructure(string folderPath)
        {
            char[] ds = {'\\'};
            string dst = folderPath + System.IO.Path.DirectorySeparatorChar + System.IO.Path.GetFileName(root.Path.TrimEnd(ds));
            ProcessExportStructure(root.Path, dst, false);
        }

        public void ExportStructureAndContent(string folderPath)
        {
            char[] ds = { '\\' };
            string dst = folderPath + System.IO.Path.DirectorySeparatorChar + System.IO.Path.GetFileName(root.Path.TrimEnd(ds));
            ProcessExportStructure(root.Path, dst, true);
        }

        public void ExportLog(string folderPath)
        {
            string dst = folderPath + System.IO.Path.DirectorySeparatorChar + LogControl.LOG_FILENAME;
            System.IO.File.Copy(LogControl.LogFileFullName, dst);
            System.Diagnostics.Process.Start(dst);
        }

        private void ProcessExportStructure(string src, string dst, bool copyContent)
        {            
            if (dst[dst.Length - 1] != System.IO.Path.DirectorySeparatorChar)
            {
                dst += System.IO.Path.DirectorySeparatorChar;
            }

            if (!Directory.Exists(dst))
            {
                Directory.CreateDirectory(dst);
            }

            String[] FSEntries = Directory.GetFileSystemEntries(src);
            if (!copyContent)
            {
                foreach (string dir in FSEntries)
                {
                    if (Directory.Exists(dir))
                    {
                        ProcessExportStructure(dir, dst + System.IO.Path.GetFileName(dir), copyContent);
                    }
                }
            }
            else
            {
                foreach (string dir in FSEntries)
                {
                    if (Directory.Exists(dir))
                    {
                        ProcessExportStructure(dir, dst + System.IO.Path.GetFileName(dir), copyContent);
                    }
                    else
                    {
                        System.IO.File.Copy(dir, dst + System.IO.Path.GetFileName(dir), true);
                    }
                }
            }
        }

        #endregion

        #region Flag Consolidation

        public void Flag(Element element, 
            bool hasStart, DateTime startDate, CustomTimeSpan startTime, bool isStartAllDay,
            bool hasDue, DateTime dueDate, CustomTimeSpan dueTime, bool isDueAllDay,
            bool addToToday, bool addToReminder, bool addToTask)
        {
            Outlook.Application outlook_app = new Microsoft.Office.Interop.Outlook.Application();
            Outlook.AppointmentItem app_item;
            Outlook.TaskItem task_item;

            if (hasStart || hasDue)
            {
                app_item = outlook_app.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem) as Microsoft.Office.Interop.Outlook.AppointmentItem;

                if (isStartAllDay)
                {
                    app_item.Start = startDate;
                    app_item.AllDayEvent = true;
                    
                    element.StartDate = startDate;
                    element.DueDate = startDate.AddDays(1);
                }
                else
                {
                    if (startTime.IsAM == false)
                    {
                        startTime.Hour += 12;
                    }
                    string startDateTime = startDate.Month.ToString()
                        + "/" + startDate.Day.ToString()
                        + "/" + startDate.Year.ToString()
                        + " " + startTime.Hour.ToString()
                        + ":" + startTime.Minutes.ToString()
                        + ":00";
                    if (dueTime.IsAM == false)
                    {
                        dueTime.Hour += 12;
                    }
                    string dueDateTime = dueDate.Month.ToString()
                        + "/" + dueDate.Day.ToString()
                        + "/" + dueDate.Year.ToString()
                        + " " + dueTime.Hour.ToString()
                        + ":" + dueTime.Minutes.ToString()
                        + ":00";
                    app_item.Start = System.DateTime.Parse(startDateTime);
                    app_item.End = System.DateTime.Parse(dueDateTime);
                    app_item.AllDayEvent = false;

                    element.StartDate = startDate;
                    element.DueDate = dueDate;
                }

                if (addToReminder)
                {
                    app_item.Body = "Start date for " + element.NoteText + @" <file:\\" + element.Path + ">";
                    app_item.Subject = element.NoteText;
                    app_item.ReminderMinutesBeforeStart = 0;

                    app_item.Save();
                    app_item = null;
                }
            }

            if (addToTask)
            {
                task_item = outlook_app.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olTaskItem) as Microsoft.Office.Interop.Outlook.TaskItem;

                task_item.Subject = element.NoteText;
                task_item.Body = "Task for " + element.NoteText + " " + element.Path;
                if (hasStart)
                {
                    task_item.StartDate = startDate;
                }
                if (hasDue)
                {
                    task_item.DueDate = dueDate;
                }
                task_item.ReminderSet = false;

                task_item.Save();
                task_item = null;
            }

            if (addToToday)
            {
                Element today = GetTodayElement();
                if (today != null)
                {
                    Element newElement = CreateNewElement(ElementType.Note, element.NoteText);
                    InsertElement(newElement, today, 0);
                    AddAssociation(newElement, element.Path, ElementAssociationType.FolderShortcut, null);
                    Promote(newElement);
                    Flag(newElement);

                    /*if (hasStart)
                    {
                        if (isStartAllDay)
                        {

                        }
                        else
                        {
                            newElement.NoteText += " Start: " + startDate.Month.ToString()
                                + "/" + startDate.Day.ToString()
                                + "/" + startDate.Year.ToString()
                                + " " + startTime.Hour.ToString()
                                + ":" + startTime.Minutes.ToString();
                        }
                    }
                    if (hasDue)
                    {
                        if (isDueAllDay)
                        {

                        }
                        else
                        {
                            newElement.NoteText += " Due: " + dueDate.Month.ToString()
                                + "/" + dueDate.Day.ToString()
                                + "/" + dueDate.Year.ToString()
                                + " " + dueTime.Hour.ToString()
                                + ":" + dueTime.Minutes.ToString();
                        }
                    }*/
                }
            }

            Flag(element);
        }

        public void Flag(Element element)
        {
            element.FlagStatus = FlagStatus.Flag;
            element.FlagImageSource = String.Format("pack://application:,,,/{0};component/{1}", "Planz", "Images/flag.gif");
            element.ShowFlag = Visibility.Visible;
            UpdateElement(element);      
        }

        public void Check(Element element, bool removeFromToday)
        {
            Element today = GetTodayElement();
            if (today != null)
            {
                Element toberemoved = null;
                foreach (Element ele in today.Elements)
                {
                    if (ele.NoteText == element.NoteText && ele.AssociationType == ElementAssociationType.FolderShortcut)
                    {
                        toberemoved = ele;
                        break;
                    }
                }
                if (toberemoved != null)
                {
                    if (removeFromToday)
                    {
                        DeleteElement(toberemoved);
                    }
                    else
                    {
                        Check(toberemoved);
                    }
                }
            }

            Check(element);
        }

        public void Check(Element element)
        {
            element.FlagStatus = FlagStatus.Check;
            element.FlagImageSource = String.Format("pack://application:,,,/{0};component/{1}", "Planz", "Images/check.gif");
            UpdateElement(element); 
        }

        public void Uncheck(Element element)
        {
            element.FlagStatus = FlagStatus.Normal;
            element.FlagImageSource = String.Format("pack://application:,,,/{0};component/{1}", "Planz", "Images/normal.gif");
            UpdateElement(element);
        }

        #endregion

        #region PowerD Operations

        public bool IsUnderToday(Element element)
        {
            if (element.Path.Contains(StartProcess.START_PATH + StartProcess.TODAY_PLUS))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public void PowerDDone(Element element, DateTime dt)
        {
            if (element == null)
            {
                return;
            }
            
            string dateFolderPath = JournalControl.GetJournalPath(dt) + System.IO.Path.DirectorySeparatorChar;

            Guid lastID = element.ID;
            Guid newID = Guid.NewGuid();
            element.ID = newID;
            element.PowerDStatus = PowerDStatus.Done;
            element.PowerDTimeStamp = dt;
            element.FontColor = ElementColor.SeaGreen.ToString();

            DatabaseControl temp_dbControl = new DatabaseControl(dateFolderPath);
            temp_dbControl.OpenConnection();
            bool isLocalFile = false;
            bool isLocalFolder = false;
            string targetFileName = element.AssociationURI;
            if (element.AssociationType == ElementAssociationType.File)
            {
                isLocalFile = true;
            }
            if (element.AssociationType == ElementAssociationType.Folder)
            {
                isLocalFolder = true;
            }
            if (isLocalFile || isLocalFolder)
            {
                string shortcutName = ShortcutNameConverter.GenerateShortcutNameFromFileName(element.AssociationURI, dateFolderPath);
                string shortcutPath = dateFolderPath + shortcutName;

                if (isLocalFile)
                {
                    element.AssociationType = ElementAssociationType.FileShortcut;
                }
                else
                {
                    element.AssociationType = ElementAssociationType.FolderShortcut;
                }
                element.AssociationURI = shortcutName;

                CreateShortcut(element, element.ParentElement.Path + targetFileName, shortcutPath);
            }
            Element parent = element.ParentElement;
            element.ParentElement = null;
            element.Position = -1;
            temp_dbControl.InsertElementIntoXML(element);
            element.ParentElement = parent;
            temp_dbControl.CloseConnection();
            if (isLocalFile)
            {
                element.AssociationType = ElementAssociationType.File;
                element.AssociationURI = targetFileName;
            }
            if (isLocalFolder)
            {
                element.AssociationType = ElementAssociationType.Folder;
                element.AssociationURI = targetFileName;
            }
            if (element.AssociationType != ElementAssociationType.None && !isLocalFile && !isLocalFolder)
            {
                System.IO.File.Copy(element.AssociationURIFullPath, dateFolderPath + element.AssociationURI);
            }

            element.ID = lastID;
            element.IsVisible = Visibility.Visible;
            element.TempData = dateFolderPath + "|" + newID.ToString();
            element.PowerDStatus = PowerDStatus.Done;
            element.PowerDTimeStamp = dt;
            element.FontColor = ElementColor.SpringGreen.ToString();

            UpdateElement(element);

            if (element.FlagStatus == FlagStatus.Flag)
            {
                Check(element);
            }
        }

        public void PowerDDefer(Element element, DateTime dt)
        {
            if (element == null)
            {
                return;
            }

            element.IsVisible = Visibility.Visible;
            element.PowerDStatus = PowerDStatus.Deferred;
            element.PowerDTimeStamp = dt;
            element.FontColor = ElementColor.Plum.ToString();

            UpdateElement(element);
        }

        public void PowerDDelegate(Element element, string folderPath)
        {
            if (Directory.Exists(folderPath) == false)
            {
                return;
            }

            if (MessageBox.Show("This association and its file or shortcut (if any) will be moved to \r\n" + folderPath, "Delegate", MessageBoxButton.OKCancel) == MessageBoxResult.Cancel)
            {
                return;
            }

            if (CheckOpenFiles(element) == true)
            {
                return;
            }

            element.PowerDStatus = PowerDStatus.Delegated;
            element.PowerDTimeStamp = DateTime.Now;

            Element parent = element.ParentElement;
            Guid id = element.ID;
            element.ID = Guid.NewGuid();
            element.Status = ElementStatus.New;
            element.FontColor = ElementColor.Blue.ToString();
            DatabaseControl temp_dbControl = new DatabaseControl(folderPath);
            temp_dbControl.OpenConnection();
            element.ParentElement = null;
            element.Position = -1;
            temp_dbControl.InsertElementIntoXML(element);
            temp_dbControl.CloseConnection();
            element.ParentElement = parent;
            element.ID = id;

            if (element.HasAssociation)
            {
                if (element.AssociationType == ElementAssociationType.Folder)
                {
                    System.IO.Directory.Move(element.AssociationURIFullPath, folderPath + element.AssociationURI);
                }
                else
                {
                    System.IO.File.Move(element.AssociationURIFullPath, folderPath + element.AssociationURI);
                }
            }

            string folderName = System.IO.Directory.GetParent(folderPath).Name;
           
            element.NoteText = "\"" + element.NoteText + "\" has been moved to: " + folderName;
            element.Status = ElementStatus.Normal;
            element.PowerDStatus = PowerDStatus.Done;
            element.PowerDTimeStamp = DateTime.Now;
            element.FontColor = ElementColor.SpringGreen.ToString();

            element.AssociationType = ElementAssociationType.FolderShortcut;
            string shortcutName = ShortcutNameConverter.GenerateShortcutNameFromFileName(folderName, parent.Path);
            string shortcutPath = parent.Path + shortcutName;
            element.AssociationURI = shortcutName;
            CreateShortcut(element, folderPath, shortcutPath);

            element.TailImageSource = FileTypeHandler.GetIcon(ElementAssociationType.FolderShortcut, element.AssociationURIFullPath);

            UpdateElement(element);
        }

        public void PowerDDelete(Element element, PowerDDeleteType pddt)
        {
            switch (pddt)
            {
                case PowerDDeleteType.Delete:
                    DeleteMessageType dmt = DeleteMessageType.Default;
                    switch (element.Type)
                    {
                        case ElementType.Heading:
                            if (element.IsRemoteHeading)
                            {
                                dmt = DeleteMessageType.InplaceExpansionHeading;
                            }
                            else
                            {
                                if (HasChildOrContent(element))
                                {
                                    dmt = DeleteMessageType.HeadingWithChildren;
                                }
                                else
                                {
                                    dmt = DeleteMessageType.HeadingWithoutChildren;
                                }
                            }
                            break;
                        case ElementType.Note:
                            if (element.HasAssociation)
                            {
                                switch (element.AssociationType)
                                {
                                    case ElementAssociationType.File:
                                        dmt = DeleteMessageType.NoteWithFileAssociation;
                                        break;
                                    case ElementAssociationType.FileShortcut:
                                    case ElementAssociationType.Web:
                                    case ElementAssociationType.Email:
                                        dmt = DeleteMessageType.NoteWithShortcutAssociation;
                                        break;
                                };
                            }
                            else
                            {
                                dmt = DeleteMessageType.NoteWithoutAssociation;
                            }
                            break;
                    };

                    if (dmt == DeleteMessageType.HeadingWithoutChildren ||
                        dmt == DeleteMessageType.NoteWithoutAssociation)
                    {
                        
                    }
                    else
                    {
                        DeleteWindow dw = new DeleteWindow(dmt);
                        if (dw.ShowDialog().Value == true)
                        {
                            
                        }
                        else
                        {
                            return;
                        }
                    }

                    DeleteElement(element);

                    break;
                case PowerDDeleteType.Undo:

                    string folderPath = String.Empty;
                    string oldGuid = String.Empty;
                    Element oldElement = null;
                    if (element.PowerDStatus == PowerDStatus.Done && element.TempData != String.Empty)
                    {
                        folderPath = element.TempData.Split('|')[0];
                        oldGuid = element.TempData.Split('|')[1];

                        DatabaseControl temp_dbControl = new DatabaseControl(folderPath);
                        temp_dbControl.OpenConnection();
                        foreach (Element ele in temp_dbControl.GetAllElementFromXML())
                        {
                            if (ele.ID.ToString() == oldGuid)
                            {
                                oldElement = ele;
                                break;
                            }
                        }

                        temp_dbControl.RemoveElementFromXML(oldElement);
                        temp_dbControl.CloseConnection();

                        if (oldElement != null && oldElement.AssociationURI != String.Empty)
                        {
                            if (oldElement.AssociationType == ElementAssociationType.File)
                            {
                                System.IO.File.Delete(element.AssociationURIFullPath);
                                System.IO.File.Move(folderPath + oldElement.AssociationURI, element.Path + oldElement.AssociationURI);

                                element.AssociationType = ElementAssociationType.File;
                                AssignICCInfo(element, element.Path + oldElement.AssociationURI);
                            }
                            else
                            {
                                System.IO.File.Delete(folderPath + oldElement.AssociationURI);
                            }
                        }
                    }

                    element.FontColor = ElementColor.Black.ToString();
                    element.TempData = String.Empty;
                    element.PowerDStatus = PowerDStatus.None;
                    element.IsVisible = Visibility.Visible;

                    break;
            };

            UpdateElement(element);
        }

        public void LabelWith(Element element, string path)
        {
            ElementControl temp_eleControl = new ElementControl(path);

            Element ele = new Element
            {
                ID = Guid.NewGuid(),
                ParentElement = temp_eleControl.Root,
                HeadImageSource = String.Empty,
                TailImageSource = String.Empty,
                NoteText = element.NoteText,
                Path = path,
                Type = ElementType.Note,
                Status = ElementStatus.New,
                FontColor = ElementColor.Blue.ToString(),
                Position = 0,
            };

            string shortcutName = ShortcutNameConverter.GenerateShortcutNameFromFileName(element.AssociationURI, path);
            string shortcutPath = path + shortcutName;

            ele.AssociationType = ElementAssociationType.FolderShortcut;
            ele.AssociationURI = shortcutName;
            CreateShortcut(ele, element.Path, shortcutPath);

            temp_eleControl.Root.Elements.Insert(0, ele);

            DatabaseControl temp_dbControl = new DatabaseControl(path);
            temp_dbControl.OpenConnection();
            temp_dbControl.InsertElementIntoXML(ele);
            temp_dbControl.CloseConnection();

            temp_eleControl.UpdateElement(ele);

            string folderName, parentName, tempPath;
            int pos;
            char[] ds = { '\\' };

            tempPath = path.TrimEnd(ds);

            pos = tempPath.LastIndexOfAny(ds);
            folderName = tempPath.Substring(pos + 1);
            tempPath = tempPath.Remove(pos);

            pos = tempPath.LastIndexOfAny(ds);
            parentName = tempPath.Substring(pos + 1);

            //string notifyText = "Labeled with \"" + System.IO.Path.GetFileName(targetPath) + "\" under \"" + System.IO.Path.Path.GetFileName(parent) + "\"";
            string notifyText = "Labeled with \"" + folderName + "\" under \"" + parentName + "\"";

            //add an note for notification
            Element notifyEle = new Element
            {
                ID = Guid.NewGuid(),
                ParentElement = element,
                NoteText = notifyText,
                Path = element.Path,
                Type = ElementType.Note,
                Status = ElementStatus.New,
                FontColor = ElementColor.SpringGreen.ToString(),
                Position = 0,
            };

            /*
            shortcutName = ShortcutNameConverter.GenerateShortcutNameFromFileName(folderName, element.Path);
            shortcutPath = element.Path + shortcutName;

            notifyEle.AssociationType = ElementAssociationType.FolderShortcut;
            notifyEle.AssociationURI = shortcutName;
            CreateShortcut(notifyEle, path, shortcutPath);
            notifyEle.TailImageSource = FileTypeHandler.GetIcon(ElementAssociationType.FolderShortcut, notifyEle.AssociationURIFullPath);
            */
            notifyEle.PowerDStatus = PowerDStatus.Done;
            notifyEle.PowerDTimeStamp = DateTime.Now;
            InsertElement(notifyEle, element, 0);

            UpdateElement(element);
        }


        #endregion

        #region Support Functions

        public Element GetHeadingElement(NavigationItem ni)
        {
            Element targetElement = null;

            Stack<string> stack = new Stack<string>();
            NavigationItem ni_iter = ni;
            while (ni_iter != null)
            {
                stack.Push(ni_iter.Name);
                ni_iter = ni_iter.Parent;
            }

            stack.Pop();
            targetElement = root;
            while (stack.Count != 0)
            {
                string name = stack.Pop();
                if (targetElement.IsExpanded == false)
                {
                    isNavigationExpansion = true;
                    targetElement.IsExpanded = true;
                    targetElement.LevelOfSynchronization = 1;
                    UpdateElement(targetElement);
                    SyncHeadingElement(targetElement);
                    return null;
                }
                foreach (Element ele in targetElement.Elements)
                {
                    if (ele.NoteText == name)
                    {
                        targetElement = ele;
                        break;
                    }
                }
            }

            return targetElement;
        }

        public bool IsCircularHeading(Element element)
        {
            if (element.IsNote)
            {
                return false;
            }
            Element parent = element.ParentElement;
            while (parent != null)
            {
                if (parent.Path == element.Path)
                {
                    return true;
                }
                parent = parent.ParentElement;
            }
            return false;
        }

        // Check if target_path is a subpath of root_path
        public bool IsSubfolder(string root_path, string target_path)
        {
            if (root_path.EndsWith(System.IO.Path.DirectorySeparatorChar.ToString()))
            {
                root_path = root_path.Substring(0, root_path.Length - 1);
            }
            if (target_path.EndsWith(System.IO.Path.DirectorySeparatorChar.ToString()))
            {
                target_path = target_path.Substring(0, target_path.Length - 1);
            }
            DirectoryInfo di_root = new DirectoryInfo(root_path);
            DirectoryInfo di = new DirectoryInfo(target_path);
            while (di != null)
            {
                if (di.FullName == di_root.FullName)
                {
                    return true;
                }
                di = Directory.GetParent(di.FullName);
            }
            return false;
        }

        public void LoadCommandNotesUnderDaysAhead()
        {
            try
            {
                Element today = GetTodayElement();
                if (today != null)
                {
                    Element daysahead = null;
                    foreach (Element ele in today.Elements)
                    {
                        if (ele.NoteText == StartProcess.DAYS_AHEAD)
                        {
                            daysahead = ele;
                            break;
                        }
                    }
                    if (daysahead != null)
                    {
                        if (daysahead.FirstChild != null && daysahead.FirstChild.IsCommandNote == false)
                        {
                            InsertCommandNote("Click the icon on the right to show Journal in new window", ElementCommand.ShowJournalInNewWindow, daysahead);
                        }
                    }
                }
            }
            catch (Exception)
            {

            }
        }

        public void InsertCommandNote(String noteText, ElementCommand command, Element target)
        {
            Element commandNote = CreateNewElement(ElementType.Note, noteText);
            commandNote.Command = command;
            commandNote.Status = ElementStatus.Special;
            commandNote.TailImageSource = String.Format("pack://application:,,,/{0};component/{1}", "Planz", "Images/command.png");
            commandNote.HasAssociation = true;
            commandNote.CanOpen = false;
            commandNote.CanExplore = false;
            commandNote.CanRename = false;
            commandNote.CanDelete = false;

            bool hasSameCommand = false;
            foreach (Element ele in target.Elements)
            {
                if (ele.IsCommandNote && ele.Command == command)
                {
                    hasSameCommand = true;
                    break;
                }
            }
            if (!hasSameCommand)
            {
                InsertElement(commandNote, target, 0);
            }
        }

        public void DisplayMoreAssociations(Element element)
        {
            try
            {
                Queue<Element> displayElements = new Queue<Element>();

                DatabaseControl temp_dbControl = new DatabaseControl(element.ParentElement.Path);
                temp_dbControl.OpenConnection();
                temp_dbControl.SHOW_ME_ALL = true;
                List<Element> eleList = temp_dbControl.GetAllElementFromXML();
                temp_dbControl.CloseConnection();

                Element eleAbove = element.ElementAboveUnderSameParent;
                int pos = 0;
                int j = 0;
                bool more = false;
                if (eleAbove != null)
                {
                    pos = eleAbove.Position;
                    for (j = 0; j < eleList.Count; j++)
                    {
                        if (eleList[j].ID == eleAbove.ID)
                        {
                            j++;
                            break;
                        }
                    }
                }
                for (; j < eleList.Count; j++)
                {
                    if (eleList[j].IsVisible == Visibility.Collapsed)
                    {
                        displayElements.Enqueue(eleList[j]);
                        if (displayElements.Count == MAX_NEWASSOCIATION_VISIBLE)
                        {
                            more = true;
                            break;
                        }
                    }
                    else
                    {
                        break;
                    }
                }

                while (displayElements.Count > 0)
                {
                    Element ele = displayElements.Dequeue();
                    ele.ParentElement = element.ParentElement;
                    ele.IsVisible = Visibility.Visible;
                    element.ParentElement.Elements.Insert(++pos, ele);
                    UpdateElement(ele);
                }

                if (!more)
                {
                    element.ParentElement.Elements.Remove(element);
                }
            }
            catch (Exception ex)
            {

            }

        }

        private Element GetTodayElement()
        {
            Element today = null;
            foreach (Element ele in root.Elements)
            {
                if (ele.NoteText == StartProcess.TODAY_PLUS)
                {
                    today = ele;
                    break;
                }
            }
            return today;
        }

        public void UpdateDaysAhead()
        {
            Element today = GetTodayElement();
            if (today != null)
            {
                Element daysahead = null;
                DatabaseControl tmp_dbControl = new DatabaseControl(today.Path);
                List<Element> eleList = new List<Element>();

                if (today.IsExpanded == true)
                {
                    foreach (Element ele in today.Elements)
                    {
                        if (ele.NoteText == StartProcess.DAYS_AHEAD)
                        {
                            daysahead = ele;
                            break;
                        }
                    }
                }
                else
                {
                    tmp_dbControl.OpenConnection();
                    eleList = tmp_dbControl.GetAllElementFromXML();
                    tmp_dbControl.CloseConnection();

                    foreach (Element ele in eleList)
                    {
                        if (ele.NoteText == StartProcess.DAYS_AHEAD)
                        {
                            daysahead = ele;
                            break;
                        }
                    }
                }
                


                if (daysahead != null)
                {
                    daysahead.Elements.Clear();
                    GC.Collect();
                    System.IO.DirectoryInfo di = new DirectoryInfo(daysahead.Path);
                    try
                    {
                        FileInfo[] fi = di.GetFiles();
                        for (int i = 0; i < fi.Length; i++)
                        {
                            Regex regex = new Regex(@"\d{4}-\d{2}-\d{2},\w*");
                            string filename = fi[i].Name;
                            if (filename == "XooML.xml" || regex.Match(filename).Success)
                            {
                                fi[i].Delete();
                            }
                        }

                        DatabaseControl temp_dbControl = new DatabaseControl(daysahead.Path);
                        temp_dbControl.OpenConnection();
                        temp_dbControl.CloseConnection();
                    }
                    catch (Exception)
                    {
                        return;
                    }
                }
                else
                {
                    daysahead = CreateNewElement(ElementType.Heading, String.Empty);
                    daysahead.Status = ElementStatus.Special;
                    InsertElement(daysahead, today, today.Elements.Count);
                    daysahead.NoteText = StartProcess.DAYS_AHEAD;
                    UpdateElement(daysahead);
                }

                const int num_daysahead = 7;
                List<string> journalPath = JournalControl.GetJournalPathDaysAhead(DateTime.Now, num_daysahead);
                for (int i = 0; i < num_daysahead; i++)
                {
                    string dayPath = journalPath[i];
                    string dayName = System.IO.Path.GetFileName(dayPath);
                    int period = dayName.IndexOf(',');
                    string part1 = dayName.Substring(0, period);
                    string part2 = dayName.Substring(period + 2);
                    Element dayElement = CreateNewElement(ElementType.Note, System.IO.Path.GetFileName(dayPath));
                    if (i == 0)
                    {
                        part2 = "Today";
                    }
                    else if (i == 1)
                    {
                        part2 = "Tomorrow";
                    }
                    dayElement.NoteText = part2 + ", " + part1;
                    InsertElement(dayElement, daysahead, daysahead.Elements.Count);
                    AddAssociation(dayElement, dayPath, ElementAssociationType.FolderShortcut, null);
                    Promote(dayElement);
                }
            }
        }

        //Create the Today+ heading when the planz runs the first time
        public void CreateToday()
        {
            Element today = GetTodayElement();
            if (today != null)
            {
                #region Days Ahead

                UpdateDaysAhead();

                #endregion
            }
        }

        //Update the Today+ heading when the planz runs the first tiem at each day
        //Move deferred notes from date folder to Today+
        public void UpdateToday()
        {
            Element todayPlus = GetTodayElement();

            string todayJournalPath = JournalControl.GetJournalPath(DateTime.Now);

            ElementControl eleControl = new ElementControl(todayJournalPath + "\\");
            Element today = eleControl.root;
            
            List<Element> eleList = new List<Element>();
            foreach (Element ele in today.Elements)
            {
                if(ele.IsNote && ele.PowerDStatus==PowerDStatus.Deferred)
                    eleList.Add(ele);
            }

            int index = 0;
            foreach (Element element in eleList)
            {
                element.PowerDStatus = PowerDStatus.None;
                eleControl.MoveElement(FindElementByGuid(today,element.ID), todayPlus, index);
                index++;
            }

        }

        public List<Element> FindHeadingElementsByLevel(int level)
        {
            List<Element> targetList = new List<Element>();

            switch (level)
            {
                case 1:
                    targetList = root.HeadingElements;
                    break;
                case 2:
                    foreach (Element element in root.HeadingElements)
                    {
                        targetList.AddRange(element.HeadingElements);
                    }
                    break;
                case 3:
                default:
                    foreach (Element element in FindAllHeadingElements(root))
                    {
                        if (element.Level >= 3)
                        {
                            targetList.Add(element);
                        }
                    }
                    break;
            }

            return targetList;
        }

        public List<Element> FindAllHeadingElements(Element element)
        {
            List<Element> targetList = new List<Element>();

            if (element.Type == ElementType.Heading)
            {
                targetList.Add(element);
                foreach (Element ele in element.HeadingElements)
                {
                    targetList.AddRange(FindAllHeadingElements(ele));
                }
            }

            return targetList;
        }

        public List<Element> FindAllNoteElements(Element element)
        {
            List<Element> targetList = new List<Element>();

            if (element.Type == ElementType.Note)
            {
                targetList.Add(element);
            }

            return targetList;
        }

        public Element FindElementByGuid(Element element, Guid guid)
        {
            Element target = null;

            foreach (Element ele in element.Elements)
            {
                if (ele.ID == guid)
                {
                    return ele;
                }

                if (ele.HasChildren)
                {
                    target = FindElementByGuid(ele, guid);
                    if (target != null)
                    {
                        return target;
                    }
                }
            }

            return target;
        } 

        public int FindElementByKeyword(string keyword, bool inViewOnly)
        {
            List<Element> target = new List<Element>();

            FindElementByKeyword(keyword, target, root, inViewOnly);

            sf.Clear();
            sf.TargetList = target;

            return target.Count;
        }

        public Element FindNext()
        {
            return sf.GetNextElement();
        }

        public Element FindNextElementAfterDeletion(Element tobeDeleted)
        {
            Element nextElement = tobeDeleted.ElementBelowUnderSameParent;
            if (nextElement == null)
            {
                nextElement = tobeDeleted.ElementAboveUnderSameParent;
            }
            if (nextElement == null)
            {
                nextElement = tobeDeleted.ElementBelow;
            }
            if (nextElement == null)
            {
                nextElement = tobeDeleted.ElementAbove; 
            }
            return nextElement;
        }

        private void FindElementByKeyword(string keyword, List<Element> target, Element element, bool inViewOnly)
        {
            if (element.NoteText.ToLower().Contains(keyword))
            {
                target.Add(element);
            }

            if (element.HasChildren && 
                ((inViewOnly && element.IsExpanded) || !inViewOnly))
            {
                foreach (Element ele in element.Elements)
                {
                    FindElementByKeyword(keyword, target, ele, inViewOnly);
                }
            }
        }

        protected virtual void OnOperationError(object sender, OperationErrorEventArgs e)
        {
            if (generalOperationErrorDelegate != null)
                generalOperationErrorDelegate(sender, e);
        }

        protected virtual void OnFindOpenFiles(object sender, EventArgs e)
        {
            if (hasOpenFileDelegate != null)
                hasOpenFileDelegate(sender, e);
        }

        public bool CheckOpenFiles(Element element)
        {
            ArrayList openFileList;
            if (HasOpenFiles(element, out openFileList))
            {
                if (hasOpenFileDelegate != null)
                {
                    OnFindOpenFiles(openFileList, new EventArgs());
                }
                return true;
            }
            return false;
        }

        private bool HasOpenFiles(Element element, out ArrayList openFileList)
        {
            openFileList = GetOpenFileList(element);

            if (openFileList.Count > 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private ArrayList GetOpenFileList(Element element)
        {
            ArrayList openFileList = new ArrayList();

            switch (element.Type)
            {
                case ElementType.Heading:
                    if (element.IsLocalHeading)
                    {
                        ElementControl temp_eleControl = new ElementControl(element.Path);
                        foreach (Element ele in temp_eleControl.Root.Elements)
                        {
                            openFileList.AddRange(GetOpenFileList(ele));
                        }
                    }
                    break;
                case ElementType.Note:
                    if (element.AssociationType == ElementAssociationType.File)
                    {
                        if (System.IO.File.Exists(element.AssociationURIFullPath))
                        {
                            try
                            {
                                FileStream fs = System.IO.File.Open(element.AssociationURIFullPath, System.IO.FileMode.OpenOrCreate, System.IO.FileAccess.Read, System.IO.FileShare.None);
                                fs.Close();
                            }
                            catch (System.IO.IOException ex)
                            {
                                openFileList.Add(element.AssociationURIFullPath);
                            }
                        }
                    }
                    break;
            };

            return openFileList;
        }



        public bool HasChildOrContent(Element element)
        {
            if (element.IsNote)
            {
                return false;
            }
            else if (element.HasChildren)
            {
                return true;
            }
            else
            {
                System.IO.DirectoryInfo di = new DirectoryInfo(element.Path);

                DatabaseControl temp_dbControl = new DatabaseControl(element.Path);
                temp_dbControl.OpenConnection();
                bool noContent = false;
                if (temp_dbControl.GetAllElementFromXML().Count == 0)
                {
                    if (di.GetFiles().Length == 1 && di.GetFiles()[0].Name == StartProcess.XOOML_XML_FILENAME)
                    {
                        noContent = true;
                    }
                }
                temp_dbControl.CloseConnection();

                return !noContent;
            }
        }

        public bool IsICCEmail(Element element)
        {
            try
            {
                WshShellClass wshShell = new WshShellClass();
                IWshRuntimeLibrary.IWshShortcut shortcut = (IWshRuntimeLibrary.IWshShortcut)wshShell.CreateShortcut(element.AssociationURIFullPath);
                if (shortcut.Description == "PlanzICC")
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception)
            {
                return false;
            }
        }

        public void FindSentEmail(Element element)
        {
            try
            {
                WshShellClass wshShell = new WshShellClass();
                IWshRuntimeLibrary.IWshShortcut shortcut = (IWshRuntimeLibrary.IWshShortcut)wshShell.CreateShortcut(element.AssociationURIFullPath);
                string entryID = shortcut.Arguments.Substring(16);// /select outlook:<entryID>

                Outlook.Application olApp = new Outlook.Application();
                Outlook.MailItem mail = olApp.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem) as Outlook.MailItem;
                Outlook.NameSpace ns = olApp.GetNamespace("mapi");
                Outlook.MAPIFolder sentFolder = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail);
                Outlook.MAPIFolder draftFolder = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDrafts);

                bool foundInDraft = false;
                foreach (Outlook.MailItem mailitem in draftFolder.Items)
                {
                    if (entryID == mailitem.EntryID)
                    {
                        foundInDraft = true;
                        if (element.NoteText == "New Email" && mailitem.Subject != "New Email")
                        {
                            element.NoteText = mailitem.Subject;
                            UpdateElement(element);
                        }
                        break;
                    }
                }
                if (foundInDraft)
                {
                    return;
                }
                else
                {
                    Outlook.UserProperties props;
                    Outlook.UserProperty prop;
                    Outlook.MailItem sentMail;

                    int sentSize = sentFolder.Items.Count;
                    for (int i = sentSize; i > 0; i--)
                    {
                        sentMail = sentFolder.Items[i] as Outlook.MailItem;
                        props = sentMail.UserProperties;
                        prop = props.Find("PlanzID", true);
                        if (prop != null && prop.Value.ToString() == element.ID.ToString())
                        {
                            shortcut.Arguments = "/select outlook:" + sentMail.EntryID;
                            shortcut.Description = "";
                            shortcut.Save();

                            if (element.NoteText == "New Email" && sentMail.Subject != "New Email")
                            {
                                element.NoteText = sentMail.Subject;
                                UpdateElement(element);
                            }

                            break;
                        }
                    }
                }
            }
            catch (Exception)
            {
                
            }
        }

        #endregion
    }
}
