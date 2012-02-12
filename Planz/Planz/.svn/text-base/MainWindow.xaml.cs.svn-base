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
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Diagnostics;
using System.ComponentModel;
using System.Collections.Specialized;
using System.Threading;
using Microsoft.Windows.Controls;

namespace Planz
{
    public partial class MainWindow : Window
    {
        private TextBox previousNavigatedTextBox = null;
        private String previousText = String.Empty;

        private string previousPowerDDelegateFolder = null;
        private string previousNewWindowFolder = null;
        private string previousChangeFocusFolder = null;
        private Queue powerDDelegateFolderCache = new Queue(5);
        private Queue newWindowFolderCache = new Queue(4);
        private Queue changeFocusFolderCache = new Queue(4);
        private const int MAX_SELECTEDFOLDER_COUNT = 3;

        private ElementControl elementControl;

        private bool isHeadImageDrag = false;
        private Point headImageDragStart;
        private Line headImageDragLine = new Line();

        private const int TEXTBOX_RIGHT_MARGIN = 70;
        private const int ACTIVEITEM_TITLE_LENGTH = 50;
        private bool SELECTALL_FOCUS_HANDLE = false;

        private bool SHOW_DONEDEFER_HANDLE = false;

        private Thread UI_THREAD = Thread.CurrentThread;

        private string startPath = StartProcess.START_PATH;

        public MainWindow()
        {
            InitializeComponent();

            // Show UI Loading
            StartLoadingUI();

            this.PlanzMainWindow.Title = "Planz - " + this.startPath;

            // Setup
            StartProcess.CheckFirstTimeRunning();

            elementControl = new ElementControl(startPath);

            elementControl.fsSyncReporter += new FSSyncReporter(ElementControl_fsSyncReporter);
            elementControl.hasOpenFileDelegate += new HasOpenFileDelegate(ElementControl_hasOpenFileDelegate);
            elementControl.generalOperationErrorDelegate += new GeneralOperationErrorDelegate(ElementControl_generalOperationError);

            // Load Plan
            this.Plan.ItemsSource = elementControl.Root.Elements;

            // Load Navigation Tab
            LoadNavigationTab();

            // Load QC
            LoadQuickCapture();
        }

        public MainWindow(string startPath)
        {
            InitializeComponent();

            // Window Text
            this.PlanzMainWindow.Title = "Planz - " + startPath;

            // Show UI Loading
            StartLoadingUI();

            // Setup
            this.startPath = startPath;

            elementControl = new ElementControl(startPath);
            elementControl.fsSyncReporter += new FSSyncReporter(ElementControl_fsSyncReporter);
            elementControl.hasOpenFileDelegate += new HasOpenFileDelegate(ElementControl_hasOpenFileDelegate);
            elementControl.generalOperationErrorDelegate += new GeneralOperationErrorDelegate(ElementControl_generalOperationError);

            // Load Plan
            this.Plan.ItemsSource = elementControl.Root.Elements;

            // Load Navigation Tab
            LoadNavigationTab();
        }

        private void LoadQuickCapture()
        {
            try
            {
                Process[] procs = Process.GetProcesses();
                Process qc_proc = null;
                foreach (Process proc in procs)
                {
                    if (proc.ProcessName == ProcessList.QuickCapture.ToString())
                    {
                        qc_proc = proc;
                        break;
                    }
                }
                if (qc_proc == null)
                {
                    string qc_path = System.Windows.Forms.Application.StartupPath + System.IO.Path.DirectorySeparatorChar + ProcessList.QuickCapture.ToString() + ".exe";
                    ProcessStartInfo psi = new ProcessStartInfo(qc_path);
                    psi.WindowStyle = ProcessWindowStyle.Minimized;
                    System.Diagnostics.Process.Start(psi);

                    LogControl.Write(
                        null,
                        LogEventAccess.QuickCapture,
                        LogEventType.Launch,
                        LogEventStatus.NULL,
                        null);
                }
                
            }
            catch (Exception ex)
            {

            }
        }

        private void UnloadQuickCapture()
        {
            try
            {
                Process[] procs = Process.GetProcesses();
                Process qc_proc = null;

                string processName = Process.GetCurrentProcess().ProcessName;

                /*
                if (System.Diagnostics.Process.GetProcessesByName(processName).Length > 1)
                {
                    return;
                }*/
                

                foreach (Process proc in procs)
                {
                    if (proc.ProcessName == ProcessList.QuickCapture.ToString())
                    {
                        qc_proc = proc;
                        break;
                    }
                }

                if (qc_proc != null)
                {
                    qc_proc.Kill();
                    LogControl.Write(
                        null,
                        LogEventAccess.QuickCapture,
                        LogEventType.Close,
                        LogEventStatus.NULL,
                        null);
                }
            }
            catch (Exception ex)
            {

            }
        }

        #region Window

        private void ElementControl_fsSyncReporter(object sender, FileSystemSyncEventArgs e)
        {
            EndLoadingUI();

            // Navigation Tree Expansion
            if (e.Message == FileSystemSyncMessage.NavigationTreeExpanded)
            {
                NavigationTextBlock_MouseDown(this.NavigationTabList.SelectedItem, null);
                return;
            }

            // Update Today+
            if (this.startPath == StartProcess.START_PATH && StartProcess.FirstTime)
            {
                StartProcess.FirstTime = false;
                elementControl.CreateToday();

                Properties.Settings.Default.LastRunningDateTime = DateTime.Now;
                Properties.Settings.Default.Save();
            }
            else
            {
                if (Properties.Settings.Default.LastRunningDateTime == null ||
                    Properties.Settings.Default.LastRunningDateTime.Year != DateTime.Now.Year ||
                    Properties.Settings.Default.LastRunningDateTime.Month != DateTime.Now.Month ||
                    Properties.Settings.Default.LastRunningDateTime.Day != DateTime.Now.Day)
                {
                    elementControl.UpdateDaysAhead();

                    //Move deferred notes
                    //elementControl.UpdateToday();

                    Properties.Settings.Default.LastRunningDateTime = DateTime.Now;
                    Properties.Settings.Default.Save();
                }

            }

            if (Properties.Settings.Default.LastFocusedElementGuid != String.Empty &&
                e.Message == FileSystemSyncMessage.HighPriorityWorkFinished &&
                elementControl.CurrentElement.ID.ToString() != Properties.Settings.Default.LastFocusedElementGuid)
            {
                Element lastFocused = elementControl.FindElementByGuid(elementControl.Root, new Guid(Properties.Settings.Default.LastFocusedElementGuid));
                if (lastFocused != null)
                {
                    elementControl.CurrentElement = lastFocused;
                    GetFocusToElementTextBox(elementControl.CurrentElement, elementControl.CurrentElement.NoteText.Length, false, false);
                }
                else
                {
                    if (elementControl.Root.HasChildren == false)
                    {
                        Element newElement = elementControl.CreateNewElement(ElementType.Note, String.Empty);
                        elementControl.InsertElement(newElement, elementControl.Root, 0);
                        elementControl.CurrentElement = newElement;
                        GetFocusToElementTextBox(elementControl.CurrentElement, elementControl.CurrentElement.NoteText.Length, false, false);
                    }
                }
            }

            if (Properties.Settings.Default.ShowOutline)
            {
                ShowOutline_Checked(null, null);
            }

            // Load Command Note
            elementControl.LoadCommandNotesUnderDaysAhead();

            SHOW_DONEDEFER_HANDLE = true;

            //elementControl.UpdateStatusChangedElementList();
        }

        private void ElementControl_hasOpenFileDelegate(object sender, EventArgs e)
        {
            string message = "The action can't be completed because the following files are open:\r\n";
            ArrayList arrayList = sender as ArrayList;

            if (arrayList != null)
            {
                foreach (string s in arrayList)
                {
                    message += "-" + System.IO.Path.GetFileName(s) + "\r\n";
                }
            }

            MessageBox.Show(message, 
                "File In Use");
        }

        private void ElementControl_generalOperationError(object sender, OperationErrorEventArgs e)
        {
            MessageBox.Show(e.Message,
                "Operation Error");
        }

        private void Plan_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                if (Properties.Settings.Default.LastFocusedElementGuid != String.Empty)
                {
                    Element lastFocused = elementControl.FindElementByGuid(elementControl.Root, new Guid(Properties.Settings.Default.LastFocusedElementGuid));
                    if (lastFocused != null)
                    {
                        elementControl.CurrentElement = lastFocused;
                        GetFocusToElementTextBox(elementControl.CurrentElement, elementControl.CurrentElement.NoteText.Length, false, true);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("Plan_Loaded\n" + ex.Message);
            }
        }

        private void Window_Activated(object sender, EventArgs e)
        {
            //GetFocusToElementTextBox(elementControl.CurrentElement, elementControl.CurrentElement.NoteText.Length);
        }

        private void Window_Loaded(object sender, EventArgs e)
        {

        }

        private void Window_Closed(object sender, EventArgs e)
        {
            //UnloadQuickCapture();
                
            if (elementControl.CurrentElement != null)
            {
                elementControl.UpdateElement(elementControl.CurrentElement);
                elementControl.UpdateStatusChangedElementList();

                Properties.Settings.Default.LastFocusedElementGuid = elementControl.CurrentElement.ID.ToString();
                Properties.Settings.Default.Save();
            }

            if (sender != null)
            {
                LogControl.Write(
                    elementControl.CurrentElement,
                    LogEventAccess.AppWindow,
                    LogEventType.Close,
                    LogEventStatus.NULL,
                    null);
            }
        }

        private void Window_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            try
            {
                Properties.Settings.Default.WindowWidth = e.NewSize.Width;
                Properties.Settings.Default.WindowHeight = e.NewSize.Height;
                Properties.Settings.Default.Save();

                ResizeElementTextBox(elementControl.Root, new Dictionary<int,double>());
            }
            catch (Exception ex)
            {
                //MessageBox.Show("Window_SizeChanged\n" + ex.Message);
            }
        }

        #endregion

        #region Ribbon Menu

        #region Application Menu

        private void RibbonCommandGeneral_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            
        }

        private void RibbonCommandSaveAsTXT_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            try
            {

                Microsoft.Win32.SaveFileDialog dialog = new Microsoft.Win32.SaveFileDialog();
                dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                dialog.FileName = System.IO.Path.GetDirectoryName(startPath);
                dialog.Filter = "Text Document (*.txt)|*.txt";
                dialog.RestoreDirectory = true;
                if (dialog.ShowDialog().Value == true)
                {
                    elementControl.SaveAsTXT(dialog.FileName);

                    LogControl.Write(
                        elementControl.Root,
                        LogEventAccess.AppMenu,
                        LogEventType.SaveAsText,
                        LogEventStatus.NULL,
                        LogEventInfo.SavePath + LogControl.COMMA + dialog.FileName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("RibbonCommandSaveAsTXT_Executed\n" + ex.Message);

                LogControl.Write(
                    elementControl.Root,
                    LogEventAccess.AppMenu,
                    LogEventType.SaveAsText,
                    LogEventStatus.Error,
                    LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void RibbonCommandSaveAsHTML_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            try
            {
                Microsoft.Win32.SaveFileDialog dialog = new Microsoft.Win32.SaveFileDialog();
                dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                dialog.FileName = System.IO.Path.GetDirectoryName(startPath);
                dialog.Filter = "Web Page (*.html)|*.html";
                dialog.RestoreDirectory = true;
                if (dialog.ShowDialog().Value == true)
                {
                    elementControl.SaveAsHTML(dialog.FileName);

                    LogControl.Write(
                        elementControl.Root,
                        LogEventAccess.AppMenu,
                        LogEventType.SaveAsHTML,
                        LogEventStatus.NULL,
                        LogEventInfo.SavePath + LogControl.COMMA + dialog.FileName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("RibbonCommandSaveAsHTML_Executed\n" + ex.Message);

                LogControl.Write(
                    elementControl.Root,
                    LogEventAccess.AppMenu,
                    LogEventType.SaveAsHTML,
                    LogEventStatus.Error,
                    LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }
        
        private void RibbonCommandExportStructure_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            try
            {
                System.Windows.Forms.FolderBrowserDialog dialog = new System.Windows.Forms.FolderBrowserDialog();
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    elementControl.ExportStructure(dialog.SelectedPath);

                    LogControl.Write(
                        elementControl.Root,
                        LogEventAccess.AppMenu,
                        LogEventType.ExportStructureOnly,
                        LogEventStatus.NULL,
                        LogEventInfo.ExportPath + LogControl.COMMA + dialog.SelectedPath);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("RibbonCommandExportStructure_Executed\n" + ex.Message);

                LogControl.Write(
                    elementControl.Root,
                    LogEventAccess.AppMenu,
                    LogEventType.ExportStructureOnly,
                    LogEventStatus.Error,
                    LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void RibbonCommandExportStructureAndContent_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            try
            {
                System.Windows.Forms.FolderBrowserDialog dialog = new System.Windows.Forms.FolderBrowserDialog();
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    elementControl.ExportStructureAndContent(dialog.SelectedPath);

                    LogControl.Write(
                       elementControl.Root,
                       LogEventAccess.AppMenu,
                       LogEventType.ExportStructureAndContent,
                       LogEventStatus.NULL,
                       LogEventInfo.ExportPath + LogControl.COMMA + dialog.SelectedPath);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("RibbonCommandExportStructureAndContent\n" + ex.Message);

                LogControl.Write(
                   elementControl.Root,
                   LogEventAccess.AppMenu,
                   LogEventType.ExportStructureAndContent,
                   LogEventStatus.Error,
                   LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }
        
        private void RibbonCommandExportLog_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            try
            {
                System.Windows.Forms.FolderBrowserDialog dialog = new System.Windows.Forms.FolderBrowserDialog();
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    elementControl.ExportLog(dialog.SelectedPath);

                    LogControl.Write(
                       elementControl.Root,
                       LogEventAccess.AppMenu,
                       LogEventType.ExportLog,
                       LogEventStatus.NULL,
                       LogEventInfo.ExportPath + LogControl.COMMA + dialog.SelectedPath);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("RibbonCommandExportLog_Executed\n" + ex.Message);

                LogControl.Write(
                   elementControl.Root,
                   LogEventAccess.AppMenu,
                   LogEventType.ExportLog,
                   LogEventStatus.Error,
                   LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void RibbonCommandCreateJournal_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            try
            {
                CreateJournalWindow cjw = new CreateJournalWindow();
                if (cjw.ShowDialog().Value)
                {
                    StartLoadingUI();

                    int year = cjw.Year;
                    JournalControl.CreateJournalFolders(year, StartProcess.JOURNAL_PATH);

                    EndLoadingUI();

                    LogControl.Write(
                       elementControl.Root,
                       LogEventAccess.AppMenu,
                       LogEventType.CreateJournalFolder,
                       LogEventStatus.NULL,
                       LogEventInfo.YearCreated + LogControl.COMMA + year.ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("RibbonCommandCreateJournal\n" + ex.Message);

                LogControl.Write(
                       elementControl.Root,
                       LogEventAccess.AppMenu,
                       LogEventType.CreateJournalFolder,
                       LogEventStatus.Error,
                       LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void RibbonCommandOptions_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            try
            {

                OptionsWindow ow = new OptionsWindow();
                if (ow.ShowDialog().Value)
                {
                    if (ow.isHeadingFontChanged)
                    {
                        Properties.Settings.Default.HeadingLevel1FontFamily = ow.headingFont;
                        Properties.Settings.Default.Save();

                        ChangeElementTextFontFamily(ElementType.Heading, new FontFamily(ow.headingFont));
                    }
                    if (ow.isHeadingSizeChanged)
                    {
                        Properties.Settings.Default.HeadingLevel1FontSize = ow.headingSize;
                        Properties.Settings.Default.Save();

                        ChangeElementTextFontSize(ElementType.Heading, ow.headingSize);
                    }
                    if (ow.isNoteFontChanged)
                    {
                        Properties.Settings.Default.NoteFontFamily = ow.noteFont;
                        Properties.Settings.Default.Save();

                        ChangeElementTextFontFamily(ElementType.Note, new FontFamily(ow.noteFont));
                    }
                    if (ow.isNoteSizeChanged)
                    {
                        Properties.Settings.Default.NoteFontSize = ow.noteSize;
                        Properties.Settings.Default.Save();

                        ChangeElementTextFontSize(ElementType.Note, ow.noteSize);
                    }

                    LogControl.Write(
                       elementControl.Root,
                       LogEventAccess.AppMenu,
                       LogEventType.Options,
                       LogEventStatus.NULL,
                       LogEventInfo.HeadingFont + LogControl.COMMA + ow.headingFont + LogControl.DELIMITER +
                       LogEventInfo.HeadingSize + LogControl.COMMA + ow.headingSize.ToString() + LogControl.DELIMITER +
                       LogEventInfo.NoteFont + LogControl.COMMA + ow.noteFont + LogControl.DELIMITER +
                       LogEventInfo.NoteSize + LogControl.COMMA + ow.noteSize.ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("RibbonCommandOptions_Executed\n" + ex.Message);

                LogControl.Write(
                       elementControl.Root,
                       LogEventAccess.AppMenu,
                       LogEventType.Options,
                       LogEventStatus.Error,
                       LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void RibbonCommandExit_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            try
            {
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("RibbonCommandExit_Executed\n" + ex.Message);
            }
        }

        #endregion

        private void RibbonCommandNewWindow_Executed(object sender, ExecutedRoutedEventArgs e)
        {

        }

        private void RibbonCommandChangeFocus_Executed(object sender, ExecutedRoutedEventArgs e)
        {

        }

        private void RibbonCommandExplore_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(elementControl.CurrentElement.Path);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("RibbonCommandExplore_Executed\n" + ex.Message);
            }
        }

        private void RibbonCommandRefresh_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            LogEventAccess lea = LogEventAccess.Ribbon;
            if (sender is TextBox)
            {
                lea = LogEventAccess.Hotkey;
            }
            else
            {
                if (e.Source is Microsoft.Windows.Controls.Ribbon.RibbonButton)
                {
                    lea = LogEventAccess.QuickAccessToolbar;
                }
                else if (e.Source is Microsoft.Windows.Controls.Ribbon.RibbonGroup)
                {
                    lea = LogEventAccess.Ribbon;
                }
            }

            try
            {

                StartLoadingUI();

                if (elementControl.CurrentElement != null)
                {
                    elementControl.UpdateElement(elementControl.CurrentElement);
                    elementControl.UpdateStatusChangedElementList();

                    Properties.Settings.Default.LastFocusedElementGuid = elementControl.CurrentElement.ID.ToString();
                    Properties.Settings.Default.Save();
                }
                
                this.Plan.ItemsSource = null;
                this.NavigationTabList.ItemsSource = null;

                GC.Collect();

                elementControl = new ElementControl(startPath);
                elementControl.fsSyncReporter += new FSSyncReporter(ElementControl_fsSyncReporter);
                elementControl.hasOpenFileDelegate += new HasOpenFileDelegate(ElementControl_hasOpenFileDelegate);
                elementControl.generalOperationErrorDelegate += new GeneralOperationErrorDelegate(ElementControl_generalOperationError);

                this.Plan.ItemsSource = elementControl.Root.Elements;

                LoadNavigationTab();

                LogControl.Write(
                   elementControl.Root,
                   lea,
                   LogEventType.Refresh,
                   LogEventStatus.NULL,
                   null);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("RibbonCommandRefresh_Executed\n" + ex.Message);

                LogControl.Write(
                   elementControl.Root,
                   lea,
                   LogEventType.Refresh,
                   LogEventStatus.Error,
                   LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void RibbonCommandDelete_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            LogEventAccess lea = LogEventAccess.Ribbon;
            if (sender is TextBox)
            {
                lea = LogEventAccess.Hotkey;
            }
            else if (sender is System.Windows.Controls.MenuItem)
            {
                lea = LogEventAccess.NodeContextMenu;
            }
            else if (sender is Planz.MainWindow)
            {
                lea = LogEventAccess.Ribbon;
            }

            Element targetElement = elementControl.CurrentElement;
            try
            {
                if (targetElement != null)
                {
                    DeleteMessageType dmt = DeleteMessageType.Default;

                    if (elementControl.SelectedElements.Count <= 1)
                    {
                        switch (targetElement.Type)
                        {
                            case ElementType.Heading:
                                if (targetElement.IsRemoteHeading)
                                {
                                    dmt = DeleteMessageType.InplaceExpansionHeading;
                                }
                                else
                                {
                                    if (elementControl.HasChildOrContent(targetElement))
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
                                if (targetElement.HasAssociation)
                                {
                                    switch (targetElement.AssociationType)
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
                            LogControl.Write(
                               targetElement,
                               lea,
                               LogEventType.Delete,
                               LogEventStatus.Begin,
                               null);

                            Element nextElement = elementControl.FindNextElementAfterDeletion(targetElement);
                            elementControl.DeleteElement(targetElement);
                            if (nextElement != null)
                            {
                                GetFocusToElementTextBox(nextElement, nextElement.NoteText.Length, false, false);
                            }

                            LogControl.Write(
                               nextElement,
                               lea,
                               LogEventType.Delete,
                               LogEventStatus.End,
                               LogEventInfo.DeleteType + LogControl.COMMA + dmt.ToString());
                        }
                        else
                        {
                            LogControl.Write(
                               targetElement,
                               lea,
                               LogEventType.Delete,
                               LogEventStatus.Begin,
                               null);

                            DeleteWindow dw = new DeleteWindow(dmt);
                            if (dw.ShowDialog().Value == true)
                            {
                                Element nextElement = elementControl.FindNextElementAfterDeletion(targetElement);
                                elementControl.DeleteElement(targetElement);
                                if (nextElement != null)
                                {
                                    GetFocusToElementTextBox(nextElement, nextElement.NoteText.Length, false, false);
                                }

                                LogControl.Write(
                                   nextElement,
                                   lea,
                                   LogEventType.Delete,
                                   LogEventStatus.End,
                                   LogEventInfo.DeleteType + LogControl.COMMA + dmt.ToString());
                            }
                            else
                            {
                                LogControl.Write(
                                   targetElement,
                                   lea,
                                   LogEventType.Delete,
                                   LogEventStatus.Cancel,
                                   null);
                            }
                        }
                    }
                    else
                    {
                        LogControl.Write(
                               targetElement,
                               lea,
                               LogEventType.Delete,
                               LogEventStatus.Begin,
                               null);

                        dmt = DeleteMessageType.MultipleItems;
                        DeleteWindow dw = new DeleteWindow(dmt);
                        if (dw.ShowDialog().Value == true)
                        {
                            Element nextElement = elementControl.FindNextElementAfterDeletion(targetElement);
                            for (int i = 0; i < elementControl.SelectedElements.Count; i++)
                            {
                                elementControl.DeleteElement(elementControl.SelectedElements[i]);
                            }
                            elementControl.SelectedElements.Clear();
                            if (nextElement != null)
                            {
                                GetFocusToElementTextBox(nextElement, nextElement.NoteText.Length, false, false);
                            }

                            LogControl.Write(
                               nextElement,
                               lea,
                               LogEventType.Delete,
                               LogEventStatus.End,
                               LogEventInfo.DeleteType + LogControl.COMMA + dmt.ToString());
                        }
                        else
                        {
                            LogControl.Write(
                               targetElement,
                               lea,
                               LogEventType.Delete,
                               LogEventStatus.Cancel,
                               null);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("RibbonCommandDelete_Executed\n" + ex.Message);

                LogControl.Write(
                   targetElement,
                   lea,
                   LogEventType.Delete,
                   LogEventStatus.Error,
                   LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void RibbonCommandHide_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            LogEventAccess lea = LogEventAccess.Ribbon;
            if (sender is Planz.MainWindow)
            {
                lea = LogEventAccess.Ribbon;
            }
            else if (sender is System.Windows.Controls.MenuItem)
            {
                lea = LogEventAccess.NodeContextMenu;
            }

            try
            {
                Element element = elementControl.CurrentElement;
                Element nextElement = element.ElementBelow;
                elementControl.HideElement(elementControl.CurrentElement);
                GetFocusToElementTextBox(nextElement, nextElement.NoteText.Length, false, true);
                
                LogControl.Write(
                    elementControl.CurrentElement,
                    lea,
                    LogEventType.Hide,
                    LogEventStatus.NULL,
                    null);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("RibbonCommandHide\n" + ex.Message);

                LogControl.Write(
                   elementControl.CurrentElement,
                   lea,
                   LogEventType.Hide,
                   LogEventStatus.Error,
                   LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void RibbonCommandFontColorClicked(object sender, RoutedEventArgs e)
        {
            try
            {
                LogControl.Write(
                   elementControl.CurrentElement,
                   LogEventAccess.Ribbon,
                   LogEventType.ChangeFontColor,
                   LogEventStatus.Begin,
                   null);

                MenuItem mi = sender as MenuItem;
                string color = mi.Header.ToString();

                elementControl.CurrentElement.FontColor = color;
                elementControl.UpdateElement(elementControl.CurrentElement);

                LogControl.Write(
                   elementControl.CurrentElement,
                   LogEventAccess.Ribbon,
                   LogEventType.ChangeFontColor,
                   LogEventStatus.End,
                   LogEventInfo.FontColor + LogControl.COMMA + color);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("RibbonCommandFontColorChecked\n" + ex.Message);

                LogControl.Write(
                   elementControl.CurrentElement,
                   LogEventAccess.Ribbon,
                   LogEventType.ChangeFontColor,
                   LogEventStatus.Error,
                   LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void RibbonSearchBox_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Enter)
                {
                    string query = ((Microsoft.Windows.Controls.Ribbon.RibbonTextBox)sender).Text;
                    query = query.ToLower();

                    int count = elementControl.FindElementByKeyword(query, true);

                    Element target;
                    if ((target = elementControl.FindNext()) != null)
                    {
                        GetFocusToElementTextBox(target, -1, false, true);
                    }

                    this.RibbonSearchStatus.Content = "Found: " + count.ToString();
                    if (count > 1)
                    {
                        this.RibbonSearchStatus.Content += "\r\nPress F3 for next occurrence";
                    }

                    LogControl.Write(
                       elementControl.Root,
                       LogEventAccess.Ribbon,
                       LogEventType.Search,
                       LogEventStatus.NULL,
                       LogEventInfo.SearchText + LogControl.COMMA + query + LogControl.DELIMITER +
                       LogEventInfo.SearchResult + LogControl.COMMA + count.ToString());
                }
                else
                {
                    this.RibbonSearchStatus.Content = String.Empty;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("RibbonSearchBox_KeyDown\n" + ex.Message);

                LogControl.Write(
                   elementControl.Root,
                   LogEventAccess.Ribbon,
                   LogEventType.Search,
                   LogEventStatus.Error,
                   LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void RibbonCommandNewHeading_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            LogEventAccess lea = LogEventAccess.Ribbon;
            if (sender is Planz.MainWindow)
            {
                lea = LogEventAccess.Ribbon;
            }
            else if (sender is System.Windows.Controls.MenuItem)
            {
                lea = LogEventAccess.NoteTextContextMenu;
            }

            try
            {
                LogControl.Write(
                   elementControl.CurrentElement,
                   lea,
                   LogEventType.CreateNewHeading,
                   LogEventStatus.Begin,
                   null);

                Element newHeading = elementControl.CreateNewElement(ElementType.Heading, String.Empty);
                elementControl.InsertElement(newHeading, elementControl.CurrentElement.ParentElement, elementControl.CurrentElement.Position + 1);
                GetFocusToElementTextBox(newHeading, 0, false, false);

                LogControl.Write(
                   elementControl.CurrentElement,
                   lea,
                   LogEventType.CreateNewHeading,
                   LogEventStatus.End,
                   null);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("RibbonCommandNewHeading_Executed\n" + ex.Message);

                LogControl.Write(
                   elementControl.CurrentElement,
                   lea,
                   LogEventType.CreateNewHeading,
                   LogEventStatus.Error,
                   LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void RibbonCommandNewNote_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            LogEventAccess lea = LogEventAccess.Ribbon;
            if (sender is Planz.MainWindow)
            {
                lea = LogEventAccess.Ribbon;
            }
            else if (sender is System.Windows.Controls.MenuItem)
            {
                lea = LogEventAccess.NoteTextContextMenu;
            }

            try
            {
                LogControl.Write(
                   elementControl.CurrentElement,
                   lea,
                   LogEventType.CreateNewNote,
                   LogEventStatus.Begin,
                   null);

                Element newNote = elementControl.CreateNewElement(ElementType.Note, String.Empty);
                elementControl.InsertElement(newNote, elementControl.CurrentElement.ParentElement, elementControl.CurrentElement.Position + 1);
                GetFocusToElementTextBox(newNote, 0, false, false);

                LogControl.Write(
                   elementControl.CurrentElement,
                   lea,
                   LogEventType.CreateNewNote,
                   LogEventStatus.End,
                   LogEventInfo.FileName + LogControl.COMMA + elementControl.CurrentElement.AssociationURI);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("RibbonCommandNewNote_Executed\n" + ex.Message);

                LogControl.Write(
                   elementControl.CurrentElement,
                   lea,
                   LogEventType.CreateNewNote,
                   LogEventStatus.Error,
                   LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void RibbonCommandNewTextDocument_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            LogEventAccess lea = LogEventAccess.Ribbon;
            if (sender is Planz.MainWindow)
            {
                lea = LogEventAccess.Ribbon;
            }
            else if (sender is System.Windows.Controls.MenuItem)
            {
                lea = LogEventAccess.NoteTextContextMenu;
            }

            try
            {
                LogControl.Write(
                   elementControl.CurrentElement,
                   lea,
                   LogEventType.CreateNewTextDocument,
                   LogEventStatus.Begin,
                   null);

                HandleICC(elementControl.CurrentElement, ICCAssociationType.NotepadTextDocument);

                LogControl.Write(
                   elementControl.CurrentElement,
                   lea,
                   LogEventType.CreateNewTextDocument,
                   LogEventStatus.End,
                   LogEventInfo.FileName + LogControl.COMMA + elementControl.CurrentElement.AssociationURI);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("RibbonCommandNewTextDocument_Executed\n" + ex.Message);

                LogControl.Write(
                   elementControl.CurrentElement,
                   lea,
                   LogEventType.CreateNewTextDocument,
                   LogEventStatus.Error,
                   LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void RibbonCommandNewWord_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            LogEventAccess lea = LogEventAccess.Ribbon;
            if (sender is Planz.MainWindow)
            {
                lea = LogEventAccess.Ribbon;
            }
            else if (sender is System.Windows.Controls.MenuItem)
            {
                lea = LogEventAccess.NoteTextContextMenu;
            }

            try
            {
                LogControl.Write(
                   elementControl.CurrentElement,
                   lea,
                   LogEventType.CreateNewWordDocument,
                   LogEventStatus.Begin,
                   null);

                HandleICC(elementControl.CurrentElement, ICCAssociationType.WordDocument);

                LogControl.Write(
                   elementControl.CurrentElement,
                   lea,
                   LogEventType.CreateNewWordDocument,
                   LogEventStatus.End,
                   null);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("RibbonCommandNewWord_Executed\n" + ex.Message);

                LogControl.Write(
                   elementControl.CurrentElement,
                   lea,
                   LogEventType.CreateNewWordDocument,
                   LogEventStatus.Error,
                   LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void RibbonCommandNewExcel_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            LogEventAccess lea = LogEventAccess.Ribbon;
            if (sender is Planz.MainWindow)
            {
                lea = LogEventAccess.Ribbon;
            }
            else if (sender is System.Windows.Controls.MenuItem)
            {
                lea = LogEventAccess.NoteTextContextMenu;
            }

            try
            {
                LogControl.Write(
                   elementControl.CurrentElement,
                   lea,
                   LogEventType.CreateNewExcelSpreadsheet,
                   LogEventStatus.Begin,
                   null);

                HandleICC(elementControl.CurrentElement, ICCAssociationType.ExcelSpreadsheet);
                
                LogControl.Write(
                   elementControl.CurrentElement,
                   lea,
                   LogEventType.CreateNewExcelSpreadsheet,
                   LogEventStatus.End,
                   LogEventInfo.FileName + LogControl.COMMA + elementControl.CurrentElement.AssociationURI);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("RibbonCommandNewExcel_Executed\n" + ex.Message);

                LogControl.Write(
                   elementControl.CurrentElement,
                   lea,
                   LogEventType.CreateNewExcelSpreadsheet,
                   LogEventStatus.Error,
                   LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void RibbonCommandNewPowerPoint_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            LogEventAccess lea = LogEventAccess.Ribbon;
            if (sender is Planz.MainWindow)
            {
                lea = LogEventAccess.Ribbon;
            }
            else if (sender is System.Windows.Controls.MenuItem)
            {
                lea = LogEventAccess.NoteTextContextMenu;
            }

            try
            {
                LogControl.Write(
                   elementControl.CurrentElement,
                   lea,
                   LogEventType.CreateNewPowerPointPresentation,
                   LogEventStatus.Begin,
                   null);

                HandleICC(elementControl.CurrentElement, ICCAssociationType.PowerPointPresentation);

                LogControl.Write(
                   elementControl.CurrentElement,
                   lea,
                   LogEventType.CreateNewPowerPointPresentation,
                   LogEventStatus.End,
                   LogEventInfo.FileName + LogControl.COMMA + elementControl.CurrentElement.AssociationURI);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("RibbonCommandNewPowerPoint_Executed\n" + ex.Message);

                LogControl.Write(
                   elementControl.CurrentElement,
                   lea,
                   LogEventType.CreateNewPowerPointPresentation,
                   LogEventStatus.Error,
                   LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void RibbonCommandNewOneNote_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            LogEventAccess lea = LogEventAccess.Ribbon;
            if (sender is Planz.MainWindow)
            {
                lea = LogEventAccess.Ribbon;
            }
            else if (sender is System.Windows.Controls.MenuItem)
            {
                lea = LogEventAccess.NoteTextContextMenu;
            }

            try
            {
                LogControl.Write(
                   elementControl.CurrentElement,
                   lea,
                   LogEventType.CreateNewOneNoteSection,
                   LogEventStatus.Begin,
                   null);

                HandleICC(elementControl.CurrentElement, ICCAssociationType.OneNoteSection);

                LogControl.Write(
                   elementControl.CurrentElement,
                   lea,
                   LogEventType.CreateNewOneNoteSection,
                   LogEventStatus.End,
                   LogEventInfo.FileName + LogControl.COMMA + elementControl.CurrentElement.AssociationURI);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("RibbonCommandNewOneNote_Executed\n" + ex.Message);

                LogControl.Write(
                   elementControl.CurrentElement,
                   lea,
                   LogEventType.CreateNewOneNoteSection,
                   LogEventStatus.Error,
                   LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void RibbonCommandNewEmail_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            LogEventAccess lea = LogEventAccess.Ribbon;
            if (sender is Planz.MainWindow)
            {
                lea = LogEventAccess.Ribbon;
            }
            else if (sender is System.Windows.Controls.MenuItem)
            {
                lea = LogEventAccess.NoteTextContextMenu;
            }

            try
            {
                LogControl.Write(
                   elementControl.CurrentElement,
                   lea,
                   LogEventType.CreateNewOutlookEmail,
                   LogEventStatus.Begin,
                   null);

                HandleICC(elementControl.CurrentElement, ICCAssociationType.OutlookEmailMessage);

                LogControl.Write(
                   elementControl.CurrentElement,
                   lea,
                   LogEventType.CreateNewOutlookEmail,
                   LogEventStatus.End,
                   LogEventInfo.FileName + LogControl.COMMA + elementControl.CurrentElement.AssociationURI);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("RibbonCommandNewEmail_Executed\n" + ex.Message);

                LogControl.Write(
                   elementControl.CurrentElement,
                   lea,
                   LogEventType.CreateNewOutlookEmail,
                   LogEventStatus.Error,
                   LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void RibbonCommandNewTwitter_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            LogEventAccess lea = LogEventAccess.Ribbon;
            if (sender is Planz.MainWindow)
            {
                lea = LogEventAccess.Ribbon;
            }
            else if (sender is System.Windows.Controls.MenuItem)
            {
                lea = LogEventAccess.NoteTextContextMenu;
            }

            try
            {
                LogControl.Write(
                   elementControl.CurrentElement,
                   lea,
                   LogEventType.CreateNewTwitterUpdate,
                   LogEventStatus.Begin,
                   null);

                HandleICC(elementControl.CurrentElement, ICCAssociationType.TwitterUpdate);

                LogControl.Write(
                   elementControl.CurrentElement,
                   lea,
                   LogEventType.CreateNewTwitterUpdate,
                   LogEventStatus.End,
                   null);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("RibbonCommandNewTwitter_Executed\n" + ex.Message);

                LogControl.Write(
                   elementControl.CurrentElement,
                   lea,
                   LogEventType.CreateNewTwitterUpdate,
                   LogEventStatus.Error,
                   LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void RibbonDropDownButton_ActiveItemList_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            try
            {
                StartLoadingUI();
                
                this.RibbonDropDownButton_ActiveItemList.Items.Clear();
                ActiveWindow aw = new ActiveWindow();
                foreach (InfoItem ii in aw.GetActiveItemList())
                {
                    MenuItem mi = new MenuItem();
                    if (ii.Title.Length > ACTIVEITEM_TITLE_LENGTH)
                    {
                        mi.Header = ii.Title.Substring(0, ACTIVEITEM_TITLE_LENGTH) + "...";
                    }
                    else
                    {
                        mi.Header = ii.Title;
                    }
                    switch (ii.Type)
                    {
                        case InfoItemType.Web:
                            mi.Icon = new Image
                            {
                                Source = new BitmapImage(new Uri(FileTypeHandler.GetIcon(ElementAssociationType.Web, ii.Uri), UriKind.Absolute))
                            };
                            break;
                        case InfoItemType.File:
                            mi.Icon = new Image
                            {
                                Source = new BitmapImage(new Uri(FileTypeHandler.GetIcon(ElementAssociationType.File, ii.Uri), UriKind.Absolute))
                            };
                            break;
                        case InfoItemType.Email:
                            mi.Icon = new Image
                            {
                                Source = new BitmapImage(new Uri(FileTypeHandler.GetIcon(ElementAssociationType.Email, ii.Uri), UriKind.Absolute))
                            };
                            break;
                    };
                    mi.Tag = ii;
                    mi.ToolTip = ii.Title;
                    mi.Click += new RoutedEventHandler(RibbonDropDownButton_ActiveItem_Click);
                    this.RibbonDropDownButton_ActiveItemList.Items.Add(mi);
                }

                if (this.RibbonDropDownButton_ActiveItemList.Items.Count == 0)
                {
                    MenuItem mi = new MenuItem();
                    mi.Header = "No active item found";
                    this.RibbonDropDownButton_ActiveItemList.Items.Add(mi);
                }

                EndLoadingUI();
            }
            catch (Exception ex)
            {
                this.RibbonDropDownButton_ActiveItemList.Items.Clear();
                //MessageBox.Show("RibbonDropDownButton_ActiveItemList_MouseEnter\n" + ex.Message);
            }
        }

        private void RibbonDropDownButton_ActiveItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                LogControl.Write(
                   elementControl.CurrentElement,
                   LogEventAccess.Ribbon,
                   LogEventType.LinkToActiveItem,
                   LogEventStatus.Begin,
                   null);

                MenuItem mi = sender as MenuItem;
                InfoItem ii = mi.Tag as InfoItem;
                switch (ii.Type)
                {
                    case InfoItemType.Web:
                        elementControl.AddAssociation(elementControl.CurrentElement, ii.Uri, ElementAssociationType.Web, ii.Title);
                        break;
                    case InfoItemType.File:
                        elementControl.AddAssociation(elementControl.CurrentElement, ii.Uri, ElementAssociationType.FileShortcut, ii.Title);
                        break;
                    case InfoItemType.Email:
                        elementControl.AddAssociation(elementControl.CurrentElement, ii.Uri, ElementAssociationType.Email, ii.Title);
                        break;
                };

                GetFocusToElementTextBox(elementControl.CurrentElement, elementControl.CurrentElement.NoteText.Length, false, false);

                LogControl.Write(
                   elementControl.CurrentElement,
                   LogEventAccess.Ribbon,
                   LogEventType.LinkToActiveItem,
                   LogEventStatus.End,
                   null);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("RibbonDropDownButton_ActiveItem_Click\n" + ex.Message);

                LogControl.Write(
                   elementControl.CurrentElement,
                   LogEventAccess.Ribbon,
                   LogEventType.LinkToActiveItem,
                   LogEventStatus.Error,
                   LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void RibbonCommandLinkToFile_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            LogEventAccess lea = LogEventAccess.Ribbon;
            if (sender is Planz.MainWindow)
            {
                lea = LogEventAccess.Ribbon;
            }
            else if (sender is System.Windows.Controls.MenuItem)
            {
                lea = LogEventAccess.NoteTextContextMenu;
            }

            try
            {
                LogControl.Write(
                   elementControl.CurrentElement,
                   lea,
                   LogEventType.LinkToFile,
                   LogEventStatus.Begin,
                   null);

                System.Windows.Forms.OpenFileDialog dialog = new System.Windows.Forms.OpenFileDialog();
                dialog.CheckFileExists = true;
                dialog.Multiselect = false;
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    elementControl.AddAssociation(elementControl.CurrentElement, dialog.FileName, ElementAssociationType.FileShortcut, null);

                    GetFocusToElementTextBox(elementControl.CurrentElement, elementControl.CurrentElement.NoteText.Length, false, false);

                    LogControl.Write(
                       elementControl.CurrentElement,
                       lea,
                       LogEventType.LinkToFile,
                       LogEventStatus.End,
                       LogEventInfo.LinkedFile + LogControl.COMMA + dialog.FileName);
                }
                else
                {
                    LogControl.Write(
                       elementControl.CurrentElement,
                       lea,
                       LogEventType.LinkToFile,
                       LogEventStatus.Cancel,
                       null);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("RibbonCommandLinkToFile_Executed\n" + ex.Message);

                LogControl.Write(
                   elementControl.CurrentElement,
                   lea,
                   LogEventType.LinkToFile,
                   LogEventStatus.Error,
                   LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void RibbonCommandLinkToFolder_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            LogEventAccess lea = LogEventAccess.Ribbon;
            if (sender is Planz.MainWindow)
            {
                lea = LogEventAccess.Ribbon;
            }
            else if (sender is System.Windows.Controls.MenuItem)
            {
                lea = LogEventAccess.NoteTextContextMenu;
            }

            try
            {
                LogControl.Write(
                   elementControl.CurrentElement,
                   lea,
                   LogEventType.LinkToFolder,
                   LogEventStatus.Begin,
                   null);

                System.Windows.Forms.FolderBrowserDialog dialog = new System.Windows.Forms.FolderBrowserDialog();
                dialog.SelectedPath = StartProcess.START_PATH;
                dialog.ShowNewFolderButton = true;
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {      
                    elementControl.AddAssociation(elementControl.CurrentElement, dialog.SelectedPath, ElementAssociationType.FolderShortcut, null);
                    elementControl.Promote(elementControl.CurrentElement);

                    GetFocusToElementTextBox(elementControl.CurrentElement, elementControl.CurrentElement.NoteText.Length, false, false);

                    LogControl.Write(
                       elementControl.CurrentElement,
                       lea,
                       LogEventType.LinkToFolder,
                       LogEventStatus.End,
                       LogEventInfo.LinkedFolder + LogControl.COMMA + dialog.SelectedPath);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("RibbonCommandLinkToFolder_Executed\n" + ex.Message);

                LogControl.Write(
                   elementControl.CurrentElement,
                   lea,
                   LogEventType.LinkToFolder,
                   LogEventStatus.Error,
                   LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void RibbonCommandLabelWith_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            LogEventAccess lea = LogEventAccess.Ribbon;
            if (sender is Planz.MainWindow)
            {
                lea = LogEventAccess.Ribbon;
            }
            else if (sender is System.Windows.Controls.MenuItem)
            {
                lea = LogEventAccess.NoteTextContextMenu;
            }


            Element element = elementControl.CurrentElement;

            try
            {
                LogControl.Write(
                    element,
                    lea,
                    LogEventType.LabelWith,
                    LogEventStatus.Begin,
                    null);
                
                if (element != null && element.IsHeading)
                {
                    //Ask user to select the folder
                    System.Windows.Forms.FolderBrowserDialog browse = new System.Windows.Forms.FolderBrowserDialog();
                    browse.Description = "Please select a heading to label with";

                    browse.RootFolder = Environment.SpecialFolder.Desktop;
                    browse.SelectedPath = StartProcess.LIFE_PATH;
                    browse.ShowNewFolderButton = false;

                    if (browse.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string path = browse.SelectedPath + System.IO.Path.DirectorySeparatorChar; ;
                        string folderName = path.Substring(path.LastIndexOf("\\") + 1);

                        //create a shortcut note at the selected folder.
                        elementControl.LabelWith(element, path);

                        LogControl.Write(
                            element,
                            lea,
                            LogEventType.LabelWith,
                            LogEventStatus.End,
                            LogEventInfo.Location + LogControl.COMMA + path);
                    }
                    else
                    {
                        LogControl.Write(
                            element,
                            lea,
                            LogEventType.LabelWith,
                            LogEventStatus.Cancel,
                            null);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("RibbonCommandLabelWith_Executed\n" + ex.Message);

                LogControl.Write(
                    element,
                    lea,
                    LogEventType.LabelWith,
                    LogEventStatus.Error,
                    LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void RibbonCommandPromote_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            try
            {
                ShiftTabKeyPressed(sender, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show("RibbonCommandPromote_Executed\n" + ex.Message);
            }
        }

        private void RibbonCommandDemote_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            try
            {
                TabKeyPressed(sender, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("RibbonCommandDemote_Executed\n" + ex.Message);
            }
        }

        private void RibbonCommandMoveUp_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            try
            {
                CtrlUpKeyPressed(sender, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("RibbonCommandMoveUp_Executed\n" + ex.Message);
            }
        }

        private void RibbonCommandMoveDown_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            try
            {
                CtrlDownKeyPressed(sender, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("RibbonCommandMoveDown_Executed\n" + ex.Message);
            }
        }

        private void RibbonCommandExpand_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            try
            {
                PlanTreeViewItem_Expanded(elementControl.CurrentElement, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("RibbonCommandExpand_Executed\n" + ex.Message);
            }
        }

        private void RibbonCommandCollapse_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            try
            {
                PlanTreeViewItem_Collapsed(elementControl.CurrentElement, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("RibbonCommandCollapse_Executed\n" + ex.Message);
            }
        }

        private void RibbonNavigationTab_Checked(object sender, RoutedEventArgs e)
        {
            try
            {
                if (this.NavigationTabList != null)
                {
                    Properties.Settings.Default.NavigationTabVisibility = true;
                    Properties.Settings.Default.Save();
                    this.NavigationTabList.Visibility = Visibility.Visible;

                    this.NavigationColumn.SetValue(ColumnDefinition.WidthProperty, new GridLength(Properties.Settings.Default.NavigationTabWidth));

                    UpdateLayout();
                    ResizeElementTextBox(elementControl.Root, new Dictionary<int, double>());

                    LogControl.Write(
                        elementControl.Root,
                        LogEventAccess.Ribbon,
                        LogEventType.ShowNavigationTree,
                        LogEventStatus.NULL,
                        null);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("RibbonNavigationTab_Checked\n" + ex.Message);

                LogControl.Write(
                    elementControl.Root,
                    LogEventAccess.Ribbon,
                    LogEventType.ShowNavigationTree,
                    LogEventStatus.Error,
                    LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void RibbonNavigationTab_Unchecked(object sender, RoutedEventArgs e)
        {
            try
            {
                if (this.NavigationTabList != null)
                {
                    Properties.Settings.Default.NavigationTabVisibility = false;
                    Properties.Settings.Default.Save();
                    this.NavigationTabList.Visibility = Visibility.Collapsed;

                    this.NavigationColumn.SetValue(ColumnDefinition.WidthProperty, new GridLength(0.0));

                    UpdateLayout();
                    ResizeElementTextBox(elementControl.Root, new Dictionary<int, double>());

                    LogControl.Write(
                        elementControl.Root,
                        LogEventAccess.Ribbon,
                        LogEventType.HideNavigationTree,
                        LogEventStatus.NULL,
                        null);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("RibbonNavigationTab_Unchecked\n" + ex.Message);

                LogControl.Write(
                    elementControl.Root,
                    LogEventAccess.Ribbon,
                    LogEventType.HideNavigationTree,
                    LogEventStatus.Error,
                    LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void ShowOutline_Checked(object sender, RoutedEventArgs e)
        {
            try
            {
                if (elementControl != null)
                {
                    Properties.Settings.Default.ShowOutline = true;
                    Properties.Settings.Default.Save();

                    ShowOutlineIcon(elementControl.Root, true);

                    LogControl.Write(
                        elementControl.Root,
                        LogEventAccess.Ribbon,
                        LogEventType.ShowOutline,
                        LogEventStatus.NULL,
                        null);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("ShowOutline_Checked\n" + ex.Message);

                LogControl.Write(
                    elementControl.Root,
                    LogEventAccess.Ribbon,
                    LogEventType.ShowOutline,
                    LogEventStatus.Error,
                    LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void ShowOutline_Unchecked(object sender, RoutedEventArgs e)
        {
            try
            {
                if (elementControl != null)
                {
                    Properties.Settings.Default.ShowOutline = false;
                    Properties.Settings.Default.Save();

                    ShowOutlineIcon(elementControl.Root, false);
                    GetFocusToElementTextBox(elementControl.CurrentElement, elementControl.CurrentElement.NoteText.Length, false, false);

                    LogControl.Write(
                        elementControl.Root,
                        LogEventAccess.Ribbon,
                        LogEventType.HideOutline,
                        LogEventStatus.NULL,
                        null);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("ShowOutline_Unchecked\n" + ex.Message);

                LogControl.Write(
                    elementControl.Root,
                    LogEventAccess.Ribbon,
                    LogEventType.HideOutline,
                    LogEventStatus.Error,
                    LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void RibbonSpellCheck_Checked(object sender, RoutedEventArgs e)
        {
            try
            {
                Properties.Settings.Default.IsSpellCheckEnabled = true;
                Properties.Settings.Default.Save();

                LogControl.Write(
                    elementControl.Root,
                    LogEventAccess.Ribbon,
                    LogEventType.EnableSpellCheck,
                    LogEventStatus.NULL,
                    null);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("RibbonSpellCheck_Checked\n" + ex.Message);

                LogControl.Write(
                    elementControl.Root,
                    LogEventAccess.Ribbon,
                    LogEventType.EnableSpellCheck,
                    LogEventStatus.Error,
                    LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void RibbonSpellCheck_Unchecked(object sender, RoutedEventArgs e)
        {
            try
            {
                Properties.Settings.Default.IsSpellCheckEnabled = false;
                Properties.Settings.Default.Save();

                LogControl.Write(
                    elementControl.Root,
                    LogEventAccess.Ribbon,
                    LogEventType.DisableSpellCheck,
                    LogEventStatus.NULL,
                    null);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("RibbonSpellCheck_Unchecked\n" + ex.Message);

                LogControl.Write(
                    elementControl.Root,
                    LogEventAccess.Ribbon,
                    LogEventType.DisableSpellCheck,
                    LogEventStatus.Error,
                    LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        /*
        private void RibbonHeadingLevel1FontFamily_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                string selected = ((Microsoft.Windows.Controls.Ribbon.RibbonComboBoxItem)((Microsoft.Windows.Controls.Ribbon.RibbonComboBox)(sender)).SelectedItem).Content.ToString();
                Properties.Settings.Default.HeadingLevel1FontFamily = selected;
                Properties.Settings.Default.Save();

                ChangeElementTextFontFamily(ElementType.Heading, new FontFamily(selected));
            }
            catch (Exception ex)
            {
                MessageBox.Show("RibbonHeadingLevel1FontFamily_SelectionChanged\n" + ex.Message);
            }
        }

        private void RibbonHeadingLevel1FontSize_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                int selected = Int32.Parse(((Microsoft.Windows.Controls.Ribbon.RibbonComboBoxItem)((Microsoft.Windows.Controls.Ribbon.RibbonComboBox)(sender)).SelectedItem).Content.ToString());
                Properties.Settings.Default.HeadingLevel1FontSize = selected;
                Properties.Settings.Default.Save();

                ChangeElementTextFontSize(ElementType.Heading, selected);
            }
            catch (Exception ex)
            {
                MessageBox.Show("RibbonHeadingLevel1FontSize_SelectionChanged\n" + ex.Message);
            }
        }

        private void RibbonNoteFontFamily_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                string selected = ((Microsoft.Windows.Controls.Ribbon.RibbonComboBoxItem)((Microsoft.Windows.Controls.Ribbon.RibbonComboBox)(sender)).SelectedItem).Content.ToString();
                Properties.Settings.Default.NoteFontFamily = selected;
                Properties.Settings.Default.Save();

                ChangeElementTextFontFamily(ElementType.Note, new FontFamily(selected));
            }
            catch (Exception ex)
            {
                MessageBox.Show("RibbonNoteFontFamily_SelectionChanged\n" + ex.Message);
            }
        }

        private void RibbonNoteFontSize_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                int selected = Int32.Parse(((Microsoft.Windows.Controls.Ribbon.RibbonComboBoxItem)((Microsoft.Windows.Controls.Ribbon.RibbonComboBox)(sender)).SelectedItem).Content.ToString());
                Properties.Settings.Default.NoteFontSize = selected;
                Properties.Settings.Default.Save();

                ChangeElementTextFontSize(ElementType.Note, selected);
            }
            catch (Exception ex)
            {
                MessageBox.Show("RibbonNoteFontSize_SelectionChanged\n" + ex.Message);
            }
        }
        */

        private void RibbonCommandUserManual_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start("http://kftf.ischool.washington.edu/planner/User_Manual/HTML/user_manual.html");
                LogControl.Write(
                    elementControl.Root,
                    LogEventAccess.AppMenu,
                    LogEventType.UserManual,
                    LogEventStatus.NULL,
                    null);

            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("RibbonCommandUserManual_Executed\n" + ex.Message);

                LogControl.Write(
                    elementControl.Root,
                    LogEventAccess.AppMenu,
                    LogEventType.UserManual,
                    LogEventStatus.Error,
                    LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void RibbonCommandFeedback_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start("http://www.talesofpim.org/forum/");
                LogControl.Write(
                    elementControl.Root,
                    LogEventAccess.AppMenu,
                    LogEventType.Feedback,
                    LogEventStatus.NULL,
                    null);

            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("RibbonCommandFeedback_Executed\n" + ex.Message);

                LogControl.Write(
                    elementControl.Root,
                    LogEventAccess.AppMenu,
                    LogEventType.Feedback,
                    LogEventStatus.Error,
                    LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void RibbonCommandAbout_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            try
            {
                AboutWindow aw = new AboutWindow();
                aw.ShowDialog();
                LogControl.Write(
                    elementControl.Root,
                    LogEventAccess.AppMenu,
                    LogEventType.AboutWindow,
                    LogEventStatus.NULL,
                    null);

            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("RibbonCommandAbout_Executed\n" + ex.Message);

                LogControl.Write(
                    elementControl.Root,
                    LogEventAccess.AppMenu,
                    LogEventType.AboutWindow,
                    LogEventStatus.Error,
                    LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void HandleICC(Element element, ICCAssociationType type)
        {
            // Note: defaultName does NOT contain file extension
            string defaultName = String.Empty;

            bool insertBelow = true;
            TextBox tb = GetTextBox(element);
            if (tb.SelectionStart == 0 && tb.Text != String.Empty)
            {
                insertBelow = false;
            }

            if (tb.SelectedText != String.Empty)
            {
                defaultName = tb.SelectedText;
                insertBelow = true;
            }
            else
            {
                defaultName = element.NoteText;
            }

            if (defaultName.Length > StartProcess.MAX_EXTRACTNAME_LENGTH)
                defaultName = defaultName.Substring(0, StartProcess.MAX_EXTRACTNAME_LENGTH);
            while (defaultName.EndsWith("."))
                defaultName = defaultName.TrimEnd('.');

            switch (type)
            {
                case ICCAssociationType.TwitterUpdate:
                    TwitterWindow tw = new TwitterWindow(defaultName);
                    if (tw.ShowDialog().Value == true)
                    {
                        Tweet tweet = new Tweet{ Username = tw.Username, Password = tw.Password, Message = tw.Tweet, };
                        elementControl.ICC(element, tweet, insertBelow, type);
                    }
                    break;
                case ICCAssociationType.OutlookEmailMessage:
                    if (defaultName.Trim() == String.Empty)
                    {
                        defaultName = "New Email";
                    }
                    elementControl.ICC(element, defaultName, insertBelow, type);
                    break;
                default:
                    string fileFullName = element.Path + ICCFileNameHandler.GenerateFileName(defaultName, type);
                    FileNameWindow fnw = new FileNameWindow(fileFullName, type);
                    if (fnw.ShowDialog().Value == true)
                    {
                        elementControl.ICC(element, fnw.FileName, insertBelow, type);
                    }
                    break;
            };

            GetFocusToElementTextBox(elementControl.CurrentElement, elementControl.CurrentElement.NoteText.Length, false, false);
        }


        #region Power D

        private void RibbonCommandFlag_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            Element currentElement = elementControl.CurrentElement;
            try
            {
                if (currentElement != null)
                {
                    if (currentElement.IsHeading)
                    {
                        FlagWindow fw = new FlagWindow(currentElement);
                        if (fw.ShowDialog().Value)
                        {
                            elementControl.Flag(currentElement,
                                fw.HasStart, fw.StartDate, fw.StartTime, fw.StartAllDay,
                                fw.HasDue, fw.DueDate, fw.DueTime, fw.DueAllDay,
                                fw.AddToToday, fw.AddToReminder, fw.AddToTask);

                            LogControl.Write(
                                currentElement,
                                LogEventAccess.Ribbon,
                                LogEventType.Flag,
                                LogEventStatus.NULL,
                                LogEventInfo.HasStart + LogControl.COMMA + fw.HasStart.ToString() + LogControl.DELIMITER +
                                LogEventInfo.StartDate + LogControl.COMMA + fw.StartDate.ToShortDateString() + LogControl.DELIMITER +
                                LogEventInfo.StartTime + LogControl.COMMA + fw.StartTime.ToString() + LogControl.DELIMITER +
                                LogEventInfo.StartAllDay + LogControl.COMMA + fw.StartAllDay.ToString() + LogControl.DELIMITER +
                                LogEventInfo.DueDate + LogControl.COMMA + fw.DueDate.ToShortDateString() + LogControl.DELIMITER +
                                LogEventInfo.DueTime + LogControl.COMMA + fw.DueTime.ToString() + LogControl.DELIMITER +
                                LogEventInfo.AddToToday + LogControl.COMMA + fw.AddToToday.ToString() + LogControl.DELIMITER +
                                LogEventInfo.AddToReminder + LogControl.COMMA + fw.AddToReminder.ToString() + LogControl.DELIMITER +
                                LogEventInfo.AddToTask + LogControl.COMMA + fw.AddToTask.ToString()
                                );
                        }
                    }
                    else
                    {
                        elementControl.Flag(currentElement);

                        LogControl.Write(
                            currentElement,
                            LogEventAccess.Ribbon,
                            LogEventType.Flag,
                            LogEventStatus.NULL,
                            null);
                    }

                    GetFocusToElementTextBox(currentElement, -1, false, false);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("RibbonCommandFlag_Executed\n" + ex.Message);

                LogControl.Write(
                    currentElement,
                    LogEventAccess.Ribbon,
                    LogEventType.Flag,
                    LogEventStatus.Error,
                    LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void RibbonCommandCheck_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            Element currentElement = elementControl.CurrentElement;
            try
            {
                if (currentElement != null)
                {
                    if (currentElement.IsHeading)
                    {
                        CheckWindow cw = new CheckWindow(currentElement);
                        if (cw.ShowDialog().Value)
                        {
                            elementControl.Check(currentElement, cw.RemoveFromToday);

                            LogControl.Write(
                                currentElement,
                                LogEventAccess.Ribbon,
                                LogEventType.Check,
                                LogEventStatus.NULL,
                                LogEventInfo.RemoveFromToday + LogControl.COMMA + cw.removeFromToday.ToString());
                        }
                    }
                    else
                    {
                        elementControl.Check(currentElement);

                        LogControl.Write(
                            currentElement,
                            LogEventAccess.Ribbon,
                            LogEventType.Check,
                            LogEventStatus.NULL,
                            null);
                    }

                    GetFocusToElementTextBox(currentElement, -1, false, false);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("RibbonCommandCheck_Executed\n" + ex.Message);

                LogControl.Write(
                    currentElement,
                    LogEventAccess.Ribbon,
                    LogEventType.Check,
                    LogEventStatus.Error,
                    LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void RibbonCommandReset_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            Element currentElement = elementControl.CurrentElement;
            try
            {
                if (currentElement != null)
                {
                    elementControl.Uncheck(currentElement);

                    LogControl.Write(
                        currentElement,
                        LogEventAccess.Ribbon,
                        LogEventType.Uncheck,
                        LogEventStatus.NULL,
                        null);

                    GetFocusToElementTextBox(currentElement, -1, false, false);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("RibbonCommandReset_Executed\n" + ex.Message);

                LogControl.Write(
                    currentElement,
                    LogEventAccess.Ribbon,
                    LogEventType.Uncheck,
                    LogEventStatus.Error,
                    LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void RibbonButton_PowerDDone_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                LogControl.Write(
                   elementControl.CurrentElement,
                   LogEventAccess.Ribbon,
                   LogEventType.PowerDDone,
                   LogEventStatus.Begin,
                   null);

                DateTime dt = DateTime.Today;
                foreach (MenuItem menuItem in RibbonButton_PowerDDone.Items)
                {
                    if (menuItem.IsChecked)
                    {
                        switch (menuItem.Header.ToString())
                        {
                            case "Today":
                                dt = DateTime.Now;
                                break;
                            case "Yesterday":
                                dt = DateTime.Today.AddDays(-1);
                                break;
                            default:
                                dt = DateTime.Parse(menuItem.Header.ToString());
                                break;
                        };
                    }
                }

                Element next = elementControl.CurrentElement.ElementBelow;

                elementControl.PowerDDone(elementControl.CurrentElement, dt);

                if (next == null)
                {
                    GetFocusToElementTextBox(elementControl.CurrentElement, -1, false, false);
                }
                else
                {
                    GetFocusToElementTextBox(next, -1, false, false);
                }

                LogControl.Write(
                   elementControl.CurrentElement,
                   LogEventAccess.Ribbon,
                   LogEventType.PowerDDone,
                   LogEventStatus.End,
                   LogEventInfo.Date + LogControl.COMMA + dt.ToString());
            }
            catch (Exception ex)
            {
                //MessageBox.Show("RibbonButton_PowerDDone_Click\n" + ex.Message);

                LogControl.Write(
                   elementControl.CurrentElement,
                   LogEventAccess.Ribbon,
                   LogEventType.PowerDDone,
                   LogEventStatus.Error,
                   LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void RibbonButton_PowerDDoneMenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                LogControl.Write(
                   elementControl.CurrentElement,
                   LogEventAccess.Ribbon,
                   LogEventType.PowerDDone,
                   LogEventStatus.Begin,
                   null);

                MenuItem mi = sender as MenuItem;
                DateTime dt;
                switch (mi.Header.ToString())
                {
                    case "Today":
                        dt = DateTime.Now;
                        break;
                    case "Yesterday":
                        dt = DateTime.Today.AddDays(-1);
                        break;
                    case "Earlier...":
                        DatePickerWindow dpw = new DatePickerWindow(DatePickerCondition.PowerDEarlier);
                        if (dpw.ShowDialog().Value == true)
                        {
                            MenuItem new_mi = new MenuItem();
                            new_mi.Header = dpw.SelectedDate.ToShortDateString();
                            new_mi.IsCheckable = true;
                            new_mi.Click += RibbonButton_PowerDDoneMenuItem_Click;

                            bool dup = false;
                            foreach (MenuItem menuItem in RibbonButton_PowerDDone.Items)
                            {
                                if (menuItem.Header.ToString() == new_mi.Header.ToString())
                                {
                                    mi = menuItem;
                                    dup = true;
                                    break;
                                }
                            }
                            if (!dup)
                            {
                                mi = new_mi;
                                this.RibbonButton_PowerDDone.Items.Add(new_mi);
                            }
                            dt = dpw.SelectedDate;
                        }
                        else
                        {
                            mi.IsChecked = false;
                            return;
                        }
                        break;
                    default:
                        dt = DateTime.Parse(mi.Header.ToString());
                        break;
                };

                foreach (MenuItem menuItem in RibbonButton_PowerDDone.Items)
                {
                    if (menuItem == mi)
                    {
                        menuItem.IsChecked = true;
                    }
                    else
                    {
                        menuItem.IsChecked = false;
                    }
                }

                elementControl.PowerDDone(elementControl.CurrentElement, dt);

                RibbonCommandPowerDNext_Executed(null, null);

                LogControl.Write(
                   elementControl.CurrentElement,
                   LogEventAccess.Ribbon,
                   LogEventType.PowerDDone,
                   LogEventStatus.End,
                   LogEventInfo.Date + LogControl.COMMA + dt.ToString());
            }
            catch (Exception ex)
            {
                //MessageBox.Show("RibbonButton_PowerDDoneMenuItem_Click\n" + ex.Message);

                LogControl.Write(
                   elementControl.CurrentElement,
                   LogEventAccess.Ribbon,
                   LogEventType.PowerDDone,
                   LogEventStatus.Error,
                   LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void RibbonButton_PowerDDefer_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                LogControl.Write(
                   elementControl.CurrentElement,
                   LogEventAccess.Ribbon,
                   LogEventType.PowerDDefer,
                   LogEventStatus.Begin,
                   null);

                DateTime dt = DateTime.Today;
                foreach (MenuItem menuItem in RibbonButton_PowerDDefer.Items)
                {
                    if (menuItem.IsChecked)
                    {
                        switch (menuItem.Header.ToString())
                        {
                            case "Tomorrow":
                                dt = DateTime.Today.AddDays(1);
                                break;
                            case "The day after tomorrow":
                                dt = DateTime.Today.AddDays(2);
                                break;
                            default:
                                dt = DateTime.Parse(menuItem.Header.ToString());
                                break;
                        };
                    }
                }

                Element next = elementControl.CurrentElement.ElementBelow;

                elementControl.PowerDDefer(elementControl.CurrentElement, dt);

                if (next == null)
                {
                    GetFocusToElementTextBox(elementControl.CurrentElement, -1, false, false);
                }
                else
                {
                    GetFocusToElementTextBox(next, -1, false, false);
                }

                LogControl.Write(
                   elementControl.CurrentElement,
                   LogEventAccess.Ribbon,
                   LogEventType.PowerDDefer,
                   LogEventStatus.End,
                   LogEventInfo.Date + LogControl.COMMA + dt.ToString());
            }
            catch (Exception ex)
            {
                //MessageBox.Show("RibbonButton_PowerDDefer_Click\n" + ex.Message);

                LogControl.Write(
                   elementControl.CurrentElement,
                   LogEventAccess.Ribbon,
                   LogEventType.PowerDDefer,
                   LogEventStatus.Error,
                   LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void RibbonButton_PowerDDeferMenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                MenuItem mi = sender as MenuItem;
                DateTime dt;
                switch (mi.Header.ToString())
                {
                    case "Tomorrow":
                        dt = DateTime.Today.AddDays(1);
                        break;
                    case "The day after tomorrow":
                        dt = DateTime.Today.AddDays(2);
                        break;
                    case "Later...":
                        DatePickerWindow dpw = new DatePickerWindow(DatePickerCondition.PowerDLater);
                        if (dpw.ShowDialog().Value == true)
                        {
                            MenuItem new_mi = new MenuItem();
                            new_mi.IsCheckable = true;
                            new_mi.Header = dpw.SelectedDate.ToShortDateString();
                            new_mi.Click += RibbonButton_PowerDDeferMenuItem_Click;

                            bool dup = false;
                            foreach (MenuItem menuItem in RibbonButton_PowerDDone.Items)
                            {
                                if (menuItem.Header.ToString() == new_mi.Header.ToString())
                                {
                                    mi = menuItem;
                                    dup = true;
                                    break;
                                }
                            }
                            if (!dup)
                            {
                                mi = new_mi;
                                this.RibbonButton_PowerDDefer.Items.Add(new_mi);
                            }
                            dt = dpw.SelectedDate;
                        }
                        else
                        {
                            mi.IsChecked = false;
                            return;
                        }
                        break;
                    default:
                        dt = DateTime.Parse(mi.Header.ToString());
                        break;
                };

                foreach (MenuItem menuItem in RibbonButton_PowerDDefer.Items)
                {
                    if (menuItem == mi)
                    {
                        menuItem.IsChecked = true;
                    }
                    else
                    {
                        menuItem.IsChecked = false;
                    }
                }

                elementControl.PowerDDefer(elementControl.CurrentElement, dt);

                RibbonCommandPowerDNext_Executed(null, null);

                LogControl.Write(
                   elementControl.CurrentElement,
                   LogEventAccess.Ribbon,
                   LogEventType.PowerDDefer,
                   LogEventStatus.End,
                   LogEventInfo.Date + LogControl.COMMA + dt.ToString());
            }
            catch (Exception ex)
            {
                //MessageBox.Show("RibbonButton_PowerDDeferMenuItem_Click\n" + ex.Message);

                LogControl.Write(
                   elementControl.CurrentElement,
                   LogEventAccess.Ribbon,
                   LogEventType.PowerDDefer,
                   LogEventStatus.Error,
                   LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void RibbonButton_PowerDDone_MouseEnter(object sender, MouseEventArgs e)
        {
            /*string text = String.Empty;
            DateTime dt = DateTime.Today;
            foreach (MenuItem menuItem in RibbonButton_PowerDDone.Items)
            {
                if (menuItem.IsChecked)
                {
                    switch (menuItem.Header.ToString())
                    {
                        case "Today":
                            dt = DateTime.Now;
                            text = "today ";
                            break;
                        case "Yesterday":
                            dt = DateTime.Today.AddDays(-1);
                            text = "yesterday ";
                            break;
                        default:
                            dt = DateTime.Parse(menuItem.Header.ToString());
                            break;
                    };
                }
            }
            ((Microsoft.Windows.Controls.Ribbon.RibbonCommand)RibbonButton_PowerDDone.Command).ToolTipDescription = "Mark as \"done\" the current association under " + text + dt.ToShortDateString() + ".";
            RibbonButton_PowerDDone.ToolTip = ((Microsoft.Windows.Controls.Ribbon.RibbonCommand)RibbonButton_PowerDDone.Command).ToolTipDescription;*/
        }

        private void RibbonButton_PowerDDefer_MouseEnter(object sender, MouseEventArgs e)
        {
            /*string text = String.Empty;
            DateTime dt = DateTime.Today;
            foreach (MenuItem menuItem in RibbonButton_PowerDDefer.Items)
            {
                if (menuItem.IsChecked)
                {
                    switch (menuItem.Header.ToString())
                    {
                        case "Tomorrow":
                            dt = DateTime.Today.AddDays(1);
                            text = "tomorrow ";
                            break;
                        case "The day after tomorrow":
                            dt = DateTime.Today.AddDays(2);
                            text = "the day after tomorrow ";
                            break;
                        default:
                            dt = DateTime.Parse(menuItem.Header.ToString());
                            break;
                    };
                }
            }
            ((Microsoft.Windows.Controls.Ribbon.RibbonCommand)RibbonButton_PowerDDefer.Command).ToolTipDescription = "Defer the current association to " + text + dt.ToShortDateString() + ". The association will reappear again (in blue) on this date.";
            RibbonButton_PowerDDefer.ToolTip = ((Microsoft.Windows.Controls.Ribbon.RibbonCommand)RibbonButton_PowerDDefer.Command).ToolTipDescription;*/
        }

        private void RibbonButton_PowerDDelegate_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                MenuItem mi_fsd = new MenuItem();
                mi_fsd.Tag = "FolderSelectDialog";
                RibbonButton_PowerDDelegateMenuItem_Click(mi_fsd, null);
            }
            catch (Exception ex)
            {
                //MessageBox.Show("RibbonButton_PowerDDelegate_Click\n" + ex.Message);

                LogControl.Write(
                   elementControl.CurrentElement,
                   LogEventAccess.Ribbon,
                   LogEventType.PowerDDelegate,
                   LogEventStatus.Error,
                   LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void RibbonButton_PowerDDelegate_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            try
            {
                StartLoadingUI();

                while (this.RibbonButton_PowerDDelegate.Items.Count > 0)
                {
                    this.RibbonButton_PowerDDelegate.Items.RemoveAt(0);
                }
                MenuItem mi_des = new MenuItem();
                mi_des.Header = "Choose a project to delegate to";
                mi_des.Tag = String.Empty;
                mi_des.IsEnabled = false;
                this.RibbonButton_PowerDDelegate.Items.Add(mi_des);
                foreach (string s in powerDDelegateFolderCache)
                {
                    MenuItem mi_project = new MenuItem();
                    mi_project.Header = System.IO.Path.GetFileName(s);
                    mi_project.Tag = s;
                    mi_project.IsCheckable = true;
                    mi_project.Click += RibbonButton_PowerDDelegateMenuItem_Click;
                    this.RibbonButton_PowerDDelegate.Items.Add(mi_project);
                }
                foreach (Element element in elementControl.Root.HeadingElements)
                {
                    if ((element.NoteText != StartProcess.TODAY_PLUS) && (element.PowerDStatus!=PowerDStatus.Done))
                    {
                        MenuItem mi_project = new MenuItem();
                        mi_project.Header = element.NoteText;
                        mi_project.Tag = element.Path;
                        mi_project.IsCheckable = true;
                        mi_project.Click += RibbonButton_PowerDDelegateMenuItem_Click;
                        this.RibbonButton_PowerDDelegate.Items.Add(mi_project);
                    }
                }
                MenuItem mi_fsd = new MenuItem();
                mi_fsd.Header = "Select a project...";
                mi_fsd.Tag = "FolderSelectDialog";
                mi_fsd.IsCheckable = true;
                mi_fsd.Click += new RoutedEventHandler(RibbonButton_PowerDDelegateMenuItem_Click);
                this.RibbonButton_PowerDDelegate.Items.Add(mi_fsd);

                if (previousPowerDDelegateFolder != null)
                {
                    foreach (MenuItem mi in RibbonButton_PowerDDelegate.Items)
                    {
                        if (mi.Tag.ToString() == previousPowerDDelegateFolder)
                        {
                            mi.IsChecked = true;
                        }
                        else
                        {
                            mi.IsChecked = false;
                        }
                    }
                }
                else
                {
                    
                }

                EndLoadingUI();
            }
            catch (Exception ex)
            {
                this.RibbonButton_PowerDDelegate.Items.Clear();
                //MessageBox.Show("RibbonButton_PowerDDelegate_PreviewMouseDown\n" + ex.Message);
            }
        }

        private void RibbonButton_PowerDDelegateMenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                LogControl.Write(
                   elementControl.CurrentElement,
                   LogEventAccess.Ribbon,
                   LogEventType.PowerDDelegate,
                   LogEventStatus.Begin,
                   null);

                string des = String.Empty;

                MenuItem mi = sender as MenuItem;
                if (System.IO.Directory.Exists(mi.Tag.ToString()))
                {
                    previousPowerDDelegateFolder = mi.Tag.ToString();
                    des = mi.Tag.ToString();
                }
                else if (mi.Tag.ToString() == "FolderSelectDialog")
                {
                    System.Windows.Forms.FolderBrowserDialog dialog = new System.Windows.Forms.FolderBrowserDialog();
                    dialog.SelectedPath = StartProcess.START_PATH;
                    dialog.ShowNewFolderButton = true;
                    if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string folderPath = dialog.SelectedPath;
                        des = folderPath;
                        
                        bool dup = false;
                        foreach (MenuItem menuItem in RibbonButton_PowerDDelegate.Items)
                        {
                            if (menuItem.Tag.ToString() == folderPath)
                            {
                                dup = true;
                                break;
                            }
                        }
                        if (!dup)
                        {
                            previousPowerDDelegateFolder = folderPath;
                            if (!this.powerDDelegateFolderCache.Contains(folderPath))
                            {
                                this.powerDDelegateFolderCache.Enqueue(folderPath);
                            }
                        }
                    }
                    else
                    {
                        LogControl.Write(
                           elementControl.CurrentElement,
                           LogEventAccess.Ribbon,
                           LogEventType.PowerDDelegate,
                           LogEventStatus.Cancel,
                           null);

                        return;
                    }
                }

                if (des.EndsWith(System.IO.Path.DirectorySeparatorChar.ToString()) == false)
                {
                    des += System.IO.Path.DirectorySeparatorChar;
                }

                if (des == elementControl.CurrentElement.Path)
                {
                    MessageBox.Show("You can not delegate a folder to itself.");
                    LogControl.Write(
                       elementControl.CurrentElement,
                       LogEventAccess.Ribbon,
                       LogEventType.PowerDDelegate,
                       LogEventStatus.Cancel,
                       null);

                    return;
                }

                elementControl.PowerDDelegate(elementControl.CurrentElement, des);

                RibbonCommandPowerDNext_Executed(null, null);

                LogControl.Write(
                   elementControl.CurrentElement,
                   LogEventAccess.Ribbon,
                   LogEventType.PowerDDelegate,
                   LogEventStatus.End,
                   LogEventInfo.Location + LogControl.COMMA + des);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("RibbonButton_PowerDDelegate_Click\n" + ex.Message);

                LogControl.Write(
                   elementControl.CurrentElement,
                   LogEventAccess.Ribbon,
                   LogEventType.PowerDDelegate,
                   LogEventStatus.Error,
                   LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void RibbonCommandPowerDDelete_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            try
            {
                LogControl.Write(
                   elementControl.CurrentElement,
                   LogEventAccess.Ribbon,
                   LogEventType.PowerDDelete,
                   LogEventStatus.Begin,
                   null);

                Element elementBelow = elementControl.CurrentElement.ElementBelow;
                Microsoft.Windows.Controls.Ribbon.RibbonCommand rc = ((Microsoft.Windows.Controls.Ribbon.RibbonButton)this.RibbonButton_PowerDDelete).Command as Microsoft.Windows.Controls.Ribbon.RibbonCommand;
                if (rc.LabelTitle == PowerDDeleteType.Undo.ToString())
                {
                    elementControl.PowerDDelete(elementControl.CurrentElement, PowerDDeleteType.Undo);
                    elementControl.CurrentElement = elementBelow;
                    GetFocusToElementTextBox(elementControl.CurrentElement, -1, false, false);
                }
                else
                {
                    elementControl.PowerDDelete(elementControl.CurrentElement, PowerDDeleteType.Delete);
                    elementControl.CurrentElement = elementBelow.ElementAbove;
                    RibbonCommandPowerDNext_Executed(null, null);
                }

                LogControl.Write(
                   elementControl.CurrentElement,
                   LogEventAccess.Ribbon,
                   LogEventType.PowerDDelete,
                   LogEventStatus.End,
                   null);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("RibbonCommandPowerDDelete_Click\n" + ex.Message);

                LogControl.Write(
                   elementControl.CurrentElement,
                   LogEventAccess.Ribbon,
                   LogEventType.PowerDDelete,
                   LogEventStatus.Error,
                   LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void RibbonCommandPowerDPrevious_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            try
            {
                Element previous = elementControl.CurrentElement.ElementAbove;
                while (previous != null && previous.IsHeading && previous.IsExpanded)
                {
                    previous = previous.ElementAbove;
                }
                if (previous == null)
                {
                    GetFocusToElementTextBox(elementControl.CurrentElement, -1, false, false);
                }
                else
                {
                    GetFocusToElementTextBox(previous, -1, false, false);
                }

                LogControl.Write(
                   elementControl.CurrentElement,
                   LogEventAccess.Ribbon,
                   LogEventType.PowerDPrevious,
                   LogEventStatus.NULL,
                   null);
            }
            catch (Exception ex)
            {
                LogControl.Write(
                   elementControl.CurrentElement,
                   LogEventAccess.Ribbon,
                   LogEventType.PowerDPrevious,
                   LogEventStatus.Error,
                   LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void RibbonCommandPowerDNext_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            try
            {
                Element next = elementControl.CurrentElement.ElementBelow;
                while (next != null && next.IsHeading && next.IsExpanded)
                {
                    next = next.ElementBelow;
                }
                if (next == null)
                {
                    GetFocusToElementTextBox(elementControl.CurrentElement, -1, false, false);
                }
                else
                {
                    GetFocusToElementTextBox(next, -1, false, false);
                }

                LogControl.Write(
                   elementControl.CurrentElement,
                   LogEventAccess.Ribbon,
                   LogEventType.PowerDNext,
                   LogEventStatus.NULL,
                   null);
            }
            catch (Exception ex)
            {
                LogControl.Write(
                   elementControl.CurrentElement,
                   LogEventAccess.Ribbon,
                   LogEventType.PowerDNext,
                   LogEventStatus.Error,
                   LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void RibbonPowerDOption_ShowAssociationMarkedDone_Checked(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!SHOW_DONEDEFER_HANDLE)
                {
                    return;
                }

                Element element = elementControl.CurrentElement.ParentElement;
                element.ShowAssociationMarkedDone = true;
                elementControl.UpdateElement(element);

                ShowOrHideMarkedAssociation(element, PowerDStatus.Done, true);

                LogControl.Write(
                   elementControl.CurrentElement,
                   LogEventAccess.Ribbon,
                   LogEventType.PowerDCheckShowAssociationMarkedDone,
                   LogEventStatus.NULL,
                   null);
            }
            catch (Exception ex)
            {
                LogControl.Write(
                   elementControl.CurrentElement,
                   LogEventAccess.Ribbon,
                   LogEventType.PowerDCheckShowAssociationMarkedDone,
                   LogEventStatus.Error,
                   LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void RibbonPowerDOption_ShowAssociationMarkedDone_Unchecked(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!SHOW_DONEDEFER_HANDLE)
                {
                    return;
                }

                Element element = elementControl.CurrentElement.ParentElement;
                element.ShowAssociationMarkedDone = false;
                elementControl.UpdateElement(element);

                ShowOrHideMarkedAssociation(element, PowerDStatus.Done, false);
                //elementControl.SyncHeadingElement(element);

                LogControl.Write(
                   elementControl.CurrentElement,
                   LogEventAccess.Ribbon,
                   LogEventType.PowerDUncheckShowAssociationMarkedDone,
                   LogEventStatus.NULL,
                   null);
            }
            catch (Exception ex)
            {
                LogControl.Write(
                   elementControl.CurrentElement,
                   LogEventAccess.Ribbon,
                   LogEventType.PowerDUncheckShowAssociationMarkedDone,
                   LogEventStatus.Error,
                   LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void RibbonPowerDOption_ShowAssociationMarkedDefer_Checked(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!SHOW_DONEDEFER_HANDLE)
                {
                    return;
                }

                Element element = elementControl.CurrentElement.ParentElement;
                element.ShowAssociationMarkedDefer = true;
                elementControl.UpdateElement(element);

                ShowOrHideMarkedAssociation(element, PowerDStatus.Deferred, true);

                LogControl.Write(
                   elementControl.CurrentElement,
                   LogEventAccess.Ribbon,
                   LogEventType.PowerDCheckShowAssociationMarkedDefer,
                   LogEventStatus.NULL,
                   null);
            }
            catch (Exception ex)
            {
                LogControl.Write(
                   elementControl.CurrentElement,
                   LogEventAccess.Ribbon,
                   LogEventType.PowerDCheckShowAssociationMarkedDefer,
                   LogEventStatus.Error,
                   LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void RibbonPowerDOption_ShowAssociationMarkedDefer_Unchecked(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!SHOW_DONEDEFER_HANDLE)
                {
                    return;
                }

                Element element = elementControl.CurrentElement.ParentElement;
                element.ShowAssociationMarkedDefer = false;
                elementControl.UpdateElement(element);

                ShowOrHideMarkedAssociation(element, PowerDStatus.Deferred, false);
                //elementControl.SyncHeadingElement(element);

                LogControl.Write(
                   elementControl.CurrentElement,
                   LogEventAccess.Ribbon,
                   LogEventType.PowerDUncheckShowAssociationMarkedDefer,
                   LogEventStatus.NULL,
                   null);
            }
            catch (Exception ex)
            {
                LogControl.Write(
                   elementControl.CurrentElement,
                   LogEventAccess.Ribbon,
                   LogEventType.PowerDUncheckShowAssociationMarkedDefer,
                   LogEventStatus.Error,
                   LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void ShowOrHideMarkedAssociation(Element element, PowerDStatus status, bool show)
        {
            if (show)
            {
                elementControl.ShowOrHidePowerDElement(element, status, show);
                elementControl.SyncHeadingElement(element);
            }
            else
            {
                elementControl.ShowOrHidePowerDElement(element, status, show);
                //elementControl.SyncHeadingElement(element);
            }

        }

        private void RibbonButton_NewWindow_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //Change the window focus and reload the window
                string newWindowStartPath = string.Empty;

                if (previousChangeFocusFolder == null)
                {
                    if (startPath != StartProcess.START_PATH)
                        newWindowStartPath = StartProcess.START_PATH;

                    if (startPath != StartProcess.JOURNAL_PATH)
                        newWindowStartPath = StartProcess.JOURNAL_PATH;
                }
                else
                    newWindowStartPath = previousChangeFocusFolder;

                previousChangeFocusFolder = this.startPath;

                MainWindow newWindow = new MainWindow(newWindowStartPath);
                newWindow.Show();

                RibbonCommandNewWindow_Executed(null, null);

                LogControl.Write(
                   elementControl.CurrentElement,
                   LogEventAccess.Ribbon,
                   LogEventType.NewWindow,
                   LogEventStatus.NULL,
                   LogEventInfo.Location + LogControl.COMMA + newWindowStartPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("RibbonButton_NewWindow_Click\n" + ex.Message);

                LogControl.Write(
                   elementControl.CurrentElement,
                   LogEventAccess.Ribbon,
                   LogEventType.NewWindow,
                   LogEventStatus.Error,
                   LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void RibbonButton_NewWindow_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            try
            {
                StartLoadingUI();

                while (this.RibbonButton_NewWindow.Items.Count > 0)
                {
                    this.RibbonButton_NewWindow.Items.RemoveAt(0);
                }
                MenuItem mi_des = new MenuItem();
                mi_des.Header = "Choose a project to open new window into";
                mi_des.Tag = String.Empty;
                mi_des.IsEnabled = false;
                this.RibbonButton_NewWindow.Items.Add(mi_des);

                char[] ds = { '\\' };

                MenuItem mi_project = new MenuItem();
                mi_project.Header = System.IO.Path.GetFileName(StartProcess.START_PATH.TrimEnd(ds));
                mi_project.Tag = StartProcess.START_PATH;
                mi_project.IsCheckable = false;
                mi_project.Click += RibbonButton_NewWindowMenuItem_Click;
                this.RibbonButton_NewWindow.Items.Add(mi_project);

                mi_project = new MenuItem();
                mi_project.Header = System.IO.Path.GetFileName(StartProcess.JOURNAL_PATH.TrimEnd(ds));
                mi_project.Tag = StartProcess.JOURNAL_PATH;
                mi_project.IsCheckable = false;
                mi_project.Click += RibbonButton_NewWindowMenuItem_Click;
                this.RibbonButton_NewWindow.Items.Add(mi_project);

                mi_project = new MenuItem();
                mi_project.Header = System.IO.Path.GetFileName(StartProcess.TODAY_PLUS_PATH.TrimEnd(ds));
                mi_project.Tag = StartProcess.TODAY_PLUS_PATH;
                mi_project.IsCheckable = false;
                mi_project.Click += RibbonButton_NewWindowMenuItem_Click;
                this.RibbonButton_NewWindow.Items.Add(mi_project);

                mi_project = new MenuItem();
                mi_project.Header = System.IO.Path.GetFileName(StartProcess.DAYS_AHEAD_PATH.TrimEnd(ds));
                mi_project.Tag = StartProcess.DAYS_AHEAD_PATH;
                mi_project.IsCheckable = false;
                mi_project.Click += RibbonButton_NewWindowMenuItem_Click;
                this.RibbonButton_NewWindow.Items.Add(mi_project);

                mi_project = new MenuItem();
                mi_project.Header = System.IO.Path.GetFileName(StartProcess.GOALS_THINGS_PATH.TrimEnd(ds));
                mi_project.Tag = StartProcess.GOALS_THINGS_PATH;
                mi_project.IsCheckable = false;
                mi_project.Click += RibbonButton_NewWindowMenuItem_Click;
                this.RibbonButton_NewWindow.Items.Add(mi_project);

                mi_project = new MenuItem();
                mi_project.Header = System.IO.Path.GetFileName(StartProcess.ROLES_PEOPLE_PATH.TrimEnd(ds));
                mi_project.Tag = StartProcess.ROLES_PEOPLE_PATH;
                mi_project.IsCheckable = false;
                mi_project.Click += RibbonButton_NewWindowMenuItem_Click;
                this.RibbonButton_NewWindow.Items.Add(mi_project);


                int count = this.RibbonButton_NewWindow.Items.Count;
                for (int i = 0; i < count; i++)
                {
                    MenuItem item = (MenuItem)this.RibbonButton_NewWindow.Items[i];
                    if (item.Tag.ToString() == startPath)
                        ((MenuItem)this.RibbonButton_NewWindow.Items[i]).IsEnabled = false;
                }

                MenuItem mi_fsd = new MenuItem();
                mi_fsd.Header = "Select a Focus...";
                mi_fsd.Tag = "FolderSelectDialog";
                mi_fsd.IsCheckable = true;
                mi_fsd.Click += new RoutedEventHandler(RibbonButton_NewWindowMenuItem_Click);
                this.RibbonButton_NewWindow.Items.Add(mi_fsd);

                EndLoadingUI();
            }
            catch (Exception ex)
            {
                this.RibbonButton_NewWindow.Items.Clear();

                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("RibbonButton_NewWindow_PreviewMouseDown\n" + ex.Message);
            }
        }

        private void RibbonButton_NewWindowMenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string des = String.Empty;

                MenuItem mi = sender as MenuItem;
                if (System.IO.Directory.Exists(mi.Tag.ToString()))
                {
                    previousNewWindowFolder = mi.Tag.ToString();
                    des = mi.Tag.ToString();
                }
                else if (mi.Tag.ToString() == "FolderSelectDialog")
                {
                    System.Windows.Forms.FolderBrowserDialog dialog = new System.Windows.Forms.FolderBrowserDialog();
                    dialog.SelectedPath = StartProcess.LIFE_PATH;
                    dialog.ShowNewFolderButton = true;
                    if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string folderPath = dialog.SelectedPath;
                        des = folderPath;

                        bool dup = false;
                        foreach (MenuItem menuItem in RibbonButton_NewWindow.Items)
                        {
                            if (menuItem.Tag.ToString() == folderPath)
                            {
                                dup = true;
                                break;
                            }
                        }
                        if (!dup)
                        {
                            previousNewWindowFolder = folderPath;
                        }
                    }
                    else
                    {
                        return;
                    }
                }

                if (des.EndsWith(System.IO.Path.DirectorySeparatorChar.ToString()) == false)
                {
                    des += System.IO.Path.DirectorySeparatorChar;
                }

                string newWindowStartPath = des;

                MainWindow newWindow = new MainWindow(newWindowStartPath);

                newWindow.Show();


                RibbonCommandNewWindow_Executed(null, null);

                LogControl.Write(
                   elementControl.CurrentElement,
                   LogEventAccess.Ribbon,
                   LogEventType.NewWindow,
                   LogEventStatus.NULL,
                   LogEventInfo.Location + LogControl.COMMA + des);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("RibbonButton_NewWindowMenuItem_Click\n" + ex.Message);

                LogControl.Write(
                   elementControl.CurrentElement,
                   LogEventAccess.Ribbon,
                   LogEventType.NewWindow,
                   LogEventStatus.Error,
                   LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void RibbonButton_ChangeFocus_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //Change the window focus and reload the window
                string newFocusPath = string.Empty;

                if (previousChangeFocusFolder == null)
                    newFocusPath = StartProcess.JOURNAL_PATH;
                else
                    newFocusPath = previousChangeFocusFolder;

                previousChangeFocusFolder = this.startPath;

                //Change the foucs
                StartLoadingUI();

                if (elementControl.CurrentElement != null)
                {
                    elementControl.UpdateElement(elementControl.CurrentElement);
                    elementControl.UpdateStatusChangedElementList();

                    Properties.Settings.Default.LastFocusedElementGuid = elementControl.CurrentElement.ID.ToString();
                    Properties.Settings.Default.Save();
                }

                this.startPath = newFocusPath;
                this.PlanzMainWindow.Title = "Planz - " + this.startPath;

                elementControl.UpdateRootPath(this.startPath);

                // Load Plan
                this.Plan.ItemsSource = elementControl.Root.Elements;

                // Load Navigation Tab
                LoadNavigationTab();
                this.Show();

                RibbonCommandChangeFocus_Executed(null, null);

                LogControl.Write(
                   elementControl.CurrentElement,
                   LogEventAccess.Ribbon,
                   LogEventType.ChangeFocus,
                   LogEventStatus.NULL,
                   LogEventInfo.Location + LogControl.COMMA + newFocusPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("RibbonButton_ChangeFocus_Click\n" + ex.Message);

                LogControl.Write(
                   elementControl.CurrentElement,
                   LogEventAccess.Ribbon,
                   LogEventType.ChangeFocus,
                   LogEventStatus.Error,
                   LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void RibbonButton_ChangeFocus_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            try
            {
                StartLoadingUI();

                while (this.RibbonButton_ChangeFocus.Items.Count > 0)
                {
                    this.RibbonButton_ChangeFocus.Items.RemoveAt(0);
                }

                MenuItem mi_des = new MenuItem();
                mi_des.Header = "Choose a project to change the focus into";
                mi_des.Tag = String.Empty;
                mi_des.IsEnabled = false;
                this.RibbonButton_ChangeFocus.Items.Add(mi_des);
                char[] ds = { '\\' };

                MenuItem mi_project;

                mi_project = new MenuItem();
                mi_project.Header = System.IO.Path.GetFileName(StartProcess.START_PATH.TrimEnd(ds));
                mi_project.Tag = StartProcess.START_PATH;
                mi_project.IsCheckable = false;
                mi_project.Click += RibbonButton_ChangeFocusMenuItem_Click;
                this.RibbonButton_ChangeFocus.Items.Add(mi_project);

                mi_project = new MenuItem();
                mi_project.Header = System.IO.Path.GetFileName(StartProcess.JOURNAL_PATH.TrimEnd(ds));
                mi_project.Tag = StartProcess.JOURNAL_PATH;
                mi_project.IsCheckable = false;
                mi_project.Click += RibbonButton_ChangeFocusMenuItem_Click;
                this.RibbonButton_ChangeFocus.Items.Add(mi_project);

                mi_project = new MenuItem();
                mi_project.Header = System.IO.Path.GetFileName(StartProcess.TODAY_PLUS_PATH.TrimEnd(ds));
                mi_project.Tag = StartProcess.TODAY_PLUS_PATH;
                mi_project.IsCheckable = false;
                mi_project.Click += RibbonButton_ChangeFocusMenuItem_Click;
                this.RibbonButton_ChangeFocus.Items.Add(mi_project);

                mi_project = new MenuItem();
                mi_project.Header = System.IO.Path.GetFileName(StartProcess.DAYS_AHEAD_PATH.TrimEnd(ds));
                mi_project.Tag = StartProcess.DAYS_AHEAD_PATH;
                mi_project.IsCheckable = false;
                mi_project.Click += RibbonButton_ChangeFocusMenuItem_Click;
                this.RibbonButton_ChangeFocus.Items.Add(mi_project);


                mi_project = new MenuItem();
                mi_project.Header = System.IO.Path.GetFileName(StartProcess.GOALS_THINGS_PATH.TrimEnd(ds));
                mi_project.Tag = StartProcess.GOALS_THINGS_PATH;
                mi_project.IsCheckable = false;
                mi_project.Click += RibbonButton_ChangeFocusMenuItem_Click;
                this.RibbonButton_ChangeFocus.Items.Add(mi_project);

                mi_project = new MenuItem();
                mi_project.Header = System.IO.Path.GetFileName(StartProcess.ROLES_PEOPLE_PATH.TrimEnd(ds));
                mi_project.Tag = StartProcess.ROLES_PEOPLE_PATH;
                mi_project.IsCheckable = false;
                mi_project.Click += RibbonButton_ChangeFocusMenuItem_Click;
                this.RibbonButton_ChangeFocus.Items.Add(mi_project);


                int count = this.RibbonButton_ChangeFocus.Items.Count;
                for (int i = 0; i < count; i++)
                {
                    MenuItem item = (MenuItem)this.RibbonButton_ChangeFocus.Items[i];
                    if (item.Tag.ToString() == startPath)
                        ((MenuItem)this.RibbonButton_ChangeFocus.Items[i]).IsEnabled = false;
                }

                MenuItem mi_fsd = new MenuItem();
                mi_fsd.Header = "Select a Focus...";
                mi_fsd.Tag = "FolderSelectDialog";
                mi_fsd.IsCheckable = true;
                mi_fsd.Click += new RoutedEventHandler(RibbonButton_ChangeFocusMenuItem_Click);
                this.RibbonButton_ChangeFocus.Items.Add(mi_fsd);

                EndLoadingUI();
            }
            catch (Exception ex)
            {
                this.RibbonButton_ChangeFocus.Items.Clear();

                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("RibbonButton_ChangeFocus_PreviewMouseDown\n" + ex.Message);
            }
        }

        private void RibbonButton_ChangeFocusMenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {

                string des = String.Empty;

                MenuItem mi = sender as MenuItem;
                if (System.IO.Directory.Exists(mi.Tag.ToString()))
                {
                    des = mi.Tag.ToString();
                }
                else if (mi.Tag.ToString() == "FolderSelectDialog")
                {
                    System.Windows.Forms.FolderBrowserDialog dialog = new System.Windows.Forms.FolderBrowserDialog();
                    dialog.SelectedPath = StartProcess.LIFE_PATH;
                    dialog.ShowNewFolderButton = true;
                    if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string folderPath = dialog.SelectedPath;
                        des = folderPath;
                    }
                    else
                    {
                        return;
                    }
                }

                if (des.EndsWith(System.IO.Path.DirectorySeparatorChar.ToString()) == false)
                {
                    des += System.IO.Path.DirectorySeparatorChar;
                }

                // if ont include current selection
                previousChangeFocusFolder = startPath;

                //Change the window focus and reload the window
                this.startPath = des;
                this.PlanzMainWindow.Title = "Planz - " + this.startPath;

                // Setup
                elementControl.UpdateRootPath(startPath);
                // Load Plan
                this.Plan.ItemsSource = elementControl.Root.Elements;
                // Load Navigation Tab
                LoadNavigationTab();
                this.Show();

                RibbonCommandChangeFocus_Executed(null, null);

                string evenInfo = "Location:"+des;
                LogControl.Write(
                   elementControl.CurrentElement,
                   LogEventAccess.Ribbon,
                   LogEventType.ChangeFocus,
                   LogEventStatus.NULL,
                   evenInfo);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("RibbonButton_ChangeFocusMenuItem_Click\n" + ex.Message);

                LogControl.Write(
                   elementControl.CurrentElement,
                   LogEventAccess.Ribbon,
                   LogEventType.ChangeFocus,
                   LogEventStatus.Error,
                   LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        #endregion

        #endregion

        #region Element TextBox

        private void ElementTextBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            Element currentElement = GetElement(sender);
            TextBox tb = sender as TextBox;
            currentElement.NoteText = tb.Text;

            switch (e.Key)
            {
                case Key.Up:
                    if (Keyboard.Modifiers == ModifierKeys.Control)
                    {
                        CtrlUpKeyPressed(sender, e);
                    }
                    else
                    {
                        UpKeyPressed(currentElement, e);
                    }
                    e.Handled = true;
                    break;
                case Key.Down:
                    if (Keyboard.Modifiers == ModifierKeys.Control)
                    {
                        CtrlDownKeyPressed(sender, e);
                    }
                    else
                    {
                        DownKeyPressed(currentElement, e);
                    }
                    e.Handled = true;
                    break;
                case Key.Left:
                    if (Keyboard.Modifiers == ModifierKeys.Control)
                    {
                        ShiftTabKeyPressed(sender, e);
                        e.Handled = true;
                    }
                    break;
                case Key.Right:
                    if (Keyboard.Modifiers == ModifierKeys.Control)
                    {
                        TabKeyPressed(sender, e);
                        e.Handled = true;
                    }
                    break;
                case Key.Back:
                    BackKeyPressed(currentElement, e);
                    break;
                case Key.Delete:
                    if (currentElement == elementControl.SelectedElements.FirstOrDefault() || elementControl.SelectedElements.Count > 1)
                    {
                        RibbonCommandDelete_Executed(sender, null);
                        e.Handled = true;
                    }
                    else
                    {
                        e.Handled = false;
                    }
                    break;
                case Key.F5:
                    RibbonCommandRefresh_Executed(sender, null);
                    break;
            };
        }

        private void ElementTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            Element currentElement = GetElement(sender);

            switch (e.Key)
            {
                case Key.Enter:
                    //if (Keyboard.Modifiers == ModifierKeys.Shift)
                    //{
                    //    ShiftEnterKeyPressed(sender, e); 
                    //}
                    //else
                    //{
                        EnterKeyPressed(currentElement, e); 
                    //}
                    e.Handled = true;
                    break;
                case Key.Tab:
                    if (Keyboard.Modifiers == ModifierKeys.Shift)
                    {
                        ShiftTabKeyPressed(sender, e);
                    }
                    else
                    {
                        TabKeyPressed(sender, e);
                    }
                    e.Handled = true;
                    break;
                case Key.F3:
                    Element target;
                    if ((target = elementControl.FindNext()) != null)
                    {
                        GetFocusToElementTextBox(target, -1, false, true);
                    }
                    e.Handled = true;
                    break;
                case Key.F:
                    if (Keyboard.Modifiers == ModifierKeys.Control)
                    {
                        this.RibbonTabHome.IsSelected = true;
                        this.RibbonSearchBox.UpdateLayout();
                        this.RibbonSearchBox.Focus();

                        e.Handled = true;
                    }
                    break;
            }
        }

        private void EnterKeyPressed(object sender, KeyEventArgs e)
        {
            try
            {
                Element currentElement = sender as Element;

                Element newElement = null;
                int index = 0;
                int selectionStart = 0;

                switch (currentElement.Type)
                {
                    case ElementType.Heading:

                        try
                        {
                            selectionStart = GetTextBoxSelectionStart(currentElement);
                            if (selectionStart == 0)
                            {
                                if (currentElement.NoteText == String.Empty)
                                {
                                    index = currentElement.ParentElement.Elements.IndexOf(currentElement) + 1;
                                }
                                else
                                {
                                    index = currentElement.ParentElement.Elements.IndexOf(currentElement);
                                }
                                newElement = elementControl.CreateNewElement(ElementType.Heading, String.Empty);
                                elementControl.InsertElement(newElement, currentElement.ParentElement, index);

                                LogControl.Write(
                                       newElement,
                                       LogEventAccess.Hotkey,
                                       LogEventType.CreateNewHeading,
                                       LogEventStatus.NULL,
                                       null);
                            }
                            else
                            {
                                if (currentElement.IsExpanded)
                                {
                                    newElement = elementControl.CreateNewElement(ElementType.Note, String.Empty);
                                    elementControl.InsertElement(newElement, currentElement, 0);

                                    LogControl.Write(
                                       newElement,
                                       LogEventAccess.Hotkey,
                                       LogEventType.CreateNewNote,
                                       LogEventStatus.NULL,
                                       null);
                                }
                                else
                                {
                                    index = currentElement.ParentElement.Elements.IndexOf(currentElement) + 1;
                                    newElement = elementControl.CreateNewElement(ElementType.Heading, String.Empty);
                                    elementControl.InsertElement(newElement, currentElement.ParentElement, index);

                                    LogControl.Write(
                                       newElement,
                                       LogEventAccess.Hotkey,
                                       LogEventType.CreateNewHeading,
                                       LogEventStatus.NULL,
                                       null);
                                }
                            }
                            GetFocusToElementTextBox(newElement, 0, false, false);
                        }
                        catch (Exception)
                        {
                            LogControl.Write(
                               elementControl.CurrentElement,
                               LogEventAccess.Hotkey,
                               LogEventType.CreateNewHeading,
                               LogEventStatus.Error,
                               null);
                        }

                        break;

                    case ElementType.Note:
                        try
                        {
                            Element focusElement = null;
                            selectionStart = GetTextBoxSelectionStart(currentElement);
                            if (currentElement.IsCommandNote)
                                selectionStart = currentElement.NoteText.Length;

                            if (selectionStart == 0 && (currentElement.NoteText != String.Empty || currentElement.HasAssociation))
                            {
                                index = currentElement.ParentElement.Elements.IndexOf(currentElement);
                                newElement = elementControl.CreateNewElement(ElementType.Note, String.Empty);
                                focusElement = newElement;
                            }
                            else
                            {
                                if (currentElement.HasAssociation && selectionStart != currentElement.NoteText.Length)
                                {
                                    index = currentElement.ParentElement.Elements.IndexOf(currentElement);
                                    newElement = elementControl.CreateNewElement(ElementType.Note, String.Empty);
                                    newElement.FontColor = currentElement.FontColor;
                                    newElement.NoteText = currentElement.NoteText.Substring(0, selectionStart);
                                    currentElement.NoteText = currentElement.NoteText.Substring(selectionStart);
                                    focusElement = currentElement;
                                }
                                else
                                {
                                    index = currentElement.ParentElement.Elements.IndexOf(currentElement) + 1;
                                    newElement = elementControl.CreateNewElement(ElementType.Note, String.Empty);
                                    if (selectionStart != currentElement.NoteText.Length)
                                    {
                                        newElement.FontColor = currentElement.FontColor;
                                    }
                                    newElement.NoteText = currentElement.NoteText.Substring(selectionStart);
                                    currentElement.NoteText = currentElement.NoteText.Substring(0, selectionStart);
                                    focusElement = newElement;
                                }
                            }
                            elementControl.InsertElement(newElement, currentElement.ParentElement, index);
                            GetFocusToElementTextBox(focusElement, 0, false, false);

                            LogControl.Write(
                                   newElement,
                                   LogEventAccess.Hotkey,
                                   LogEventType.CreateNewNote,
                                   LogEventStatus.NULL,
                                   null);
                        }
                        catch (Exception)
                        {
                            LogControl.Write(
                               elementControl.CurrentElement,
                               LogEventAccess.Hotkey,
                               LogEventType.CreateNewNote,
                               LogEventStatus.Error,
                               null);
                        }

                        break;
                };
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("EnterKeyPressed\n" + ex.Message);
            }
        }

        private void ShiftEnterKeyPressed(object sender, KeyEventArgs e)
        {
            try
            {
                Element element = GetElement(sender);
                if (element != null)
                {
                    if (element.IsNote)
                    {
                        TextBox tb = sender as TextBox;
                        string upper = tb.Text.Substring(0, tb.SelectionStart);
                        string lower = tb.Text.Substring(tb.SelectionStart);
                        tb.Text = upper + "\r\n" + lower;
                        tb.SelectionStart = tb.Text.Length;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("ShiftEnterKeyPressed\n" + ex.Message);
            }
        }

        private void BackKeyPressed(object sender, KeyEventArgs e)
        {
            try
            {
                Element currentElement = sender as Element;

                int selectionStart = 0;
                string selectedText = String.Empty;

                Element elementAbove = currentElement.ElementAbove;
                int length = elementAbove.NoteText.Length;
                DeleteMessageType dmt = DeleteMessageType.Default;


                switch (currentElement.Type)
                {
                    case ElementType.Heading:
                        //if (elementControl.HasChildOrContent(currentElement) == false)
                        //{
                            selectionStart = GetTextBoxSelectionStart(currentElement);
                            if (selectionStart != 0)
                            {
                                return;
                            }
                            selectedText = GetTextBoxSelectedText(currentElement);
                            if (selectedText != String.Empty)
                            {
                                return;
                            }
                            
                            e.Handled = true;
                            dmt = DeleteMessageType.HeadingWithoutChildrenOrContent;

                            if (elementAbove.NoteText == String.Empty &&
                                elementControl.HasChildOrContent(elementAbove) == false)
                            {
                                

                                LogControl.Write(
                                    elementAbove,
                                    LogEventAccess.Hotkey,
                                    LogEventType.Delete,
                                    LogEventStatus.Begin,
                                    LogEventInfo.DeleteType + LogControl.COMMA + dmt);
                                
                                elementControl.DeleteElement(elementAbove);
                                
                                GetFocusToElementTextBox(currentElement, length, false, false);
                                LogControl.Write(
                                    currentElement,
                                    LogEventAccess.Hotkey,
                                    LogEventType.Delete,
                                    LogEventStatus.End,
                                    null);
                            }
                            else if (currentElement.NoteText == String.Empty &&
                                elementControl.HasChildOrContent(currentElement) == false)
                            {
                                LogControl.Write(
                                    currentElement,
                                    LogEventAccess.Hotkey,
                                    LogEventType.Delete,
                                    LogEventStatus.Begin,
                                    LogEventInfo.DeleteType + LogControl.COMMA + dmt);

                                elementControl.DeleteElement(currentElement);

                                GetFocusToElementTextBox(elementAbove, length, false, false);
                                LogControl.Write(
                                    elementAbove,
                                    LogEventAccess.Hotkey,
                                    LogEventType.Delete,
                                    LogEventStatus.End,
                                    null);
                            }
                            else
                            {
                            }
                            
                        //}
                        //else
                        //{
                        //    return;
                        //}
                        break;
                    case ElementType.Note:
                        selectionStart = GetTextBoxSelectionStart(currentElement);
                        if (selectionStart != 0)
                        {
                            return;
                        }
                        selectedText = GetTextBoxSelectedText(currentElement);
                        if (selectedText != String.Empty)
                        {
                            return;
                        }
                        if (currentElement == elementControl.Root.FirstChild && currentElement.NoteText != String.Empty) 
                        {
                            return;
                        }
                        if (currentElement.HasChildren)
                        {
                            return;
                        }
                        if (currentElement == currentElement.ParentElement.FirstChild && currentElement.HasAssociation)
                        {
                            return;
                        }

                        dmt = DeleteMessageType.NoteWithoutAssociation;

                        if (!currentElement.HasAssociation)
                        {
                            LogControl.Write(
                                currentElement,
                                LogEventAccess.Hotkey,
                                LogEventType.Delete,
                                LogEventStatus.Begin,
                                LogEventInfo.DeleteType + LogControl.COMMA + dmt);

                            elementAbove.NoteText += currentElement.NoteText;
                            elementControl.RemoveElement(currentElement, currentElement.ParentElement);

                            e.Handled = true;

                            GetFocusToElementTextBox(elementAbove, length, false, false);

                            LogControl.Write(
                                elementAbove,
                                LogEventAccess.Hotkey,
                                LogEventType.Delete,
                                LogEventStatus.End,
                                null);
                        }
                        else if (!elementAbove.HasAssociation)
                        {
                            LogControl.Write(
                                elementAbove,
                                LogEventAccess.Hotkey,
                                LogEventType.Delete,
                                LogEventStatus.Begin,
                                LogEventInfo.DeleteType + LogControl.COMMA + dmt);

                            currentElement.NoteText = elementAbove.NoteText + currentElement.NoteText;
                            elementControl.RemoveElement(elementAbove, currentElement.ParentElement);

                            e.Handled = true;

                            GetFocusToElementTextBox(currentElement, length, false, false);

                            LogControl.Write(
                                currentElement,
                                LogEventAccess.Hotkey,
                                LogEventType.Delete,
                                LogEventStatus.End,
                                null);
                        }
                        break;
                };
                
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("BackKeyPressed\n" + ex.Message);

                LogControl.Write(
                   elementControl.CurrentElement,
                   LogEventAccess.Hotkey,
                   LogEventType.Delete,
                   LogEventStatus.Error,
                   LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);

            }
        }

        private void UpKeyPressed(object sender, KeyEventArgs e)
        {
            try
            {
                Element currentElement = sender as Element;

                Element elementAbove = currentElement.ElementAbove;

                if (elementAbove != null)
                {
                    GetFocusToElementTextBox(elementAbove, elementAbove.NoteText.Length, false, false);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("UpKeyPressed\n" + ex.Message);
            }
        }

        private void DownKeyPressed(object sender, KeyEventArgs e)
        {
            try
            {
                Element currentElement = sender as Element;

                Element elementBelow = currentElement.ElementBelow;

                if (elementBelow != null)
                {
                    GetFocusToElementTextBox(elementBelow, 0, false, false);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("DownKeyPressed\n" + ex.Message);
            }
        }

        private void TabKeyPressed(object sender, KeyEventArgs e)
        {
            LogEventAccess lea = LogEventAccess.Ribbon;
            if (sender is Planz.MainWindow)
            {
                lea = LogEventAccess.Ribbon;
            }
            else if (sender is System.Windows.Controls.MenuItem)
            {
                lea = LogEventAccess.NodeContextMenu;
            }
            else if (sender is TextBox)
            {
                lea = LogEventAccess.Hotkey;
            }

            Element currentElement = elementControl.CurrentElement;

            try
            {
                LogControl.Write(
                    currentElement,
                    lea,
                    LogEventType.Demote,
                    LogEventStatus.Begin,
                    null);

                elementControl.Demote(currentElement);
                ChangeElementTextFontFamily(currentElement);
                ChangeElementTextFontSize(currentElement);
                GetFocusToElementTextBox(currentElement, currentElement.NoteText.Length, false, false);

                LogControl.Write(
                    currentElement,
                    lea,
                    LogEventType.Demote,
                    LogEventStatus.End,
                    null);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("TabKeyPressed\n" + ex.Message);

                LogControl.Write(
                    currentElement,
                    lea,
                    LogEventType.Demote,
                    LogEventStatus.Error,
                    LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void ShiftTabKeyPressed(object sender, KeyEventArgs e)
        {
            LogEventAccess lea = LogEventAccess.Ribbon;
            if (sender is Planz.MainWindow)
            {
                lea = LogEventAccess.Ribbon;
            }
            else if (sender is System.Windows.Controls.MenuItem)
            {
                lea = LogEventAccess.NodeContextMenu;
            }
            else if (sender is TextBox)
            {
                lea = LogEventAccess.Hotkey;
            }

            Element currentElement = elementControl.CurrentElement;

            try
            {
                LogControl.Write(
                    currentElement,
                    lea,
                    LogEventType.Promote,
                    LogEventStatus.Begin,
                    null);

                elementControl.Promote(currentElement);
                ChangeElementTextFontFamily(currentElement);
                ChangeElementTextFontSize(currentElement);
                GetFocusToElementTextBox(currentElement, currentElement.NoteText.Length, false, false);

                LogControl.Write(
                    currentElement,
                    lea,
                    LogEventType.Promote,
                    LogEventStatus.End,
                    null);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("ShiftTabKeyPressed\n" + ex.Message);

                LogControl.Write(
                    currentElement,
                    lea,
                    LogEventType.Promote,
                    LogEventStatus.End,
                    LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void CtrlUpKeyPressed(object sender, KeyEventArgs e)
        {
            LogEventAccess lea = LogEventAccess.Ribbon;
            if (sender is Planz.MainWindow)
            {
                lea = LogEventAccess.Ribbon;
            }
            else if (sender is System.Windows.Controls.MenuItem)
            {
                lea = LogEventAccess.NodeContextMenu;
            }
            else if (sender is TextBox)
            {
                lea = LogEventAccess.Hotkey;
            }

            Element currentElement = elementControl.CurrentElement;

            try
            {
                LogControl.Write(
                    currentElement,
                    lea,
                    LogEventType.MoveUp,
                    LogEventStatus.Begin,
                    null);

                if (currentElement != null)
                {
                    elementControl.MoveUp(currentElement);
                    GetFocusToElementTextBox(currentElement, -1, false, false);
                }

                if (e != null)
                {
                    e.Handled = true;
                }

                LogControl.Write(
                    currentElement,
                    lea,
                    LogEventType.MoveUp,
                    LogEventStatus.End,
                    null);
            }
            catch (Exception ex)
            {
                const string duplexFileName = "Cannot create a file when that file already exists.\r\n";
                if (ex.Message == duplexFileName)
                {
                    MessageBox.Show("There is already a file/folder with the same name in this location.");
                }
                else
                {
                    MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                }
                //MessageBox.Show("CtrlUpKeyPressed\n" + ex.Message);

                LogControl.Write(
                    currentElement,
                    lea,
                    LogEventType.MoveUp,
                    LogEventStatus.Error,
                    LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void CtrlDownKeyPressed(object sender, KeyEventArgs e)
        {
            LogEventAccess lea = LogEventAccess.Ribbon;
            if (sender is Planz.MainWindow)
            {
                lea = LogEventAccess.Ribbon;
            }
            else if (sender is System.Windows.Controls.MenuItem)
            {
                lea = LogEventAccess.NodeContextMenu;
            }
            else if (sender is TextBox)
            {
                lea = LogEventAccess.Hotkey;
            }

            Element currentElement = elementControl.CurrentElement;

            try
            {
                LogControl.Write(
                    currentElement,
                    lea,
                    LogEventType.MoveDown,
                    LogEventStatus.Begin,
                    null);

                if (currentElement != null)
                {
                    elementControl.MoveDown(currentElement);
                    GetFocusToElementTextBox(currentElement, -1, false, false);
                }
                
                if (e != null)
                {
                    e.Handled = true;
                }

                LogControl.Write(
                    currentElement,
                    lea,
                    LogEventType.MoveDown,
                    LogEventStatus.End,
                    null);
            }
            catch (Exception ex)
            {
                const string duplexFileName = "Cannot create a file when that file already exists.\r\n";
                if (ex.Message == duplexFileName)
                {
                    MessageBox.Show("There is already a file/folder with the same name in this location.");
                }
                else
                {
                    MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                }
                //MessageBox.Show("CtrlDownKeyPressed\n" + ex.Message);

                LogControl.Write(
                    currentElement,
                    lea,
                    LogEventType.MoveDown,
                    LogEventStatus.Error,
                    LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void ElementTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            Element element = GetElement(sender);
            if (element != null)
            {
                elementControl.UpdateElement(element);

                if (SELECTALL_FOCUS_HANDLE)
                {
                    e.Handled = true;
                }

                if (element.NoteText != previousText)
                {
                    LogControl.Write(
                        element,
                        LogEventAccess.NoteTextBox,
                        LogEventType.TextChanged,
                        LogEventStatus.NULL,
                        LogEventInfo.PreviousName + LogControl.COMMA + previousText + LogControl.DELIMITER + 
                        LogEventInfo.NewName + LogControl.COMMA + element.NoteText);
                }
            }
        }

        private void ElementTextBox_GotFocus(object sender, RoutedEventArgs e)
        {
            Element element = GetElement(sender);
            if (element != null)
            {
                elementControl.PreviousElement = elementControl.CurrentElement;
                elementControl.CurrentElement = element;
                previousText = element.NoteText;

                ResetPowerDButtons(element);

                if (element.IsCommandNote)
                {
                    ((TextBox)sender).IsReadOnly = true;
                    ((TextBox)sender).Select(element.NoteText.Length, 0);
                }
                //else
                //    elementControl.SelectedElements.Add(element);
            }
        }

        private void ResetPowerDButtons(Element element)
        {
            try
            {
                Microsoft.Windows.Controls.Ribbon.RibbonCommand pd_done = ((Microsoft.Windows.Controls.Ribbon.RibbonSplitButton)this.RibbonButton_PowerDDone).Command as Microsoft.Windows.Controls.Ribbon.RibbonCommand;
                Microsoft.Windows.Controls.Ribbon.RibbonCommand pd_defer = ((Microsoft.Windows.Controls.Ribbon.RibbonSplitButton)this.RibbonButton_PowerDDefer).Command as Microsoft.Windows.Controls.Ribbon.RibbonCommand;
                Microsoft.Windows.Controls.Ribbon.RibbonCommand pd_delegate = ((Microsoft.Windows.Controls.Ribbon.RibbonSplitButton)this.RibbonButton_PowerDDelegate).Command as Microsoft.Windows.Controls.Ribbon.RibbonCommand;
                Microsoft.Windows.Controls.Ribbon.RibbonCommand pd_delete = ((Microsoft.Windows.Controls.Ribbon.RibbonButton)this.RibbonButton_PowerDDelete).Command as Microsoft.Windows.Controls.Ribbon.RibbonCommand;

                switch (element.Type)
                {
                    case ElementType.Heading:
                        if (element.IsExpanded)
                        {
                            RibbonButton_PowerDDone.IsEnabled = false;
                            RibbonButton_PowerDDefer.IsEnabled = false;
                            RibbonButton_PowerDDelegate.IsEnabled = false;
                            RibbonButton_PowerDDelete.IsEnabled = false;
                        }
                        else
                        {
                            RibbonButton_PowerDDone.IsEnabled = true;
                            RibbonButton_PowerDDefer.IsEnabled = true;
                            RibbonButton_PowerDDelegate.IsEnabled = true;
                            RibbonButton_PowerDDelete.IsEnabled = true;
                        }
                        break;
                    case ElementType.Note:
                        if (element.IsCommandNote)
                        {
                            RibbonButton_PowerDDone.IsEnabled = false;
                            RibbonButton_PowerDDefer.IsEnabled = false;
                            RibbonButton_PowerDDelegate.IsEnabled = false;
                            RibbonButton_PowerDDelete.IsEnabled = false;
                        }
                        else
                        {
                            RibbonButton_PowerDDone.IsEnabled = true;
                            RibbonButton_PowerDDefer.IsEnabled = true;
                            RibbonButton_PowerDDelegate.IsEnabled = true;
                            RibbonButton_PowerDDelete.IsEnabled = true;
                        }
                        break;
                }

                if (element.PowerDStatus == PowerDStatus.Done && element.FontColor == ElementColor.SpringGreen.ToString())
                {
                    RibbonButton_PowerDDone.IsEnabled = false;
                    RibbonButton_PowerDDefer.IsEnabled = false;
                    RibbonButton_PowerDDelegate.IsEnabled = false;
                    RibbonButton_PowerDDelete.IsEnabled = true;

                    pd_delete.LabelTitle = PowerDDeleteType.Undo.ToString();
                    BitmapImage bi = new BitmapImage(new Uri("Images/powerd_undo.png", UriKind.Relative));
                    pd_delete.LargeImageSource = bi;
                }
                else if (element.PowerDStatus == PowerDStatus.Deferred && element.FontColor == ElementColor.Plum.ToString())
                {
                    RibbonButton_PowerDDone.IsEnabled = false;
                    RibbonButton_PowerDDefer.IsEnabled = false;
                    RibbonButton_PowerDDelegate.IsEnabled = false;
                    RibbonButton_PowerDDelete.IsEnabled = true;

                    pd_delete.LabelTitle = PowerDDeleteType.Undo.ToString();
                    BitmapImage bi = new BitmapImage(new Uri("Images/powerd_undo.png", UriKind.Relative));
                    pd_delete.LargeImageSource = bi;
                }
                else
                {
                    pd_delete.LabelTitle = PowerDDeleteType.Delete.ToString();
                    BitmapImage bi = new BitmapImage(new Uri("Images/powerd_delete.png", UriKind.Relative));
                    pd_delete.LargeImageSource = bi;
                }

                SHOW_DONEDEFER_HANDLE = false;
                RibbonPowerDOption_ShowAssociationMarkedDone.IsChecked = element.ParentElement.ShowAssociationMarkedDone;
                RibbonPowerDOption_ShowAssociationMarkedDefer.IsChecked = element.ParentElement.ShowAssociationMarkedDefer;
                SHOW_DONEDEFER_HANDLE = true;
            }
            catch (Exception ex)
            {

            }
        }

        private void ElementTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            /*
            Element currentElement = GetElement(sender);
            TextBox tb = sender as TextBox;
            currentElement.NoteText = tb.Text;

            LogControl.Write(
               elementControl.CurrentElement,
               LogEventAccess.Hotkey,
               LogEventType.TextChanged,
               LogEventStatus.NULL,
               null);
            */
        }

        private void ElementTextBox_Loaded(object sender, RoutedEventArgs e)
        {
            TextBox tb = sender as TextBox;
            Rect textRext = tb.GetRectFromCharacterIndex(tb.Text.Length);
            Point p = tb.TranslatePoint(new Point(0, 0), this.Plan);
            tb.Width = this.Plan.ActualWidth - Math.Max(0, textRext.Width) - p.X - TEXTBOX_RIGHT_MARGIN;
        }

        private void ElementTextBox_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            try
            {
                if (SELECTALL_FOCUS_HANDLE)
                {
                    if (elementControl.SelectedElements.Count <= 1)
                    {
                        ElementUnselectAll(elementControl.PreviousElement.ParentElement);
                    }
                    else
                    {
                        foreach (Element ele in elementControl.SelectedElements)
                        {
                            ElementUnselectAll(ele);
                        }
                    }
                }
            }
            catch (Exception)
            {

            }
            finally
            {
                //Element currentElement = GetElement(sender);
                elementControl.SelectedElements.Clear();
                //elementControl.SelectedElements.Add(currentElement);
            }
        }

        #region Context Menu

        private void ElementTextBox_ContextMenu_Opened(object sender, RoutedEventArgs e)
        {
            ContextMenu cm = sender as ContextMenu;
            TextBox tb = cm.PlacementTarget as TextBox;
            if (tb.SelectedText == String.Empty)
            {
                tb.SelectAll();
            }
        }

        private void ElementTextBox_ContextMenu_Closed(object sender, RoutedEventArgs e)
        {
            
        }

        private void ElementTextBox_ContextMenuItem_Explore(object sender, RoutedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(elementControl.CurrentElement.Path);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("ElementTextBox_ContextMenuItem_Explore\n" + ex.Message);
            }
        }

        private void ElementTextBox_ContextMenuItem_NewHeading(object sender, RoutedEventArgs e)
        {
            try
            {
                RibbonCommandNewHeading_Executed(sender, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("ElementTextBox_ContextMenuItem_NewHeading\n" + ex.Message);
            }
        }

        private void ElementTextBox_ContextMenuItem_NewNote(object sender, RoutedEventArgs e)
        {
            try
            {
                RibbonCommandNewNote_Executed(sender, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("ElementTextBox_ContextMenuItem_NewNote\n" + ex.Message);
            }
        }

        private void ElementTextBox_ContextMenuItem_NewTextDocument(object sender, RoutedEventArgs e)
        {
            try
            {
                RibbonCommandNewTextDocument_Executed(sender, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("ElementTextBox_ContextMenuItem_NewTextDocument\n" + ex.Message);
            }
        }

        private void ElementTextBox_ContextMenuItem_NewWord(object sender, RoutedEventArgs e)
        {
            try
            {
                RibbonCommandNewWord_Executed(sender, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("ElementTextBox_ContextMenuItem_NewWordDocument\n" + ex.Message);
            }
        }

        private void ElementTextBox_ContextMenuItem_NewExcel(object sender, RoutedEventArgs e)
        {
            try
            {
                RibbonCommandNewExcel_Executed(sender, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("ElementTextBox_ContextMenuItem_NewExcel\n" + ex.Message);
            }
        }

        private void ElementTextBox_ContextMenuItem_NewPowerPoint(object sender, RoutedEventArgs e)
        {
            try
            {
                RibbonCommandNewPowerPoint_Executed(sender, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("ElementTextBox_ContextMenuItem_NewPowerPoint\n" + ex.Message);
            }
        }

        private void ElementTextBox_ContextMenuItem_NewOneNote(object sender, RoutedEventArgs e)
        {
            try
            {
                RibbonCommandNewOneNote_Executed(sender, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("ElementTextBox_ContextMenuItem_NewOneNote\n" + ex.Message);
            }
        }

        private void ElementTextBox_ContextMenuItem_NewEmail(object sender, RoutedEventArgs e)
        {
            try
            {
                MenuItem mi = sender as MenuItem;
                //mi.ReleaseMouseCapture();
                
                RibbonCommandNewEmail_Executed(sender, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("ElementTextBox_ContextMenuItem_NewEmail\n" + ex.Message);
            }
        }

        private void ElementTextBox_ContextMenuItem_NewTwitter(object sender, RoutedEventArgs e)
        {
            try
            {
                RibbonCommandNewTwitter_Executed(sender, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("ElementTextBox_ContextMenuItem_NewTwitter\n" + ex.Message);
            }
        }

        private void ElementTextBox_ContextMenuItem_Link_MouseEnter(object sender, MouseEventArgs e)
        {
            try
            {
                StartLoadingUI();

                MenuItem mi_link = sender as MenuItem;
                while (mi_link.Items.Count > 2)
                {
                    mi_link.Items.RemoveAt(0);
                }

                int index = 0;
                
                ActiveWindow aw = new ActiveWindow();
                foreach (InfoItem ii in aw.GetActiveItemList())
                {
                    MenuItem mi = new MenuItem();
                    if (ii.Title.Length > ACTIVEITEM_TITLE_LENGTH)
                    {
                        mi.Header = ii.Title.Substring(0, ACTIVEITEM_TITLE_LENGTH) + "...";
                    }
                    else
                    {
                        mi.Header = ii.Title;
                    }
                    switch (ii.Type)
                    {
                        case InfoItemType.Web:
                            mi.Icon = new Image
                            {
                                Source = new BitmapImage(new Uri(FileTypeHandler.GetIcon(ElementAssociationType.Web, ii.Uri), UriKind.Absolute)),
                                Width = 18,
                                Height = 18,
                            };
                            break;
                        case InfoItemType.File:
                            mi.Icon = new Image
                            {
                                Source = new BitmapImage(new Uri(FileTypeHandler.GetIcon(ElementAssociationType.File, ii.Uri), UriKind.Absolute)),
                                Width = 18,
                                Height = 18,
                            };
                            break;
                        case InfoItemType.Email:
                            mi.Icon = new Image
                            {
                                Source = new BitmapImage(new Uri(FileTypeHandler.GetIcon(ElementAssociationType.Email, ii.Uri), UriKind.Absolute)),
                                Width = 18,
                                Height = 18,
                            };
                            break;
                    };
                    mi.Tag = ii;
                    mi.ToolTip = ii.Title;
                    mi.Click += new RoutedEventHandler(ElementTextBox_ContextMenuItem_ActiveItem_Click);
                    
                    mi_link.Items.Insert(index++, mi);
                }

                EndLoadingUI();
            }
            catch (Exception ex)
            {
                MenuItem mi_link = sender as MenuItem;
                while (mi_link.Items.Count > 2)
                {
                    mi_link.Items.RemoveAt(0);
                }

                //MessageBox.Show("ElementTextBox_ContextMenuItem_Link_MouseEnter\n" + ex.Message);
            }
        }

        private void ElementTextBox_ContextMenuItem_ActiveItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                LogControl.Write(
                   elementControl.CurrentElement,
                   LogEventAccess.NoteTextContextMenu,
                   LogEventType.LinkToActiveItem,
                   LogEventStatus.Begin,
                   null);

                MenuItem mi = sender as MenuItem;
                InfoItem ii = mi.Tag as InfoItem;
                switch (ii.Type)
                {
                    case InfoItemType.Web:
                        elementControl.AddAssociation(elementControl.CurrentElement, ii.Uri, ElementAssociationType.Web, ii.Title);
                        break;
                    case InfoItemType.File:
                        elementControl.AddAssociation(elementControl.CurrentElement, ii.Uri, ElementAssociationType.FileShortcut, ii.Title);
                        break;
                    case InfoItemType.Email:
                        elementControl.AddAssociation(elementControl.CurrentElement, ii.Uri, ElementAssociationType.Email, ii.Title);
                        break;
                };

                GetFocusToElementTextBox(elementControl.CurrentElement, elementControl.CurrentElement.NoteText.Length, false, false);

                LogControl.Write(
                   elementControl.CurrentElement,
                   LogEventAccess.NoteTextContextMenu,
                   LogEventType.LinkToActiveItem,
                   LogEventStatus.End,
                   null);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("ActiveItem_Click\n" + ex.Message);

                LogControl.Write(
                   elementControl.CurrentElement,
                   LogEventAccess.NoteTextContextMenu,
                   LogEventType.LinkToActiveItem,
                   LogEventStatus.Error,
                   LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void ElementTextBox_ContextMenuItem_LinkToFile(object sender, RoutedEventArgs e)
        {
            try
            {
                RibbonCommandLinkToFile_Executed(null, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("ElementTextBox_ContextMenuItem_LinkFile\n" + ex.Message);
            }
        }

        private void ElementTextBox_ContextMenuItem_LinkToFolder(object sender, RoutedEventArgs e)
        {
            try
            {
                RibbonCommandLinkToFolder_Executed(null, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("ElementTextBox_ContextMenuItem_LinkFolder\n" + ex.Message);
            }
        }

        private void ElementTextBox_ContextMenuItem_Delete(object sender, RoutedEventArgs e)
        {
            try
            {
                RibbonCommandDelete_Executed(null, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("ElementTextBox_ContextMenuItem_Delete\n" + ex.Message);
            }
        }

        private void ElementTextBox_ContextMenuItem_Import(object sender, RoutedEventArgs e)
        {
            Element element = elementControl.CurrentElement;

            try
            {
                if (element != null & element.IsHeading)
                {
                    DateTime dt = JournalControl.GetDateTime(element.Path);
                    elementControl.AddAppointments(element, dt);

                    PlanTreeViewItem_Expanded(elementControl.FindElementByGuid(elementControl.Root, element.ID), null);

                    LogControl.Write(
                        element,
                        LogEventAccess.NoteTextContextMenu,
                        LogEventType.Import,
                        LogEventStatus.NULL,
                        null);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("AssociationDropDownMenu_Open\n" + ex.Message);

                LogControl.Write(
                    element,
                    LogEventAccess.NoteTextContextMenu,
                    LogEventType.Import,
                    LogEventStatus.Error,
                    LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void ElementTextBox_ContextMenuItem_LabelWith(object sender, RoutedEventArgs e)
        {
            try
            {
                RibbonCommandLabelWith_Executed(null, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("ElementTextBox_ContextMenuItem_LinkFolder\n" + ex.Message);
            }

        }

        #endregion

        #endregion

        #region PlanTreeViewItem

        private void PlanTreeViewItem_Expanded(object sender, RoutedEventArgs e)
        {
            LogEventAccess lea = LogEventAccess.Ribbon;

            Element element;
            if (sender is Element)
            {
                element = sender as Element;
                lea = LogEventAccess.Ribbon;
            }
            else if (sender is MenuItem)
            {
                element = ((MenuItem)sender).Tag as Element;
                lea = LogEventAccess.NodeContextMenu;
            }
            else
            {
                element = ((TreeViewItem)e.OriginalSource).Header as Element;
                lea = LogEventAccess.NodeIcon;
            }

            try
            {
                if (element != null && element.IsHeading)
                {
                    if (elementControl.IsCircularHeading(element))
                    {
                        MessageBox.Show("You cannot expand \"" + element.NoteText + "\" circularly.");
                        return;
                    }
                 
                    element.IsExpanded = true;
                    element.LevelOfSynchronization = 1;
                    elementControl.UpdateElement(element);

                    GetFocusToElementTextBox(element, -1, false, false);

                    StartLoadingUI();

                    elementControl.SyncHeadingElement(element);

                    // Keep only one heading expanded
                    foreach (Element ele in elementControl.FindAllHeadingElements(elementControl.Root))
                    {
                        if (ele.Path.ToLower() == element.Path.ToLower() && ele.ID != element.ID && ele.IsExpanded)
                        {
                            ele.IsExpanded = false;
                        }
                    }

                    LogControl.Write(
                        element,
                        lea,
                        LogEventType.Expand,
                        LogEventStatus.NULL,
                        null);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("PlanTreeViewItem_Expanded\n" + ex.Message);

                LogControl.Write(
                    element,
                    lea,
                    LogEventType.Expand,
                    LogEventStatus.Error,
                    LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void PlanTreeViewItem_Collapsed(object sender, RoutedEventArgs e)
        {
            LogEventAccess lea = LogEventAccess.Ribbon;

            Element element;
            if (sender is Element)
            {
                element = sender as Element;
                lea = LogEventAccess.Ribbon;
            }
            else if (sender is MenuItem)
            {
                element = ((MenuItem)sender).Tag as Element;
                lea = LogEventAccess.NodeContextMenu;
            }
            else
            {
                element = ((TreeViewItem)e.OriginalSource).Header as Element;
                lea = LogEventAccess.NodeIcon;
            }

            try
            {
                if (element != null && element.IsHeading)
                {
                    element.IsExpanded = false;
                    element.LevelOfSynchronization = 0;
                    elementControl.UpdateElement(element);

                    element.Elements.Clear();
                    GC.Collect();

                    LogControl.Write(
                        element,
                        lea,
                        LogEventType.Collapse,
                        LogEventStatus.NULL,
                        null);
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show("PlanTreeViewItem_Collapsed\n" + ex.Message);

                LogControl.Write(
                    element,
                    lea,
                    LogEventType.Collapse,
                    LogEventStatus.Error,
                    LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void PlanTreeView_MouseLeave(object sender, MouseEventArgs e)
        {
            if (elementControl.HoverElement != null)
            {
                if (Properties.Settings.Default.ShowOutline == false)
                {
                    elementControl.HoverElement.ShowExpander = Visibility.Hidden;
                    elementControl.HoverElement.HeadImageSource = HeadImageControl.Note_Empty;
                    if (elementControl.HoverElement.FlagStatus == FlagStatus.Normal)
                    {
                        elementControl.HoverElement.ShowFlag = Visibility.Hidden;
                    }
                }
            }
        }

        private void Expander_MouseEnter(object sender, MouseEventArgs e)
        {
            System.Windows.Controls.Primitives.ToggleButton toggbleButton = sender as System.Windows.Controls.Primitives.ToggleButton;
            Element element = toggbleButton.DataContext as Element;
            if (element != null)
            {
                toggbleButton.ToolTip = "Heading Level: " + element.Level.ToString();
            }
        }

        private void Expander_MouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            try
            {
                System.Windows.Controls.Primitives.ToggleButton toggbleButton = sender as System.Windows.Controls.Primitives.ToggleButton;
                Element element = toggbleButton.DataContext as Element;
                //TreeViewItem tvi = toggbleButton.TemplatedParent as TreeViewItem;
                
                if (element != null)
                {
                    elementControl.CurrentElement = element;

                    //TextBox tb = null;
                    //FindTextBox(tvi as DependencyObject, element.ID, ref tb);
                    //if (tb != null)
                    //{
                    //    tb.SelectAll();
                    //}

                    ContextMenu cm = new ContextMenu();
                    toggbleButton.ContextMenu = cm;

                    MenuItem select = new MenuItem();
                    select.Header = "Select";
                    select.Tag = element;
                    select.Click += HeadImageContextMenu_Select;

                    MenuItem selectall = new MenuItem();
                    selectall.Header = "Select All";
                    selectall.Tag = element;
                    selectall.Click += HeadImageContextMenu_SelectAll;

                    MenuItem hide = new MenuItem();
                    hide.Focusable = false;
                    hide.Header = "Hide";
                    hide.Click += new RoutedEventHandler(HeadImageContextMenu_Hide);

                    MenuItem expand = new MenuItem();
                    expand.Header = "Expand";
                    expand.Icon = new Image
                    {
                        Source = new BitmapImage(new Uri("Images/tool_expand.png", UriKind.Relative)),
                        Width = 18,
                        Height = 18,
                    };
                    expand.Tag = element;
                    expand.Click += PlanTreeViewItem_Expanded;

                    MenuItem collapse = new MenuItem();
                    collapse.Header = "Collapse";
                    collapse.Icon = new Image
                    {
                        Source = new BitmapImage(new Uri("Images/tool_collapse.png", UriKind.Relative)),
                        Width = 18,
                        Height = 18,
                    };
                    collapse.Tag = element;
                    collapse.Click += PlanTreeViewItem_Collapsed;

                    MenuItem promote = new MenuItem();
                    promote.Header = "Promote";
                    promote.Icon = new Image
                    {
                        Source = new BitmapImage(new Uri("Images/tool_promote.png", UriKind.Relative)),
                        Width = 18,
                        Height = 18,
                    };
                    promote.Click += new RoutedEventHandler(HeadImageContextMenu_Promote);

                    MenuItem demote = new MenuItem();
                    demote.Header = "Demote";
                    demote.Icon = new Image
                    {
                        Source = new BitmapImage(new Uri("Images/tool_demote.png", UriKind.Relative)),
                        Width = 18,
                        Height = 18,
                    };
                    demote.Click += new RoutedEventHandler(HeadImageContextMenu_Demote);

                    MenuItem moveup = new MenuItem();
                    moveup.Header = "Move Up";
                    moveup.Icon = new Image
                    {
                        Source = new BitmapImage(new Uri("Images/tool_moveup.png", UriKind.Relative)),
                        Width = 18,
                        Height = 18,
                    };
                    moveup.Click += new RoutedEventHandler(HeadImageContextMenu_MoveUp);

                    MenuItem movedown = new MenuItem();
                    movedown.Header = "Move Down";
                    movedown.Icon = new Image
                    {
                        Source = new BitmapImage(new Uri("Images/tool_movedown.png", UriKind.Relative)),
                        Width = 18,
                        Height = 18,
                    };
                    movedown.Click += new RoutedEventHandler(HeadImageContextMenu_MoveDown);

                    MenuItem delete = new MenuItem();
                    delete.Header = "Delete";
                    delete.Icon = new Image
                    {
                        Source = new BitmapImage(new Uri("Images/tool_delete.png", UriKind.Relative)),
                        Width = 18,
                        Height = 18,
                    };
                    delete.Click += new RoutedEventHandler(HeadImageContextMenu_Delete);

                    cm.Items.Add(select);
                    //cm.Items.Add(selectall);
                    //cm.Items.Add(hide);
                    cm.Items.Add(new Separator());
                    cm.Items.Add(expand);
                    cm.Items.Add(collapse);
                    cm.Items.Add(promote);
                    cm.Items.Add(demote);
                    cm.Items.Add(moveup);
                    cm.Items.Add(movedown);
                    cm.Items.Add(new Separator());
                    cm.Items.Add(delete);
                }

                e.Handled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("Expander_MouseRightButtonDown\n" + ex.Message);
            }
        }

        #endregion

        #region Element Tail Image & Head Image

        private void AssociationDropDownMenu_Open(object sender, RoutedEventArgs e)
        {
            MenuItem mi = sender as MenuItem;
            Element element = mi.DataContext as Element;
            try
            {
                if (element != null)
                {
                    if (element.IsHeading)
                    {
                        MainWindow newWindow = new MainWindow(element.Path);
                        newWindow.Show();
                    }
                    else
                    {
                        if (element.AssociationType == ElementAssociationType.Email && elementControl.IsICCEmail(element))
                        {
                            elementControl.FindSentEmail(element);
                        }
                        System.Diagnostics.Process.Start(element.AssociationURIFullPath);
                    }

                    LogControl.Write(
                        element,
                        LogEventAccess.AssociationContextMenu,
                        LogEventType.Open,
                        LogEventStatus.NULL,
                        null);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("AssociationDropDownMenu_Open\n" + ex.Message);

                LogControl.Write(
                    element,
                    LogEventAccess.AssociationContextMenu,
                    LogEventType.Open,
                    LogEventStatus.Error,
                    LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void AssociationDropDownMenu_Explore(object sender, RoutedEventArgs e)
        {
            MenuItem mi = sender as MenuItem;
            Element element = mi.DataContext as Element;
            try
            {
                if (element != null)
                {
                    if (element.AssociationType == ElementAssociationType.FileShortcut)
                    {
                        IWshRuntimeLibrary.IWshShortcut shortcut = elementControl.GetShortcut(element);
                        if (shortcut != null)
                        {
                            System.Diagnostics.Process.Start(System.IO.Path.GetDirectoryName(shortcut.TargetPath));
                        }
                    }
                    else if (element.AssociationType == ElementAssociationType.FolderShortcut)
                    {
                        IWshRuntimeLibrary.IWshShortcut shortcut = elementControl.GetShortcut(element);
                        if (shortcut != null)
                        {
                            System.Diagnostics.Process.Start(shortcut.TargetPath);
                        }
                    }
                    else
                    {
                        System.Diagnostics.Process.Start(element.Path);
                    }

                    LogControl.Write(
                        element,
                        LogEventAccess.AssociationContextMenu,
                        LogEventType.Explore,
                        LogEventStatus.NULL,
                        null);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("AssociationDropDownMenu_Explore\n" + ex.Message);

                LogControl.Write(
                    element,
                    LogEventAccess.AssociationContextMenu,
                    LogEventType.Explore,
                    LogEventStatus.Error,
                    LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void AssociationDropDownMenu_Rename(object sender, RoutedEventArgs e)
        {
            MenuItem mi = sender as MenuItem;
            Element element = mi.DataContext as Element;
            try
            {
                if (element != null)
                {
                    FileNameWindow fnw = new FileNameWindow(element.AssociationURIFullPath, "Please rename:");
                    if (fnw.ShowDialog().Value)
                    {
                        if (element.IsLocalHeading)
                        {
                            if (elementControl.CheckOpenFiles(element))
                            {
                                return;
                            }

                            string newFolderPath = System.IO.Path.GetDirectoryName(System.IO.Path.GetDirectoryName(element.Path)) + System.IO.Path.DirectorySeparatorChar +
                                fnw.FileName + System.IO.Path.DirectorySeparatorChar;
                            string previousPath = element.Path;
                            element.Path = newFolderPath;
                            elementControl.RenameFolder(element, previousPath);
                        }
                        else
                        {
                            string newFileNameFullPath = System.IO.Path.GetDirectoryName(element.AssociationURIFullPath) + System.IO.Path.DirectorySeparatorChar +
                                fnw.FileName + System.IO.Path.GetExtension(element.AssociationURIFullPath);
                            System.IO.File.Move(element.AssociationURIFullPath, newFileNameFullPath);
                            element.AssociationURI = System.IO.Path.GetFileName(newFileNameFullPath);
                            elementControl.UpdateElement(element);
                        }

                        LogControl.Write(
                            element,
                            LogEventAccess.AssociationContextMenu,
                            LogEventType.Rename,
                            LogEventStatus.NULL,
                            null);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("AssociationDropDownMenu_Rename\n" + ex.Message);

                LogControl.Write(
                    element,
                    LogEventAccess.AssociationContextMenu,
                    LogEventType.Rename,
                    LogEventStatus.Error,
                    LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void AssociationDropDownMenu_NewVersion(object sender, RoutedEventArgs e)
        {
            MenuItem mi = sender as MenuItem;
            Element element = mi.DataContext as Element;
            try
            {
                if (element != null)
                {
                    int index = 2;
                    string fileFullName = element.AssociationURIFullPath;
                    string fileFolder = System.IO.Path.GetDirectoryName(fileFullName) + System.IO.Path.DirectorySeparatorChar;
                    string fileName = System.IO.Path.GetFileNameWithoutExtension(fileFullName);
                    string fileExt = System.IO.Path.GetExtension(fileFullName);
                    System.Text.RegularExpressions.Regex regex = new System.Text.RegularExpressions.Regex(@"\w*\([0-9]+\)");
                    if (regex.IsMatch(fileName))
                    {
                        int lastIndex = fileName.LastIndexOf('(');
                        index = 1 + Int32.Parse(fileName.Substring(lastIndex + 1, fileName.LastIndexOf(')') - lastIndex - 1));
                        fileName = fileName.Substring(0, lastIndex - 1);
                    }
                    while (System.IO.File.Exists(fileFullName))
                    {
                        fileFullName = fileFolder + fileName + " (" + index.ToString() + ")" + fileExt;
                        index++;
                    }

                    LogControl.Write(
                        element,
                        LogEventAccess.AssociationContextMenu,
                        LogEventType.NewVersion,
                        LogEventStatus.Begin,
                        null);

                    FileNameWindow fnw = new FileNameWindow(fileFullName, "Please enter the file name of a new version:");
                    if (fnw.ShowDialog().Value)
                    {
                        string newFileNameFullPath = fileFolder + fnw.FileName + fileExt;
                        System.IO.File.Copy(element.AssociationURIFullPath, newFileNameFullPath);

                        Element newElement = elementControl.CreateNewElement(ElementType.Note, " --- " + fnw.FileName);
                        newElement.AssociationType = ElementAssociationType.File;
                        newElement.AssociationURI = fnw.FileName + fileExt;
                        newElement.TailImageSource = FileTypeHandler.GetIcon(newElement.AssociationType, newElement.AssociationURIFullPath);

                        elementControl.InsertElement(newElement, element.ParentElement, element.Position);

                        LogControl.Write(
                            newElement,
                            LogEventAccess.AssociationContextMenu,
                            LogEventType.NewVersion,
                            LogEventStatus.End,
                            LogEventInfo.NewName + LogControl.COMMA + fnw.FileName + fileExt);
                    }
                    else
                    {
                        LogControl.Write(
                            element,
                            LogEventAccess.AssociationContextMenu,
                            LogEventType.NewVersion,
                            LogEventStatus.Cancel,
                            null);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("AssociationDropDownMenu_NewVersion\n" + ex.Message);

                LogControl.Write(
                    element,
                    LogEventAccess.AssociationContextMenu,
                    LogEventType.NewVersion,
                    LogEventStatus.Error,
                    LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void AssociationDropDownMenu_Delete(object sender, RoutedEventArgs e)
        {
            MenuItem mi = sender as MenuItem;
            Element element = mi.DataContext as Element;
            try
            {
                if (element != null)
                {
                    LogControl.Write(
                        element,
                        LogEventAccess.AssociationContextMenu,
                        LogEventType.Delete,
                        LogEventStatus.Begin,
                        null);

                    DeleteWindow dw;
                    DeleteMessageType dmt = DeleteMessageType.Default;
                    if (element.IsLocalHeading)
                    {
                        dmt = DeleteMessageType.HeadingWithChildren;
                    }
                    else if (element.AssociationType == ElementAssociationType.File)
                    {
                        dmt = DeleteMessageType.File;
                    }
                    else
                    {
                        dmt = DeleteMessageType.Shortcut;
                    }
                    dw = new DeleteWindow(dmt);
                    if (dw.ShowDialog().Value == true)
                    {
                        if (element.IsLocalHeading)
                        {
                            elementControl.DeleteElement(element);
                        }
                        else
                        {
                            elementControl.Demote(element);
                            elementControl.RemoveAssociation(element);
                        }

                        LogControl.Write(
                            element,
                            LogEventAccess.AssociationContextMenu,
                            LogEventType.Delete,
                            LogEventStatus.End,
                            LogEventInfo.DeleteType + LogControl.COMMA + dmt.ToString());
                    }
                    else
                    {
                        LogControl.Write(
                            element,
                            LogEventAccess.AssociationContextMenu,
                            LogEventType.Delete,
                            LogEventStatus.Cancel,
                            null);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("AssociationDropDownMenu_Delete\n" + ex.Message);

                LogControl.Write(
                    element,
                    LogEventAccess.AssociationContextMenu,
                    LogEventType.Delete,
                    LogEventStatus.Error,
                    LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void AssociationDropDownMenu_Reply(object sender, RoutedEventArgs e)
        {
            MenuItem mi = sender as MenuItem;
            Element element = mi.DataContext as Element;
            try
            {
                if (element != null)
                {
                    elementControl.ReplyEmail(element);

                    LogControl.Write(
                        element,
                        LogEventAccess.AssociationContextMenu,
                        LogEventType.ReplyToAll,
                        LogEventStatus.NULL,
                        null);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("AssociationDropDownMenu_Reply\n" + ex.Message);

                LogControl.Write(
                    element,
                    LogEventAccess.AssociationContextMenu,
                    LogEventType.ReplyToAll,
                    LogEventStatus.Error,
                    LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void AssociationDropDownMenu_Execute(object sender, RoutedEventArgs e)
        {
            Element element = null;
            if (sender is MenuItem)
            {
                MenuItem mi = sender as MenuItem;
                element = mi.DataContext as Element;
            }
            else if (sender is Wpf.Controls.SplitButton)
            {
                Wpf.Controls.SplitButton sb = sender as Wpf.Controls.SplitButton;
                Image image = sb.Content as Image;
                element = GetElement(image);
            }

            try
            {
                if (element != null)
                {
                    switch (element.Command)
                    {
                        case ElementCommand.None:
                            break;
                        case ElementCommand.ShowJournalInNewWindow:
                            MainWindow newWindow = new MainWindow(StartProcess.JOURNAL_PATH);
                            newWindow.Show();                            
                            break;
                        case ElementCommand.ImportMeetingsAndAppointmentsFromOutlook:
                            DateTime dt = JournalControl.GetDateTime(element.ParentElement.Path);
                            elementControl.AddAppointments(element.ParentElement, dt);
                            //PlanTreeViewItem_Expanded(elementControl.FindElementByGuid(elementControl.Root, element.ParentElement.ID), null);
                            break;
                        case ElementCommand.DisplayMoreAssociations:
                            //Display more association under current command note
                            elementControl.DisplayMoreAssociations(element);
                            break;
                        case ElementCommand.ShowHiddenAssociations:
                            element.ParentElement.ShowAssociationMarkedDone = true;
                            element.ParentElement.ShowAssociationMarkedDefer = true;
                            elementControl.UpdateElement(element.ParentElement);
                            ShowOrHideMarkedAssociation(element.ParentElement, PowerDStatus.Done, true);
                            ShowOrHideMarkedAssociation(element.ParentElement, PowerDStatus.Deferred, true);
                            break;
                    };

                    LogControl.Write(
                        element,
                        LogEventAccess.AssociationContextMenu,
                        LogEventType.CommandNote,
                        LogEventStatus.NULL,
                        null);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("AssociationDropDownMenu_Execute\n" + ex.Message);

                LogControl.Write(
                    element,
                    LogEventAccess.AssociationContextMenu,
                    LogEventType.CommandNote,
                    LogEventStatus.Error,
                    LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message + LogControl.DELIMITER +
                        LogEventInfo.Command + LogControl.COMMA + element.Command.ToString());
            }
        }

        private void AssociationDropDownMenu_RelatedMessgages(object sender, RoutedEventArgs e)
        {
            MenuItem mi = sender as MenuItem;
            Element element = mi.DataContext as Element;
            try
            {
                if (element != null)
                {
                    elementControl.RelatedMessages(element);

                    LogControl.Write(
                        element,
                        LogEventAccess.AssociationContextMenu,
                        LogEventType.RelatedMessages,
                        LogEventStatus.NULL,
                        null);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("AssociationDropDownMenu_RelatedMessgages\n" + ex.Message);

                LogControl.Write(
                    element,
                    LogEventAccess.AssociationContextMenu,
                    LogEventType.RelatedMessages,
                    LogEventStatus.Error,
                    LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void AssociationDropDownMenu_Click(object sender, RoutedEventArgs e)
        {
            Wpf.Controls.SplitButton sb = sender as Wpf.Controls.SplitButton;
            Image image = sb.Content as Image;
            Element element = GetElement(image);
            try
            {
                if (element != null)
                {
                    if (element.IsCommandNote == false)
                    {
                        if (element.AssociationType == ElementAssociationType.Email && elementControl.IsICCEmail(element))
                        {
                            elementControl.FindSentEmail(element);
                        }
                        System.Diagnostics.Process.Start(element.AssociationURIFullPath);

                        LogControl.Write(
                            element,
                            LogEventAccess.AssociationIcon,
                            LogEventType.Open,
                            LogEventStatus.NULL,
                            null);
                    }
                    else
                    {
                        AssociationDropDownMenu_Execute(sender, e);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("AssociationDropDownMenu_Click\n" + ex.Message);

                if (element.IsCommandNote == false)
                {
                    if (element.HasEmailAssociation == false)
                    {
                        LogControl.Write(
                            element,
                            LogEventAccess.AssociationIcon,
                            LogEventType.Open,
                            LogEventStatus.Error,
                            LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
                    }
                    else
                    {
                        LogControl.Write(
                            element,
                            LogEventAccess.AssociationIcon,
                            LogEventType.RelatedMessages,
                            LogEventStatus.Error,
                            LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
                    }
                }
                else
                {
                    LogControl.Write(
                        element,
                        LogEventAccess.AssociationIcon,
                        LogEventType.CommandNote,
                        LogEventStatus.Error,
                        LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message + LogControl.DELIMITER +
                            LogEventInfo.Command + LogControl.COMMA + element.Command.ToString());
                }
            }
        }

        private void AssociationDropDownMenu_LabelWith(object sender, RoutedEventArgs e)
        {
            MenuItem mi = sender as MenuItem;
            Element element = mi.DataContext as Element;

            try
            {
                if (element != null && element.IsHeading)
                {
                    //Ask user to select the folder
                    System.Windows.Forms.FolderBrowserDialog browse = new System.Windows.Forms.FolderBrowserDialog();
                    browse.Description = "Please select a heading to label with";

                    browse.RootFolder = Environment.SpecialFolder.Desktop;
                    browse.SelectedPath = StartProcess.LIFE_PATH;
                    browse.ShowNewFolderButton = false;

                    if (browse.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string path = browse.SelectedPath + System.IO.Path.DirectorySeparatorChar; ;
                        string folderName = path.Substring(path.LastIndexOf("\\") + 1);

                        //create a shortcut note at the selected folder.
                        elementControl.LabelWith(element, path);

                        LogControl.Write(
                            element,
                            LogEventAccess.AssociationContextMenu,
                            LogEventType.LabelWith,
                            LogEventStatus.NULL,
                            null);
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("ElementTextBox_ContextMenuItem_LabeWith\n" + ex.Message);

                LogControl.Write(
                   element,
                   LogEventAccess.AssociationContextMenu,
                   LogEventType.LabelWith,
                   LogEventStatus.Error,
                   LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void AssociationDropDownMenu_Import(object sender, RoutedEventArgs e)
        {
            MenuItem mi = sender as MenuItem;
            Element element = mi.DataContext as Element;
            try
            {
                if (element != null && element.IsHeading)
                {
                    DateTime dt = JournalControl.GetDateTime(element.Path);
                    elementControl.AddAppointments(element, dt);

                    PlanTreeViewItem_Expanded(elementControl.FindElementByGuid(elementControl.Root,element.ID), null);

                    LogControl.Write(
                        element,
                        LogEventAccess.AssociationContextMenu,
                        LogEventType.Import,
                        LogEventStatus.NULL,
                        null);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("AssociationDropDownMenu_Open\n" + ex.Message);

                LogControl.Write(
                    element,
                    LogEventAccess.AssociationContextMenu,
                    LogEventType.Import,
                    LogEventStatus.Error,
                    LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void TailImage_MouseEnter(object sender, MouseEventArgs e)
        {
            try
            {
                Mouse.OverrideCursor = Cursors.Hand;

                Element element = GetElement(sender);
                if (element != null && element.IsCommandNote == false)
                {
                    ((Image)sender).ToolTip = elementControl.GetAssociationDescription(element);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("TailImage_MouseEnter\n" + ex.Message);
            }
        }

        private void TailImage_MouseLeave(object sender, MouseEventArgs e)
        {
            Mouse.OverrideCursor = null;
        }

        private void HeadImage_MouseEnter(object sender, MouseEventArgs e)
        {
            Mouse.OverrideCursor = Cursors.Hand;

            StackPanel sp = VisualTreeHelper.GetParent(sender as DependencyObject) as StackPanel;
            TextBox tb = sp.FindName("ElementTextBox") as TextBox;
            Element element = GetElement(tb);
            if (element != null)
            {
                ((Image)sender).ToolTip = "Note Level: " + element.Level.ToString();
            }
        }

        private void HeadImage_MouseLeave(object sender, MouseEventArgs e)
        {
            Mouse.OverrideCursor = null;
        }

        private void HeadImage_MouseMove(object sender, MouseEventArgs e)
        {
            /*Point headImageDragEnd = e.GetPosition(null);
            Vector diff = headImageDragStart - headImageDragEnd;

            if (e.LeftButton == MouseButtonState.Pressed &&
                Math.Abs(diff.X) > SystemParameters.MinimumHorizontalDragDistance &&
                Math.Abs(diff.Y) > SystemParameters.MinimumVerticalDragDistance)
            {
                StackPanel sp = VisualTreeHelper.GetParent(sender as DependencyObject) as StackPanel;
                TextBox tb = sp.FindName("ElementTextBox") as TextBox;
                Element element = GetElement(tb);

                if (element != null)
                {
                    DataObject dragData = new DataObject("Element", element);
                    DragDrop.DoDragDrop(sp, dragData, DragDropEffects.Move);

                    isHeadImageDrag = true;
                    
                }
            }*/
        }

        private void HeadImage_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            //headImageDragStart = e.GetPosition(null);
        }

        private void HeadImage_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            try
            {
                StackPanel sp = VisualTreeHelper.GetParent(sender as DependencyObject) as StackPanel;
                TextBox tb = sp.FindName("ElementTextBox") as TextBox;
                Element element = GetElement(tb);

                if (Keyboard.IsKeyDown(Key.LeftCtrl))
                {
                    elementControl.SelectedElements.Add(element);
                    SELECTALL_FOCUS_HANDLE = true;
                    tb.Focus();
                    tb.SelectAll();
                }
                else
                {
                    elementControl.SelectedElements.Clear();
                    elementControl.SelectedElements.Add(element);
                    tb.Focus();
                    tb.SelectAll();
                }

                e.Handled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("HeadImage_MouseLeftButtonDown\n" + ex.Message);
            }
        }

        private void HeadImage_ContextMenu_Opened(object sender, RoutedEventArgs e)
        {
            try
            {
                StackPanel sp = VisualTreeHelper.GetParent(((ContextMenu)sender).PlacementTarget as DependencyObject) as StackPanel;
                TextBox tb = sp.FindName("ElementTextBox") as TextBox;
                tb.SelectAll();
                SELECTALL_FOCUS_HANDLE = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("HeadImage_ContextMenu_Opened\n" + ex.Message);
            }
        }

        private void HeadImage_ContextMenu_Closed(object sender, RoutedEventArgs e)
        {

        }

        private void HeadImage_MouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            try
            {
                StackPanel sp = VisualTreeHelper.GetParent(sender as DependencyObject) as StackPanel;
                TextBox tb = sp.FindName("ElementTextBox") as TextBox;
                Element element = GetElement(sender);
                
                if (element != null)
                {
                    elementControl.CurrentElement = element;

                    Image image = sender as Image;
                    ContextMenu cm = new ContextMenu();
                    cm.Focusable = false;
                    cm.Opened += HeadImage_ContextMenu_Opened;
                    cm.Closed += HeadImage_ContextMenu_Closed;
                    image.ContextMenu = cm;

                    MenuItem select = new MenuItem();
                    select.Focusable = false;
                    select.Header = "Select";
                    select.Tag = element;
                    select.Click += HeadImageContextMenu_Select;

                    MenuItem hide = new MenuItem();
                    hide.Focusable = false;
                    hide.Header = "Hide";
                    hide.Click += new RoutedEventHandler(HeadImageContextMenu_Hide);

                    MenuItem promote = new MenuItem();
                    promote.Focusable = false;
                    promote.Header = "Promote";
                    promote.Icon = new Image
                    {
                        Source = new BitmapImage(new Uri("Images/tool_promote.png", UriKind.Relative)),
                        Width = 18,
                        Height = 18,
                    };
                    promote.Click += new RoutedEventHandler(HeadImageContextMenu_Promote);

                    MenuItem demote = new MenuItem();
                    demote.Focusable = false;
                    demote.Header = "Demote";
                    demote.Icon = new Image
                    {
                        Source = new BitmapImage(new Uri("Images/tool_demote.png", UriKind.Relative)),
                        Width = 18,
                        Height = 18,
                    };
                    demote.Click += new RoutedEventHandler(HeadImageContextMenu_Demote);

                    MenuItem moveup = new MenuItem();
                    moveup.Focusable = false;
                    moveup.Header = "Move Up";
                    moveup.Icon = new Image
                    {
                        Source = new BitmapImage(new Uri("Images/tool_moveup.png", UriKind.Relative)),
                        Width = 18,
                        Height = 18,
                    };
                    moveup.Click += new RoutedEventHandler(HeadImageContextMenu_MoveUp);

                    MenuItem movedown = new MenuItem();
                    movedown.Focusable = false;
                    movedown.Header = "Move Down";
                    movedown.Icon = new Image
                    {
                        Source = new BitmapImage(new Uri("Images/tool_movedown.png", UriKind.Relative)),
                        Width = 18,
                        Height = 18,
                    };
                    movedown.Click += new RoutedEventHandler(HeadImageContextMenu_MoveDown);

                    MenuItem delete = new MenuItem();
                    delete.Focusable = false;
                    delete.Header = "Delete";
                    delete.Icon = new Image
                    {
                        Source = new BitmapImage(new Uri("Images/tool_delete.png", UriKind.Relative)),
                        Width = 18,
                        Height = 18,
                    };
                    delete.Click += new RoutedEventHandler(HeadImageContextMenu_Delete);

                    if (elementControl.SelectedElements.Count <= 1)
                    {
                        cm.Items.Add(select);
                        //cm.Items.Add(hide);
                        cm.Items.Add(new Separator());
                        cm.Items.Add(promote);
                        cm.Items.Add(demote);
                        cm.Items.Add(moveup);
                        cm.Items.Add(movedown);
                        cm.Items.Add(new Separator());
                        cm.Items.Add(delete);
                    }
                    else
                    {
                        //cm.Items.Add(hide);
                        cm.Items.Add(delete);
                    }
                }

                e.Handled = true; 
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("HeadImage_MouseRightButtonDown\n" + ex.Message);
            }
        }

        #region Context Menu

        private void HeadImageContextMenu_Select(object sender, RoutedEventArgs e)
        {
            MenuItem mi = sender as MenuItem;
            Element element = mi.Tag as Element;

            try
            {
                GetFocusToElementTextBox(element, -1, false, false);

                LogControl.Write(
                    element,
                    LogEventAccess.NodeContextMenu,
                    LogEventType.Select,
                    LogEventStatus.NULL,
                    null);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("HeadImageContextMenu_Select\n" + ex.Message);

                LogControl.Write(
                    element,
                    LogEventAccess.NodeContextMenu,
                    LogEventType.Select,
                    LogEventStatus.Error,
                    LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void HeadImageContextMenu_SelectAll(object sender, RoutedEventArgs e)
        {
            MenuItem mi = sender as MenuItem;
            Element element = mi.Tag as Element;

            try
            {
                GetFocusToElementTextBox(element, -1, true, false);

                LogControl.Write(
                    element,
                    LogEventAccess.NodeContextMenu,
                    LogEventType.SelectAll,
                    LogEventStatus.NULL,
                    null);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("HeadImageContextMenu_SelectAll\n" + ex.Message);

                LogControl.Write(
                    element,
                    LogEventAccess.NodeContextMenu,
                    LogEventType.SelectAll,
                    LogEventStatus.Error,
                    LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void HeadImageContextMenu_Hide(object sender, RoutedEventArgs e)
        {
            try
            {
                RibbonCommandHide_Executed(sender, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("HeadImageContextMenu_Hide\n" + ex.Message);
            }
        }

        private void HeadImageContextMenu_Promote(object sender, RoutedEventArgs e)
        {
            try
            {
                RibbonCommandPromote_Executed(sender, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("HeadImageContextMenu_Promote\n" + ex.Message);
            }
        }

        private void HeadImageContextMenu_Demote(object sender, RoutedEventArgs e)
        {
            try
            {
                RibbonCommandDemote_Executed(sender, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("HeadImageContextMenu_Demote\n" + ex.Message);
            }
        }

        private void HeadImageContextMenu_MoveUp(object sender, RoutedEventArgs e)
        {
            try
            {
                RibbonCommandMoveUp_Executed(sender, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show("HeadImageContextMenu_MoveUp\n" + ex.Message);
            }
        }

        private void HeadImageContextMenu_MoveDown(object sender, RoutedEventArgs e)
        {
            try
            {
                RibbonCommandMoveDown_Executed(sender, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("HeadImageContextMenu_MoveDown\n" + ex.Message);
            }
        }

        private void HeadImageContextMenu_Delete(object sender, RoutedEventArgs e)
        {
            try
            {
                RibbonCommandDelete_Executed(sender, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("HeadImageContextMenu_Delete\n" + ex.Message);
            }
        }

        #endregion

        private void FlagImage_MouseDown(object sender, MouseButtonEventArgs e)
        {
            StartLoadingUI();

            StackPanel sp = VisualTreeHelper.GetParent(sender as DependencyObject) as StackPanel;
            TextBox tb = sp.FindName("ElementTextBox") as TextBox;
            Element currentElement = GetElement(tb);
            try
            {
                if (currentElement != null)
                {
                    if (tb.Text != currentElement.NoteText)
                    {
                        currentElement.NoteText = tb.Text;

                        LogControl.Write(
                           elementControl.CurrentElement,
                           LogEventAccess.Hotkey,
                           LogEventType.TextChanged,
                           LogEventStatus.NULL,
                           null);
                    }

                    if (currentElement.IsHeading)
                    {
                        if (currentElement.FlagStatus == FlagStatus.Normal)
                        {
                            FlagWindow fw = new FlagWindow(currentElement);
                            if (fw.ShowDialog().Value)
                            {
                                elementControl.Flag(currentElement,
                                    fw.HasStart, fw.StartDate, fw.StartTime, fw.StartAllDay,
                                    fw.HasDue, fw.DueDate, fw.DueTime, fw.DueAllDay,
                                    fw.AddToToday, fw.AddToReminder, fw.AddToTask);

                                foreach (Element ele in elementControl.FindAllHeadingElements(elementControl.Root))
                                {
                                    if (ele.Path.ToLower() == currentElement.Path.ToLower() && ele.ID != currentElement.ID)
                                    {
                                        elementControl.Flag(ele);
                                    }
                                }

                                LogControl.Write(
                                    currentElement,
                                    LogEventAccess.FlagIcon,
                                    LogEventType.Flag,
                                    LogEventStatus.NULL,
                                    LogEventInfo.HasStart + LogControl.COMMA + fw.HasStart.ToString() + LogControl.DELIMITER +
                                    LogEventInfo.StartDate + LogControl.COMMA + fw.StartDate.ToShortDateString() + LogControl.DELIMITER +
                                    LogEventInfo.StartTime + LogControl.COMMA + fw.StartTime.ToString() + LogControl.DELIMITER +
                                    LogEventInfo.StartAllDay + LogControl.COMMA + fw.StartAllDay.ToString() + LogControl.DELIMITER +
                                    LogEventInfo.HasDue + LogControl.COMMA + fw.HasDue.ToString() + LogControl.DELIMITER +
                                    LogEventInfo.DueDate + LogControl.COMMA + fw.DueDate.ToShortDateString() + LogControl.DELIMITER +
                                    LogEventInfo.DueTime + LogControl.COMMA + fw.DueTime.ToString() + LogControl.DELIMITER +
                                    LogEventInfo.DueAllDay + LogControl.COMMA + fw.DueAllDay.ToString() + LogControl.DELIMITER +
                                    LogEventInfo.AddToToday + LogControl.COMMA + fw.AddToToday.ToString() + LogControl.DELIMITER +
                                    LogEventInfo.AddToReminder + LogControl.COMMA + fw.AddToReminder.ToString() + LogControl.DELIMITER +
                                    LogEventInfo.AddToTask + LogControl.COMMA + fw.AddToTask.ToString()
                                    );
                            }
                            else
                            {
                                LogControl.Write(
                                    currentElement,
                                    LogEventAccess.FlagIcon,
                                    LogEventType.Flag,
                                    LogEventStatus.Cancel,
                                    null);
                            }
                        }
                        else if (currentElement.FlagStatus == FlagStatus.Flag)
                        {
                            CheckWindow cw = new CheckWindow(currentElement);
                            if (cw.ShowDialog().Value)
                            {
                                elementControl.Check(currentElement, cw.RemoveFromToday);

                                foreach (Element ele in elementControl.FindAllHeadingElements(elementControl.Root))
                                {
                                    if (ele.Path.ToLower() == currentElement.Path.ToLower() && ele.ID != currentElement.ID && ele.IsExpanded)
                                    {
                                        elementControl.Check(ele);
                                    }
                                }

                                LogControl.Write(
                                    currentElement,
                                    LogEventAccess.FlagIcon,
                                    LogEventType.Check,
                                    LogEventStatus.NULL,
                                    LogEventInfo.RemoveFromToday + LogControl.COMMA + cw.RemoveFromToday.ToString());
                            }
                            else
                            {
                                LogControl.Write(
                                    currentElement,
                                    LogEventAccess.FlagIcon,
                                    LogEventType.Check,
                                    LogEventStatus.Cancel,
                                    null);
                            }
                        }
                        else if (currentElement.FlagStatus == FlagStatus.Check)
                        {
                            elementControl.Uncheck(currentElement);

                            foreach (Element ele in elementControl.FindAllHeadingElements(elementControl.Root))
                            {
                                if (ele.Path.ToLower() == currentElement.Path.ToLower() && ele.ID != currentElement.ID)
                                {
                                    elementControl.Uncheck(ele);
                                }
                            }

                            LogControl.Write(
                                currentElement,
                                LogEventAccess.FlagIcon,
                                LogEventType.Uncheck,
                                LogEventStatus.NULL,
                                null);
                        }
                    }
                    else
                    {
                        if (currentElement.FlagStatus == FlagStatus.Normal)
                        {
                            elementControl.Flag(currentElement);

                            LogControl.Write(
                                currentElement,
                                LogEventAccess.FlagIcon,
                                LogEventType.Flag,
                                LogEventStatus.NULL,
                                null);
                        }
                        else if (currentElement.FlagStatus == FlagStatus.Flag)
                        {
                            elementControl.Check(currentElement);

                            LogControl.Write(
                                currentElement,
                                LogEventAccess.FlagIcon,
                                LogEventType.Check,
                                LogEventStatus.NULL,
                                null);
                        }
                        else if (currentElement.FlagStatus == FlagStatus.Check)
                        {
                            elementControl.Uncheck(currentElement);

                            LogControl.Write(
                                currentElement,
                                LogEventAccess.FlagIcon,
                                LogEventType.Uncheck,
                                LogEventStatus.NULL,
                                null);
                        }
                    }
                }

                GetFocusToElementTextBox(currentElement, -1, false, false);

                e.Handled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("FlagImage_MouseDown\n" + ex.Message);

                LogControl.Write(
                    currentElement,
                    LogEventAccess.FlagIcon,
                    LogEventType.Flag,
                    LogEventStatus.Error,
                    LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
            finally
            {
                EndLoadingUI();
            }
        }

        private void FlagImage_MouseEnter(object sender, MouseEventArgs e)
        {
            try
            {
                Element element = GetElement(sender);
                string tooltip = elementControl.GetFlagDescription(element);
                if (tooltip != String.Empty)
                {
                    ((Image)sender).ToolTip = tooltip;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("FlagImage_MouseEnter\n" + ex.Message);
            }
        }

        #endregion

        #region Navigation Tab

        private void LoadNavigationTab()
        {
            this.NavigationTabList.ItemsSource = elementControl.NI_Root.Items;

            if (Properties.Settings.Default.NavigationTabVisibility == true)
            {
                this.NavigationTabList.Visibility = Visibility.Visible;
                RibbonNavigationTab_Checked(null, null);
            }
            else
            {
                this.NavigationTabList.Visibility = Visibility.Collapsed;
                RibbonNavigationTab_Unchecked(null, null);
            }
        }

        private void NavigationTextBlock_MouseDown(object sender, MouseButtonEventArgs e)
        {
            NavigationItem ni = null;
            if (sender is NavigationItem)
            {
                ni = sender as NavigationItem;
            }
            else if (sender is TextBlock)
            {
                ni = ((TextBlock)sender).Tag as NavigationItem;
            }
            Element targetElement = elementControl.GetHeadingElement(ni);

            try
            {
                if (targetElement != null)
                {
                    TextBox tb = GetTextBox(targetElement);
                    if (tb != null)
                    {
                        if (previousNavigatedTextBox != null)
                        {
                            previousNavigatedTextBox.BorderBrush = Brushes.White;
                        }
                        previousNavigatedTextBox = tb;
                        tb.BorderBrush = Brushes.SkyBlue;
                        GetFocusToElementTextBox(targetElement, -1, false, true);
                    }

                    LogControl.Write(
                        targetElement,
                        LogEventAccess.NavigationTree,
                        LogEventType.ClickNavigationItem,
                        LogEventStatus.NULL,
                        null);
                }
                else
                {
                    StartLoadingUI();
                }
            }
            catch (Exception ex)
            {
                LogControl.Write(
                    targetElement,
                    LogEventAccess.NavigationTree,
                    LogEventType.ClickNavigationItem,
                    LogEventStatus.Error,
                    LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
        }

        private void NavigationTreeViewItem_Collapsed(object sender, RoutedEventArgs e)
        {
            try
            {

            }
            catch (Exception ex)
            {

            }
        }

        private void NavigationTreeViewItem_Expanded(object sender, RoutedEventArgs e)
        {
            try
            {
                NavigationItem ni_root = ((TreeViewItem)e.OriginalSource).Header as NavigationItem;

                elementControl.SyncNavigationItem(ni_root);
            }
            catch (Exception ex)
            {

            }
        }

        private void NavigationTabList_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            try
            {
                double width = Double.Parse(this.NavigationColumn.GetValue(ColumnDefinition.WidthProperty).ToString());
                Properties.Settings.Default.NavigationTabWidth = width;
                Properties.Settings.Default.Save();
            }
            catch (Exception ex)
            {
                //MessageBox.Show("NavigationTabList_SizeChanged\n" + ex.Message);
            }
        }

        #endregion
        
        #region Support Function

        private static TreeViewItem GetParentTreeViewItem(DependencyObject item)
        {
            if (item != null)
            {
                DependencyObject parent = VisualTreeHelper.GetParent(item);
                TreeViewItem parentTreeViewItem = parent as TreeViewItem;
                return (parentTreeViewItem != null) ? parentTreeViewItem : GetParentTreeViewItem(parent);
            }
            return null;
        }

        private Element GetElement(object sender)
        {
            TreeViewItem item = GetParentTreeViewItem((DependencyObject)sender);
            if (item == null)
            {
                return null;
            }
            else
            {
                return item.Header as Element;
            }
        }

        private void FindTextBox(DependencyObject d_obj, Guid id, ref TextBox target)
        {
            if (d_obj == null)
            {
                return;
            }
            else
            {
                if (d_obj.GetType() == typeof(TextBox) &&
                    ((TextBox)d_obj).Tag.ToString() == id.ToString() && 
                    ((TextBox)d_obj).IsVisible)
                {
                    target = (TextBox)d_obj;
                }
                else
                {
                    int count = VisualTreeHelper.GetChildrenCount(d_obj);
                    for (int i = 0; i < count; i++)
                    {
                        FindTextBox(VisualTreeHelper.GetChild(d_obj, i), id, ref target);
                    }
                }
            }
        }

        private void FindScrollViewer(DependencyObject d_obj, ref ScrollViewer target)
        {
            if (d_obj == null)
            {
                return;
            }
            else
            {
                if (d_obj.GetType() == typeof(ScrollViewer))
                {
                    target = (ScrollViewer)d_obj;
                }
                else
                {
                    int count = VisualTreeHelper.GetChildrenCount(d_obj);
                    for (int i = 0; i < count; i++)
                    {
                        FindScrollViewer(VisualTreeHelper.GetChild(d_obj, i), ref target);
                    }
                }
            }
        }

        private TextBox GetTextBox(Element element)
        {
            TextBox target = null;
            FindTextBox(this.Plan, element.ID, ref target);

            return target;
        }

        private int GetTextBoxSelectionStart(Element element)
        {
            TextBox target = null;
            FindTextBox(this.Plan, element.ID, ref target);

            return target.SelectionStart;
        }

        private string GetTextBoxSelectedText(Element element)
        {
            TextBox target = null;
            FindTextBox(this.Plan, element.ID, ref target);

            return target.SelectedText;
        }

        private void GetFocusToElementTextBox(
            Element element, 
            int cursor_pos, 
            bool highlightChildren, 
            bool bringToView)
        {
            if (element == null)
            {
                return;
            }

            UpdateLayout();

            TextBox target = null;
            FindTextBox(this.Plan, element.ID, ref target);
            if (target != null)
            {
                target.Focus();
                //Mouse.Capture(target, CaptureMode.Element);

                if (cursor_pos != -1)
                {
                    target.SelectionStart = cursor_pos;
                }
                else
                {
                    if (highlightChildren)
                    {
                        ElementUnselectAll(elementControl.PreviousElement);
                        target.Focus();
                        ElementSelectAll(element);
                    }
                    else
                    {
                        target.Focus();
                        target.SelectAll();
                    }
                }

                if (bringToView)
                {
                    Point p = target.TranslatePoint(new Point(0, 0), this.Plan);
                    double offset = p.Y - this.Plan.ActualHeight;

                    if ((offset <= 0 && Math.Abs(offset) <= this.Plan.ActualHeight / 3) ||
                        offset > 0)
                    {
                        ScrollViewer sv = null;
                        FindScrollViewer(this.Plan, ref sv);

                        sv.ScrollToVerticalOffset(sv.VerticalOffset + p.Y - this.Plan.ActualHeight / 3);
                    }
                }
            }
        }

        private void ElementSelectAll(Element element)
        {
            SELECTALL_FOCUS_HANDLE = true;

            TextBox tb = GetTextBox(element);
            if (tb != null)
            {
                switch (element.Type)
                {
                    case ElementType.Heading:
                        foreach (Element ele in element.Elements)
                        {
                            ElementSelectAll(ele);
                        }
                        tb.Focus();
                        tb.SelectAll();
                        break;
                    case ElementType.Note:
                        tb.Focus();
                        tb.SelectAll();
                        break;
                }
            }
        }

        private void ElementUnselectAll(Element element)
        {
            SELECTALL_FOCUS_HANDLE = false;

            TextBox tb = GetTextBox(element);
            if (tb != null)
            {
                switch (element.Type)
                {
                    case ElementType.Heading:
                        foreach (Element ele in element.Elements)
                        {
                            ElementUnselectAll(ele);
                        }
                        tb.Focus();
                        tb.SelectionStart = tb.Text.Length;
                        ElementTextBox_LostFocus(tb, new RoutedEventArgs());
                        break;
                    case ElementType.Note:
                        tb.Focus();
                        tb.SelectionStart = tb.Text.Length;
                        ElementTextBox_LostFocus(tb, new RoutedEventArgs());
                        break;
                }
            }
        }

        private void ShowOutlineIcon(Element element, bool show)
        {
            if (show)
            {
                element.ShowFlag = Visibility.Visible;

                if (element.IsHeading)
                {
                    element.ShowExpander = Visibility.Visible;

                    foreach (Element ele in element.Elements)
                    {
                        ShowOutlineIcon(ele, show);
                    }
                }
                else
                {
                    element.HeadImageSource = HeadImageControl.Note_Unselected;
                }
            }
            else
            {
                if (element.FlagStatus == FlagStatus.Normal)
                {
                    element.ShowFlag = Visibility.Hidden;
                }
                else
                {
                    element.ShowFlag = Visibility.Visible;
                }

                if (element.IsHeading)
                {
                    element.ShowExpander = Visibility.Hidden;

                    foreach (Element ele in element.Elements)
                    {
                        ShowOutlineIcon(ele, show);
                    }
                }
                else
                {
                    element.HeadImageSource = HeadImageControl.Note_Empty;
                }
            }
        }

        private void ResizeElementTextBox(Element element, Dictionary<int, double> dict)
        {
            TextBox tb;
            Rect textRext;
            Point p;

            int level = element.Level;

            tb = GetTextBox(element);
            if (tb != null)
            {
                if (dict.ContainsKey(level))
                {
                    tb.Width = dict[level];
                }
                else
                {
                    textRext = tb.GetRectFromCharacterIndex(tb.Text.Length);
                    p = tb.TranslatePoint(new Point(0, 0), this.Plan);
                    tb.Width = this.Plan.ActualWidth - textRext.Width - p.X - TEXTBOX_RIGHT_MARGIN;
                    dict.Add(level, tb.Width);
                }
            }

            switch (element.Type)
            {
                case ElementType.Heading:
                    foreach (Element ele in element.Elements)
                    {
                        ResizeElementTextBox(ele, dict);
                    }
                    break;
            }
        }

        private void ChangeElementTextFontFamily(Element element)
        {
            TextBox tb = GetTextBox(element);
            if (tb != null)
            {
                tb.FontFamily = new FontFamily(element.FontFamily);
            }
        }

        private void ChangeElementTextFontFamily(ElementType elementType, FontFamily ff)
        {
            switch (elementType)
            {
                case ElementType.Heading:
                    foreach (Element ele in elementControl.FindAllHeadingElements(elementControl.Root))
                    {
                        TextBox tb = GetTextBox(ele);
                        if (tb != null)
                        {
                            tb.FontFamily = ff;
                        }
                    }
                    break;
                case ElementType.Note:
                    ChangeNoteElementTextFontFamily(elementControl.Root, ff);
                    break;
            }
        }

        private void ChangeNoteElementTextFontFamily(Element element, FontFamily ff)
        {
            TextBox tb = GetTextBox(element);
            if (tb != null && element.Type == ElementType.Note)
            {
                tb.FontFamily = ff;
            }

            switch (element.Type)
            {
                case ElementType.Heading:
                    foreach (Element ele in element.Elements)
                    {
                        ChangeNoteElementTextFontFamily(ele, ff);
                    }
                    break;
            }
        }

        private void ChangeElementTextFontSize(Element element)
        {
            TextBox tb = GetTextBox(element);
            if (tb != null)
            {
                tb.FontSize = element.FontSize;
            }
        }

        private void ChangeElementTextFontSize(ElementType elementType, int size)
        {
            Typeface myTypeface = new Typeface(elementControl.Root.FontFamily);
            FormattedText ft = new FormattedText("Test",
                System.Globalization.CultureInfo.CurrentCulture,
                FlowDirection.LeftToRight,
                myTypeface, size, Brushes.Black);
            Properties.Settings.Default.TailImageHeight = ft.Height;
            Properties.Settings.Default.Save();

            switch (elementType)
            {
                case ElementType.Heading:
                    foreach (Element ele in elementControl.FindAllHeadingElements(elementControl.Root))
                    {
                        TextBox tb = GetTextBox(ele);
                        if (tb != null)
                        {
                            tb.FontSize = ele.FontSize;
                        }
                    }
                    break;
                case ElementType.Note:
                    ChangeNoteElementTextFontSize(elementControl.Root, size, ft.Height);
                    break;
            }
        }

        private void ChangeNoteElementTextFontSize(Element element, int size, double tailIconHeight)
        {
            element.TailImageHeight = tailIconHeight;
            element.TailImageWidth = tailIconHeight;

            TextBox tb = GetTextBox(element);
            if (tb != null && element.Type == ElementType.Note)
            {
                tb.FontSize = size;
            }

            switch (element.Type)
            {
                case ElementType.Heading:
                    foreach (Element ele in element.Elements)
                    {
                        ChangeNoteElementTextFontSize(ele, size, tailIconHeight);
                    }
                    break;
            }
        }

        private void StartLoadingUI()
        {
            this.Cursor = Cursors.Wait;
        }

        private void EndLoadingUI()
        {
            this.Cursor = Cursors.Arrow;
        }

        #endregion

        #region Element StackPanel

        private void ElementStackPanel_MouseEnter(object sender, MouseEventArgs e)
        {
            Element element = GetElement(sender);
            if (element != null)
            {
                if (Properties.Settings.Default.ShowOutline == false)
                {
                    if (element.IsHeading == false)
                    {
                        element.HeadImageSource = HeadImageControl.Note_Unselected;
                    }
                    else
                    {
                        element.ShowExpander = Visibility.Visible;
                    }

                    if (elementControl.HoverElement != null && elementControl.HoverElement != element)
                    {
                        if (elementControl.HoverElement.FlagStatus == FlagStatus.Flag ||
                        elementControl.HoverElement.FlagStatus == FlagStatus.Check)
                        {

                        }
                        else
                        {
                            elementControl.HoverElement.ShowFlag = Visibility.Hidden;
                        }

                        elementControl.HoverElement.HeadImageSource = HeadImageControl.Note_Empty;
                        elementControl.HoverElement.ShowExpander = Visibility.Hidden;
                    }

                    element.ShowFlag = Visibility.Visible;
                }

                TextBox tb = ((StackPanel)sender).FindName("ElementTextBox") as TextBox;
                tb.BorderBrush = Brushes.SkyBlue;

                //
                elementControl.HoverElement = element;
            }
        }

        private void ElementStackPanel_MouseLeave(object sender, MouseEventArgs e)
        {
            TextBox tb = ((StackPanel)sender).FindName("ElementTextBox") as TextBox;
            if (tb != null)
            {
                tb.BorderBrush = Brushes.White;
            }
        }

        private void ElementStackPanel_PreviewDragEnter(object sender, DragEventArgs e)
        {
            e.Effects = DragDropEffects.Link;
            e.Handled = true;
        }

        private void ElementStackPanel_PreviewDrop(object sender, DragEventArgs e)
        {
            Element dropElement = GetElement(sender);
            try
            {
                LogControl.Write(
                    dropElement,
                    LogEventAccess.AppWindow,
                    LogEventType.DragLink,
                    LogEventStatus.Begin,
                    null);

                if (dropElement != null)
                {
                    // Notice the order
                    if (e.Data.GetDataPresent(DataFormats.FileDrop))
                    {
                        #region File and Folder

                        string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                        string firstFilePath = files[0];

                        if (System.IO.Directory.Exists(firstFilePath))
                        {
                            // Folder
                            elementControl.AddAssociation(dropElement, firstFilePath, ElementAssociationType.FolderShortcut, null);
                        }
                        else
                        {
                            // File
                            elementControl.AddAssociation(dropElement, firstFilePath, ElementAssociationType.FileShortcut, null);
                        }

                        #endregion
                    }
                    else if (e.Data.GetDataPresent("Object Descriptor"))
                    {
                        #region Outlook

                        System.IO.MemoryStream mems = (System.IO.MemoryStream)e.Data.GetData("Object Descriptor");
                        string strDescriptor = OutlookSupportFunction.GetObjectDescriptor(mems);

                        if (strDescriptor.ToLower() == "outlook")
                        {
                            elementControl.AddAssociation(dropElement, null, ElementAssociationType.Email, null);
                        }

                        #endregion

                        #region Word Text

                        if (strDescriptor.ToLower() == "microsoft office word document")
                        {
                            string htmlText = e.Data.GetData("HTML Format").ToString();
                            string srcText = e.Data.GetData("Text").ToString();

                            srcText = srcText.Replace("\r\n", " ").Replace("\t", " ").Trim();
                            if (srcText.Length > StartProcess.MAX_EXTRACTTEXT_LENGTH)
                                srcText = srcText.Substring(0, StartProcess.MAX_EXTRACTTEXT_LENGTH) + "...";

                            // Word
                            // Sample htmlText = 
                            //Version:1.0
                            //StartHTML:0000000196
                            //EndHTML:0000021629
                            //StartFragment:0000021546
                            //EndFragment:0000021589
                            //SourceURL:file:///C:\Users\Sheng%20Bi\Desktop\INSC544\Assignments\bisheng.hw5.answer.docx
                            //
                            //<html......>
                            //......

                            const string word_symbol = @"SourceURL:file:///";
                            if (htmlText.Contains(word_symbol))
                            {
                                int url_index = htmlText.IndexOf(word_symbol) + word_symbol.Length;
                                int url_length = htmlText.IndexOf("\r\n\r\n<html") - url_index;
                                string url = htmlText.Substring(url_index, url_length);
                                url = url.Replace("%20", " ");

                                elementControl.AddAssociation(dropElement, url, ElementAssociationType.FileShortcut, srcText);
                            }
                        }

                        #endregion
                    }
                    else if (e.Data.GetDataPresent("FileGroupDescriptor"))
                    {
                        #region Web

                        string url = (string)e.Data.GetData(DataFormats.StringFormat);

                        elementControl.AddAssociation(dropElement, url, ElementAssociationType.Web, null);

                        #endregion
                    }
                    else if (e.Data.GetDataPresent("HTML Format"))
                    {
                        #region FireFox Text

                        string htmlText = e.Data.GetData("HTML Format").ToString();
                        string srcText = e.Data.GetData("UnicodeText").ToString();

                        srcText = srcText.Replace("\r\n", " ").Replace("\t", " ").Trim();
                        if (srcText.Length > StartProcess.MAX_EXTRACTTEXT_LENGTH)
                            srcText = srcText.Substring(0, StartProcess.MAX_EXTRACTTEXT_LENGTH) + "...";

                        // FireFox
                        // Sample htmlText = 
                        //Version:0.9
                        //StartHTML:00000186
                        //EndHTML:00000295
                        //StartFragment:00000220
                        //EndFragment:00000259
                        //SourceURL:http://www.google.com/firefox?client=firefox-a&rls=org.mozilla:en-US:official
                        //<html><body>
                        //<!--StartFragment--><font size="-1">Change the color</font><!--EndFragment-->
                        //</body>
                        //</html>
                        const string ff_symbol = "SourceURL:";
                        if (htmlText.Contains(ff_symbol))
                        {
                            int url_index = htmlText.IndexOf(ff_symbol) + ff_symbol.Length;
                            int url_length = htmlText.IndexOf("\r\n<html>") - url_index;
                            string url = htmlText.Substring(url_index, url_length);

                            elementControl.AddAssociation(dropElement, url, ElementAssociationType.Web, srcText);
                        }

                        #endregion
                    }

                    GetFocusToElementTextBox(elementControl.CurrentElement, elementControl.CurrentElement.NoteText.Length, false, false);
                    LogControl.Write(
                        elementControl.CurrentElement,
                        LogEventAccess.AppWindow,
                        LogEventType.DragLink,
                        LogEventStatus.End,
                        null);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops!\nSomething unexpected happened, please close this Planz window and reopen.");
                //MessageBox.Show("ElementStackPanel_PreviewDrop\n" + ex.Message);

                LogControl.Write(
                    dropElement,
                    LogEventAccess.AppWindow,
                    LogEventType.DragLink,
                    LogEventStatus.Error,
                    LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }
            finally
            {
                e.Handled = true;
            }
        }

        #endregion

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {

            if (sender != null)
            {
                string eventInfo = LogEventInfo.Version + LogControl.COMMA + StartProcess.PLANZ_VERSION + LogControl.DELIMITER
                                    + LogEventInfo.Build + LogControl.COMMA + StartProcess.BUILD_VERSION + LogControl.DELIMITER
                                    + LogEventInfo.Date + LogControl.COMMA + StartProcess.BUILD_DATE;

                LogControl.Write(
                    elementControl.CurrentElement,
                    LogEventAccess.AppWindow,
                    LogEventType.Launch,
                    LogEventStatus.NULL,
                    eventInfo);

            }
        }

    }
}
