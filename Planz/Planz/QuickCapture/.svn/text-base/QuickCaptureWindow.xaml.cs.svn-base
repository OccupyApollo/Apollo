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
* (a copy is included in the Planz\LICENSE.txt file that accompanied this code).
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
using System.Collections.Specialized;
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
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;
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
using GlobalTextSelHook;
using GlobalKeyboardHook;
using Planz;


namespace QuickCapture
{
    /// <summary>
    /// Interaction logic for Window1.xaml
    /// </summary>
    public partial class QuickCaptureWindow : Window
    {
        private string startPath = StartProcess.START_PATH;
        private InfoItem newInfoItem = null;
        private string newNoteText = String.Empty;
        private const int ACTIVEITEM_TITLE_LENGTH = 50;
        private String xmlFileFullDir = String.Empty;
        private bool bNoteModified = false;
        private bool bSkipModifyConfirmation = false;
        private const int MAX_SELECTEDLOC_COUNT = 9;
        private Queue selectedLocCache = new Queue(10);
        private ComboBoxItem selectedItem = null;
        private ComboBoxItem projectItem = null;
        private ComboBoxItem journalItem = null;
        private int posSel = 0;
        private int posJournal = 0;
        private int posProject = 0;

        private KeyboardHandler hotkey_handler;

        public QuickCaptureWindow()
        {
            InitializeComponent();

            ReInitilize();

            if (Directory.Exists(StartProcess.APPDATA_PLANZ_PATH) == false)
            {
                Directory.CreateDirectory(StartProcess.APPDATA_PLANZ_PATH);
            }

            textBox_Note.Text = newNoteText;
            this.Activate();
            textBox_Note.Select(newNoteText.Length, 0);
            textBox_Note.Focus();

            checkBox_URI.Content = "No link available";

            hotkey_handler = new KeyboardHandler(this);
            updateComboBox_SaveLoc();
            if(comboBox_SaveLoc.Items.Count>0 && comboBox_SaveLoc.SelectedIndex ==-1)
                comboBox_SaveLoc.SelectedIndex = 0;
            this.Title = "Quick Capture (Windows logo key + C)";                                                                                                                                                                                                                                                                
        }

        private void button_OK_Click(object sender, RoutedEventArgs e)
        {
           // return;
            // Click to save the note
            //none of URI or note is available, there is nothing to save

            if (newNoteText.Length == 0 && newInfoItem == null)
                return;

            bNoteModified = false;
            bSkipModifyConfirmation = false;

            //use the title as note, if only URI is available
            if (newInfoItem!=null && newNoteText.Length == 0 )
                newNoteText = newInfoItem.Title;

            ComboBoxItem cbi_sel = (ComboBoxItem) comboBox_SaveLoc.SelectedItem;

            string xmlFileFullPath = cbi_sel.Tag.ToString();
            
            Element parentElement = new Element
            {
                ParentElement = null,
                HeadImageSource = String.Empty,
                TailImageSource = String.Empty,
                NoteText = String.Empty,
                IsExpanded = true,
                Path = xmlFileFullPath,
                Type = ElementType.Heading,
            };

            Element newElement = new Element
            {
                ParentElement = parentElement,
                HeadImageSource = String.Empty,
                TailImageSource = String.Empty,
                NoteText = newNoteText,
                IsExpanded = false,
                Path = xmlFileFullPath,
                Type = ElementType.Note,
                FontColor = ElementColor.Blue.ToString(),
                Status = ElementStatus.New,
            };

            newElement.ParentElement = parentElement;
            newElement.Position = 0;

            if ((newInfoItem != null) && (bool)checkBox_URI.IsChecked)
            {
                ElementAssociationType newType;
                switch (newInfoItem.Type)
                {
                    case InfoItemType.Email:
                        newType = ElementAssociationType.Email;
                        break;
                    case InfoItemType.File:
                        newType = ElementAssociationType.FileShortcut;
                        break;
                    case InfoItemType.Web:
                        newType = ElementAssociationType.Web;
                        break;
                    default:
                        newType = ElementAssociationType.None;
                        break;
                }
                newElement.AssociationType = newType;
            }
            

            try
            {
                newElement.ParentElement.Elements.Insert(0, newElement);

                DatabaseControl temp_dbControl = new DatabaseControl(newElement.ParentElement.Path);
                temp_dbControl.OpenConnection();
                temp_dbControl.InsertElementIntoXML(newElement);
                temp_dbControl.CloseConnection();

                ElementControl elementControl = new ElementControl(newElement.ParentElement.Path);
                elementControl.CurrentElement = newElement;

                //if URI is available and selected, association will be added together with the note
                if ((newInfoItem!= null) && (bool)checkBox_URI.IsChecked)
                {
                    elementControl.AddAssociation(newElement, newInfoItem.Uri, newElement.AssociationType, newNoteText);
                }
                
                string eventInfo = LogEventInfo.NoteText + LogControl.COMMA + newElement.NoteText;

                if ((bool)checkBox_URI.IsChecked)
                    eventInfo += LogControl.DELIMITER + LogEventInfo.LinkStatus + LogControl.COMMA + "check";
                else
                    eventInfo += LogControl.DELIMITER + LogEventInfo.LinkStatus + LogControl.COMMA + "unCheck";

                if(newInfoItem!=null && newInfoItem.Uri!=null)
                    eventInfo += LogControl.DELIMITER + LogEventInfo.LinkedFile + LogControl.COMMA + newInfoItem.Uri;

                eventInfo += LogControl.DELIMITER + LogEventInfo.PutUnder + LogControl.COMMA + newElement.Path;

                LogControl.Write(
                elementControl.CurrentElement,
                LogEventAccess.QuickCapture,
                LogEventType.CreateNewNote,
                LogEventStatus.NULL,
                eventInfo);

                newInfoItem = null;

            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("QuickCapture_button_OK_Click\n" + ex.Message);

                LogControl.Write(
                   newElement,
                   LogEventAccess.QuickCapture,
                   LogEventType.CreateNewNote,
                   LogEventStatus.Error,
                   LogEventInfo.ErrorMessage + LogControl.COMMA + ex.Message);
            }

            ReInitilize();

            this.Visibility = Visibility.Hidden;
            this.ShowInTaskbar = true;
        }

        private void Window_Deactivated(object sender, EventArgs e)
        {
            if (bSkipModifyConfirmation == false)
            {
                System.Windows.MessageBoxButton buttons = System.Windows.MessageBoxButton.YesNo;
                string message = "Are you sure to leave Quick Capture? Your note will be discarded.";
                string caption = "";
                if (bNoteModified == true && (System.Windows.MessageBox.Show(message, caption, buttons, System.Windows.MessageBoxImage.Warning, System.Windows.MessageBoxResult.No) == System.Windows.MessageBoxResult.No))
                {
                    this.Visibility = Visibility.Visible;
                    this.Focus();
                    this.Topmost = true;
                    return;   
                }
                else
                {
                    if (IsSingleInstance())
                        this.Visibility = Visibility.Hidden;
                    else
                        System.Diagnostics.Process.GetCurrentProcess().Kill();
                }
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            this.WindowStartupLocation = WindowStartupLocation.Manual;
            this.Left = SystemParameters.WorkArea.Width - this.ActualWidth - this.Margin.Right;
            this.Top = SystemParameters.WorkArea.Height - this.ActualHeight - this.Margin.Bottom;
            this.Visibility = Visibility.Visible;
        }

        public void On_Activate()
        {
            ActiveWindow aw = new ActiveWindow();
            TextSelHandler textsel = new TextSelHandler(this);
            
            IntPtr hWnd = IntPtr.Zero;
            InfoItem activeItem = null;
            System.Windows.Clipboard.Clear();

            try
            {
                //support for synchronizing droplist with filesystem dynamically 
                

                hWnd = aw.GetActiveWindowHandle();

                textsel.SendCtrlC(hWnd);

                //wait for the copy to be completed
                Thread.Sleep(100);

                activeItem = aw.GetActiveItem();

                newNoteText = System.Windows.Clipboard.GetText().Replace("\r\n", " ").Replace("\t", " ").Trim();
                if (newNoteText.Length > StartProcess.MAX_EXTRACTTEXT_LENGTH)
                    newNoteText = newNoteText.Substring(0, StartProcess.MAX_EXTRACTTEXT_LENGTH) + "...";

                if ((activeItem != null) && (activeItem.Uri!=null) && (activeItem.Uri.Length!=0))
                {
                    newInfoItem = activeItem;

                    if (newNoteText.Length != 0)
                        newNoteText = "\"" + newNoteText + "\" ";
                    else if(newInfoItem.Title!=null)
                        newNoteText = "\"" + newInfoItem.Title + "\" ";

                    string title;
                    if (newInfoItem.Title == null || newInfoItem.Title.Length == 0)
                    {
                        title = newNoteText.Replace("/", "").Replace("\\", "").Replace("*", "").Replace("?", "").Replace("\"", "").Replace("<", "").Replace(">", "").Replace("|", "").Replace(":", "");
                        if (title.Length > StartProcess.MAX_EXTRACTNAME_LENGTH)
                            title = title.Substring(0, StartProcess.MAX_EXTRACTNAME_LENGTH);
                    }
                    else
                        title = newInfoItem.Title;

                    checkBox_URI.IsEnabled = true;
                    checkBox_URI.IsChecked = true;

                    if (title.Length > ACTIVEITEM_TITLE_LENGTH)
                        checkBox_URI.Content = "Include link to: " + "\"" + title.Substring(0, ACTIVEITEM_TITLE_LENGTH) + "..." + "\"";
                    else
                        checkBox_URI.Content = "Include link to: " + "\"" + title + "\"";
                }
                textBox_Note.Text = newNoteText;
            }
            catch (Exception)
            {
                ReInitilize();
            }

            updateComboBox_SaveLoc();
            int selectedIndex = 0;
            int i;
            for (i = 0; i < comboBox_SaveLoc.Items.Count; i++)
            {
                ComboBoxItem cbi_Item = (ComboBoxItem)comboBox_SaveLoc.Items[i];
                if (cbi_Item.Tag.ToString() == selectedItem.Tag.ToString() && cbi_Item.Content.ToString() == selectedItem.Content.ToString())
                    selectedIndex = i;
            }
            comboBox_SaveLoc.SelectedIndex = selectedIndex;

            this.Visibility = Visibility.Visible;
            this.Activate();
            textBox_Note.Focus();
            textBox_Note.Select(newNoteText.Length, 0);
            bNoteModified = false;
            this.Topmost = true;
            System.Windows.Clipboard.Clear();


            aw = null;
            textsel = null;
            hWnd = IntPtr.Zero;
            activeItem = null;
            GC.Collect();
        }

        private void checkBox_URI_Checked(object sender, RoutedEventArgs e)
        {

        }

        private void textBox_Note_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (newNoteText != textBox_Note.Text)
            {
                newNoteText = textBox_Note.Text;
                bNoteModified = true;
            }
        }
  
        private void comboBox_SaveLoc_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ComboBoxItem cbi_Sel = (ComboBoxItem) comboBox_SaveLoc.SelectedItem;
            int selectedIndex = comboBox_SaveLoc.SelectedIndex;

            if (cbi_Sel == null)
                return;

            string new_loc = null;
            if (projectItem!=null && cbi_Sel.Tag.ToString() == projectItem.Tag.ToString() && selectedIndex == posProject)
            {
                System.Windows.Forms.FolderBrowserDialog browse = new System.Windows.Forms.FolderBrowserDialog();
                browse.ShowNewFolderButton = false;
                browse.Description = "Please select a heading";
                //
                bSkipModifyConfirmation = true;
                browse.RootFolder = Environment.SpecialFolder.Desktop;
                browse.SelectedPath = StartProcess.START_PATH;
                browse.ShowNewFolderButton = true;

                if (browse.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    string path = browse.SelectedPath;
                    new_loc = path + System.IO.Path.DirectorySeparatorChar;
                }
                else
                    comboBox_SaveLoc.SelectedIndex = 0;
                bSkipModifyConfirmation = false;
            }

            if (journalItem!=null && cbi_Sel.Tag.ToString() == journalItem.Tag.ToString() && selectedIndex == posJournal)
            {
                bSkipModifyConfirmation = true;
                DatePickerWindow_QC dpw = new DatePickerWindow_QC();
                if (dpw.ShowDialog().Value == true)
                {
                    DateTime selDate = dpw.SelectedDate;
                    string date_path = JournalControl.GetJournalPath(selDate);
                    if(date_path!=null)
                       new_loc = date_path +System.IO.Path.DirectorySeparatorChar;
                    else
                    {
                        System.Windows.MessageBox.Show("Selected date folder doesn't exist!");
                        comboBox_SaveLoc.SelectedIndex = 0;
                    }
                }
                else
                    comboBox_SaveLoc.SelectedIndex = 0;
                bSkipModifyConfirmation = false;
            }

            if (new_loc != null)
            {
                bool bUpdate = false;
                for (int i = 0; i < posProject; i++)
                {
                    ComboBoxItem cbi = (ComboBoxItem)comboBox_SaveLoc.Items[i];
                    if (cbi.Tag.ToString() == new_loc)
                    {
                        comboBox_SaveLoc.SelectedIndex = i;
                        bUpdate = true;
                    }
                }

                if (bUpdate == false && !selectedLocCache.Contains(new_loc))
                {
                    selectedLocCache.Enqueue(new_loc);
                    
                    while (selectedLocCache.Count > MAX_SELECTEDLOC_COUNT)
                        selectedLocCache.Dequeue();
                    updateComboBox_SaveLoc();
                    comboBox_SaveLoc.SelectedIndex = posSel;
                }

                bSkipModifyConfirmation = false;
                this.Visibility = Visibility.Visible;
                this.Focus();
                this.Topmost = true;
            }
            selectedItem = (ComboBoxItem) comboBox_SaveLoc.SelectedItem;
        }

        private void updateComboBox_SaveLoc()
        {
            comboBox_SaveLoc.Items.Clear();
            GC.Collect();

            char[] ds = { '\\' };

            int count = 0;
            //add note_path
            if(Directory.Exists(StartProcess.INBOX_PATH))
            {
                ComboBoxItem cbi_Notes = new ComboBoxItem();
                cbi_Notes.Tag = StartProcess.INBOX_PATH;
                cbi_Notes.Content = System.IO.Path.GetFileName(StartProcess.INBOX_PATH.TrimEnd(ds));
                comboBox_SaveLoc.Items.Add(cbi_Notes);
                count++;
            }


            //add Today+
            if(Directory.Exists(StartProcess.TODAY_PLUS_PATH))
            {
                ComboBoxItem cbi_Today = new ComboBoxItem();
                cbi_Today.Tag = StartProcess.TODAY_PLUS_PATH;
                cbi_Today.Content = System.IO.Path.GetFileName(StartProcess.TODAY_PLUS.TrimEnd(ds));
                comboBox_SaveLoc.Items.Add(cbi_Today);
                count++;
            }

            ComboBoxItem cbi_Desktop = new ComboBoxItem();
            cbi_Desktop.Tag = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + System.IO.Path.DirectorySeparatorChar;
            cbi_Desktop.Content = Environment.SpecialFolder.Desktop.ToString();
            comboBox_SaveLoc.Items.Add(cbi_Desktop);
            count++;

            posSel = count;
            foreach (string loc in selectedLocCache)
            {
                if(Directory.Exists(loc))
                {
                    ComboBoxItem cbi_loc = new ComboBoxItem();
                    cbi_loc.Tag = loc;
                    cbi_loc.Content = System.IO.Path.GetFileName(loc.TrimEnd(ds));
                    comboBox_SaveLoc.Items.Insert(posSel,cbi_loc);
                    count++;
                }
            }

            
            //add Project_path
            posProject = count;
            ComboBoxItem cbi_Project = new ComboBoxItem();
            cbi_Project.Tag = StartProcess.START_PATH;
            cbi_Project.Content = "Select a Location...";
            comboBox_SaveLoc.Items.Add(cbi_Project);
            count++;
            projectItem = cbi_Project;


            //add Journal_path
            posJournal = count;
            string thisYear = StartProcess.JOURNAL_PATH + DateTime.Now.Year.ToString() + System.IO.Path.DirectorySeparatorChar;
            if (Directory.Exists(StartProcess.JOURNAL_PATH) && Directory.Exists(thisYear))
            {
                ComboBoxItem cbi_Journal = new ComboBoxItem();
                cbi_Journal.Tag = StartProcess.JOURNAL_PATH;
                cbi_Journal.Content = "Select by Date...";
                comboBox_SaveLoc.Items.Add(cbi_Journal);
                count++;
                journalItem = cbi_Journal;
            }
            
            
        }

        private void ReInitilize()
        {
            newNoteText = "";
            textBox_Note.Text = newNoteText;

            checkBox_URI.Content = "No link available";
            checkBox_URI.IsEnabled = false;
            checkBox_URI.IsChecked = false;
        }

        protected override void OnClosing(CancelEventArgs e)
        {   
            base.OnClosing(e);
            e.Cancel = true;
            this.Visibility = Visibility.Hidden;
        }

        private bool IsSingleInstance()
        {
            int qcProcCnt = Process.GetProcessesByName(ProcessList.QuickCapture.ToString()).Count();

            if (qcProcCnt > 1)
                return false;
            return true;
        }
    }

    public class CustomTimeSpan
    {
        public int Hour
        {
            get;
            set;
        }

        public int Minutes
        {
            get;
            set;
        }

        public bool IsAM
        {
            get;
            set;
        }

        public override string ToString()
        {
            string returnString = String.Empty;
            returnString += Hour.ToString();
            returnString += ":";
            if (Minutes >= 10)
            {
                returnString += Minutes.ToString();
            }
            else
            {
                returnString += "0" + Minutes.ToString();
            }
            if (IsAM)
            {
                returnString += " AM";
            }
            else
            {
                returnString += " PM";
            }
            return returnString;
        }
    }
}
