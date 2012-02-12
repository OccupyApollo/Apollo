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
using IWshRuntimeLibrary;

namespace Planz
{
    public enum LogEventAccess
    {
        AppWindow, // Launch and close
        QuickAccessToolbar,
        AppMenu,
        Ribbon,
        NavigationTree,
        NoteTextBox,
        NoteTextContextMenu,
        AssociationIcon,
        AssociationContextMenu,
        NodeIcon, // Both expander and node icons
        NodeContextMenu,
        FlagIcon,
        Hotkey,
        QuickCapture,
    }

    public enum LogEventType
    {
        Launch, // open Planz
        Close, // close Planz
        SaveAsText,
        SaveAsHTML,
        ExportStructureOnly,
        ExportStructureAndContent,
        ExportLog,
        CreateJournalFolder,
        Options,
        Refresh,
        Expand,
        Collapse,
        Promote,
        Demote,
        MoveUp,
        MoveDown,
        Delete,
        Search,
        CreateNewHeading,
        CreateNewNote,
        CreateNewTextDocument,
        CreateNewWordDocument,
        CreateNewExcelSpreadsheet,
        CreateNewPowerPointPresentation,
        CreateNewOneNoteSection,
        CreateNewOutlookEmail,
        CreateNewTwitterUpdate,
        LinkToActiveItem,
        LinkToFile,
        LinkToFolder,
        ChangeFontColor,
        ShowNavigationTree,
        HideNavigationTree,
        ShowOutline,
        HideOutline,
        ClickNavigationItem,
        EnableSpellCheck,
        DisableSpellCheck,
        Select,
        SelectAll,
        Open,
        Explore,
        Rename,
        NewVersion,
        ReplyToAll,
        Flag,
        Check,
        Uncheck,
        OK, // dialog OK
        Cancel, // dialog Cancel
        RenameHeading,
        PowerDDo,
        PowerDDone,
        PowerDDelegate,
        PowerDDefer,
        PowerDDelete,
        PowerDPrevious,
        PowerDNext,
        PowerDCheckShowAssociationMarkedDone,
        PowerDUncheckShowAssociationMarkedDone,
        PowerDCheckShowAssociationMarkedDefer,
        PowerDUncheckShowAssociationMarkedDefer,
        Import, //import appointments from outlook calendar
        RelatedMessages, //display conversation in outlook
        NewWindow,
        CommandNote,
        LabelWith,
        ChangeFocus,
        Hide,
        Feedback,
        AboutWindow,
        UserManual,
        BackKeyPress,
        TextChanged,
        DragLink,
    }

    public enum LogEventStatus
    {
        Begin,
        End,
        NULL,
        Cancel,
        Error,
    }

    public enum LogEventInfo
    {
        ErrorMessage,
        SavePath,
        ExportPath,
        YearCreated,
        HeadingFont,
        HeadingSize,
        NoteFont,
        NoteSize,
        DeleteType,
        FontColor,
        SearchText,
        SearchResult,
        PreviousName,
        NewName,
        LinkedFile,
        LinkedFolder,
        RemoveFromToday,
        HasStart,
        StartDate,
        StartTime,
        StartAllDay,
        HasDue,
        DueDate,
        DueTime,
        DueAllDay,
        AddToToday,
        AddToReminder,
        AddToTask,
        Command,
        Location,
        FileName,
        Date,
        NoteText,
        LinkStatus,
        Version,
        Build,
        PutUnder,
    }

    public class LogControl
    {
        public const string LOG_FILENAME = "PlanzLog.txt";
        private static string logPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + 
            System.IO.Path.DirectorySeparatorChar + StartProcess.APPDATA_FOLDERNAME + System.IO.Path.DirectorySeparatorChar 
            + LOG_FILENAME;
        private static System.IO.StreamWriter logWriter;

        public const string COMMA = ":";
        public const string DELIMITER = "|";

        public static void Initialize()
        {
            if (System.IO.File.Exists(logPath) == false)
            {
                logWriter = new StreamWriter(System.IO.File.Create(logPath));
            }
            else
            {
                logWriter = System.IO.File.AppendText(logPath);
            }  
        }

        public static void Close()
        {
            logWriter.Flush();
            logWriter.Close();
        }

        public static string LogFileFullName
        {
            get { return logPath; }
        }

        public static void Write(
            Element element, 
            LogEventAccess eventAccess,
            LogEventType eventType,
            LogEventStatus eventStatus,
            String eventInfo)
        {
            Initialize();

            try
            {
                //lock (logWriter)
                //{
                    string logEntry = String.Empty;

                    const string SPACE = "\t";

                    // Date Time
                    var unixTime = DateTime.Now.ToUniversalTime() - new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
                    logEntry += System.Convert.ToInt64(unixTime.TotalSeconds) + SPACE
                        + DateTime.Now.ToShortDateString() + " "
                        + DateTime.Now.ToLongTimeString() + SPACE;

                    // Log Op
                    logEntry += eventAccess.ToString() + SPACE +
                        eventType.ToString() + SPACE +
                        eventStatus.ToString() + SPACE;
                    
                    // Element 
                    if (element != null)
                    {
                        logEntry +=
                            element.ID.ToString() + SPACE +
                            element.Path.ToString() + SPACE +
                            element.Type.ToString() + SPACE +
                            element.Position.ToString() + SPACE +
                            element.NoteText.Substring(0, Math.Min(element.NoteText.Length, 50)).Replace("\r\n", "#EOL") + SPACE +
                            element.AssociationType.ToString() + SPACE;

                        String url = String.Empty;
                        try
                        {
                            switch (element.AssociationType)
                            {
                                case ElementAssociationType.Email:
                                    url = element.AssociationURIFullPath;
                                    break;
                                case ElementAssociationType.File:
                                    url = "file:///" + element.AssociationURIFullPath;
                                    break;
                                case ElementAssociationType.FileShortcut:
                                    url = "file:///" + GetShortcut(element).TargetPath;
                                    break;
                                case ElementAssociationType.FolderShortcut:
                                    url = "file:///" + GetShortcut(element).TargetPath;
                                    break;
                                case ElementAssociationType.Tweet:
                                    url = GetShortcut(element).Description;
                                    break;
                                case ElementAssociationType.Web:
                                    url = GetShortcut(element).Description;
                                    break;
                            }
                        }
                        catch (Exception)
                        {
                        }

                        logEntry += url + SPACE;
                    }

                    // Log Comment
                    logEntry += eventInfo;

                    if (logWriter != null)
                    {
                        logWriter.WriteLine(logEntry);
                        logWriter.Flush();
                    }
                //}
            }
            catch (Exception ex)
            {

            }
            Close();
        }

        public static IWshRuntimeLibrary.IWshShortcut GetShortcut(Element element)
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
    }
}
