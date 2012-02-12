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

namespace Planz
{
    public enum ElementType
    {
        Heading = 0,
        Note,
    }

    public enum FlagStatus
    {
        Normal,
        Flag,
        Check,
    }

    public enum ElementStatus
    {
        Normal = 0,
        New,
        MissingAssociation,
        TempStatus_HasEmailAttachment,
        Special,
        ToBeDeleted,
    }

    public enum PowerDStatus
    {
        None,
        Done,
        Deferred,
        Delegated,
        ToBeDeleted,
    }

    public enum ElementAssociationType
    {
        None = 0,
        File,
        FileShortcut,
        Folder,
        FolderShortcut,
        Web,
        Email,
        Tweet,
        Missing,
        Appointment,
    }

    public enum ICCAssociationType
    {
        NotepadTextDocument = 0,
        WordDocument,
        ExcelSpreadsheet,
        PowerPointPresentation,
        OneNoteSection,
        OutlookEmailMessage,
        TwitterUpdate,
    }

    public enum ViewMode
    {
        Outline = 0,
        Document,
    }

    public enum DeleteMessageType
    {
        NoteWithoutAssociation = 0,
        NoteWithFileAssociation,
        NoteWithShortcutAssociation,
        HeadingWithoutChildren,
        HeadingWithChildren,
        HeadingWithoutNoteText,
        InplaceExpansionHeading,
        HeadingWithoutChildrenOrContent,
        Shortcut,
        File,
        MultipleItems,
        Default,
    }

    public enum ButtonStyle
    {
        YESNO = 0,
        OKCANCEL = 1,
    }

    public enum FileSystemSyncMessage
    {
        HighPriorityWorkFinished = 0,
        LowPriorityWorkFinished,
        NavigationTreeExpanded,
    }

    public enum InfoItemType
    {
        Web = 0,
        File,
        Email,
    }

    public enum ProcessList
    {
        WINWORD = 0, 
        EXCEL,
        POWERPNT,
        OUTLOOK,
        iexplore,
        AcroRd32,
        firefox,
        QuickCapture,
        Planz,
    }

    public enum ElementColor
    {
        Red = 0,
        Black,
        Blue,
        Gray,
        DarkBlue,
        Purple,
        SeaGreen,
        SpringGreen,
        Plum,
    }

    public enum PowerDDeleteType
    {
        Delete = 0,
        Undo,
    }

    public enum ElementCommand
    {
        None,
        ShowJournalInNewWindow,
        ImportMeetingsAndAppointmentsFromOutlook,
        DisplayMoreAssociations,
        ShowOrHideDeferAssociations,
        ShowOrHideDoneAssociations,
        ShowHiddenAssociations,
    }
}
