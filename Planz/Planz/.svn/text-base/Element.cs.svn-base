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
using System.Net;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Collections.ObjectModel;
using System.ComponentModel;

namespace Planz
{
    public class Element : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private void NotifyPropertyChanged(String propertyName)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        protected virtual void OnPropertyChanged(PropertyChangedEventArgs e)
        {
            if (this.PropertyChanged != null)
            {
                this.PropertyChanged(this, e);
            }
        }

        private ObservableCollection<Element> children = new ObservableCollection<Element>();

        private Guid Id;
        private ElementType type = ElementType.Note;
        private bool isHeading = false;
        private bool isNote = true;
        private string noteText = String.Empty;
        private ElementStatus status = ElementStatus.Normal;
        private PowerDStatus powerdStatus = PowerDStatus.None;
        private FlagStatus flagStatus = FlagStatus.Normal;
        private bool isExpanded = false;
        private string path = String.Empty;
        private int position = 0;
        private ElementAssociationType associationType = ElementAssociationType.None;
        private string associationURI = String.Empty;
        private int levelOfSynchronization = 0;
        private string fontColor = ElementColor.Black.ToString();
        private string tag = String.Empty;
        private string description = String.Empty;
        private Visibility isVisible = Visibility.Visible;
        private DateTime createdOn = DateTime.Now;
        private DateTime modifiedOn = DateTime.Now;
        private DateTime powerdTimeStamp = DateTime.Now;
        private DateTime startDate = new DateTime();
        private DateTime dueDate = new DateTime();
        private string associationFragment = String.Empty;
        private string tempData = String.Empty;
        private ElementCommand command = ElementCommand.None;

        private string headImageSource = HeadImageControl.Note_Empty;
        private string flagImageSource = String.Empty;
        private string tailImageSource = String.Empty;
        private bool hasAssociation = false;
        private bool canOpen = true;
        private bool canExplore = true;
        private bool canRename = true;
        private bool canDelete = true;
        private bool canCreateNewVersion = false;
        private bool canImportCalendar = false;
        private bool canLabelWith = false;
        private bool hasEmailAssociation = false;
        private bool canExecute = false;
        private bool showAssociationMarkedDone = true;
        private bool showAssociationMarkedDefer = true;
        private Visibility showExpander = Visibility.Hidden;
        private Visibility showFlag = Visibility.Hidden;
        private double tailImageWidth = Properties.Settings.Default.TailImageHeight;
        private double tailImageHeight = Properties.Settings.Default.TailImageHeight;

        public Element()
        {
            Id = Guid.NewGuid();

            Elements = new ObservableCollection<Element>();

            if (Properties.Settings.Default.ShowOutline)
            {
                ShowFlag = Visibility.Visible;
                if (this.isHeading)
                {
                    ShowExpander = Visibility.Visible;
                }
                else
                {
                    HeadImageSource = HeadImageControl.Note_Unselected;
                }
            }
        }

        #region Property

        public Guid ID
        {
            get
            {
                return Id;
            }
            set
            {
                Id = value;
            }
        }

        public ElementType Type
        {
            get
            {
                return this.type;
            }
            set
            {
                this.type = value;

                if (this.type == ElementType.Heading)
                {
                    IsHeading = true;
                    IsNote = false;

                    if (this.fontColor == ElementColor.Black.ToString())
                    {
                        FontColor = ElementColor.DarkBlue.ToString();
                    }

                    this.showAssociationMarkedDefer = false;
                    this.showAssociationMarkedDone = false;
                }
                else if (this.type == ElementType.Note)
                {
                    IsHeading = false;
                    IsNote = true;

                    if (this.fontColor == ElementColor.DarkBlue.ToString())
                    {
                        FontColor = ElementColor.Black.ToString();
                    }

                    this.showAssociationMarkedDefer = true;
                    this.showAssociationMarkedDone = true;
                }
            }
        }

        public int Position
        {
            get
            {
                if (this.ParentElement != null)
                {
                    return this.ParentElement.Elements.IndexOf(this);
                }
                else
                {
                    return this.position;
                }
            }
            set
            {
                this.position = value;
            }
        }

        public int Level
        {
            get
            {
                int level = 0;
                Element curr = this;
                while (curr.ParentElement != null)
                {
                    curr = curr.ParentElement;
                    level++;
                }
                return level;
            }
        }

        public string NoteText
        {
            get
            {
                return this.noteText;
            }
            set
            {
                this.noteText = value;
                this.OnPropertyChanged(new PropertyChangedEventArgs("NoteText"));
            }
        }

        public bool IsExpanded
        {
            get
            {
                return this.isExpanded;
            }
            set
            {
                this.isExpanded = value;
                OnPropertyChanged(new PropertyChangedEventArgs("IsExpanded"));
            }
        }

        public bool IsCollapsed
        {
            get { return !IsExpanded; }
        }

        public ElementStatus Status
        {
            get
            {
                return this.status;
            }
            set
            {
                this.status = value;
            }
        }

        public FlagStatus FlagStatus
        {
            get
            {
                return this.flagStatus;
            }
            set
            {
                this.flagStatus = value;
            }
        }

        public PowerDStatus PowerDStatus
        {
            get
            {
                return this.powerdStatus;
            }
            set
            {
                this.powerdStatus = value;
            }
        }

        // Data binding for Expander.Visibility
        public bool IsHeading
        {
            get
            {
                return this.isHeading;
            }
            set
            {
                this.isHeading = value;
                OnPropertyChanged(new PropertyChangedEventArgs("IsHeading"));
            }
        }

        // Data binding for HeadImage.Visibility
        public bool IsNote
        {
            get
            {
                return this.isNote;
            }
            set
            {
                this.isNote = value;
                OnPropertyChanged(new PropertyChangedEventArgs("IsNote"));
            }
        }

        public bool IsLocalHeading
        {
            get
            {
                if (this.type == ElementType.Heading && this.associationType != ElementAssociationType.FolderShortcut)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
        }

        public bool IsRemoteHeading
        {
            get
            {
                if (this.type == ElementType.Heading && this.associationType == ElementAssociationType.FolderShortcut)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
        }

        public bool IsCommandNote
        {
            get
            {
                if (this.command == ElementCommand.None)
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
        }

        // Path ends with System.IO.Path.DirectorySeparatorChar
        public string Path
        {
            get
            {
                if (this.type == (int)ElementType.Heading)
                {
                    return this.path;
                }
                else
                {
                    return this.ParentElement.Path;
                }
            }
            set
            {
                this.path = value;
                if (path.Contains(StartProcess.JOURNAL_PATH))
                {
                    canImportCalendar = true;
                    try
                    {
                        JournalControl.GetDateTime(Path);
                    }
                    catch
                    {
                        canImportCalendar = false;
                    }
                }

            }
        }

        public bool HasChildren
        {
            get { return (Elements.Count > 0) ? true : false; }
        }

        public ElementAssociationType AssociationType
        {
            get
            {
                return this.associationType;
            }
            set
            {
                this.associationType = value;
                switch (this.associationType)
                {
                    case ElementAssociationType.Email:
                        HasEmailAssociation = true;
                        break;
                    case ElementAssociationType.File:
                        CanCreateNewVersion = true;
                        break;
                };
            }
        }

        // Relative path
        public string AssociationURI
        {
            get
            {
                return this.associationURI;
            }
            set
            {
                this.associationURI = value;
                if (value == String.Empty)
                {
                    HasAssociation = false;
                }
                else
                {
                    HasAssociation = true;
                }
            }
        }

        public string AssociationURIFullPath
        {
            get
            {
                if (IsLocalHeading)
                {
                    return Path;
                }
                else if (IsRemoteHeading)
                {
                    return this.ParentElement.Path + associationURI;
                }
                else
                {
                    return Path + AssociationURI;
                }
            }
        }

        public string Tag
        {
            get
            {
                return this.tag;
            }
            set
            {
                this.tag = value;
            }
        }

        public string Description
        {
            get
            {
                return this.description;
            }
            set
            {
                this.description = value;
            }
        }

        public int LevelOfSynchronization
        {
            get
            {
                return this.levelOfSynchronization;
            }
            set
            {
                this.levelOfSynchronization = value;
            }
        }

        public DateTime CreatedOn
        {
            get
            {
                return this.createdOn;
            }
            set
            {
                this.createdOn = value;
            }
        }

        public DateTime ModifiedOn
        {
            get
            {
                return this.modifiedOn;
            }
            set
            {
                this.modifiedOn = value;
            }
        }

        public DateTime PowerDTimeStamp
        {
            get
            {
                return this.powerdTimeStamp;
            }
            set
            {
                this.powerdTimeStamp = value;
            }
        }

        public DateTime StartDate
        {
            get
            {
                return this.startDate;
            }
            set
            {
                this.startDate = value;
            }
        }

        public DateTime DueDate
        {
            get
            {
                return this.dueDate;
            }
            set
            {
                this.dueDate = value;
            }
        }

        public string AssociationFragment
        {
            get
            {
                return this.associationFragment;
            }
            set
            {
                this.associationFragment = value;
            }
        }

        public string TempData
        {
            get
            {
                return this.tempData;
            }
            set
            {
                this.tempData = value;
            }
        }

        public ElementCommand Command
        {
            get
            {
                return this.command;
            }
            set
            {
                this.command = value;
                if (this.command == ElementCommand.None)
                {
                    this.canExecute = false;
                }
                else
                {
                    this.canExecute = true;
                }
            }
        }

        #endregion

        #region UI

        public Visibility IsVisible
        {
            get
            {
                return this.isVisible;
            }
            set
            {
                this.isVisible = value;
                if (this.IsVisible == Visibility.Collapsed)
                {
                    if (this.ParentElement != null)
                    {
                        this.ParentElement.Elements.Remove(this);
                    }
                }
                else
                {
                    
                }
                OnPropertyChanged(new PropertyChangedEventArgs("IsVisible"));
            }
        }

        public bool HasAssociation
        {
            get
            {
                return this.hasAssociation;
            }
            set
            {
                this.hasAssociation = value;
                OnPropertyChanged(new PropertyChangedEventArgs("HasAssociation"));
            }
        }

        public string HeadImageSource
        {
            get
            {
                return this.headImageSource;
            }
            set
            {
                this.headImageSource = value;
                this.OnPropertyChanged(new PropertyChangedEventArgs("HeadImageSource"));
            }
        }

        public string FlagImageSource
        {
            get
            {
                return this.flagImageSource;
            }
            set
            {
                this.flagImageSource = value;
                this.OnPropertyChanged(new PropertyChangedEventArgs("FlagImageSource"));
            }
        }

        public Visibility ShowFlag
        {
            get
            {
                return this.showFlag;
            }
            set
            {
                this.showFlag = value;
                OnPropertyChanged(new PropertyChangedEventArgs("ShowFlag"));
            }
        }

        public string TailImageSource
        {
            get
            {
                return this.tailImageSource;
            }
            set
            {
                this.tailImageSource = value;
                this.OnPropertyChanged(new PropertyChangedEventArgs("TailImageSource"));
            }
        }

        // Data binding for MenuItem Open
        public bool CanOpen
        {
            get
            {
                return this.canOpen;
            }
            set
            {
                this.canOpen = value;
                OnPropertyChanged(new PropertyChangedEventArgs("CanOpen"));
            }
        }

        // Data binding for MenuItem Explore
        public bool CanExplore
        {
            get
            {
                return this.canExplore;
            }
            set
            {
                this.canExplore = value;
                OnPropertyChanged(new PropertyChangedEventArgs("CanExplore"));
            }
        }

        // Data binding for MenuItem Rename
        public bool CanRename
        {
            get
            {
                return this.canRename;
            }
            set
            {
                this.canRename = value;
                OnPropertyChanged(new PropertyChangedEventArgs("CanRename"));
            }
        }

        // Data binding for MenuItem Delete
        public bool CanDelete
        {
            get
            {
                return this.canDelete;
            }
            set
            {
                this.canDelete = value;
                OnPropertyChanged(new PropertyChangedEventArgs("CanDelete"));
            }
        }

        // Data binding for MenuItem New Version
        public bool CanCreateNewVersion
        {
            get
            {
                return this.canCreateNewVersion;
            }
            set
            {
                this.canCreateNewVersion = value;
                OnPropertyChanged(new PropertyChangedEventArgs("CanCreateNewVersion"));
            }
        }

        // Data binding for MenuItem Import Calendar
        public bool CanImportCalendar
        {
            get
            {
                return this.canImportCalendar;
            }
            set
            {
                this.canImportCalendar = value;
                OnPropertyChanged(new PropertyChangedEventArgs("CanImportCalendar"));
            }
        }

        // Data binding for MenuItem LabelWith
        public bool CanLabelWith
        {
            get
            {
                return this.canLabelWith;
            }
            set
            {
                this.canLabelWith = value;
                OnPropertyChanged(new PropertyChangedEventArgs("CanLabelWith"));
            }
        }

        // Data binding for MenuItem Reply
        public bool HasEmailAssociation
        {
            get
            {
                return this.hasEmailAssociation;
            }
            set
            {
                this.hasEmailAssociation = value;
                OnPropertyChanged(new PropertyChangedEventArgs("HasEmailAssociation"));
            }
        }

        // Data binding for MenuItem Reply
        public bool CanExecute
        {
            get
            {
                return this.canExecute;
            }
            set
            {
                this.canExecute = value;
                OnPropertyChanged(new PropertyChangedEventArgs("CanExecute"));
            }
        }

        public bool ShowAssociationMarkedDone
        {
            get
            {
                return this.showAssociationMarkedDone;
            }
            set
            {
                this.showAssociationMarkedDone = value;
            }
        }

        public bool ShowAssociationMarkedDefer
        {
            get
            {
                return this.showAssociationMarkedDefer;
            }
            set
            {
                this.showAssociationMarkedDefer = value;
            }
        }

        public double TailImageWidth
        {
            get
            {
                return this.tailImageWidth;
            }
            set
            {
                this.tailImageWidth = value;
                this.OnPropertyChanged(new PropertyChangedEventArgs("TailImageWidth"));
            }
        }

        public double TailImageHeight
        {
            get
            {
                return this.tailImageHeight;
            }
            set
            {
                this.tailImageHeight = value;
                this.OnPropertyChanged(new PropertyChangedEventArgs("TailImageHeight"));
            }
        }

        public Visibility ShowExpander
        {
            get
            {
                return this.showExpander;
            }
            set
            {
                this.showExpander = value;
                OnPropertyChanged(new PropertyChangedEventArgs("ShowExpander"));
            }
        }

        public string FontFamily
        {
            get
            {
                string fontFamily = String.Empty;

                switch (type)
                {
                    case ElementType.Heading:
                        fontFamily = Properties.Settings.Default.HeadingLevel1FontFamily;
                        break;
                    case ElementType.Note:
                        fontFamily = Properties.Settings.Default.NoteFontFamily;
                        break;
                }

                return fontFamily;
            }
        }

        public int FontSize
        {
            get
            {
                int fontSize = 0;

                switch (type)
                {
                    case ElementType.Heading:
                        fontSize = Properties.Settings.Default.HeadingLevel1FontSize;
                        break;
                    case ElementType.Note:
                        fontSize = Properties.Settings.Default.NoteFontSize;
                        break;
                }

                return fontSize;
            }
        }

        public string FontColor
        {
            get
            {
                return this.fontColor;
            }
            set
            {
                this.fontColor = value;
                this.OnPropertyChanged(new PropertyChangedEventArgs("FontColor"));
            }
        }

        #endregion

        #region Element Friends

        public Element ParentElement
        {
            get;
            set;
        }

        public ObservableCollection<Element> Elements
        {
            get
            {
                return this.children;
            }
            set
            {
                this.children = value;
                OnPropertyChanged(new PropertyChangedEventArgs("Elements"));
            }
        }

        public List<Element> HeadingElements
        {
            get
            {
                List<Element> headingChildren = new List<Element>();
                foreach (Element ele in Elements)
                {
                    if (ele.Type == ElementType.Heading)
                    {
                        headingChildren.Add(ele);
                    }
                }
                return headingChildren;
            }
        }

        public Element FirstChild
        {
            get
            {
                if (Elements.Count == 0)
                {
                    return null;
                }
                else
                {
                    return Elements[0];
                }
            }
        }

        public Element LastChild
        {
            get
            {
                if (Elements.Count == 0)
                {
                    return null;
                }
                else
                {
                    return Elements[Elements.Count - 1];
                }
            }
        }

        public Element ElementAbove
        {
            get
            {
                if (this.ParentElement != null)
                {
                    int index = this.ParentElement.Elements.IndexOf(this);
                    if (index > 0)
                    {
                        Element targetElement = this.ParentElement.Elements[index - 1];
                        if (targetElement.IsExpanded && targetElement.HasChildren)
                        {
                            targetElement = targetElement.LastChild;
                            while (targetElement.IsExpanded && targetElement.HasChildren)
                            {
                                targetElement = targetElement.LastChild;
                            }
                            return targetElement;
                        }
                        else
                        {
                            return targetElement;
                        }
                    }
                    else
                    {
                        return this.ParentElement;
                    }
                }
                else
                {
                    return null;
                }
            }
        }

        public Element ElementAboveUnderSameParent
        {
            get
            {
                if (this.ParentElement != null)
                {
                    int index = this.ParentElement.Elements.IndexOf(this);
                    if (index > 0)
                    {
                        return this.ParentElement.Elements[index - 1];
                    }
                    else
                    {
                        return null;
                    }
                }
                else
                {
                    return null;
                }
            }
        }

        public Element ElementBelow
        {
            get
            {
                if (this.ParentElement != null)
                {
                    if (this.IsExpanded && this.HasChildren)
                    {
                        return this.FirstChild;
                    }
                    else
                    {
                        int index = this.ParentElement.Elements.IndexOf(this);
                        if (index < this.ParentElement.Elements.Count - 1)
                        {
                            return this.ParentElement.Elements[index + 1];
                        }
                        else
                        {
                            Element targetElementAboveSibling = this.ParentElement;
                            while (targetElementAboveSibling.ElementBelowUnderSameParent == null)
                            {
                                targetElementAboveSibling = targetElementAboveSibling.ParentElement;
                                if (targetElementAboveSibling == null)
                                {
                                    break;
                                }
                            }
                            if (targetElementAboveSibling == null)
                            {
                                return null;
                            }
                            else
                            {
                                return targetElementAboveSibling.ElementBelowUnderSameParent;
                            }
                        }
                    }
                }
                else
                {
                    return null;
                }
            }
        }

        public Element ElementBelowUnderSameParent
        {
            get
            {
                if (this.ParentElement != null)
                {
                    int index = this.ParentElement.Elements.IndexOf(this);
                    if (index < this.ParentElement.Elements.Count - 1)
                    {
                        return this.ParentElement.Elements[index + 1];
                    }
                    else
                    {
                        return null;
                    }
                }
                else
                {
                    return null;
                }
            }
        }

        #endregion
    }
}
