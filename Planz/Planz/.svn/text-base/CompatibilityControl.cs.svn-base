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
using System.Xml;

namespace Planz
{
    public static class PlannerXMLConverter
    {
        //convert planner.xml into xooml.xml

        private const string XML_ROOT = "Label";
        private const string XML_ROOT_VIEW = "View";
        private const string XML_HEADING = "ChildLabel";
        private const string XML_NOTE = "Note";
        private const string XML_LINK = "Link";
        private const string XML_MAIL = "Mail";
        private const string XML_WORD = "Word";
        private const string XML_EXCEL = "Excel";
        private const string XML_PPT = "Powerpoint";
        private const string XML_ELEMENT_GUID = "GUID";
        private const string XML_ELEMENT_NOTETEXT = "Name";
        private const string XML_ELEMENT_POSITION = "Order";
        private const string XML_ELEMENT_ASSOCIATIONTYPE = "LinkType";
        private const string XML_ELEMENT_ASSOCIATIONURI = "LinkTarget";
        private const string XML_ELEMENT_ISCOLLAPSED = "IsCollapsed";
        private const string XML_ELEMENT_FOLDERLINK = "FolderLink";
        private const string XML_ELEMENT_ASSOCIATIONTYPE_FILE = "File";
        private const string XML_ELEMENT_ASSOCIATIONTYPE_FILESHORTCUT = "FileShortcut";
        private const string XML_ELEMENT_ASSOCIATIONTYPE_WEB = "WebShortcut";
        private const string XML_ELEMENT_ASSOCIATIONTYPE_FOLDERSHORTCUT = "FolderShortcut";
        private const string XML_ELEMENT_ASSOCIATIONTYPE_NONE = "None";
        private const string XML_MAIL_ENTRYID = "EntryID";

        public static List<Element> LoadXML(string path, XmlDocument xmlDoc)
        {
            xmlDoc = new XmlDocument();
            xmlDoc.Load(path);
            string dirPath = System.IO.Path.GetDirectoryName(path) + System.IO.Path.DirectorySeparatorChar;
            return ConvertXML(dirPath, xmlDoc);
        }

        private static List<Element> ConvertXML(string dirPath, XmlDocument xmlDoc)
        {
            List<Element> elementList = new List<Element>();

            XmlNode xmlRootNode = xmlDoc.GetElementsByTagName(XML_ROOT)[0];

            foreach (XmlNode xmlElement in xmlRootNode.ChildNodes)
            {
                try
                {
                    Element element = new Element();

                    element.ID = new Guid(xmlElement.Attributes[XML_ELEMENT_GUID].Value);
                    element.NoteText = xmlElement.Attributes[XML_ELEMENT_NOTETEXT].Value;
                    element.Position = Int32.Parse(xmlElement.Attributes[XML_ELEMENT_POSITION].Value);
                    element.IsExpanded = !Boolean.Parse(xmlElement.Attributes[XML_ELEMENT_ISCOLLAPSED].Value);
                    element.Status = ElementStatus.Normal;

                    switch (xmlElement.Name)
                    {
                        case XML_HEADING:
                            element.Type = ElementType.Heading;
                            element.AssociationURI = HeadingNameConverter.ConvertFromHeadingNameToFolderName(element);
                            element.AssociationType = ElementAssociationType.Folder;
                            element.FontColor = ElementColor.DarkBlue.ToString();
                            if (element.IsExpanded)
                            {
                                element.LevelOfSynchronization = 1;
                            }
                            else
                            {
                                element.LevelOfSynchronization = 0;
                            }
                            break;
                        case XML_NOTE:
                        case XML_LINK:
                        case XML_MAIL:
                        case XML_WORD:
                        case XML_EXCEL:
                        case XML_PPT:
                        default:
                            element.Type = ElementType.Note;
                            element.LevelOfSynchronization = 0;
                            break;
                    };

                    if (xmlElement.Attributes[XML_ELEMENT_ASSOCIATIONURI] != null)
                    {
                        element.AssociationURI = xmlElement.Attributes[XML_ELEMENT_ASSOCIATIONURI].Value;
                    }

                    if (xmlElement.Attributes[XML_ELEMENT_ASSOCIATIONTYPE] != null)
                    {
                        switch (xmlElement.Attributes[XML_ELEMENT_ASSOCIATIONTYPE].Value)
                        {
                            case XML_ELEMENT_ASSOCIATIONTYPE_FOLDERSHORTCUT:
                                element.AssociationType = ElementAssociationType.FolderShortcut;
                                break;
                            case XML_ELEMENT_ASSOCIATIONTYPE_FILESHORTCUT:
                                element.AssociationType = ElementAssociationType.FileShortcut;
                                break;
                            case XML_ELEMENT_ASSOCIATIONTYPE_WEB:
                                element.AssociationType = ElementAssociationType.Web;
                                break;
                            case XML_ELEMENT_ASSOCIATIONTYPE_FILE:
                            case XML_ELEMENT_ASSOCIATIONTYPE_NONE:
                            default:
                                element.AssociationType = ElementAssociationType.File;
                                break;
                        };
                    }
                    else
                    {
                        if (xmlElement.Name == XML_WORD ||
                            xmlElement.Name == XML_EXCEL ||
                            xmlElement.Name == XML_PPT)
                        {
                            element.AssociationType = ElementAssociationType.File;
                        }
                    }

                    // In-place expansion will be converted to Note with folder shortcut
                    if (xmlElement.Attributes[XML_ELEMENT_FOLDERLINK] != null)
                    {
                        element.Type = ElementType.Heading;
                        element.AssociationURI = xmlElement.Attributes[XML_ELEMENT_ASSOCIATIONURI].Value;
                        element.IsExpanded = false;
                        element.AssociationType = ElementAssociationType.FolderShortcut;
                        element.LevelOfSynchronization = 1;
                    }

                    // Email
                    if (xmlElement.Name == XML_MAIL)
                    {
                        string shortcutName = ShortcutNameConverter.GenerateShortcutNameFromEmailSubject(element.NoteText, dirPath);
                        string shortcutPath = dirPath + shortcutName;
                        IWshRuntimeLibrary.WshShellClass wshShell = new IWshRuntimeLibrary.WshShellClass();
                        IWshRuntimeLibrary.IWshShortcut shortcut = (IWshRuntimeLibrary.IWshShortcut)wshShell.CreateShortcut(shortcutPath);

                        string targetPath = OutlookSupportFunction.GenerateShortcutTargetPath();
                        shortcut.TargetPath = targetPath;
                        shortcut.Arguments = "/select outlook:" + xmlElement.Attributes[XML_MAIL_ENTRYID].Value;
                        shortcut.Description = targetPath;

                        shortcut.Save();

                        element.AssociationType = ElementAssociationType.Email;
                        element.AssociationURI = shortcutName;
                    }

                    elementList.Add(element);
                }
                catch (Exception ex)
                {
                    continue;
                }
            }

            elementList.Sort(new SortByPosition());

            return elementList;
        }
    }

    public static class PlanzXMLConverter
    {
        //convert planz.xml into xooml.xml
        private const string XML_ROOT = "Root";
        private const string XML_ELEMENT = "Element";
        private const string XML_ELEMENT_GUID = "Guid";
        private const string XML_ELEMENT_TYPE = "Type";
        private const string XML_ELEMENT_NOTETEXT = "NoteText";
        private const string XML_ELEMENT_POSITION = "Position";
        private const string XML_ELEMENT_STATUS = "Status";
        private const string XML_ELEMENT_ASSOCIATIONTYPE = "AssociationType";
        private const string XML_ELEMENT_ASSOCIATIONURI = "AssociationURI";
        private const string XML_ELEMENT_ISEXPANDED = "IsExpanded";
        private const string XML_ELEMENT_HIGHLIGHT = "Highlight";
        private const string XML_ELEMENT_TAG = "Tag";
        private const string XML_ELEMENT_EXTRA = "Extra";
        private const string XML_ELEMENT_DISCOVERABLE = "Discoverable";

        public static List<Element> LoadXML(string path, XmlDocument xmlDoc)
        {
            xmlDoc = new XmlDocument();
            xmlDoc.Load(path);
            string dirPath = System.IO.Path.GetDirectoryName(path) + System.IO.Path.DirectorySeparatorChar;
            return ConvertXML(dirPath, xmlDoc);
        }

        private static List<Element> ConvertXML(string dirPath, XmlDocument xmlDoc)
        {
            List<Element> elementList = new List<Element>();

            XmlNode xmlRootNode = xmlDoc.GetElementsByTagName(XML_ROOT)[0];
            int position = 0;
            String currentDir = dirPath;

            foreach (XmlNode xmlElement in xmlRootNode.ChildNodes)
            {
                try
                {
                    //add new element into the XML file
                    Element element = new Element();
                    element.ID = new Guid(xmlElement.Attributes[XML_ELEMENT_GUID].Value);
                    element.AssociationURI = xmlElement.Attributes[XML_ELEMENT_ASSOCIATIONURI].Value;
                    bool isDiscoverable = Boolean.Parse(xmlElement.Attributes[XML_ELEMENT_DISCOVERABLE].Value);
                    if (isDiscoverable)
                    {
                        element.LevelOfSynchronization = 1;
                    }
                    else
                    {
                        element.LevelOfSynchronization = 0;
                    }
                    element.NoteText = xmlElement.Attributes[XML_ELEMENT_NOTETEXT].Value;
                    element.AssociationType = (ElementAssociationType)Enum.Parse(typeof(ElementAssociationType), xmlElement.Attributes[XML_ELEMENT_ASSOCIATIONTYPE].Value);
                    element.IsVisible = System.Windows.Visibility.Visible;
                    element.Type = (ElementType)Enum.Parse(typeof(ElementType), xmlElement.Attributes[XML_ELEMENT_TYPE].Value);
                    element.IsExpanded = Boolean.Parse(xmlElement.Attributes[XML_ELEMENT_ISEXPANDED].Value);
                    switch (xmlElement.Attributes[XML_ELEMENT_STATUS].Value)
                    {
                        case "Flag":
                            element.Status = ElementStatus.Normal;
                            element.FlagStatus = FlagStatus.Flag;
                            element.ShowFlag = System.Windows.Visibility.Visible;
                            element.FlagImageSource = String.Format("pack://application:,,,/{0};component/{1}", "Planz", "Images/flag.gif");
                            break;
                        case "Check":
                            element.Status = ElementStatus.Normal;
                            element.FlagStatus = FlagStatus.Check;
                            element.ShowFlag = System.Windows.Visibility.Visible;
                            element.FlagImageSource = String.Format("pack://application:,,,/{0};component/{1}", "Planz", "Images/check.gif");
                            break;
                        default:
                            element.Status = (ElementStatus)Enum.Parse(typeof(ElementStatus), xmlElement.Attributes[XML_ELEMENT_STATUS].Value);
                            element.FlagImageSource = String.Format("pack://application:,,,/{0};component/{1}", "Planz", "Images/normal.gif");
                            element.FlagStatus = FlagStatus.Normal;
                            break;
                    };
                    element.PowerDStatus = PowerDStatus.None;
                    element.FontColor = xmlElement.Attributes[XML_ELEMENT_HIGHLIGHT].Value;
                    element.Position = position++;
                    element.Tag = xmlElement.Attributes[XML_ELEMENT_TAG].Value;

                    //to be corrected
                    if (element.AssociationURI != String.Empty)
                    {
                        if (element.IsLocalHeading && System.IO.Directory.Exists(currentDir + element.AssociationURI))
                        {
                            // If the associated folder exists
                            element.TailImageSource = FileTypeHandler.GetIcon((ElementAssociationType)element.AssociationType, currentDir + element.AssociationURI);
                        }
                        else if (System.IO.File.Exists(currentDir + element.AssociationURI))
                        {
                            // If the associated file exists
                            element.TailImageSource = FileTypeHandler.GetIcon((ElementAssociationType)element.AssociationType, currentDir + element.AssociationURI);
                        }
                        else
                        {
                            // If the associated file is missing
                            element.TailImageSource = FileTypeHandler.GetMissingFileIcon();
                            element.FontColor = ElementColor.Gray.ToString();
                            element.Status = ElementStatus.MissingAssociation;
                        }
                    }

                    switch (element.Type)
                    {
                        case ElementType.Heading:
                            if (element.IsRemoteHeading)
                            {
                                IWshRuntimeLibrary.WshShellClass wshShell = new IWshRuntimeLibrary.WshShellClass();
                                IWshRuntimeLibrary.IWshShortcut shortcut = (IWshRuntimeLibrary.IWshShortcut)wshShell.CreateShortcut(currentDir + element.AssociationURI);
                                if (shortcut != null)
                                {
                                    element.Path = shortcut.TargetPath + System.IO.Path.DirectorySeparatorChar;
                                }
                            }
                            else
                            {
                                element.Path = currentDir + HeadingNameConverter.ConvertFromHeadingNameToFolderName(element) + System.IO.Path.DirectorySeparatorChar;
                                element.TailImageSource = FileTypeHandler.GetIcon(ElementAssociationType.Folder, element.Path);
                            }

                            if (System.IO.Directory.Exists(element.Path))
                            {
                            }
                            else
                            {
                                // If the associated folder is missing
                                element.Type = ElementType.Note;
                                element.IsExpanded = false;
                                element.FontColor = ElementColor.Gray.ToString();
                                element.Status = ElementStatus.MissingAssociation;
                                element.AssociationType = ElementAssociationType.None;
                                element.AssociationURI = String.Empty;
                                element.TailImageSource = String.Empty;
                            }
                            break;
                        case ElementType.Note:
                            break;
                    };
                    elementList.Add(element);
                }
                catch (Exception ex)
                {
                    continue;
                }
            }
            elementList.Sort(new SortByPosition());
            return elementList;
        }
    }
}
