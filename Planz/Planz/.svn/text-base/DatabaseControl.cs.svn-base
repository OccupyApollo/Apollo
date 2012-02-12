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
using System.IO;

namespace Planz
{
    public delegate void NewDatabaseControlHandler(object sender, EventArgs e);
    public delegate void ElementStatusChangedDelegate(object sender, EventArgs e);
    public delegate void ElementUpdateDelegate(object sender, EventArgs e);
    public delegate void ElementDeleteDelegate(object sender, EventArgs e);
    public delegate void NewXooMLCreateDelegate(object sender, EventArgs e);

    public class DatabaseControl
    {
        private const string DB_FILE_PLANNER7 = "planner.xml";
        private const string DB_FILE_PLANZ8 = "Planz.xml";
        private const string DB_FILE = "XooML.xml";

        // Namespace Prefix
        private const string PREFIX_XOOML = "xooml";
        private const string PREFIX_PLANZ = "planz";
        private const string PREFIX_XSI = "xsi";
        private const string XMLNS = "xmlns";

        // Fragment Level
        private const string FRAGMENT = "fragment";
        private const string FRAGMENT_TOOL = "fragmentToolAttributes";

        private const string NAMESPACE_XOOML = "http://kftf.ischool.washington.edu/xmlns/xooml";
        private const string NAMESPACE_PLANZ = "http://kftf.ischool.washington.edu/xmlns/planz";
        private const string NAMESPACE_XSI = "http://www.w3.org/2001/XMLSchema-instance";
        
        private const string SCHEMA_LOCATION = "schemaLocation";
        private const string SCHEMA_LOCATION_VALUE = "http://kftf.ischool.washington.edu/xmlns/xooml http://kftf.ischool.washington.edu/XMLschema/0.41/XooML.xsd";
        
        private const string SCHEMA_VERSION = "schemaVersion";
        private const string SCHEMA_VERSION_VALUE = "0.41";
        private const string DEFAULT_APPLICATION = "defaultApplication";
        private const string RELATED_ITEM = "relatedItem";
        
        private const string TOOL_VERSION = "toolVersion";
        private const string TOOL_NAME = "toolName";
        private const string START_DATE = "startDate";
        private const string DUE_DATE = "dueDate";
        private const string SHOW_ASSOCIATION_MARKED_DONE = "showAssociationMarkedDone";
        private const string SHOW_ASSOCIATION_MARKED_DEFER = "showAssociationMarkedDefer";

        private const string TOOL_PLANZ_VERSION = "1.0.7";
        private const string TOOL_PLANZ_NAME = "Planz";

        // Association Level
        private const string ASSOCIATION = "association";
        private const string ASSOCIATION_TOOL = "associationToolAttributes";

        private const string ASSOCIATION_ID = "ID";
        private const string ASSOCIATION_ITEM = "associatedItem";
        private const string ASSOCIATION_ITEM_TYPE = "associatedItemType";
        private const string ASSOCIATION_ITEM_ICON = "associatedIcon";
        private const string ASSOCIATION_FRAGMENT = "associatedXooMLFragment";
        private const string LEVEL_OF_SYNCHRONIZATION = "levelOfSynchronization";
        private const string DISPLAY_TEXT = "displayText";
        private const string OPEN_WITH_DEFAULT = "openWithDefault";
        private const string CREATED_BY = "createdBy";
        private const string MODIFIED_BY = "modifiedBy";
        private const string CREATED_ON = "createdOn";
        private const string MODIFIED_ON = "modifiedOn";

        private const string TYPE = "type";
        private const string ISVISIBLE = "isVisible";
        private const string ISCOLLAPSED = "isCollapsed";
        private const string STATUS = "status";
        private const string FLAG_STATUS = "flagStatus";
        private const string POWERD_STATUS = "powerDStatus";
        private const string POWERD_TIMESTAMP = "powerDTimeStamp";
        private const string FONT_COLOR = "fontColor";
        private const string POSITION = "position";

        private string currentDir = null;
        private string xmlFilePath = null;
        private XmlNode xmlRootNode = null;

        private XmlDocument xmlDoc = null;
        XmlNamespaceManager xmlnsManager = null;

        private List<string> fileList = new List<string>();
        private List<string> subfolderList = new List<string>();

        public event NewDatabaseControlHandler newDBControlHighP;
        public event NewDatabaseControlHandler newDBControlLowP;
        public event ElementUpdateDelegate elementUpdate;
        public event ElementStatusChangedDelegate elementStatusChangedDelegate;
        public event ElementDeleteDelegate elementDelete;
        public event NewXooMLCreateDelegate newXooMLCreate;

        public bool SHOW_ME_ALL = false;
        public bool DO_NOT_UPDATE_POWERDREGION = false;

        // Note: param "currentDir" should be ending with System.IO.Path.DirectorySeparatorChar
        public DatabaseControl(string currentDir)
        {
            this.currentDir = currentDir;
            this.xmlFilePath = currentDir + DB_FILE;

            try
            {
                DirectoryInfo dInfo = new DirectoryInfo(this.currentDir);
                List<FileInfo> sortedFiles = dInfo.GetFiles().ToList();
                sortedFiles.Sort(new SortByLastModified());
                foreach (FileInfo fi in sortedFiles)
                {
                    if (fi.Attributes == (FileAttributes.Archive | FileAttributes.Hidden) ||
                        fi.Name.StartsWith("~$"))
                    {
                        continue;
                    }

                    if (fi.Name != DB_FILE && 
                        fi.Name != DB_FILE_PLANNER7 &&
                        fi.Name != DB_FILE_PLANZ8)
                    {
                        fileList.Add(fi.Name);
                    }
                }
                foreach (DirectoryInfo di in dInfo.GetDirectories())
                {
                    subfolderList.Add(di.Name);
                }
                
            }
            catch (Exception)
            {

            }
        }

        public void OpenConnection()
        { 
            if (File.Exists(xmlFilePath))
            {
                LoadXML();
            }
            else
            {
                // Support Planner 7 series and Planz 8 beta
                string plannerXMLPath = currentDir + DB_FILE_PLANNER7;
                string planzXMLPath = currentDir + DB_FILE_PLANZ8;
                //if planz.xml exist
                if (File.Exists(planzXMLPath))
                {
                    List<Element> elementList = PlanzXMLConverter.LoadXML(planzXMLPath, xmlDoc);
                    CreateXML(false);
                    foreach (Element element in elementList)
                    {
                        InsertElementIntoXML(element);
                    }
                }else if (File.Exists(plannerXMLPath)) //if planner.xml exist
                {
                    List<Element> elementList = PlannerXMLConverter.LoadXML(plannerXMLPath, xmlDoc);
                    CreateXML(false);
                    foreach (Element element in elementList)
                    {
                        InsertElementIntoXML(element);
                    }
                }else
                {
                    CreateXML(false);
                    //OnNewXooMLCreate(this, new EventArgs());
                }
            }
        }

        public void CloseConnection()
        {
            xmlDoc = null;

            //GC.Collect();
            //GC.WaitForPendingFinalizers();
        }

        public bool IsConnected
        {
            get
            {
                if (xmlDoc == null)
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
        }

        protected virtual void OnNewDBControlHighP(object sender, EventArgs e)
        {
            if (newDBControlHighP != null)
                newDBControlHighP(sender, e);
        }

        protected virtual void OnNewDBControlLowP(object sender, EventArgs e)
        {
            if (newDBControlLowP != null)
                newDBControlLowP(sender, e);
        }

        protected virtual void OnElementUpdate(object sender, EventArgs e)
        {
            if (elementUpdate != null)
                elementUpdate(sender, e);
        }

        protected virtual void OnElementStatusChanged(object sender, EventArgs e)
        {
            if (elementStatusChangedDelegate != null)
            {
                elementStatusChangedDelegate(sender, e);
            }
        }

        protected virtual void OnElementDelete(object sender, EventArgs e)
        {
            if (elementDelete != null)
                elementDelete(sender, e);
        }

        protected virtual void OnNewXooMLCreate(object sender, EventArgs e)
        {
            if (newXooMLCreate != null)
                newXooMLCreate(sender, e);
        }

        private void LoadXML()
        {
            try
            {
                xmlDoc = new XmlDocument();
                xmlDoc.Load(xmlFilePath);
                xmlRootNode = xmlDoc.GetElementsByTagName(FRAGMENT, NAMESPACE_XOOML)[0];

                LoadNamespace();
            }
            catch (Exception ex)
            {

            }
        }

        private void LoadNamespace()
        {
            xmlnsManager = new XmlNamespaceManager(xmlDoc.NameTable);
            xmlnsManager.AddNamespace(PREFIX_XOOML, NAMESPACE_XOOML);
            xmlnsManager.AddNamespace(PREFIX_PLANZ, NAMESPACE_PLANZ);
            xmlnsManager.AddNamespace(PREFIX_XSI, NAMESPACE_XSI);
        }

        private void CreateXML(bool hasFirstNote)
        {
            try
            {
                xmlDoc = new XmlDocument();

                LoadNamespace();

                XmlNode declaration = xmlDoc.CreateXmlDeclaration("1.0", "UTF-8", null);
                xmlDoc.AppendChild(declaration);

                xmlRootNode = xmlDoc.CreateElement(PREFIX_XOOML, FRAGMENT, NAMESPACE_XOOML);
                
                XmlAttribute x1 = xmlDoc.CreateAttribute(PREFIX_XSI, SCHEMA_LOCATION, NAMESPACE_XSI);
                x1.Value = SCHEMA_LOCATION_VALUE;
                XmlAttribute x2 = xmlDoc.CreateAttribute(SCHEMA_VERSION);
                x2.Value = SCHEMA_VERSION_VALUE;
                XmlAttribute x3 = xmlDoc.CreateAttribute(DEFAULT_APPLICATION);
                XmlAttribute x4 = xmlDoc.CreateAttribute(RELATED_ITEM);
                x4.Value = currentDir;

                XmlNode fragmentToolNode = xmlDoc.CreateElement(PREFIX_XOOML, FRAGMENT_TOOL, NAMESPACE_XOOML);

                XmlAttribute ns = xmlDoc.CreateAttribute(XMLNS);
                ns.Value = NAMESPACE_PLANZ;
                XmlAttribute y1 = xmlDoc.CreateAttribute(TOOL_VERSION);
                y1.Value = TOOL_PLANZ_VERSION;
                XmlAttribute y2 = xmlDoc.CreateAttribute(TOOL_NAME);
                y2.Value = TOOL_PLANZ_NAME;
                XmlAttribute y3 = xmlDoc.CreateAttribute(START_DATE);
                XmlAttribute y4 = xmlDoc.CreateAttribute(DUE_DATE);
                XmlAttribute y5 = xmlDoc.CreateAttribute(SHOW_ASSOCIATION_MARKED_DONE);
                y5.Value = "False";
                XmlAttribute y6 = xmlDoc.CreateAttribute(SHOW_ASSOCIATION_MARKED_DEFER);
                y6.Value = "False";

                xmlRootNode.Attributes.Append(x1);
                xmlRootNode.Attributes.Append(x2);
                xmlRootNode.Attributes.Append(x3);
                xmlRootNode.Attributes.Append(x4);
                fragmentToolNode.Attributes.Append(ns);
                fragmentToolNode.Attributes.Append(y1);
                fragmentToolNode.Attributes.Append(y2);
                fragmentToolNode.Attributes.Append(y3);
                fragmentToolNode.Attributes.Append(y4);
                fragmentToolNode.Attributes.Append(y5);
                fragmentToolNode.Attributes.Append(y6);

                xmlRootNode.AppendChild(fragmentToolNode);
                xmlDoc.AppendChild(xmlRootNode);

                if (hasFirstNote)
                {
                    Element newElement = new Element
                    {
                        NoteText = String.Empty,
                        IsExpanded = false,
                        Status = ElementStatus.New,
                        Type = ElementType.Note,
                        //Position = 0,
                    };
                    InsertElementIntoXML(newElement);
                }

                SaveXML();
                
            }
            catch (Exception ex)
            {

            }
        }

        public void SaveXML()
        {
            try
            {
                xmlDoc.Save(xmlFilePath);
            }
            catch (Exception ex)
            {

            }
        }

        public Element GetFragmentElementFromXML()
        {
            Element element = new Element();

            XmlElement fragmentToolNode = FindXMLElement(xmlRootNode as XmlElement, NAMESPACE_PLANZ);

            if (fragmentToolNode != null)
            {
                element.ShowAssociationMarkedDone = Boolean.Parse(fragmentToolNode.Attributes[SHOW_ASSOCIATION_MARKED_DONE].Value);
                element.ShowAssociationMarkedDefer = Boolean.Parse(fragmentToolNode.Attributes[SHOW_ASSOCIATION_MARKED_DEFER].Value);

                if (fragmentToolNode.Attributes[START_DATE].Value != String.Empty)
                {
                    element.StartDate = DateTime.Parse(fragmentToolNode.Attributes[START_DATE].Value);
                }
                if (fragmentToolNode.Attributes[DUE_DATE].Value != String.Empty)
                {
                    element.DueDate = DateTime.Parse(fragmentToolNode.Attributes[DUE_DATE].Value);
                }
            }

            return element;
        }

        public List<Element> GetAllElementFromXML()
        {
            List<Element> allElement = new List<Element>();

            bool showDone = false;
            bool showDeferred = false;

            double tailIconHeight = Properties.Settings.Default.TailImageHeight;

            bool removePositionAttribute = false;

            bool isMoveDefer = false;
            int posMoveDefer = 0;

            foreach (XmlElement xmlElement in xmlRootNode.ChildNodes)
            {
                try
                {
                    if (xmlElement.LocalName != ASSOCIATION)
                    {
                        showDone = Boolean.Parse(xmlElement.Attributes[SHOW_ASSOCIATION_MARKED_DONE].Value);
                        showDeferred = Boolean.Parse(xmlElement.Attributes[SHOW_ASSOCIATION_MARKED_DEFER].Value);

                        continue;
                    }

                    Element element = new Element
                    {
                        ID = new Guid(xmlElement.Attributes[ASSOCIATION_ID].Value),
                        AssociationURI = xmlElement.Attributes[ASSOCIATION_ITEM].Value,
                        LevelOfSynchronization = Int32.Parse(xmlElement.Attributes[LEVEL_OF_SYNCHRONIZATION].Value),
                        NoteText = xmlElement.Attributes[DISPLAY_TEXT].Value,
                        TailImageHeight = tailIconHeight,
                        TailImageWidth = tailIconHeight,
                    };

                    XmlElement associationToolNode = FindXMLElement(xmlElement, NAMESPACE_PLANZ);
                    if (associationToolNode != null)
                    {
                        element.AssociationType = (ElementAssociationType)Enum.Parse(typeof(ElementAssociationType), associationToolNode.Attributes[ASSOCIATION_ITEM_TYPE].Value);
                        element.Type = (ElementType)Enum.Parse(typeof(ElementType), associationToolNode.Attributes[TYPE].Value);
                        element.IsExpanded = !Boolean.Parse(associationToolNode.Attributes[ISCOLLAPSED].Value);
                        element.Status = (ElementStatus)Enum.Parse(typeof(ElementStatus), associationToolNode.Attributes[STATUS].Value);
                        element.PowerDStatus = (PowerDStatus)Enum.Parse(typeof(PowerDStatus), associationToolNode.Attributes[POWERD_STATUS].Value);
                        element.FlagStatus = (FlagStatus)Enum.Parse(typeof(FlagStatus), associationToolNode.Attributes[FLAG_STATUS].Value);
                        element.FontColor = associationToolNode.Attributes[FONT_COLOR].Value;

                        if (Boolean.Parse(associationToolNode.Attributes[ISVISIBLE].Value) == true)
                        {
                            element.IsVisible = System.Windows.Visibility.Visible;
                        }
                        else
                        {
                            element.IsVisible = System.Windows.Visibility.Collapsed;
                        }

                        if (element.PowerDStatus == PowerDStatus.Deferred)
                        {
                            if (element.FontColor == ElementColor.Plum.ToString())
                            {
                                element.FontColor = ElementColor.Purple.ToString();
                            }
                            if (showDeferred)
                            {
                                element.IsVisible = System.Windows.Visibility.Visible;
                            }
                            else
                            {
                                element.IsVisible = System.Windows.Visibility.Collapsed;
                            }
                        }
                        if (element.PowerDStatus == PowerDStatus.Done)
                        {
                            if (element.FontColor == ElementColor.SpringGreen.ToString())
                            {
                                element.FontColor = ElementColor.SeaGreen.ToString();
                            }
                            if (showDone)
                            {
                                element.IsVisible = System.Windows.Visibility.Visible;
                            }
                            else
                            {
                                element.IsVisible = System.Windows.Visibility.Collapsed;
                            }
                        }

                        isMoveDefer = false;
                        if (associationToolNode.Attributes[POWERD_TIMESTAMP].Value != String.Empty)
                        {
                            element.PowerDTimeStamp = DateTime.Parse(associationToolNode.Attributes[POWERD_TIMESTAMP].Value);
                            if (element.PowerDTimeStamp <= DateTime.Now.Date && element.PowerDStatus == PowerDStatus.Deferred)
                            {
                                element.FontColor = ElementColor.Blue.ToString();
                                element.PowerDStatus = PowerDStatus.None;
                                element.IsVisible = System.Windows.Visibility.Visible;
                                element.Status = ElementStatus.New;
                                isMoveDefer = true;
                            }
                        }
                        
                        if (associationToolNode.HasAttribute(POSITION))
                        {
                            element.Position = Int32.Parse(associationToolNode.Attributes[POSITION].Value);
                        }
                        
                    }

                    if (!removePositionAttribute && associationToolNode.HasAttribute(POSITION))
                    {
                        removePositionAttribute = true;
                    }

                    if (element.Status == ElementStatus.ToBeDeleted)
                    {
                        OnElementDelete(element, new EventArgs());
                    }

                    if ((element.Status == ElementStatus.New) && (element.IsVisible == System.Windows.Visibility.Visible))
                    {
                        OnElementStatusChanged(element, new EventArgs());
                    }

                    if (xmlElement.Attributes[CREATED_ON].Value != String.Empty)
                    {
                        element.CreatedOn = DateTime.Parse(xmlElement.Attributes[CREATED_ON].Value);
                    }
                    if (xmlElement.Attributes[MODIFIED_ON].Value != String.Empty)
                    {
                        element.ModifiedOn = DateTime.Parse(xmlElement.Attributes[MODIFIED_ON].Value);
                    }

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
                            //OnElementStatusChanged(element, new EventArgs());
                        }

                        fileList.Remove(element.AssociationURI);
                    }
                    //if (element.IsHeading)
                    //{
                    switch (element.FlagStatus)
                    {
                        case FlagStatus.Normal:
                            element.FlagImageSource = String.Format("pack://application:,,,/{0};component/{1}", "Planz", "Images/normal.gif");
                            break;
                        case FlagStatus.Flag:
                            element.ShowFlag = System.Windows.Visibility.Visible;
                            element.FlagImageSource = String.Format("pack://application:,,,/{0};component/{1}", "Planz", "Images/flag.gif");
                            break;
                        case FlagStatus.Check:
                            element.ShowFlag = System.Windows.Visibility.Visible;
                            element.FlagImageSource = String.Format("pack://application:,,,/{0};component/{1}", "Planz", "Images/check.gif");
                            break;
                    };
                    //}
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

                            if (Directory.Exists(element.Path))
                            {
                                // If the associated folder exists
                                if (element.IsExpanded)
                                {
                                    OnNewDBControlHighP(element, new EventArgs());
                                }
                                else
                                {
                                    //if (element.Discoverable)
                                    //{
                                    //    OnNewDBControlLowP(element, new EventArgs());
                                    //}
                                }
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

                                UpdateElementIntoXML(element);

                                //OnElementStatusChanged(element, new EventArgs());
                            }

                            string toRemove = HeadingNameConverter.ConvertFromHeadingNameToFolderName(element);
                            for (int i = 0; i < subfolderList.Count; i++)
                            {
                                if (toRemove.ToLower() == subfolderList[i].ToLower())
                                {
                                    subfolderList.RemoveAt(i);
                                    break;
                                }
                            }

                            break;
                        case ElementType.Note:
                            break;
                    };

                    if (isMoveDefer)
                        allElement.Insert(posMoveDefer++, element);
                    else
                        allElement.Add(element);
                }
                catch (Exception ex)
                {

                }
            }

            // Done/Defer regions
            if (!DO_NOT_UPDATE_POWERDREGION)
            {
                UpdatePowerDRegion(allElement); 
                UpdatePosition(allElement);
            }

            // Ellipsis regions
            //if (!SHOW_ME_ALL)
            //{
            //    UpdateEllipsisRegions(allElement);
            //}

            // New files and subfolders
            List<Element> newElementList = new List<Element>();
            if (fileList.Count != 0)
            {
                foreach (string fileName in fileList)
                {
                    string fileFullName = currentDir + fileName;
                    ElementAssociationType eat = (ElementAssociationType)FileTypeHandler.GetFileType(fileFullName);

                    Element newElement = new Element
                    {
                        AssociationType = eat,
                        AssociationURI = fileName,
                        FontColor = ElementColor.Blue.ToString(),
                        NoteText = System.IO.Path.GetFileNameWithoutExtension(fileName),
                        TailImageSource = FileTypeHandler.GetIcon(eat, currentDir + fileName),
                        Type = ElementType.Note,
                        //Position = 0,
                        Status = ElementStatus.New,
                        LevelOfSynchronization = 0,
                        FlagImageSource = String.Format("pack://application:,,,/{0};component/{1}", "Planz", "Images/normal.gif"),
                    };
                    newElementList.Add(newElement);

                    OnElementStatusChanged(newElement, new EventArgs());
                }
            }
            if (subfolderList.Count != 0)
            {
                foreach (string subfolderName in subfolderList)
                {
                    Element newElement = new Element
                    {
                        IsExpanded = false,
                        FontColor = ElementColor.Blue.ToString(),
                        NoteText = HeadingNameConverter.ConvertFromFolderNameToHeadingName(subfolderName),
                        Type = ElementType.Heading,
                        //Position = 0,
                        Status = ElementStatus.New,
                        LevelOfSynchronization = 0,
                        AssociationURI = subfolderName,
                        AssociationType = ElementAssociationType.Folder,
                        TailImageSource = FileTypeHandler.GetIcon(ElementAssociationType.Folder, currentDir + subfolderName),
                        FlagImageSource = String.Format("pack://application:,,,/{0};component/{1}", "Planz", "Images/normal.gif"),
                    };
                    newElement.Path = currentDir + HeadingNameConverter.ConvertFromHeadingNameToFolderName(newElement) + System.IO.Path.DirectorySeparatorChar;

                    if (HeadingNameConverter.IsValidGuid(newElement.NoteText))
                    {
                        newElement.Path = currentDir + newElement.NoteText + System.IO.Path.DirectorySeparatorChar;
                        newElement.NoteText = String.Empty;
                        
                    }
                    
                    newElementList.Add(newElement);

                    OnElementStatusChanged(newElement, new EventArgs());

                    //OnNewDBControlLowP(newElement, new EventArgs());
                }
            }

            if (removePositionAttribute)
            {
                allElement.Sort(new SortByPosition());

                UpdatePosition(allElement);
            }

            int insert_p = 0;
            foreach (Element newElement in newElementList)
            {
                newElement.Position = insert_p;
                InsertElementIntoXML(newElement);
                allElement.Insert(insert_p++, newElement);
                //UpdatePosition(allElement);
            }

            return allElement;
        }

        private XmlElement FindXMLElement(Element element)
        {
            XmlElement xmlElement = null;
            foreach (XmlNode xmlNode in xmlRootNode.ChildNodes)
            {
                if (xmlNode.LocalName != ASSOCIATION)
                {
                    continue;
                }

                if (xmlNode.Attributes[ASSOCIATION_ID].Value == element.ID.ToString())
                {
                    xmlElement = xmlNode as XmlElement;
                    break;
                }
            }
            return xmlElement;
        }

        public int GetElementPosition(Element element)
        {
            int index = 0;
            XmlElement xmlElement = null;
            foreach (XmlNode xmlNode in xmlRootNode.ChildNodes)
            {
                if (xmlNode.LocalName != ASSOCIATION)
                {
                    continue;
                }

                if (xmlNode.Attributes[ASSOCIATION_ID].Value == element.ID.ToString())
                {
                    xmlElement = xmlNode as XmlElement;
                    return index;
                }

                index++;
            }
            return -1;
        }

        private XmlElement FindXMLElement(XmlElement parent, string namespaceURI)
        {
            foreach (XmlElement node in parent.ChildNodes)
            {
                if (node.Attributes[XMLNS].Value == namespaceURI)
                {
                    return node;
                }
            }
            return null;
        }

        private XmlElement FindXMLElementByPosition(int position)
        {
            return xmlRootNode.ChildNodes.Item(position + 1) as XmlElement;
        }

        private XmlAttribute FindXMLElementAttribute(Element element, string attributeName, string namespaceURI)
        {
            XmlElement xmlElement = FindXMLElement(element);
            if (xmlElement == null)
            {
                return null;
            }
            else
            {
                if (IsCommonAttribute(attributeName))
                {
                    return xmlElement.Attributes[attributeName];
                }
                else
                {
                    XmlElement associationToolNode = FindXMLElement(xmlElement, namespaceURI);
                    if (associationToolNode != null)
                    {
                        return associationToolNode.Attributes[attributeName];
                    }
                    else
                    {
                        return null;
                    }
                }
            }
        }

        private bool IsCommonAttribute(string localName)
        {
            switch (localName)
            {
                case ASSOCIATION_ID:
                case ASSOCIATION_ITEM:
                case ASSOCIATION_ITEM_ICON:
                case ASSOCIATION_FRAGMENT:
                case LEVEL_OF_SYNCHRONIZATION:
                case DISPLAY_TEXT:
                case OPEN_WITH_DEFAULT:
                case CREATED_BY:
                case CREATED_ON:
                case MODIFIED_BY:
                case MODIFIED_ON:
                    return true;
                default:
                    return false;
            };
        }

        private string AssignXmlAttributeValue(string localName, Element element)
        {
            switch (localName)
            {
                case ASSOCIATION_ID:
                    return element.ID.ToString();
                case ASSOCIATION_ITEM:
                    return element.AssociationURI.ToString();
                case ASSOCIATION_ITEM_TYPE:
                    return element.AssociationType.ToString();
                case ASSOCIATION_ITEM_ICON:
                    return String.Empty;
                case ASSOCIATION_FRAGMENT:
                    return String.Empty;
                case LEVEL_OF_SYNCHRONIZATION:
                    return element.LevelOfSynchronization.ToString();
                case DISPLAY_TEXT:
                    return element.NoteText.ToString();
                case OPEN_WITH_DEFAULT:
                    return String.Empty;
                case CREATED_BY:
                    return PREFIX_PLANZ.ToString();
                case CREATED_ON:
                    return element.CreatedOn.ToShortDateString() + " " + element.CreatedOn.ToShortTimeString();
                case MODIFIED_ON:
                    return element.ModifiedOn.ToShortDateString() + " " + element.ModifiedOn.ToShortTimeString();
                case TYPE:
                    return element.Type.ToString();
                case ISVISIBLE:
                    if (element.IsVisible == System.Windows.Visibility.Visible)
                    {
                        return "True";
                    }
                    else
                    {
                        return "False";
                    }
                case ISCOLLAPSED:
                    return element.IsCollapsed.ToString();
                case STATUS:
                    return element.Status.ToString();
                case POWERD_STATUS:
                    return element.PowerDStatus.ToString();
                case POWERD_TIMESTAMP:
                    return element.PowerDTimeStamp.ToShortDateString() + " " + element.PowerDTimeStamp.ToShortTimeString();
                case FONT_COLOR:
                    return element.FontColor.ToString();
                case FLAG_STATUS:
                    return element.FlagStatus.ToString();
                //case POSITION:
                //    return element.Position.ToString();
                case XMLNS:
                    return NAMESPACE_PLANZ;
                default:
                    return String.Empty;
            };
        }

        public void InsertElementIntoXML(Element element)
        {
            try
            {
                if (FindXMLElement(element) == null)
                {
                    XmlElement xmlElement = xmlDoc.CreateElement(PREFIX_XOOML, ASSOCIATION, NAMESPACE_XOOML);
                    XmlElement xmlElementTool = xmlDoc.CreateElement(PREFIX_XOOML, ASSOCIATION_TOOL, NAMESPACE_XOOML);
                    
                    XmlAttribute associationID = xmlDoc.CreateAttribute(ASSOCIATION_ID);
                    XmlAttribute associationItem = xmlDoc.CreateAttribute(ASSOCIATION_ITEM);
                    XmlAttribute associationItemIcon = xmlDoc.CreateAttribute(ASSOCIATION_ITEM_ICON);
                    XmlAttribute associationFragment = xmlDoc.CreateAttribute(ASSOCIATION_FRAGMENT);
                    XmlAttribute levelOfSynchronization = xmlDoc.CreateAttribute(LEVEL_OF_SYNCHRONIZATION);
                    XmlAttribute displayText = xmlDoc.CreateAttribute(DISPLAY_TEXT);
                    XmlAttribute openWithDefault = xmlDoc.CreateAttribute(OPEN_WITH_DEFAULT);
                    XmlAttribute createdBy = xmlDoc.CreateAttribute(CREATED_BY);
                    XmlAttribute createdOn = xmlDoc.CreateAttribute(CREATED_ON);
                    XmlAttribute modifiedBy = xmlDoc.CreateAttribute(MODIFIED_BY);
                    XmlAttribute modifiedOn = xmlDoc.CreateAttribute(MODIFIED_ON);

                    XmlAttribute ns = xmlDoc.CreateAttribute(XMLNS);
                    XmlAttribute type = xmlDoc.CreateAttribute(TYPE);
                    XmlAttribute associationItemType = xmlDoc.CreateAttribute(ASSOCIATION_ITEM_TYPE);
                    XmlAttribute isVisible = xmlDoc.CreateAttribute(ISVISIBLE);
                    XmlAttribute isCollapsed = xmlDoc.CreateAttribute(ISCOLLAPSED);
                    XmlAttribute status = xmlDoc.CreateAttribute(STATUS);
                    XmlAttribute powerDStatus = xmlDoc.CreateAttribute(POWERD_STATUS);
                    XmlAttribute powerDTimeStamp = xmlDoc.CreateAttribute(POWERD_TIMESTAMP);
                    XmlAttribute fontColor = xmlDoc.CreateAttribute(FONT_COLOR);
                    XmlAttribute flagStatus = xmlDoc.CreateAttribute(FLAG_STATUS);
                    //XmlAttribute position = xmlDoc.CreateAttribute(POSITION);

                    associationID.Value = AssignXmlAttributeValue(ASSOCIATION_ID, element);
                    associationItem.Value = AssignXmlAttributeValue(ASSOCIATION_ITEM, element);
                    associationItemIcon.Value = AssignXmlAttributeValue(ASSOCIATION_ITEM_ICON, element);
                    associationFragment.Value = AssignXmlAttributeValue(ASSOCIATION_FRAGMENT, element);
                    levelOfSynchronization.Value = AssignXmlAttributeValue(LEVEL_OF_SYNCHRONIZATION, element);
                    displayText.Value = AssignXmlAttributeValue(DISPLAY_TEXT, element);
                    openWithDefault.Value = AssignXmlAttributeValue(OPEN_WITH_DEFAULT, element);
                    createdBy.Value = AssignXmlAttributeValue(CREATED_BY, element);
                    createdOn.Value = AssignXmlAttributeValue(CREATED_ON, element);
                    modifiedBy.Value = AssignXmlAttributeValue(MODIFIED_BY, element);
                    modifiedOn.Value = AssignXmlAttributeValue(MODIFIED_ON, element);

                    ns.Value = AssignXmlAttributeValue(XMLNS, element);
                    type.Value = AssignXmlAttributeValue(TYPE, element);
                    associationItemType.Value = AssignXmlAttributeValue(ASSOCIATION_ITEM_TYPE, element);
                    isVisible.Value = AssignXmlAttributeValue(ISVISIBLE, element);
                    isCollapsed.Value = AssignXmlAttributeValue(ISCOLLAPSED, element);
                    status.Value = AssignXmlAttributeValue(STATUS, element);
                    powerDStatus.Value = AssignXmlAttributeValue(POWERD_STATUS, element);
                    powerDTimeStamp.Value = AssignXmlAttributeValue(POWERD_TIMESTAMP, element);
                    fontColor.Value = AssignXmlAttributeValue(FONT_COLOR, element);
                    flagStatus.Value = AssignXmlAttributeValue(FLAG_STATUS, element);
                    //position.Value = AssignXmlAttributeValue(POSITION, element);

                    xmlElement.Attributes.Append(associationID);
                    xmlElement.Attributes.Append(associationItem);
                    xmlElement.Attributes.Append(associationItemIcon);
                    xmlElement.Attributes.Append(associationFragment);
                    xmlElement.Attributes.Append(levelOfSynchronization);
                    xmlElement.Attributes.Append(displayText);
                    xmlElement.Attributes.Append(openWithDefault);
                    xmlElement.Attributes.Append(createdBy);
                    xmlElement.Attributes.Append(createdOn);
                    xmlElement.Attributes.Append(modifiedBy);
                    xmlElement.Attributes.Append(modifiedOn);

                    xmlElementTool.Attributes.Append(ns);
                    xmlElementTool.Attributes.Append(type);
                    xmlElementTool.Attributes.Append(associationItemType);
                    xmlElementTool.Attributes.Append(isVisible);
                    xmlElementTool.Attributes.Append(isCollapsed);
                    xmlElementTool.Attributes.Append(status);
                    xmlElementTool.Attributes.Append(powerDStatus);
                    xmlElementTool.Attributes.Append(powerDTimeStamp);
                    xmlElementTool.Attributes.Append(fontColor);
                    xmlElementTool.Attributes.Append(flagStatus);
                    //xmlElementTool.Attributes.Append(position);

                    xmlElement.AppendChild(xmlElementTool);
                    
                    //xmlRootNode.AppendChild(xmlElement);

                    if (element.Position == -1)
                    {
                        xmlRootNode.InsertBefore(xmlElement, FindXMLElementByPosition(0));
                    }
                    else
                    {
                        xmlRootNode.InsertBefore(xmlElement, FindXMLElementByPosition(element.Position));
                    }

                    SaveXML();

                    //UpdatePosition(element);
                }
                else
                {
                    UpdateElementIntoXML(element);
                }
            }
            catch (Exception ex)
            {

            }
        }

        public void UpdateFragmentElementIntoXML(Element element)
        {
            try
            {
                XmlElement fragmentToolNode = FindXMLElement(xmlRootNode as XmlElement, NAMESPACE_PLANZ);

                if (fragmentToolNode != null)
                {
                    fragmentToolNode.Attributes[SHOW_ASSOCIATION_MARKED_DONE].Value = element.ShowAssociationMarkedDone.ToString();
                    fragmentToolNode.Attributes[SHOW_ASSOCIATION_MARKED_DEFER].Value = element.ShowAssociationMarkedDefer.ToString();
                    fragmentToolNode.Attributes[START_DATE].Value = element.StartDate.ToShortDateString() + " " + element.StartDate.ToShortTimeString();
                    fragmentToolNode.Attributes[DUE_DATE].Value = element.DueDate.ToShortDateString() + " " + element.DueDate.ToShortTimeString();

                    SaveXML();
                }
            }
            catch (Exception ex)
            {

            }
        }

        public void UpdateElementIntoXML(Element element)
        {
            try
            {
                element.ModifiedOn = DateTime.Now;

                XmlElement xmlElement = FindXMLElement(element);

                if (xmlElement != null)
                {
                    XmlElement associationToolNode = FindXMLElement(xmlElement, NAMESPACE_PLANZ);

                    foreach (XmlAttribute xmlAttribute in xmlElement.Attributes)
                    {
                        xmlAttribute.Value = AssignXmlAttributeValue(xmlAttribute.LocalName, element);
                    }

                    if (associationToolNode != null)
                    {
                        foreach (XmlAttribute xmlAttribute in associationToolNode.Attributes)
                        {
                            xmlAttribute.Value = AssignXmlAttributeValue(xmlAttribute.LocalName, element);
                        }
                    }

                    SaveXML();

                    UpdatePosition(element);
                }
            }
            catch (Exception ex)
            {

            }
        }

        // Update all positions of element.ParenetElement.Elements according to Element.Position
        private void UpdatePosition(Element element)
        {
            if (element.ParentElement != null)
            {
                int index = 0;
                for (int i = 1; i < xmlRootNode.ChildNodes.Count; i++)
                {
                    Element ele = element.ParentElement.Elements[index++];
                    while (ele.IsCommandNote)
                    {
                        ele = element.ParentElement.Elements[index++];
                    }
                    XmlElement xmlEle = xmlRootNode.ChildNodes[i] as XmlElement;

                    if (xmlEle.HasAttribute(ASSOCIATION_ID))
                    {
                        if (ele.ID.ToString() != xmlEle.Attributes[ASSOCIATION_ID].Value.ToString())
                        {
                            XmlElement newXmlEle = FindXMLElement(ele);
                            xmlRootNode.RemoveChild(newXmlEle);
                            xmlRootNode.InsertBefore(newXmlEle, FindXMLElementByPosition(i - 1));
                        }
                    }

                    if (index == element.ParentElement.Elements.Count)
                    {
                        break;
                    }
                }

                SaveXML();
            }
        }

        // Update all positions of List<Element> according to its position in the list
        private void UpdatePosition(List<Element> elementList)
        {
            while (xmlRootNode.ChildNodes.Count > 1)
            {
                xmlRootNode.RemoveChild(xmlRootNode.LastChild);
            }

            SaveXML();

            for (int i = elementList.Count - 1; i >= 0; i--)
            {
                Element ele = elementList[i];
                InsertElementIntoXML(ele);
            }

            //for (int i = 0; i < elementList.Count; i++)
            //{
            //    Element ele = elementList[i];
            //    InsertElementIntoXML(ele);
            //}
        }

        public void RemoveElementFromXML(Element element)
        {
            try
            {
                XmlElement xmlElement = FindXMLElement(element);

                if (xmlElement != null)
                {
                    xmlRootNode.RemoveChild(xmlElement);

                    SaveXML();

                    UpdatePosition(element);
                }
            }
            catch (Exception ex)
            {

            }
        }

        public bool CreateFolder(Element element)
        {
            try
            {
                // If duplex folder name
                string folderName = element.NoteText;
                if (Directory.Exists(element.Path))
                {
                    int index = 2;
                    while (Directory.Exists(element.Path))
                    {
                        folderName = element.NoteText + " (" + index.ToString() + ")";
                        element.NoteText = folderName;
                        element.Path = element.ParentElement.Path + HeadingNameConverter.ConvertFromHeadingNameToFolderName(element) + System.IO.Path.DirectorySeparatorChar;
                        index++;
                    }
                    element.NoteText = folderName;
                }

                element.AssociationType = ElementAssociationType.Folder;
                element.AssociationURI = folderName;
                element.TailImageSource = FileTypeHandler.GetIcon(ElementAssociationType.Folder, element.Path);

                Directory.CreateDirectory(element.Path);

                //UpdateElementIntoXML(element);
                OnElementUpdate(element, new EventArgs());

                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        public void RemoveFolder(Element element)
        {
            try
            {
                if (Directory.Exists(element.Path))
                {
                    Directory.Delete(element.Path, true);
                }
            }
            catch (Exception ex)
            {

            }
        }

        public void RenameFolder(Element element, string previousPath)
        {
            try
            {
                if (Directory.Exists(element.Path) == false)
                {
                    string folderName = System.IO.Path.GetFileName(System.IO.Path.GetDirectoryName(element.Path));
                    if (element.NoteText != String.Empty)
                    {
                        element.NoteText = HeadingNameConverter.ConvertFromFolderNameToHeadingName(folderName);
                    }
                    element.AssociationURI = folderName;

                    Directory.Move(previousPath, element.Path);

                    //UpdateElementIntoXML(element);
                    OnElementUpdate(element, new EventArgs());
                }
            }
            catch (Exception ex)
            {

            }
        }

        public void MoveFolder(Element element, string previousPath)
        {
            try
            {
                if (Directory.Exists(element.Path) == false)
                {
                    Directory.Move(previousPath, element.Path);
                }
            }
            catch (Exception ex)
            {

            }
        }

        // For NavigationTab use only
        public List<NavigationItem> GetAllHeadingElementFromXML()
        {
            List<NavigationItem> allHeadingElement = new List<NavigationItem>();

            bool showDone = false;
            bool showDeferred = false;

            foreach (XmlElement xmlElement in xmlRootNode.ChildNodes)
            {
                try
                {
                    if (xmlElement.LocalName != ASSOCIATION)
                    {
                        showDone = Boolean.Parse(xmlElement.Attributes[SHOW_ASSOCIATION_MARKED_DONE].Value);
                        showDeferred = Boolean.Parse(xmlElement.Attributes[SHOW_ASSOCIATION_MARKED_DEFER].Value);

                        continue;
                    }

                    Element element = new Element
                    {
                        ID = new Guid(xmlElement.Attributes[ASSOCIATION_ID].Value),
                        AssociationURI = xmlElement.Attributes[ASSOCIATION_ITEM].Value,
                        LevelOfSynchronization = Int32.Parse(xmlElement.Attributes[LEVEL_OF_SYNCHRONIZATION].Value),
                        NoteText = xmlElement.Attributes[DISPLAY_TEXT].Value,
                    };

                    XmlElement associationToolNode = FindXMLElement(xmlElement, NAMESPACE_PLANZ);
                    if (associationToolNode != null)
                    {
                        element.AssociationType = (ElementAssociationType)Enum.Parse(typeof(ElementAssociationType), associationToolNode.Attributes[ASSOCIATION_ITEM_TYPE].Value);
                        element.Type = (ElementType)Enum.Parse(typeof(ElementType), associationToolNode.Attributes[TYPE].Value);
                        element.IsExpanded = !Boolean.Parse(associationToolNode.Attributes[ISCOLLAPSED].Value);
                        element.Status = (ElementStatus)Enum.Parse(typeof(ElementStatus), associationToolNode.Attributes[STATUS].Value);
                        element.PowerDStatus = (PowerDStatus)Enum.Parse(typeof(PowerDStatus), associationToolNode.Attributes[POWERD_STATUS].Value);
                        element.FlagStatus = (FlagStatus)Enum.Parse(typeof(FlagStatus), associationToolNode.Attributes[FLAG_STATUS].Value);
                        element.FontColor = associationToolNode.Attributes[FONT_COLOR].Value;
                        //element.Position = Int32.Parse(associationToolNode.Attributes[POSITION].Value);
                        if (associationToolNode.Attributes[POWERD_TIMESTAMP].Value != String.Empty)
                        {
                            element.PowerDTimeStamp = DateTime.Parse(associationToolNode.Attributes[POWERD_TIMESTAMP].Value);
                        }
                    }

                    if (Boolean.Parse(associationToolNode.Attributes[ISVISIBLE].Value) == true)
                    {
                        element.IsVisible = System.Windows.Visibility.Visible;
                    }
                    else
                    {
                        element.IsVisible = System.Windows.Visibility.Collapsed;
                    }

                    if (element.IsHeading)
                    {
                        if((element.PowerDStatus == PowerDStatus.Deferred && !showDeferred)
                            || (element.PowerDStatus == PowerDStatus.Done && !showDone))
                            continue;

                        NavigationItem ni = new NavigationItem
                        {
                            Name = element.NoteText,
                            //Position = element.Position,
                        };

                        if (element.AssociationType == ElementAssociationType.FolderShortcut)
                        {
                            IWshRuntimeLibrary.WshShellClass wshShell = new IWshRuntimeLibrary.WshShellClass();
                            IWshRuntimeLibrary.IWshShortcut shortcut = (IWshRuntimeLibrary.IWshShortcut)wshShell.CreateShortcut(currentDir + element.AssociationURI);
                            if (shortcut != null)
                            {
                                ni.Path = shortcut.TargetPath + System.IO.Path.DirectorySeparatorChar;
                            }
                        }
                        else
                        {
                            ni.Path = currentDir + HeadingNameConverter.ConvertFromHeadingNameToFolderName(element) + System.IO.Path.DirectorySeparatorChar;
                        }

                        allHeadingElement.Add(ni);
                    }
                }
                catch (Exception)
                {
                    continue;
                }
            }

            //allHeadingElement.Sort(new SortByHeadingPosition());
            return allHeadingElement;
        }

        public void UpdateEllipsisRegions(List<Element> elements)
        {
            int hiddenAssociationCount = 0;
            int lastEllipsis = -2;
            for (int i = 0; i < elements.Count; i++)
            {
                if (elements[i].PowerDStatus != PowerDStatus.Done && 
                    elements[i].PowerDStatus != PowerDStatus.Deferred)
                {
                    if (elements[i].IsVisible == System.Windows.Visibility.Collapsed)
                    {
                        if (lastEllipsis == i - 1)
                        {
                            elements[i - 1].NoteText = String.Format("Click the icon on the right to show the next {0} association(s)", ++hiddenAssociationCount);
                            
                            elements.RemoveAt(i);
                            
                            i--;
                        }
                        else
                        {
                            lastEllipsis = i;
                            hiddenAssociationCount = 0;

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
                            
                            elements[i] = commandNote;
                        }
                    }
                }
            }
        }

        private void UpdatePowerDRegion(List<Element> elements)
        {
            // Move deferred-to-today association to top
            int todo_start = 0;
            for (int i = elements.Count - 1; i >= todo_start; i--)
            {
                if (elements[i].PowerDStatus == PowerDStatus.Deferred &&
                    elements[i].PowerDTimeStamp.Date <= DateTime.Now.Date)
                {
                    for (int j = 0; j < i - todo_start; j++)
                    {
                        Element temp = elements[i - j - 1]; 
                        elements[i - j - 1] = elements[i - j];
                        elements[i - j] = temp;
                    }
  
                    todo_start++;
                    i++;
                }
            }

            // Move deferred association to deferred region
            int defer_region_start = elements.Count;
            for (int i = elements.Count - 1; i >= todo_start; i--)
            {
                if (elements[i].PowerDStatus == PowerDStatus.Deferred ||
                    elements[i].PowerDStatus == PowerDStatus.Done)
                {
                    defer_region_start = i;
                }
                else
                {
                    break;
                }
            }
            for (int i = defer_region_start - 1; i >= todo_start; i--)
            {
                if (elements[i].PowerDStatus == PowerDStatus.Deferred)
                {
                    for (int j = 0; j < defer_region_start - i - 1; j++)
                    {
                        Element temp = elements[i + j + 1];
                        elements[i + j + 1] = elements[i + j];
                        elements[i + j] = temp;
                    }

                    --defer_region_start;
                }
            }

            // Move done association to done region
            int done_region_start = elements.Count;
            for (int i = elements.Count - 1; i >= todo_start; i--)
            {
                if (elements[i].PowerDStatus == PowerDStatus.Done)
                {
                    done_region_start = i;
                }
                else
                {
                    break;
                }
            }
            for (int i = done_region_start - 1; i >= todo_start; i--)
            {
                if (elements[i].PowerDStatus == PowerDStatus.Done)
                {
                    for (int j = 0; j < done_region_start - i - 1; j++)
                    {
                        Element temp = elements[i + j + 1];
                        elements[i + j + 1] = elements[i + j];
                        elements[i + j] = temp;
                    }

                    --done_region_start;
                }
            }
        }
    }


    #region Support Function

    internal class SortByPosition : IComparer<Element>
    {
        public int Compare(Element x, Element y)
        {
            if (x.Position > y.Position)
            {
                return 1;
            }
            else if (x.Position == y.Position)
            {
                return 0;
            }
            else
            {
                return -1;
            }
        }
    }

    internal class SortByHeadingPosition : IComparer<NavigationItem>
    {
        public int Compare(NavigationItem x, NavigationItem y)
        {
            if (x.Position > y.Position)
            {
                return 1;
            }
            else if (x.Position == y.Position)
            {
                return 0;
            }
            else
            {
                return -1;
            }
        }
    }

    internal class SortByLastModified : IComparer<FileInfo>
    {
        public int Compare(FileInfo x, FileInfo y)
        {
            return y.LastWriteTime.CompareTo(x.LastWriteTime);
        }
    }



    #endregion
}
