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
using System.Net;
using System.IO;
using System.Web;
using System.Xml;

namespace Planz
{
    public class Tweet
    {
        public string Username
        {
            get;
            set;
        }

        public string Password
        {
            get;
            set;
        }

        public string Message
        {
            get;
            set;
        }

        public string MessageID
        {
            get;
            set;
        }

        public string AccountName
        {
            get;
            set;
        }
    }

    public class TwitterControl
    {
        private Tweet tweet;
        private const string source = "Planz";

        public TwitterControl(Tweet tweet)
        {
            this.tweet = tweet;
        }

        public bool Update()
        {
            try
            {
                Uri actionUri = new Uri(string.Format("http://twitter.com/statuses/update.xml?status={0}&source={1}", HttpUtility.UrlEncode(tweet.Message), source));

                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(actionUri);
                request.Method = "POST";
                request.MaximumAutomaticRedirections = 4;
                request.MaximumResponseHeadersLength = 4;
                request.ContentLength = 0;
                request.ServicePoint.Expect100Continue = false;
                request.Credentials = new NetworkCredential(tweet.Username, tweet.Password);

                HttpWebResponse response = (HttpWebResponse)request.GetResponse();

                Stream receiveStream = response.GetResponseStream();
                StreamReader readStream = new StreamReader(receiveStream, Encoding.UTF8);
                string responseString = readStream.ReadToEnd();

                XmlDocument responseXML = new XmlDocument();
                responseXML.LoadXml(responseString);
                if (responseXML.DocumentElement != null)
                {
                    switch (responseXML.DocumentElement.Name.ToLower())
                    {
                        case "status":
                            ParseStatusNode(responseXML.DocumentElement);
                            break;
                        default:
                            throw new Exception("Invalid Response.");
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        private void ParseStatusNode(XmlNode xmlNode)
        {
            tweet.MessageID = xmlNode["id"].InnerText;
            ParseUserNode(xmlNode["user"]);
        }

        private void ParseUserNode(XmlNode xmlNode)
        {
            tweet.AccountName = xmlNode["screen_name"].InnerText;
        }

        public string GetTweetURL()
        {
            return "http://twitter.com/" + tweet.AccountName + "/status/" + tweet.MessageID;
        }
    }
}
