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

namespace Planz
{
    public static class JournalControl
    {
        public static void CreateJournalFolders(int year, string path)
        {
            string journalFolder = path + year.ToString() + System.IO.Path.DirectorySeparatorChar;
            Directory.CreateDirectory(journalFolder);

            DateTime year0101 = new DateTime(year, 1, 1);

            bool isLeapYear = false;
            if (year % 4 == 0)
            {
                isLeapYear = true;
            }

            int daysOfMonth = 30;
            int startDayOfMonth = 1;

            #region Create month folder

            for (int i = 1; i <= 12; i++)
            {
                // Format: 2009-01, January
                string month = year.ToString() + "-";
                if (i < 10)
                {
                    month += "0" + i.ToString();
                }
                else
                {
                    month += i.ToString();
                }
                month += ", ";
                switch (i)
                {
                    case 1:
                        month += "January";
                        daysOfMonth = 31;
                        break;
                    case 2:
                        month += "Februray";
                        if (isLeapYear)
                        {
                            daysOfMonth = 29;
                        }
                        else
                        {
                            daysOfMonth = 28;
                        }
                        break;
                    case 3:
                        month += "March";
                        daysOfMonth = 31;
                        break;
                    case 4:
                        month += "April";
                        daysOfMonth = 30;
                        break;
                    case 5:
                        month += "May";
                        daysOfMonth = 31;
                        break;
                    case 6:
                        month += "June";
                        daysOfMonth = 30;
                        break;
                    case 7:
                        month += "July";
                        daysOfMonth = 31;
                        break;
                    case 8:
                        month += "August";
                        daysOfMonth = 31;
                        break;
                    case 9:
                        month += "September";
                        daysOfMonth = 30;
                        break;
                    case 10:
                        month += "October";
                        daysOfMonth = 31;
                        break;
                    case 11:
                        month += "November";
                        daysOfMonth = 30;
                        break;
                    case 12:
                        month += "December";
                        daysOfMonth = 31;
                        break;
                }

                Directory.CreateDirectory(journalFolder + month + System.IO.Path.DirectorySeparatorChar);

                #region Create week folder

                // Format: Week of 2009-01-04
                int start = startDayOfMonth, end = 1;
                if (i == 1)
                {
                    switch (year0101.DayOfWeek)
                    {
                        case DayOfWeek.Monday:
                            end = 6;
                            break;
                        case DayOfWeek.Tuesday:
                            end = 5;
                            break;
                        case DayOfWeek.Wednesday:
                            end = 4;
                            break;
                        case DayOfWeek.Thursday:
                            end = 3;
                            break;
                        case DayOfWeek.Friday:
                            end = 2;
                            break;
                        case DayOfWeek.Saturday:
                            end = 1;
                            break;
                        case DayOfWeek.Sunday:
                            end = 0;
                            break;
                    }

                    if (end != 0)
                    {
                        string weekPath = journalFolder + month + System.IO.Path.DirectorySeparatorChar +
                            "Week of " + year.ToString() + "-01-01" + System.IO.Path.DirectorySeparatorChar;
                        Directory.CreateDirectory(weekPath);

                        int start_p = (int)year0101.DayOfWeek;
                        for (int j = 1; j <= end; j++)
                        {
                            string day = year.ToString() + "-01-0" + j.ToString() + ", ";
                            switch (start_p++ % 7)
                            {
                                case 0:
                                    day += "Sunday";
                                    break;
                                case 1:
                                    day += "Monday";
                                    break;
                                case 2:
                                    day += "Tuesday";
                                    break;
                                case 3:
                                    day += "Wednesday";
                                    break;
                                case 4:
                                    day += "Thursday";
                                    break;
                                case 5:
                                    day += "Friday";
                                    break;
                                case 6:
                                    day += "Saturday";
                                    break;
                            }

                            string dayPath = weekPath + day + System.IO.Path.DirectorySeparatorChar;
                            Directory.CreateDirectory(dayPath);
                        }
                    }

                    start += end;
                }
                while (start <= daysOfMonth)
                {
                    string week = "Week of " + year.ToString() + "-";
                    if (i < 10)
                    {
                        week += "0" + i.ToString();
                    }
                    else
                    {
                        week += i.ToString();
                    }
                    week += "-";

                    if (start < 10)
                    {
                        week += "0" + start.ToString();
                    }
                    else
                    {
                        week += start.ToString();
                    }

                    string weekPath = journalFolder + month + System.IO.Path.DirectorySeparatorChar +
                        week + System.IO.Path.DirectorySeparatorChar;
                    Directory.CreateDirectory(weekPath);

                    #region Create day folder

                    for (int j = 0; j < 7; j++)
                    {
                        string day = year.ToString() + "-";
                        if (start + j > daysOfMonth)
                        {
                            if (i == 12)
                            {
                                break;
                            }

                            int nextMonth = i + 1;
                            if (nextMonth < 10)
                            {
                                day += "0" + nextMonth.ToString();
                            }
                            else
                            {
                                day += nextMonth.ToString();
                            }
                        }
                        else
                        {
                            if (i < 10)
                            {
                                day += "0" + i.ToString();
                            }
                            else
                            {
                                day += i.ToString();
                            }
                        }
                        day += "-";
                        if (start + j < 10)
                        {
                            day += "0" + (start + j).ToString();
                        }
                        else
                        {
                            if (start + j > daysOfMonth)
                            {
                                day += "0" + (start + j - daysOfMonth).ToString();
                            }
                            else
                            {
                                day += (start + j).ToString();
                            }
                        }
                        day += ", ";

                        switch (j)
                        {
                            case 0:
                                day += "Sunday";
                                break;
                            case 1:
                                day += "Monday";
                                break;
                            case 2:
                                day += "Tuesday";
                                break;
                            case 3:
                                day += "Wednesday";
                                break;
                            case 4:
                                day += "Thursday";
                                break;
                            case 5:
                                day += "Friday";
                                break;
                            case 6:
                                day += "Saturday";
                                break;
                        }

                        string dayPath = weekPath + day + System.IO.Path.DirectorySeparatorChar;
                        Directory.CreateDirectory(dayPath);
                    }

                    #endregion

                    start += 7;
                }

                startDayOfMonth = start - daysOfMonth;

                #endregion
            }

            #endregion
        }

        public static string GetJournalPath(DateTime dt)
        {
            try
            {
                // Algorithm needs to be changed, this one is O(500)

                string path = StartProcess.JOURNAL_PATH;
                path += dt.Year.ToString() + System.IO.Path.DirectorySeparatorChar;

                string target = dt.Year.ToString() + "-";
                if (dt.Month < 10)
                {
                    target += "0" + dt.Month.ToString() + "-";
                }
                else
                {
                    target += dt.Month.ToString() + "-";
                }
                if (dt.Day < 10)
                {
                    target += "0" + dt.Day.ToString() + ", ";
                }
                else
                {
                    target += dt.Day.ToString() + ", ";
                }
                target += dt.DayOfWeek.ToString();

                DirectoryInfo root = new DirectoryInfo(path);
                foreach (DirectoryInfo di_month in root.GetDirectories())
                {
                    foreach (DirectoryInfo di_week in di_month.GetDirectories())
                    {
                        foreach (DirectoryInfo di_day in di_week.GetDirectories())
                        {
                            if (di_day.Name == target)
                            {
                                return di_day.FullName;
                            }
                        }
                    }
                }
                
                return null;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public static bool JournalFolderExists(int year, string path)
        {
            string journalFolder = path + year.ToString() + System.IO.Path.DirectorySeparatorChar;
            return Directory.Exists(journalFolder);
        }

        public static List<string> GetJournalPathDaysAhead(DateTime dt, int days)
        {
            List<string> daysAhead = new List<string>();

            if (dt.AddDays(days - 1).Year == dt.Year + 1 &&
                JournalFolderExists(dt.Year + 1, StartProcess.JOURNAL_PATH) == false)
            {
                CreateJournalFolders(dt.Year + 1, StartProcess.JOURNAL_PATH);
            }

            for (int i = 0; i < days; i++)
            {
                string path = GetJournalPath(dt.AddDays(i));
                if (path != null)
                {
                    daysAhead.Add(path);
                }
            }

            return daysAhead;
        }

        public static DateTime GetDateTime(string path)
        {
            char[] ds = { '\\' };

            int index = path.TrimEnd(ds).LastIndexOf('\\');
            String date = path.Substring(index+1,10);
            DateTime dt = DateTime.Parse(date);

            return dt;
        }
    }
}
