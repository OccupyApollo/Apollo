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
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace Planz
{
    public partial class FlagWindow : Window
    {
        private Element currentElement;

        private List<CustomTimeSpan> timeSpanList = new List<CustomTimeSpan>();

        public FlagWindow(Element currentElement)
        {
            InitializeComponent();

            this.currentElement = currentElement;

            this.headingName.Text = currentElement.NoteText;

            AssignTimeSpan(true);
            AssignTimeSpan(false);

            foreach (CustomTimeSpan ts in timeSpanList)
            {
                this.startTimePicker.Items.Add(ts);
                this.dueTimePicker.Items.Add(ts);
            }

            AssignDefaultTimeSpanValue();

            this.startDatePicker.SelectedDate = DateTime.Now;
            this.dueDatePicker.SelectedDate = DateTime.Now;

            this.addReminder.IsEnabled = false;
            this.startTimePicker.IsEnabled = false;
            this.dueTimePicker.IsEnabled = false;
            this.startAllDay.IsChecked = true;
            this.dueAllDay.IsChecked = true;
            this.start.IsChecked = false;
            this.due.IsChecked = false;
        }

        private void AssignTimeSpan(bool isAM)
        {
            timeSpanList.Add(new CustomTimeSpan { Hour = 12, Minutes = 0, IsAM = isAM, });
            timeSpanList.Add(new CustomTimeSpan { Hour = 12, Minutes = 30, IsAM = isAM, });
            for (int i = 1; i <= 11; i++)
            {
                timeSpanList.Add(new CustomTimeSpan { Hour = i, Minutes = 0, IsAM = isAM, });
                timeSpanList.Add(new CustomTimeSpan { Hour = i, Minutes = 30, IsAM = isAM, });
            }
        }

        private void AssignDefaultTimeSpanValue()
        {
            DateTime current = DateTime.Now;
            bool isAM = true;
            if (current.Hour >= 12)
            {
                isAM = false;
            }
            int hour = current.Hour;
            int minutes = 30;
            if (current.Minute >= 30)
            {
                minutes = 0;
                hour++;
                if (hour == 24)
                {
                    hour = 0;
                    isAM = true;
                }
            }
            if (hour == 0 || hour == 12)
            {
                hour = 12;
            }
            else if (hour > 12)
            {
                hour = hour - 12;
            }

            for (int i = 0; i < timeSpanList.Count; i++)
            {
                if (timeSpanList[i].Hour == hour &&
                    timeSpanList[i].Minutes == minutes &&
                    timeSpanList[i].IsAM == isAM)
                {
                    this.startTimePicker.SelectedIndex = i;
                    this.dueTimePicker.SelectedIndex = i;
                    break;
                }
            }
        }

        public bool AddToToday
        {
            get;
            set;
        }

        public bool AddToReminder
        {
            get;
            set;
        }

        public bool AddToTask
        {
            get;
            set;
        }

        public CustomTimeSpan StartTime
        {
            get;
            set;
        }

        public CustomTimeSpan DueTime
        {
            get;
            set;
        }

        public DateTime StartDate
        {
            get;
            set;
        }

        public DateTime DueDate
        {
            get;
            set;
        }

        public bool StartAllDay
        {
            get;
            set;
        }

        public bool DueAllDay
        {
            get;
            set;
        }

        public bool HasStart
        {
            get;
            set;
        }

        public bool HasDue
        {
            get;
            set;
        }

        private void OK_Click(object sender, RoutedEventArgs e)
        {
            StartTime = this.startTimePicker.SelectedItem as CustomTimeSpan;
            DueTime = this.dueTimePicker.SelectedItem as CustomTimeSpan;

            StartDate = this.startDatePicker.SelectedDate.Value;
            DueDate = this.dueDatePicker.SelectedDate.Value;

            this.DialogResult = true;
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
        }

        private void startAllDay_Checked(object sender, RoutedEventArgs e)
        {
            this.startTimePicker.IsEnabled = false;
            StartAllDay = true;
        }

        private void startAllDay_Unchecked(object sender, RoutedEventArgs e)
        {
            this.startTimePicker.IsEnabled = true;
            StartAllDay = false;
        }

        private void dueAllDay_Checked(object sender, RoutedEventArgs e)
        {
            this.dueTimePicker.IsEnabled = false;
            DueAllDay = true;
        }

        private void dueAllDay_Unchecked(object sender, RoutedEventArgs e)
        {
            this.dueTimePicker.IsEnabled = true;
            DueAllDay = false;
        }

        private void start_Checked(object sender, RoutedEventArgs e)
        {
            this.addReminder.IsEnabled = true;
            HasStart = true;
        }

        private void start_Unchecked(object sender, RoutedEventArgs e)
        {
            if (this.due.IsChecked == false)
            {
                this.addReminder.IsEnabled = false;
            }
            HasStart = false;
        }

        private void due_Checked(object sender, RoutedEventArgs e)
        {
            this.addReminder.IsEnabled = true;
            HasDue = true;
        }

        private void due_Unchecked(object sender, RoutedEventArgs e)
        {
            if (this.start.IsChecked == false)
            {
                this.addReminder.IsEnabled = false;
            }
            HasDue = false;
        }

        private void addToToday_Checked(object sender, RoutedEventArgs e)
        {
            AddToToday = true;
        }

        private void addToToday_Unchecked(object sender, RoutedEventArgs e)
        {
            AddToToday = false;
        }

        private void addReminder_Checked(object sender, RoutedEventArgs e)
        {
            AddToReminder = true;
        }

        private void addReminder_Unchecked(object sender, RoutedEventArgs e)
        {
            AddToReminder = false;
        }

        private void addTask_Checked(object sender, RoutedEventArgs e)
        {
            AddToTask = true;
        }

        private void addTask_Unchecked(object sender, RoutedEventArgs e)
        {
            AddToTask = false;
        }

        private void startDatePicker_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            this.start.IsChecked = true;
        }

        private void dueDatePicker_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            this.due.IsChecked = true;
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