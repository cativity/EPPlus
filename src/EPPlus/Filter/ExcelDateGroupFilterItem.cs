/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Filter
{
    /// <summary>
    /// A date group for filters
    /// </summary>
    public class ExcelFilterDateGroupItem : ExcelFilterItem
    {
        /// <summary>
        /// Filter out the specified year
        /// </summary>
        /// <param name="year">The year</param>
        public ExcelFilterDateGroupItem(int year)
        {
            this.Grouping = eDateTimeGrouping.Year;
            this.Year = year;
            this.Validate();
        }

        /// <summary>
        /// Filter out the specified year and month
        /// </summary>
        /// <param name="year">The year</param>
        /// <param name="month">The month</param>
        public ExcelFilterDateGroupItem(int year, int month)
        {
            this.Grouping = eDateTimeGrouping.Month;
            this.Year = year;
            this.Month = month;
            this.Validate();
        }
        /// <summary>
        /// Filter out the specified year, month and day
        /// </summary>
        /// <param name="year">The year</param>
        /// <param name="month">The month</param>
        /// <param name="day">The day</param>
        public ExcelFilterDateGroupItem(int year, int month, int day)
        {
            this.Grouping = eDateTimeGrouping.Day;
            this.Year = year;
            this.Month = month;
            this.Day = day;
            this.Validate();
        }
        /// <summary>
        /// Filter out the specified year, month, day and hour
        /// </summary>
        /// <param name="year">The year</param>
        /// <param name="month">The month</param>
        /// <param name="day">The day</param>
        /// <param name="hour">The hour</param>
        public ExcelFilterDateGroupItem(int year, int month, int day, int hour)
        {
            this.Grouping = eDateTimeGrouping.Hour;
            this.Year = year;
            this.Month = month;
            this.Day = day;
            this.Hour = hour;
            this.Validate();
        }
        /// <summary>
        /// Filter out the specified year, month, day, hour and and minute
        /// </summary>
        /// <param name="year">The year</param>
        /// <param name="month">The month</param>
        /// <param name="day">The day</param>
        /// <param name="hour">The hour</param>
        /// <param name="minute">The minute</param>
        public ExcelFilterDateGroupItem(int year, int month, int day, int hour, int minute)
        {
            this.Grouping = eDateTimeGrouping.Minute;
            this.Year = year;
            this.Month = month;
            this.Day = day;
            this.Hour = hour;
            this.Minute = minute;
            this.Validate();
        }
        /// <summary>
        /// Filter out the specified year, month, day, hour and and minute
        /// </summary>
        /// <param name="year">The year</param>
        /// <param name="month">The month</param>
        /// <param name="day">The day</param>
        /// <param name="hour">The hour</param>
        /// <param name="minute">The minute</param>
        /// <param name="second">The second</param>
        public ExcelFilterDateGroupItem(int year, int month, int day, int hour, int minute, int second)
        {
            this.Grouping = eDateTimeGrouping.Second;
            this.Year = year;
            this.Month = month;
            this.Day = day;
            this.Hour = hour;
            this.Minute = minute;
            this.Second = second;
            this.Validate();
        }
        private void Validate()
        {
            if (this.Year < 0 && this.Year > 9999)
            {
                throw (new ArgumentException("Year out of range(0-9999)"));
            }

            if (this.Grouping == eDateTimeGrouping.Year)
            {
                return;
            }

            if (this.Month < 1 && this.Month > 12)
            {
                throw (new ArgumentException("Month out of range(1-12)"));
            }
            if (this.Grouping == eDateTimeGrouping.Month)
            {
                return;
            }

            if (this.Day < 1 && this.Day > 31)
            {
                throw (new ArgumentException("Month out of range(1-31)"));
            }
            if (this.Grouping == eDateTimeGrouping.Day)
            {
                return;
            }

            if (this.Hour < 0 && this.Hour > 23)
            {
                throw (new ArgumentException("Hour out of range(0-23)"));
            }
            if (this.Grouping == eDateTimeGrouping.Hour)
            {
                return;
            }

            if (this.Minute < 0 && this.Minute > 59)
            {
                throw (new ArgumentException("Minute out of range(0-59)"));
            }
            if (this.Grouping == eDateTimeGrouping.Minute)
            {
                return;
            }

            if (this.Second < 0 && this.Second > 59)
            {
                throw (new ArgumentException("Second out of range(0-59)"));
            }
        }

        internal void AddNode(XmlNode node)
        {
            XmlElement? e = node.OwnerDocument.CreateElement("dateGroupItem", ExcelPackage.schemaMain);
            e.SetAttribute("dateTimeGrouping", this.Grouping.ToString().ToLower());
            e.SetAttribute("year", this.Year.ToString(CultureInfo.InvariantCulture));

            if (this.Month.HasValue)
            {
                e.SetAttribute("month", this.Month.Value.ToString(CultureInfo.InvariantCulture));
                if (this.Day.HasValue)
                {
                    e.SetAttribute("day", this.Day.Value.ToString(CultureInfo.InvariantCulture));
                    if (this.Hour.HasValue)
                    {
                        e.SetAttribute("hour", this.Hour.Value.ToString(CultureInfo.InvariantCulture));
                        if (this.Minute.HasValue)
                        {
                            e.SetAttribute("minute", this.Minute.Value.ToString(CultureInfo.InvariantCulture));
                            if (this.Second.HasValue)
                            {
                                e.SetAttribute("second", this.Second.Value.ToString(CultureInfo.InvariantCulture));
                            }
                        }
                    }
                }
            }

            node.AppendChild(e);
        }

        internal bool Match(DateTime value)
        {
            bool match = value.Year == this.Year;

            if(match && this.Month.HasValue)
            {
                match = value.Month == this.Month;
                if(match && this.Day.HasValue)
                {
                    match = value.Day == this.Day;
                    if (match && this.Hour.HasValue)
                    {
                        match = value.Hour == this.Hour;
                        if (match && this.Minute.HasValue)
                        {
                            match = value.Minute == this.Minute;
                            if (match && this.Second.HasValue)
                            {
                                match = value.Second == this.Second;
                            }
                        }
                    }
                }
            }
            return match;
        }
        /// <summary>
        /// The grouping. Is set depending on the selected constructor
        /// </summary>
        public eDateTimeGrouping Grouping{ get; }
        /// <summary>
        /// Year to filter on
        /// </summary>
        public int Year { get; }
        /// <summary>
        /// Month to filter on
        /// </summary>
        public int? Month { get; }
        /// <summary>
        /// Day to filter on
        /// </summary>
        public int? Day { get; }
        /// <summary>
        /// Hour to filter on
        /// </summary>
        public int? Hour { get; }
        /// <summary>
        /// Minute to filter on
        /// </summary>
        public int? Minute { get;  }
        /// <summary>
        /// Second to filter on
        /// </summary>
        public int? Second { get;  }
    }
}