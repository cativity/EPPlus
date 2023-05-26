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
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;

namespace OfficeOpenXml.DataValidation;

/// <summary>
/// Represents a time between 00:00:00 and 23:59:59
/// </summary>
public class ExcelTime
{
    private event EventHandler _timeChanged;
    private readonly decimal SecondsPerDay = 3600 * 24;
    private readonly decimal SecondsPerHour = 3600;
    private readonly decimal SecondsPerMinute = 60;
    /// <summary>
    /// Max number of decimals when rounding.
    /// </summary>
    public const int NumberOfDecimals = 15;

    /// <summary>
    /// Default constructor
    /// </summary>
    public ExcelTime()
    {

    }

    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="value">An existing time for initialization</param>
    public ExcelTime(decimal value)
    {
        if (value < 0M)
        {
            throw new ArgumentException("Value cannot be less than 0");
        }
        else if (value >= 1M)
        {
            throw new ArgumentException("Value cannot be greater or equal to 1");
        }

        this.Init(value);
    }

    private void Init(decimal value)
    {
        // handle hour
        decimal totalSeconds = value * this.SecondsPerDay;
        decimal hour = Math.Floor(totalSeconds / this.SecondsPerHour);
        this.Hour = (int)hour;

        // handle minute
        decimal remainingSeconds = totalSeconds - (hour * this.SecondsPerHour);
        decimal minute = Math.Floor(remainingSeconds / this.SecondsPerMinute);
        this.Minute = (int)minute;

        // handle second
        remainingSeconds = totalSeconds - (hour * this.SecondsPerHour) - (minute * this.SecondsPerMinute);
        decimal second = Math.Round(remainingSeconds, MidpointRounding.AwayFromZero);
        // Second might be rounded to 60... the SetSecond method handles that.
        this.SetSecond((int)second);
    }

    /// <summary>
    /// If we are unlucky second might be rounded up to 60. This will have the minute to be raised and might affect the hour.
    /// </summary>
    /// <param name="value"></param>
    private void SetSecond(int value)
    {
        if (value == 60)
        {
            this.Second = 0;
            int minute = this.Minute + 1;
            this.SetMinute(minute);
        }
        else
        {
            this.Second = value;
        }
    }

    private void SetMinute(int value)
    {
        if (value == 60)
        {
            this.Minute = 0;
            int hour = this.Hour + 1;
            this.SetHour(hour);
        }
        else
        {
            this.Minute = value;
        }
    }

    private void SetHour(int value)
    {
        if (value == 24)
        {
            this.Hour = 0;
        }
        else
        {
            this.Hour = value;
        }
    }

    internal event EventHandler TimeChanged
    {
        add { this._timeChanged += value; }
        remove { this._timeChanged -= value; }
    }

    private void OnTimeChanged()
    {
        if (this._timeChanged != null)
        {
            this._timeChanged(this, EventArgs.Empty);
        }
    }

    private int _hour;
    /// <summary>
    /// Hour between 0 and 23
    /// </summary>
    public int Hour 
    {
        get
        {
            return this._hour;
        }
        set
        {
            if (value < 0)
            {
                throw new InvalidOperationException("Value for hour cannot be negative");
            }
            if (value > 23)
            {
                throw new InvalidOperationException("Value for hour cannot be greater than 23");
            }

            this._hour = value;
            this.OnTimeChanged();
        }
    }

    private int _minute;
    /// <summary>
    /// Minute between 0 and 59
    /// </summary>
    public int Minute
    {
        get
        {
            return this._minute;
        }
        set
        {
            if (value < 0)
            {
                throw new InvalidOperationException("Value for minute cannot be negative");
            }
            if (value > 59)
            {
                throw new InvalidOperationException("Value for minute cannot be greater than 59");
            }

            this._minute = value;
            this.OnTimeChanged();
        }
    }

    private int? _second;
    /// <summary>
    /// Second between 0 and 59
    /// </summary>
    public int? Second
    {
        get
        {
            return this._second;
        }
        set
        {
            if (value < 0)
            {
                throw new InvalidOperationException("Value for second cannot be negative");
            }
            if (value > 59)
            {
                throw new InvalidOperationException("Value for second cannot be greater than 59");
            }

            this._second = value;
            this.OnTimeChanged();
        }
    }

    private static decimal Round(decimal value)
    {
        return Math.Round(value, NumberOfDecimals);
    }

    private decimal ToSeconds()
    {
        decimal result = this.Hour * this.SecondsPerHour;
        result += this.Minute * this.SecondsPerMinute;
        result += this.Second ?? 0;
        return (decimal)result;
    }

    /// <summary>
    /// Returns the excel decimal representation of a time.
    /// </summary>
    /// <returns></returns>
    public decimal ToExcelTime()
    {
        decimal seconds = this.ToSeconds();
        return Round(seconds / (decimal)this.SecondsPerDay);
    }

    /// <summary>
    /// Returns the excel decimal representation of a time as a string.
    /// </summary>
    /// <returns></returns>
    public string ToExcelString()
    {
        return this.ToExcelTime().ToString(CultureInfo.InvariantCulture);
    }
    /// <summary>
    /// Converts the object to a string
    /// </summary>
    /// <returns>The string</returns>
    public override string ToString()
    {
        int second = this.Second ?? 0;
        return string.Format("{0}:{1}:{2}",
                             this.Hour < 10 ? "0" + this.Hour.ToString() : this.Hour.ToString(),
                             this.Minute < 10 ? "0" + this.Minute.ToString() : this.Minute.ToString(),
                             second < 10 ? "0" + second.ToString() : second.ToString());
    }

}