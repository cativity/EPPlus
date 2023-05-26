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
using System.Text;

namespace OfficeOpenXml.Utils.TypeConversion
{
    internal class ValueWrapper
    {
        private readonly object _object;

        public ValueWrapper(object obj)
        {
            this._object = obj;
        }

        public bool IsString
        {
            get
            {
                if (this._object == null)
                {
                    return false;
                }

                return this._object is string;
            }
        }

        public bool IsEmptyString
        {
            get
            {
                if (this._object == null)
                {
                    return false;
                }

                return this._object is string && this._object.ToString().Trim() == string.Empty;
            }
        }

        public bool IsNumeric
        {
            get
            {
                if(this._object == null)
                {
                    return false;
                }

                return NumericTypeConversions.IsNumeric(this._object.GetType());
            }
        }

        public bool IsDateTime
        {
            get
            {
                return this._object is DateTime;
            }
        }

        public bool IsTimeSpan
        {
            get
            {
                return this._object is TimeSpan;
            }
        }

        public DateTime ToDateTime()
        {
            return (DateTime)this._object;
        }

        public TimeSpan ToTimeSpan()
        {
            return (TimeSpan)this._object;
        }

        public double ToDouble()
        {
            return Convert.ToDouble(this._object);
        }

        public override string ToString()
        {
            return this._object.ToString();
        }

        public object Object
        {
            get { return this._object; }
        }
    }
}
