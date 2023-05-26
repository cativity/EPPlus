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
    internal class TypeConvertUtil<TReturnType>
    {
        internal TypeConvertUtil(object o)
        {
            this.Value = new ValueWrapper(o);
            this.ReturnType = new ReturnTypeWrapper<TReturnType>();
        }

        public ReturnTypeWrapper<TReturnType> ReturnType
        {
            get;
            private set;
        }

        public ValueWrapper Value
        {
            get;
            private set;
        }

        public object ConvertToReturnType()
        {
            if (this.ReturnType.IsNullable && this.Value.IsEmptyString)
            {
                return null;
            }
            if (NumericTypeConversions.IsNumeric(this.ReturnType.Type))
            {
                if(NumericTypeConversions.TryConvert(this.Value.Object, out object convertedObj, this.ReturnType.Type))
                {
                    return convertedObj;
                }
                return default(TReturnType);
            }
            return this.Value.Object;
        }

        public bool TryGetDateTime(out object returnDate)
        {
            returnDate = default;
            if (!this.ReturnType.IsDateTime)
            {
                return false;
            }

            if (this.Value.Object is double)
            {
                returnDate = DateTime.FromOADate(this.Value.ToDouble());
                return true;
            }
            if (this.Value.IsTimeSpan)
            {
                returnDate = new DateTime(this.Value.ToTimeSpan().Ticks);
                return true;
            }
            if (this.Value.IsString)
            {
                if (DateTime.TryParse(this.Value.ToString(), out DateTime dt))
                {
                    returnDate = dt;
                    return true;
                }
            }
            return false;
        }

        public bool TryGetTimeSpan(out object timeSpan)
        {
            timeSpan = default;
            if (!this.ReturnType.IsTimeSpan)
            {
                return false;
            }

            if (this.Value.Object is long)
            {
                timeSpan = new TimeSpan(Convert.ToInt64(this.Value.Object));
                return true;
            }
            if(this.Value.Object is double)
            {
                timeSpan = new TimeSpan(DateTime.FromOADate((double)this.Value.Object).Ticks);
                return true;
            }
            if (this.Value.IsDateTime)
            {
                timeSpan = new TimeSpan(this.Value.ToDateTime().Ticks);
                return true;
            }
            if (this.Value.IsString)
            {
                if (TimeSpan.TryParse(this.Value.ToString(), out TimeSpan ts))
                {
                    timeSpan = ts;
                    return true;
                }
                throw new FormatException(this.Value.ToString() + " could not be parsed to a TimeSpan");
            }
            return false;
        }
    }
}
