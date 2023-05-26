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
using OfficeOpenXml.Drawing.Style;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Drawing.Style.Effect
{

    /// <summary>
    /// The reflection effect
    /// </summary>
    public class ExcelDrawingReflectionEffect  : ExcelDrawingShadowEffectBase
    {
        private readonly string _directionPath = "{0}/@dir";
        private readonly string _startPositionPath = "{0}/@stPos";
        private readonly string _startOpacityPath = "{0}/@stA";
        private readonly string _endPositionPath = "{0}/@endPos";
        private readonly string _endOpacityPath = "{0}/@endA";
        private readonly string _fadeDirectionPath = "{0}/@fadeDir";
        private readonly string _shadowAlignmentPath = "{0}/@algn";
        private readonly string _rotateWithShapePath = "{0}/@rotWithShape";
        private readonly string _verticalSkewAnglePath = "{0}/@ky";
        private readonly string _horizontalSkewAnglePath = "{0}/@kx";
        private readonly string _verticalScalingFactorPath = "{0}/@sy";
        private readonly string _horizontalScalingFactorPath = "{0}/@sx";
        private readonly string _blurRadPath = "{0}/@blurRad";
        internal ExcelDrawingReflectionEffect(XmlNamespaceManager nameSpaceManager, XmlNode topNode, string[] schemaNodeOrder, string path) : base(nameSpaceManager, topNode, schemaNodeOrder, path)
        {
            this._startPositionPath = string.Format(this._startPositionPath, path);
            this._startOpacityPath = string.Format(this._startOpacityPath, path);
            this._endPositionPath = string.Format(this._endPositionPath, path);
            this._endOpacityPath = string.Format(this._endOpacityPath, path);
            this._fadeDirectionPath = string.Format(this._fadeDirectionPath, path);
            this._shadowAlignmentPath = string.Format(this._shadowAlignmentPath, path);
            this._rotateWithShapePath = string.Format(this._rotateWithShapePath, path);
            this._verticalSkewAnglePath = string.Format(this._verticalSkewAnglePath, path);
            this._horizontalSkewAnglePath = string.Format(this._horizontalSkewAnglePath, path);
            this._verticalScalingFactorPath = string.Format(this._verticalScalingFactorPath, path);
            this._horizontalScalingFactorPath = string.Format(this._horizontalScalingFactorPath, path);
            this._directionPath = string.Format(this._directionPath, path);
            this._blurRadPath = string.Format(this._blurRadPath, path);
        }
        /// <summary>
        /// The start position along the alpha gradient ramp of the alpha value.
        /// </summary>
        public double? StartPosition
        {
            get
            {
                return this.GetXmlNodePercentage(this._startPositionPath) ?? 0;
            }
            set
            {
                this.SetXmlNodePercentage(this._startPositionPath, value, false);
            }
        }
        /// <summary>
        /// The starting reflection opacity
        /// </summary>
        public double? StartOpacity
        {
            get
            {
                return this.GetXmlNodePercentage(this._startOpacityPath) ?? 100;
            }
            set
            {
                this.SetXmlNodePercentage(this._startOpacityPath, value, false);
            }
        }

        /// <summary>
        /// The end position along the alpha gradient ramp of the alpha value.
        /// </summary>
        public double? EndPosition
        {
            get
            {
                return this.GetXmlNodePercentage(this._endPositionPath) ?? 100;
            }
            set
            {
                this.SetXmlNodePercentage(this._endPositionPath, value, false);
            }
        }
        /// <summary>
        /// The ending reflection opacity
        /// </summary>
        public double? EndOpacity
        {
            get
            {
                return this.GetXmlNodePercentage(this._endOpacityPath) ?? 0;
            }
            set
            {
                this.SetXmlNodePercentage(this._endOpacityPath, value, false);
            }
        }
        /// <summary>
        /// The direction to offset the reflection
        /// </summary>
        public double? FadeDirection
        {
            get
            {
                return this.GetXmlNodeAngel(this._fadeDirectionPath, 90);
            }
            set
            {
                this.SetXmlNodeAngel(this._fadeDirectionPath, value, "FadeDirection", -90, 90);
            }
        }
        /// <summary>
        /// Alignment
        /// </summary>
        public eRectangleAlignment Alignment
        {
            get
            {
                return this.GetXmlNodeString(this._shadowAlignmentPath).TranslateRectangleAlignment();
            }
            set
            {
                if (value == eRectangleAlignment.Bottom)
                {
                    this.DeleteNode(this._shadowAlignmentPath);
                }
                else
                {
                    this.SetXmlNodeString(this._shadowAlignmentPath, value.TranslateString());
                }
            }
        }
        /// <summary>
        /// If the shadow rotates with the shape
        /// </summary>
        public bool RotateWithShape
        {
            get
            {
                return this.GetXmlNodeBool(this._rotateWithShapePath, true);
            }
            set
            {
                this.SetXmlNodeBool(this._rotateWithShapePath, value, true);
            }
        }
        /// <summary>
        /// Horizontal skew angle.
        /// Ranges from -90 to 90 degrees 
        /// </summary>
        public double? HorizontalSkewAngle
        {
            get
            {
                return this.GetXmlNodeAngel(this._horizontalSkewAnglePath);
            }
            set
            {
                this.SetXmlNodeAngel(this._horizontalSkewAnglePath, value, "HorizontalSkewAngle", -90, 90);
            }
        }
        /// <summary>
        /// Vertical skew angle.
        /// Ranges from -90 to 90 degrees 
        /// </summary>
        public double? VerticalSkewAngle
        {
            get
            {
                return this.GetXmlNodeAngel(this._verticalSkewAnglePath);
            }
            set
            {
                this.SetXmlNodeAngel(this._verticalSkewAnglePath, value, "HorizontalSkewAngle", -90, 90);
            }
        }
        /// <summary>
        /// Horizontal scaling factor in percentage .
        /// A negative value causes a flip.
        /// </summary>
        public double? HorizontalScalingFactor
        {
            get
            {
                return this.GetXmlNodePercentage(this._horizontalScalingFactorPath) ?? 100;
            }
            set
            {
                this.SetXmlNodePercentage(this._horizontalScalingFactorPath, value, true, 10000);
            }
        }
        /// <summary>
        /// Vertical scaling factor in percentage .
        /// A negative value causes a flip.
        /// </summary>
        public double? VerticalScalingFactor
        {
            get
            {
                return this.GetXmlNodePercentage(this._verticalScalingFactorPath) ?? 100;
            }
            set
            {
                this.SetXmlNodePercentage(this._verticalScalingFactorPath, value, true, 10000);
            }
        }
        /// <summary>
        /// The direction to offset the shadow
        /// </summary>
        public double? Direction
        {
            get
            {
                return this.GetXmlNodeAngel(this._directionPath);
            }
            set
            {
                this.SetXmlNodeAngel(this._directionPath, value, "Direction");
            }
        }
        /// <summary>
        /// The blur radius.
        /// </summary>
        public double? BlurRadius
        {
            get
            {
                return this.GetXmlNodeEmuToPt(this._blurRadPath);
            }
            set
            {
                this.SetXmlNodeEmuToPt(this._blurRadPath, value);
            }
        }
    }
}