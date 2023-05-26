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
using System.Xml;
using OfficeOpenXml.Table.PivotTable;

namespace OfficeOpenXml.Drawing.Chart;

/// <summary>
/// A Surface chart
/// </summary>
public sealed class ExcelSurfaceChart : ExcelChartStandard
{
    #region "Constructors"
    internal ExcelSurfaceChart(ExcelDrawings drawings, XmlNode node, eChartType? type, ExcelChart topChart, ExcelPivotTable PivotTableSource, XmlDocument chartXml, ExcelGroupShape parent = null) :
        base(drawings, node, type, topChart, PivotTableSource, chartXml, parent)
    {
        this.Init();
    }
    internal ExcelSurfaceChart(ExcelDrawings drawings, XmlNode node, Uri uriChart, Packaging.ZipPackagePart part, XmlDocument chartXml, XmlNode chartNode, ExcelGroupShape parent = null) :
        base(drawings, node, uriChart, part, chartXml, chartNode, parent)
    {
        this.Init();
    }

    internal ExcelSurfaceChart(ExcelChart topChart, XmlNode chartNode, ExcelGroupShape parent = null) : 
        base(topChart, chartNode, parent)
    {
        this.Init();
    }
    private void Init()
    {
        this.SetTypeProperties();
    }
    #endregion
    internal override void InitSeries(ExcelChart chart, XmlNamespaceManager ns, XmlNode node, bool isPivot, List<ExcelChartSerie> list = null)
    {
        base.InitSeries(chart, ns, node, isPivot, list);
        this.Series.Init(chart, ns, node, isPivot, base.Series._list);
    }
    const string WIREFRAME_PATH = "c:wireframe/@val";
    /// <summary>
    /// The surface chart is drawn as a wireframe
    /// </summary>
    public bool Wireframe
    {
        get
        {
            return this._chartXmlHelper.GetXmlNodeBool(WIREFRAME_PATH);
        }
        set
        {
            this._chartXmlHelper.SetXmlNodeBool(WIREFRAME_PATH, value);
        }
    }        
    internal void SetTypeProperties()
    {
        if(this.ChartType==eChartType.SurfaceWireframe || this.ChartType==eChartType.SurfaceTopViewWireframe)
        {
            this.Wireframe=true;
        }
        else 
        {
            this.Wireframe=false;
        }

        if(this.ChartType==eChartType.SurfaceTopView || this.ChartType==eChartType.SurfaceTopViewWireframe)
        {
            this.View3D.RotY = 0;
            this.View3D.RotX = 90;
        }
        else
        {
            this.View3D.RotY = 20;
            this.View3D.RotX = 15;
        }

        this.View3D.RightAngleAxes = false;
        this.View3D.Perspective = 0;
        this.Axis[1].CrossBetween = eCrossBetween.MidCat;
    }
    internal override eChartType GetChartType(string name)
    {
        if(this.Wireframe)
        {
            if (name == "surfaceChart")
            {
                return eChartType.SurfaceTopViewWireframe;
            }
            else
            {
                return eChartType.SurfaceWireframe;
            }
        }
        else
        {
            if (name == "surfaceChart")
            {
                return eChartType.SurfaceTopView;
            }
            else
            {
                return eChartType.Surface;
            }
        }
    }
    /// <summary>
    /// A collection of series for a Surface Chart
    /// </summary>
    public new ExcelChartSeries<ExcelSurfaceChartSerie> Series { get; } = new ExcelChartSeries<ExcelSurfaceChartSerie>();
}