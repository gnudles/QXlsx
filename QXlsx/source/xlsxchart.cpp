// xlsxchart.cpp

#include <QtGlobal>
#include <QString>
#include <QIODevice>
#include <QXmlStreamReader>
#include <QXmlStreamWriter>
#include <QDebug>
#include <QDateTime>
#include <QDate>
#include <QTime>

#include "xlsxchart_p.h"
#include "xlsxworksheet.h"
#include "xlsxcellrange.h"
#include "xlsxutility_p.h"

#include "xdattr.h"
#include "xdnode.h"
#include "xdxmldomreader.h"

QT_BEGIN_NAMESPACE_XLSX

/*!
 * \internal
 */
Chart::Chart(AbstractSheet *parent, CreateFlag flag)
    : AbstractOOXmlFile(new ChartPrivate(this, flag))
{
    d_func()->sheet = parent;
}

/*!
 * Destroys the chart.
 */
Chart::~Chart()
{
}

ChartPrivate::ChartPrivate(Chart *q, Chart::CreateFlag flag)
    : AbstractOOXmlFilePrivate(q, flag), chartType(static_cast<Chart::ChartType>(0))
{

}

ChartPrivate::~ChartPrivate()
{
}


/*!
 * Add the data series which is in the range \a range of the \a sheet.
 */
void Chart::addSeries(const CellRange &range, AbstractSheet *sheet)
{
    Q_D(Chart);

    if (!range.isValid())
        return;
    if (sheet && sheet->sheetType() != AbstractSheet::ST_WorkSheet)
        return;
    if (!sheet && d->sheet->sheetType() != AbstractSheet::ST_WorkSheet)
        return;

    QString sheetName = sheet ? sheet->sheetName() : d->sheet->sheetName();
    //In case sheetName contains space or '
    sheetName = escapeSheetName(sheetName);

    if (range.columnCount() == 1 || range.rowCount() == 1)
    {
        QSharedPointer<XlsxSeries> series = QSharedPointer<XlsxSeries>(new XlsxSeries);
        series->numberDataSource_numRef = sheetName + QLatin1String("!") + range.toString(true, true);
        d->seriesList.append(series);
    }
    else if (range.columnCount() < range.rowCount())
    {
        //Column based series
        int firstDataColumn = range.firstColumn();
        QString axDataSouruce_numRef;
        if (d->chartType == CT_ScatterChart || d->chartType == CT_BubbleChart)
        {
            firstDataColumn += 1;
            CellRange subRange(range.firstRow(), range.firstColumn(), range.lastRow(), range.firstColumn());
            axDataSouruce_numRef = sheetName + QLatin1String("!") + subRange.toString(true, true);
        }

        for (int col=firstDataColumn; col<=range.lastColumn(); ++col)
        {
            CellRange subRange(range.firstRow(), col, range.lastRow(), col);
            QSharedPointer<XlsxSeries> series = QSharedPointer<XlsxSeries>(new XlsxSeries);
            series->axDataSource_numRef = axDataSouruce_numRef;
            series->numberDataSource_numRef = sheetName + QLatin1String("!") + subRange.toString(true, true);
            d->seriesList.append(series);
        }
    }
    else
    {
        //Row based series
        int firstDataRow = range.firstRow();
        QString axDataSouruce_numRef;
        if (d->chartType == CT_ScatterChart || d->chartType == CT_BubbleChart)
        {
            firstDataRow += 1;
            CellRange subRange(range.firstRow(), range.firstColumn(), range.firstRow(), range.lastColumn());
            axDataSouruce_numRef = sheetName + QLatin1String("!") + subRange.toString(true, true);
        }

        for (int row=firstDataRow; row<=range.lastRow(); ++row)
        {
            CellRange subRange(row, range.firstColumn(), row, range.lastColumn());
            QSharedPointer<XlsxSeries> series = QSharedPointer<XlsxSeries>(new XlsxSeries);
            series->axDataSource_numRef = axDataSouruce_numRef;
            series->numberDataSource_numRef = sheetName + QLatin1String("!") + subRange.toString(true, true);
            d->seriesList.append(series);
        }
    }
}

/*!
 * Set the type of the chart to \a type
 */
void Chart::setChartType(ChartType type)
{
    Q_D(Chart);

    d->chartType = type;
}

/*!
 * \internal
 *
 */
void Chart::setChartStyle(int id)
{
    Q_UNUSED(id)
    //!Todo
}

void Chart::setAxisTitle(Chart::ChartAxisPos pos, QString axisTitle)
{
    Q_D(Chart);

    if ( axisTitle.isEmpty() )
        return;

    // dev24 : fixed for old compiler
    if ( pos == Chart::Left )
    {
        d->axisNames[ XlsxAxis::Left ] = axisTitle;
    }
    else if ( pos == Chart::Top )
    {
        d->axisNames[ XlsxAxis::Top ] = axisTitle;
    }
    else if ( pos == Chart::Right )
    {
        d->axisNames[ XlsxAxis::Right ] = axisTitle;
    }
    else if ( pos == Chart::Bottom )
    {
        d->axisNames[ XlsxAxis::Bottom ] = axisTitle;
    }

}

// dev25
void Chart::setChartTitle(QString strchartTitle)
{
    Q_D(Chart);

    d->chartTitle = strchartTitle;
}
/*
    <chartSpace>
        <chart>
            <view3D>
                <perspective val="30"/>
            </view3D>
            <plotArea>
                <layout/>
                <barChart>
                ...
                </barChart>
                <catAx/>
                <valAx/>
            </plotArea>
            <legend>
            ...
            </legend>
        </chart>
        <printSettings>
        </printSettings>
    </chartSpace>
*/
void Chart::saveToXmlFile(QIODevice *device) const
{
    Q_D(const Chart);

    QXmlStreamWriter writer(device);

    writer.writeStartDocument(QStringLiteral("1.0"), true);

    // L.4.13.2.2 Chart
    //
    //  chartSpace is the root node, which contains an element defining the chart,
    // and an element defining the print settings for the chart.
    writer.writeStartElement(QStringLiteral("c:chartSpace"));

    writer.writeAttribute(QStringLiteral("xmlns:c"),
                          QStringLiteral("http://schemas.openxmlformats.org/drawingml/2006/chart"));
    writer.writeAttribute(QStringLiteral("xmlns:a"),
                          QStringLiteral("http://schemas.openxmlformats.org/drawingml/2006/main"));
    writer.writeAttribute(QStringLiteral("xmlns:r"),
                          QStringLiteral("http://schemas.openxmlformats.org/officeDocument/2006/relationships"));

    /*
    * chart is the root element for the chart. If the chart is a 3D chart,
    * then a view3D element is contained, which specifies the 3D view.
    * It then has a plot area, which defines a layout and contains an element
    * that corresponds to, and defines, the type of chart.
    */
    d->saveXmlChart(writer);

    writer.writeEndElement();// c:chartSpace
    writer.writeEndDocument();
}

/*
    <chartSpace>
        <chart>
            <view3D>
                <perspective val="30"/>
            </view3D>
            <plotArea>
                <layout/>
                <barChart>
                ...
                </barChart>
                <catAx/>
                <valAx/>
            </plotArea>
            <legend>
            ...
            </legend>
        </chart>
        <printSettings>
        </printSettings>
    </chartSpace>
*/

// #define LOADING_CHART_TYPE_1 // 1 is old type.

bool Chart::loadFromXmlFile(QIODevice *device)
{
    Q_D(Chart);

#ifdef LOADING_CHART_TYPE_1

    QXmlStreamReader reader(device);
    while (!reader.atEnd())
    {
        reader.readNextStartElement();
        if (reader.tokenType() == QXmlStreamReader::StartElement)
        {
            if (reader.name() == QLatin1String("chartSpace"))
            {
                // qDebug() << " root of chart";
            }
            else if (reader.name() == QLatin1String("chart"))
            {
                if (!d->loadXmlChart(reader))
                {
                    return false;
                }
            }
            else  if (reader.name() == QLatin1String("lang"))
            {
                //!TODO : language
            }
            else  if (reader.name() == QLatin1String("printSettings"))
            {
                //!TODO : print settings
            }
            else
            {
            }
        }
    }

    return true;

#else

    bool ret;
    ret = d->loadFromXmlFile( device );
    return ret;

#endif
}

bool ChartPrivate::loadXmlChart(QXmlStreamReader &reader)
{
    Q_ASSERT(reader.name() == QLatin1String("chart"));

    while (!reader.atEnd())
    {
        reader.readNextStartElement();
        if (reader.tokenType() == QXmlStreamReader::StartElement)
        {
            if (reader.name() == QLatin1String("plotArea"))
            {
                if ( !loadXmlPlotArea(reader) )
                {
                    return false;
                }
            }
            else if (reader.name() == QLatin1String("title"))
            {
                if ( !loadXmlChartTitle(reader) )
                {
                    // return false;
                }
            }
            else if (reader.name() == QLatin1String("legend"))
            {
                //!Todo
            }
            else
            {
            }
        }
        else if (reader.tokenType() == QXmlStreamReader::EndElement &&
                 reader.name() == QLatin1String("chart") )
        {
            break;
        }
    }

    return true;
}

/*
<xsd:complexType name="CT_ChartSpace">
    <xsd:sequence>
...
    </xsd:sequence>
</xsd:complexType>
*/
bool ChartPrivate::loadFromXmlFile(QIODevice *device)
{
    XMLDOM::XMLDOMReader domReader;
    if ( !domReader.load(device) )
    {
        // qDebug() << "failed to load";
        return false;
    }

    // <xsd:element name="chart" type="CT_Chart" minOccurs="1" maxOccurs="1"/>
    XMLDOM::Node* ptrChart = domReader.findNode( 1, "c:chart" );
    if ( NULL == ptrChart )
    {
        return false; // 'chart' is mandatory field
    }
    if ( ! load1Chart(&domReader, ptrChart) )
    {
        return false;
    }

    // <xsd:element name="lang" type="CT_TextLanguageID" minOccurs="0" maxOccurs="1"/>
    XMLDOM::Node* ptrLang = domReader.findNode( 1, "c:lang" );
    if ( NULL != ptrLang )
    {
        load1Lang(&domReader, ptrLang);
    }

    // <xsd:element name="printSettings" type="CT_PrintSettings" minOccurs="0" maxOccurs="1"/>
    XMLDOM::Node* ptrPrinterSetting = domReader.findNode( 1, "c:printSettings" );
    if ( NULL != ptrPrinterSetting )
    {
        load1PrinterSettings(&domReader, ptrPrinterSetting);
    }

    //! TODO:
    // <xsd:element name="date1904" type="CT_Boolean" minOccurs="0" maxOccurs="1"/>
    // <xsd:element name="roundedCorners" type="CT_Boolean" minOccurs="0" maxOccurs="1"/>
    // <xsd:element name="style" type="CT_Style" minOccurs="0" maxOccurs="1"/>
    // <xsd:element name="clrMapOvr" type="a:CT_ColorMapping" minOccurs="0" maxOccurs="1"/>
    // <xsd:element name="pivotSource" type="CT_PivotSource" minOccurs="0" maxOccurs="1"/>
    // <xsd:element name="protection" type="CT_Protection" minOccurs="0" maxOccurs="1"/>
    // <xsd:element name="spPr" type="a:CT_ShapeProperties" minOccurs="0" maxOccurs="1"/>
    // <xsd:element name="txPr" type="a:CT_TextBody" minOccurs="0" maxOccurs="1"/>
    // <xsd:element name="externalData" type="CT_ExternalData" minOccurs="0" maxOccurs="1"/>
    // <xsd:element name="userShapes" type="CT_RelId" minOccurs="0" maxOccurs="1"/>
    // <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>

    domReader.clear();

    return true;
}

/*
<xsd:complexType name="CT_Chart">
    <xsd:sequence>
    ...
    </xsd:sequence>
</xsd:complexType>
*/
bool ChartPrivate::load1Chart(XMLDOM::XMLDOMReader *pReader, XMLDOM::Node* ptrChart)
{
    if ( NULL == ptrChart )
        return false;

    // <xsd:element name="plotArea" type="CT_PlotArea" minOccurs="1" maxOccurs="1"/>
    XMLDOM::Node* ptrPlotArea = pReader->findNode( ptrChart, "c:plotArea" );
    if ( NULL == ptrPlotArea )
    {
        return false; // 'plotArea' is mandatory field
    }
    if ( ! load2PlotArea( pReader, ptrPlotArea ) )
    {
        return false;
    }

    // <xsd:element name="title" type="CT_Title" minOccurs="0" maxOccurs="1"/>
    XMLDOM::Node* ptrTitle = pReader->findNode( ptrChart, "c:title" );
    if ( NULL != ptrTitle )
    {
        load2Title( pReader, ptrTitle );
    }

    // <xsd:element name="legend" type="CT_Legend" minOccurs="0" maxOccurs="1"/>
    XMLDOM::Node* ptrLegend = pReader->findNode( ptrChart, "c:legend" );
    if ( NULL != ptrLegend )
    {
        load2Legend( pReader, ptrLegend );
    }

    // <xsd:element name="plotVisOnly" type="CT_Boolean" minOccurs="0" maxOccurs="1"/>
    XMLDOM::Node* ptrPlotVisOnly = pReader->findNode( ptrChart, "c:plotVisOnly" );
    if ( NULL != ptrPlotVisOnly )
    {
        load2PlotVisOnly( pReader, ptrPlotVisOnly );
    }

    //! TODO:
    // <xsd:element name="autoTitleDeleted" type="CT_Boolean"       minOccurs="0" maxOccurs="1"/>
    // <xsd:element name="pivotFmts"        type="CT_PivotFmts"     minOccurs="0" maxOccurs="1"/>
    // <xsd:element name="view3D"           type="CT_View3D"        minOccurs="0" maxOccurs="1"/>
    // <xsd:element name="floor"            type="CT_Surface"       minOccurs="0" maxOccurs="1"/>
    // <xsd:element name="sideWall"         type="CT_Surface"       minOccurs="0" maxOccurs="1"/>
    // <xsd:element name="backWall"         type="CT_Surface"       minOccurs="0" maxOccurs="1"/>
    // <xsd:element name="dispBlanksAs"     type="CT_DispBlanksAs"  minOccurs="0" maxOccurs="1"/>
    // <xsd:element name="showDLblsOverMax" type="CT_Boolean"       minOccurs="0" maxOccurs="1"/>
    // <xsd:element name="extLst"           type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>

    return true;
}

/*
<xsd:complexType name="CT_Title">
    <xsd:sequence>
...
    </xsd:sequence>
</xsd:complexType>
*/
bool ChartPrivate::load2Title(XMLDOM::XMLDOMReader* pReader, XMLDOM::Node* ptrTitle )
{
    // c:title

    // <xsd:element name="tx" type="CT_Tx" minOccurs="0" maxOccurs="1"/>
    // <xsd:element name="layout" type="CT_Layout" minOccurs="0" maxOccurs="1"/>
    // <xsd:element name="overlay" type="CT_Boolean" minOccurs="0" maxOccurs="1"/>
    // <xsd:element name="spPr" type="a:CT_ShapeProperties" minOccurs="0" maxOccurs="1"/>
    // <xsd:element name="txPr" type="a:CT_TextBody" minOccurs="0" maxOccurs="1"/>
    // <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>

    return true;
}

/*
<xsd:complexType name="CT_PlotArea">
    <xsd:sequence>
        ...
    </xsd:sequence>
</xsd:complexType>
*/
bool ChartPrivate::load2PlotArea(XMLDOM::XMLDOMReader* pReader, XMLDOM::Node* ptrPlotArea)
{
    // c:plotarea

    // debugging code
    /*
    for (int cindex = 0 ; cindex < ptrPlotArea->childNode.size() ; cindex ++)
    {
        XMLDOM::Node* pNode = ptrPlotArea->childNode.at( cindex );
        if ( NULL == pNode )
            continue;

        qDebug() << pNode->level << pNode->nodeName << pNode->nodeText ;
    }
    // */

    // <xsd:element name="layout" type="CT_Layout" minOccurs="0" maxOccurs="1"/>
    XMLDOM::Node* ptrLayout = pReader->findNode( ptrPlotArea, "c:layout" );
    if ( NULL != ptrLayout )
    {
        // bool lLayout = load3Layout( pReader, ptrLayout );
    }

    bool loadingChart = false;

    XMLDOM::Node* ptrAreaChart = pReader->findNode( ptrPlotArea, "c:areaChart" );
    if ( NULL != ptrAreaChart )
    {
        loadingChart = load3AreaChart( pReader, ptrAreaChart );
    }

    XMLDOM::Node* ptrArea3DChart = pReader->findNode( ptrPlotArea, "c:area3DChart" );
    if ( NULL != ptrArea3DChart )
    {
        loadingChart = load3Aread3DChart( pReader, ptrArea3DChart );
    }

    XMLDOM::Node* ptrLineChart = pReader->findNode( ptrPlotArea, "c:lineChart" );
    if ( NULL != ptrLineChart )
    {
        loadingChart = load3LineChart( pReader, ptrLineChart );
    }

    XMLDOM::Node* ptrLine3DChart = pReader->findNode( ptrPlotArea, "c:line3DChart" );
    if ( NULL != ptrLine3DChart )
    {
        loadingChart = load3Line3DChart( pReader, ptrLine3DChart );
    }

    XMLDOM::Node* ptrStockChart = pReader->findNode( ptrPlotArea, "c:stockChart" );
    if ( NULL != ptrStockChart )
    {
        loadingChart = load3StockChart( pReader, ptrStockChart );
    }

    XMLDOM::Node* ptrRadarChart = pReader->findNode( ptrPlotArea, "c:radarChart" );
    if ( NULL != ptrRadarChart )
    {
        loadingChart = load3RadarChart( pReader, ptrRadarChart );
    }

    XMLDOM::Node* ptrScatterChart = pReader->findNode( ptrPlotArea, "c:scatterChart" );
    if ( NULL != ptrScatterChart )
    {
        loadingChart = load3SactterChart( pReader, ptrScatterChart );
    }

    XMLDOM::Node* ptrPiChart = pReader->findNode( ptrPlotArea, "c:pieChart" );
    if ( NULL != ptrPiChart )
    {
        loadingChart = load3PieChart( pReader, ptrPiChart );
    }

    XMLDOM::Node* ptrPie3DChart = pReader->findNode( ptrPlotArea, "c:pie3DChart" );
    if ( NULL != ptrPie3DChart )
    {
        loadingChart = load3Pie3DChart( pReader, ptrPie3DChart );
    }

    XMLDOM::Node* ptrDoughnutChart = pReader->findNode( ptrPlotArea, "c:doughnutChart" );
    if ( NULL != ptrDoughnutChart )
    {
        loadingChart = load3DoughnutChart( pReader, ptrDoughnutChart );
    }

    XMLDOM::Node* ptrBarChart = pReader->findNode( ptrPlotArea, "c:barChart" );
    if ( NULL != ptrBarChart )
    {
        loadingChart = load3BarChart( pReader, ptrBarChart );
    }

    XMLDOM::Node* ptrBar3DChart = pReader->findNode( ptrPlotArea, "c:bar3DChart" );
    if ( NULL != ptrBar3DChart )
    {
        loadingChart = load3Bar3DChart( pReader, ptrBar3DChart );
    }

    XMLDOM::Node* ptrOfPieChart = pReader->findNode( ptrPlotArea, "c:ofPieChart" );
    if ( NULL != ptrOfPieChart )
    {
        loadingChart = load3OfPieChart( pReader, ptrOfPieChart );
    }

    XMLDOM::Node* ptrSurfaceChart = pReader->findNode( ptrPlotArea, "c:surfaceChart" );
    if ( NULL != ptrSurfaceChart )
    {
        loadingChart = load3SurfaceChart( pReader, ptrSurfaceChart );
    }

    XMLDOM::Node* ptrSurface3DChart = pReader->findNode( ptrPlotArea, "c:surface3DChart" );
    if ( NULL != ptrSurface3DChart )
    {
        loadingChart = load3Surface3DChart( pReader, ptrSurface3DChart );
    }

    XMLDOM::Node* ptrBubbleChart = pReader->findNode( ptrPlotArea, "c:bubbleChart" );
    if ( NULL != ptrBubbleChart )
    {
        loadingChart = load3BubbleChart( pReader, ptrBubbleChart );
    }

    if (!loadingChart)
    {
        qDebug() << "[debug] failed to load any chart";
        return false;
    }

    // <xsd:element name="dTable" type="CT_DTable" minOccurs="0" maxOccurs="1"/>
    XMLDOM::Node* ptrDTable = pReader->findNode( ptrPlotArea, "c:dTable" );
    //!TODO: load3DTable()

    // <xsd:element name="spPr" type="a:CT_ShapeProperties" minOccurs="0" maxOccurs="1"/>
    XMLDOM::Node* ptrSpPr = pReader->findNode( ptrPlotArea, "c:spPr" );
    //!TODO: load3SpPr()

    // <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
    XMLDOM::Node* ptrExtLst = pReader->findNode( ptrPlotArea, "c:extLst" );
    //!TODO: load3ExtLst()

    return true;
}

bool ChartPrivate::load3AreaChart(XMLDOM::XMLDOMReader* pReader, XMLDOM::Node* ptrAreaChart )
{
/*
<xsd:complexType name="CT_AreaChart">
    <xsd:sequence>
        <xsd:group   ref="EG_AreaChartShared"              minOccurs="1" maxOccurs="1"/>
        <xsd:element name="axId"   type="CT_UnsignedInt"   minOccurs="2" maxOccurs="2"/>
        <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
</xsd:complexType>
*/


    return true;
}

bool ChartPrivate::load3Aread3DChart(XMLDOM::XMLDOMReader* pReader, XMLDOM::Node* ptrArea3DChart )
{
/*
<xsd:complexType name="CT_Area3DChart">
    <xsd:sequence>
        <xsd:group ref="EG_AreaChartShared" minOccurs="1" maxOccurs="1"/>
        <xsd:element name="gapDepth" type="CT_GapAmount"     minOccurs="0" maxOccurs="1"/>
        <xsd:element name="axId"     type="CT_UnsignedInt"   minOccurs="2" maxOccurs="3"/>
        <xsd:element name="extLst"   type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
</xsd:complexType>
*/


    return true;
}

bool ChartPrivate::load3LineChart(XMLDOM::XMLDOMReader* pReader, XMLDOM::Node* ptrLineChart )
{
/*
<xsd:complexType name="CT_LineChart">
    <xsd:sequence>
        <xsd:group   ref="EG_LineChartShared" minOccurs="1" maxOccurs="1"/>
        <xsd:element name="hiLowLines" type="CT_ChartLines" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="upDownBars" type="CT_UpDownBars" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="marker"     type="CT_Boolean" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="smooth"     type="CT_Boolean" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="axId"       type="CT_UnsignedInt" minOccurs="2" maxOccurs="2"/>
        <xsd:element name="extLst"     type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
</xsd:complexType>
*/

    return true;
}

bool ChartPrivate::load3Line3DChart(XMLDOM::XMLDOMReader* pReader, XMLDOM::Node* ptrLine3DChart )
{
/*
<xsd:complexType name="CT_Line3DChart">
    <xsd:sequence>
        <xsd:group   ref="EG_LineChartShared" minOccurs="1" maxOccurs="1"/>
        <xsd:element name="gapDepth" type="CT_GapAmount"     minOccurs="0" maxOccurs="1"/>
        <xsd:element name="axId"     type="CT_UnsignedInt"   minOccurs="3" maxOccurs="3"/>
        <xsd:element name="extLst"   type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
</xsd:complexType>
*/

    return true;
}

bool ChartPrivate::load3StockChart(XMLDOM::XMLDOMReader* pReader, XMLDOM::Node* ptrStockChart )
{
/*
<xsd:complexType name="CT_StockChart">
    <xsd:sequence>
        <xsd:element name="ser"        type="CT_LineSer"       minOccurs="3" maxOccurs="4"/>
        <xsd:element name="dLbls"      type="CT_DLbls"         minOccurs="0" maxOccurs="1"/>
        <xsd:element name="dropLines"  type="CT_ChartLines"    minOccurs="0" maxOccurs="1"/>
        <xsd:element name="hiLowLines" type="CT_ChartLines"    minOccurs="0" maxOccurs="1"/>
        <xsd:element name="upDownBars" type="CT_UpDownBars"    minOccurs="0" maxOccurs="1"/>
        <xsd:element name="axId"       type="CT_UnsignedInt"   minOccurs="2" maxOccurs="2"/>
        <xsd:element name="extLst"     type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
</xsd:complexType>
*/

    return true;
}

bool ChartPrivate::load3RadarChart(XMLDOM::XMLDOMReader* pReader, XMLDOM::Node* ptrRadarChart )
{
/*
<xsd:complexType name="CT_RadarChart">
    <xsd:sequence>
        <xsd:element name="radarStyle" type="CT_RadarStyle"    minOccurs="1" maxOccurs="1"/>
        <xsd:element name="varyColors" type="CT_Boolean"       minOccurs="0" maxOccurs="1"/>
        <xsd:element name="ser"        type="CT_RadarSer"      minOccurs="0" maxOccurs="unbounded"/>
        <xsd:element name="dLbls"      type="CT_DLbls"         minOccurs="0" maxOccurs="1"/>
        <xsd:element name="axId"       type="CT_UnsignedInt"   minOccurs="2" maxOccurs="2"/>
        <xsd:element name="extLst"     type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
</xsd:complexType>
*/

    return true;
}

bool ChartPrivate::load3SactterChart(XMLDOM::XMLDOMReader* pReader, XMLDOM::Node* ptrScatterChart )
{
/*
<xsd:complexType name="CT_ScatterChart">
    <xsd:sequence>
        <xsd:element name="scatterStyle" type="CT_ScatterStyle" minOccurs="1" maxOccurs="1"/>
        <xsd:element name="varyColors" type="CT_Boolean" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="ser" type="CT_ScatterSer" minOccurs="0" maxOccurs="unbounded"/>
        <xsd:element name="dLbls" type="CT_DLbls" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="axId" type="CT_UnsignedInt" minOccurs="2" maxOccurs="2"/>
        <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
</xsd:complexType>
*/

    return true;
}

bool ChartPrivate::load3PieChart(XMLDOM::XMLDOMReader* pReader, XMLDOM::Node* ptrPiChart )
{
/*
<xsd:complexType name="CT_PieChart">
    <xsd:sequence>
        <xsd:group   ref="EG_PieChartShared" minOccurs="1" maxOccurs="1"/>
        <xsd:element name="firstSliceAng"   type="CT_FirstSliceAng" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="extLst"          type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
</xsd:complexType>
*/

    return true;
}

bool ChartPrivate::load3Pie3DChart(XMLDOM::XMLDOMReader* pReader, XMLDOM::Node* ptrPie3DChart )
{
/*
<xsd:complexType name="CT_Pie3DChart">
    <xsd:sequence>
        <xsd:group   ref="EG_PieChartShared" minOccurs="1" maxOccurs="1"/>
        <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
</xsd:complexType>
*/

    return true;
}

bool ChartPrivate::load3DoughnutChart(XMLDOM::XMLDOMReader* pReader, XMLDOM::Node* ptrDoughnutChart )
{
/*
<xsd:complexType name="CT_DoughnutChart">
    <xsd:sequence>
        <xsd:group   ref="EG_PieChartShared" minOccurs="1" maxOccurs="1"/>
        <xsd:element name="firstSliceAng" type="CT_FirstSliceAng" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="holeSize"      type="CT_HoleSize"      minOccurs="0" maxOccurs="1"/>
        <xsd:element name="extLst"        type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
</xsd:complexType>
*/

    return true;
}

/*
<xsd:group name="EG_BarChartShared">
    <xsd:sequence>
        <xsd:element name="barDir"     type="CT_BarDir"      minOccurs="1" maxOccurs="1"/>
        <xsd:element name="grouping"   type="CT_BarGrouping" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="varyColors" type="CT_Boolean"     minOccurs="0" maxOccurs="1"/>
        <xsd:element name="ser"        type="CT_BarSer"      minOccurs="0" maxOccurs="unbounded"/>
        <xsd:element name="dLbls"      type="CT_DLbls"       minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
</xsd:group>

<xsd:complexType name="CT_BarGrouping">
    <xsd:attribute name="val" type="ST_BarGrouping" default="clustered"/>
</xsd:complexType>

<xsd:simpleType name="ST_BarGrouping">
    <xsd:restriction base="xsd:string">
        <xsd:enumeration value="percentStacked"/>
        <xsd:enumeration value="clustered"/>
        <xsd:enumeration value="standard"/>
        <xsd:enumeration value="stacked"/>
    </xsd:restriction>
</xsd:simpleType>

<xsd:complexType name="CT_BarSer">
    <xsd:sequence>
        <xsd:group   ref="EG_SerShared" minOccurs="1" maxOccurs="1"/>
        <xsd:element name="invertIfNegative" type="CT_Boolean"        minOccurs="0" maxOccurs="1"/>
        <xsd:element name="pictureOptions"   type="CT_PictureOptions" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="dPt"              type="CT_DPt"            minOccurs="0" maxOccurs="unbounded"/>
        <xsd:element name="dLbls"            type="CT_DLbls"          minOccurs="0" maxOccurs="1"/>
        <xsd:element name="trendline"        type="CT_Trendline"      minOccurs="0" maxOccurs="unbounded"/>
        <xsd:element name="errBars"          type="CT_ErrBars"        minOccurs="0" maxOccurs="1"/>
        <xsd:element name="cat"              type="CT_AxDataSource"   minOccurs="0" maxOccurs="1"/>
        <xsd:element name="val"              type="CT_NumDataSource"  minOccurs="0" maxOccurs="1"/>
        <xsd:element name="shape"            type="CT_Shape"          minOccurs="0" maxOccurs="1"/>
        <xsd:element name="extLst"           type="CT_ExtensionList"  minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
</xsd:complexType>

<xsd:complexType name="CT_DLbls">
    <xsd:sequence>
        <xsd:element name="dLbl" type="CT_DLbl" minOccurs="0" maxOccurs="unbounded"/>
        <xsd:choice>
            <xsd:element name="delete" type="CT_Boolean" minOccurs="1" maxOccurs="1"/>
            <xsd:group   ref="Group_DLbls" minOccurs="1" maxOccurs="1"/>
        </xsd:choice>
        <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
</xsd:complexType>
*/
bool ChartPrivate::load3BarChart(XMLDOM::XMLDOMReader* pReader, XMLDOM::Node* ptrBarChart )
{
/*
<xsd:complexType name="CT_BarChart">
    <xsd:sequence>
        <xsd:group   ref="EG_BarChartShared" minOccurs="1" maxOccurs="1"/>
        ...
    </xsd:sequence>
</xsd:complexType>
*/

    bool bEG_BarChartShared = load3EG_BarChartShared( pReader, ptrBarChart );

    // <xsd:element name="gapWidth" type="CT_GapAmount"     minOccurs="0" maxOccurs="1"/>
    // <xsd:element name="overlap"  type="CT_Overlap"       minOccurs="0" maxOccurs="1"/>
    // <xsd:element name="serLines" type="CT_ChartLines"    minOccurs="0" maxOccurs="unbounded"/>
    // <xsd:element name="axId"     type="CT_UnsignedInt"   minOccurs="2" maxOccurs="2"/>
    // <xsd:element name="extLst"   type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>

    return true;
}

bool ChartPrivate::load3EG_BarChartShared(XMLDOM::XMLDOMReader* pReader, XMLDOM::Node* ptrBarChart)
{
/*
<xsd:group name="EG_BarChartShared">
    <xsd:sequence>
        <xsd:element name="barDir"     type="CT_BarDir"      minOccurs="1" maxOccurs="1"/>
        ...
    </xsd:sequence>
</xsd:group>

<xsd:complexType name="CT_BarDir">
    <xsd:attribute name="val" type="ST_BarDir" default="col"/>
</xsd:complexType>

<xsd:simpleType name="ST_BarDir">
    <xsd:restriction base="xsd:string">
        <xsd:enumeration value="bar"/>
        <xsd:enumeration value="col"/>
    </xsd:restriction>
</xsd:simpleType>
*/

    // /*
    for (int ci = 0 ; ci < ptrBarChart->childNode.size() ; ci ++)
    {
        XMLDOM::Node* pNode = ptrBarChart->childNode.at( ci );
        if ( NULL == pNode )
            continue;

        qDebug() << pNode->level << pNode->nodeName << pNode->nodeText ;
    }
    // */

    XMLDOM::Node* ptrBarDir = pReader->findNode( ptrBarChart, "c:barDir" );
    if ( NULL == ptrBarDir )
    {
        qDebug() << "[debug] c:barDir is mandatory field.";
        return false;
    }
    // val (attribute)

    XMLDOM::Attr* ptrAttr = pReader->findAttr( ptrBarDir, "val" );
    if ( NULL == ptrAttr )
    {
        qDebug() << "[debug] c:barDir has no val attribute.";
        return false;
    }
    // qDebug() << ptrAttr->name << ptrAttr->value;

    if ( ptrAttr->value == "bar" )
    {
        //! TODO: barchart type is bar
    }
    else if ( ptrAttr->value == "col" )
    {
        //! TODO: barchart type is col
    }
    else
    {
        qDebug() << "[debug] c:barDir val is invalid." << ptrAttr->value ;
        return false;
    }

    // <xsd:element name="grouping"   type="CT_BarGrouping" minOccurs="0" maxOccurs="1"/>
    XMLDOM::Node* ptrGrouping = pReader->findNode( ptrBarChart, "c:grouping" );
    if ( NULL == ptrGrouping )
    {
    }

    // <xsd:element name="varyColors" type="CT_Boolean"     minOccurs="0" maxOccurs="1"/>
    XMLDOM::Node* ptrVaryColors = pReader->findNode( ptrBarChart, "c:varyColors" );
    if ( NULL == ptrVaryColors )
    {
    }

    // <xsd:element name="ser"        type="CT_BarSer"      minOccurs="0" maxOccurs="unbounded"/>
    XMLDOM::Node* ptrSer = pReader->findNode( ptrBarChart, "c:ser" );
    if ( NULL == ptrSer )
    {
    }

    // <xsd:element name="dLbls"      type="CT_DLbls"       minOccurs="0" maxOccurs="1"/>


    return true;
}

bool ChartPrivate::load3Bar3DChart(XMLDOM::XMLDOMReader* pReader, XMLDOM::Node* ptrBar3DChart)
{
/*
<xsd:complexType name="CT_Bar3DChart">
    <xsd:sequence>
        <xsd:group   ref="EG_BarChartShared"                 minOccurs="1" maxOccurs="1"/>
        <xsd:element name="gapWidth" type="CT_GapAmount"     minOccurs="0" maxOccurs="1"/>
        <xsd:element name="gapDepth" type="CT_GapAmount"     minOccurs="0" maxOccurs="1"/>
        <xsd:element name="shape"    type="CT_Shape"         minOccurs="0" maxOccurs="1"/>
        <xsd:element name="axId"     type="CT_UnsignedInt"   minOccurs="2" maxOccurs="3"/>
        <xsd:element name="extLst"   type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
</xsd:complexType>
*/

    bool bEG_BarChartShared = load3EG_BarChartShared( pReader, ptrBar3DChart );

    return true;
}

bool ChartPrivate::load3OfPieChart(XMLDOM::XMLDOMReader* pReader, XMLDOM::Node* ptrOfPieChart )
{
/*
<xsd:complexType name="CT_OfPieChart">
    <xsd:sequence>
        <xsd:element name="ofPieType"     type="CT_OfPieType" minOccurs="1" maxOccurs="1"/>
        <xsd:group   ref="EG_PieChartShared" minOccurs="1" maxOccurs="1"/>
        <xsd:element name="gapWidth"      type="CT_GapAmount" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="splitType"     type="CT_SplitType" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="splitPos"      type="CT_Double"    minOccurs="0" maxOccurs="1"/>
        <xsd:element name="custSplit"     type="CT_CustSplit" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="secondPieSize" type="CT_SecondPieSize" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="serLines"      type="CT_ChartLines"    minOccurs="0" maxOccurs="unbounded"/>
        <xsd:element name="extLst"        type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
</xsd:complexType>
*/

    return true;
}

bool ChartPrivate::load3SurfaceChart(XMLDOM::XMLDOMReader* pReader, XMLDOM::Node* ptrSurfaceChart )
{
/*
<xsd:complexType name="CT_SurfaceChart">
    <xsd:sequence>
    <xsd:group   ref="EG_SurfaceChartShared" minOccurs="1" maxOccurs="1"/>
    <xsd:element name="axId"   type="CT_UnsignedInt"   minOccurs="2" maxOccurs="3"/>
    <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
</xsd:complexType>
*/

    return true;
}

bool ChartPrivate::load3Surface3DChart(XMLDOM::XMLDOMReader* pReader, XMLDOM::Node*ptrSurface3DChart )
{
/*
<xsd:complexType name="CT_Surface3DChart">
    <xsd:sequence>
        <xsd:group   ref="EG_SurfaceChartShared"           minOccurs="1" maxOccurs="1"/>
        <xsd:element name="axId"   type="CT_UnsignedInt"   minOccurs="3" maxOccurs="3"/>
        <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
</xsd:complexType>
*/


    return true;
}

bool ChartPrivate::load3BubbleChart(XMLDOM::XMLDOMReader* pReader, XMLDOM::Node* ptrBubbleChart )
{
/*
<xsd:complexType name="CT_BubbleChart">
    <xsd:sequence>
        <xsd:element name="varyColors"     type="CT_Boolean"        minOccurs="0" maxOccurs="1"/>
        <xsd:element name="ser"            type="CT_BubbleSer"      minOccurs="0" maxOccurs="unbounded"/>
        <xsd:element name="dLbls"          type="CT_DLbls"          minOccurs="0" maxOccurs="1"/>
        <xsd:element name="bubble3D"       type="CT_Boolean"        minOccurs="0" maxOccurs="1"/>
        <xsd:element name="bubbleScale"    type="CT_BubbleScale"    minOccurs="0" maxOccurs="1"/>
        <xsd:element name="showNegBubbles" type="CT_Boolean"        minOccurs="0" maxOccurs="1"/>
        <xsd:element name="sizeRepresents" type="CT_SizeRepresents" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="axId"           type="CT_UnsignedInt"    minOccurs="2" maxOccurs="2"/>
        <xsd:element name="extLst"         type="CT_ExtensionList"  minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
</xsd:complexType>
*/

    return true;
}

bool ChartPrivate::load2Legend(XMLDOM::XMLDOMReader* pReader, XMLDOM::Node* ptrLegend )
{
    // c:legend

    return true;
}

bool ChartPrivate::load2PlotVisOnly(XMLDOM::XMLDOMReader* pReader, XMLDOM::Node* ptrPlotVisOnly )
{
    // plotVisOnly

    return true;
}

bool ChartPrivate::load1Lang(XMLDOM::XMLDOMReader *pReader, XMLDOM::Node* ptrLang)
{
    if ( NULL == ptrLang )
        return false;

    return true;
}

bool ChartPrivate::load1PrinterSettings(XMLDOM::XMLDOMReader *pReader, XMLDOM::Node* ptrPrinterSetting)
{
    if ( NULL == ptrPrinterSetting )
        return false;

    return true;
}

// TO DEBUG: loop is not work, when i looping second element.
/*
dchrt_CT_PlotArea =
    element layout { dchrt_CT_Layout }?,
    (element areaChart { dchrt_CT_AreaChart }
        | element area3DChart { dchrt_ CT_Area3DChart }
        | element lineChart { dchrt_CT_LineChart }
        | element line3DChart { dchrt_CT_Line3DChart }
        | element stockChart { dchrt_CT_StockChart }
        | element radarChart { dchrt_CT_RadarChart }
        | element scatterChart { dchrt_CT_ScatterChart }
        | element pieChart { dchrt_CT_PieChart }
        | element pie3DChart { dchrt_CT_Pie3DChart }
        | element doughnutChart { dchrt_CT_DoughnutChart }
        | element barChart { dchrt_CT_BarChart }
        | element bar3DChart { dchrt_CT_Bar3DChart }
        | element ofPieChart { dchrt_CT_OfPieChart }
        | element surfaceChart { dchrt_CT_SurfaceChart }
        | element surface3DChart { dchrt_CT_Surface3DChart }
        | element bubbleChart { dchrt_CT_BubbleChart })+,
    (element valAx { dchrt_CT_ValAx }
        | element catAx { dchrt_CT_CatAx }
        | element dateAx { dchrt_CT_DateAx }
        | element serAx { dchrt_CT_SerAx })*,
    element dTable { dchrt_CT_DTable }?,
    element spPr { a_CT_ShapeProperties }?,
    element extLst { dchrt_CT_ExtensionList }?
 */
bool ChartPrivate::loadXmlPlotArea(QXmlStreamReader &reader)
{
    Q_ASSERT(reader.name() == QLatin1String("plotArea"));

    // TO DEBUG:

    /*
    reader.readNext();

    while (!reader.atEnd())
    {
        if (reader.isStartElement())
        {
            if (loadXmlPlotAreaElement(reader))
            {

            }
            else
            {
                qDebug() << "[debug] failed to load plotarea element.";
                return false;
            }

            reader.readNext();
        }
        else
        {

            reader.readNext();
        }
    }
    */

    while (!reader.atEnd())
    {
        reader.readNextStartElement();
        if (reader.tokenType() == QXmlStreamReader::StartElement)
        {
            // qDebug() << " [loadXmlPlotArea] " << reader.name();

            if (loadXmlPlotAreaElement(reader))
            {
            }
            else
            {
                qDebug() << "[debug] failed to load plot area element.";
                return false;
            }

        }
        else if (reader.tokenType() == QXmlStreamReader::EndElement &&
                 reader.name() == "plotArea")
        {
            break;
        }
    }

    return true;
}

bool ChartPrivate::loadXmlPlotAreaElement(QXmlStreamReader &reader)
{
    if (reader.name() == QLatin1String("layout"))
    {
        //!ToDo
        // layout
    }
    else if (reader.name().endsWith(QLatin1String("Chart")))
    {
        // for pieChart, barChart, ... (choose one)
        if ( ! loadXmlXxxChart(reader) )
        {
            qDebug() << "[debug] failed to load chart";
            return false;
        }
    }
    else if (reader.name() == QLatin1String("catAx")) // catAx, dateAx, serAx, valAx
    {
        loadXmlAxisCatAx(reader);
    }
    else if (reader.name() == QLatin1String("dateAx")) // catAx, dateAx, serAx, valAx
    {
        loadXmlAxisDateAx(reader);
    }
    else if (reader.name() == QLatin1String("serAx")) // catAx, dateAx, serAx, valAx
    {
        loadXmlAxisSerAx(reader);
    }
    else if (reader.name() == QLatin1String("valAx")) // catAx, dateAx, serAx, valAx
    {
        loadXmlAxisValAx(reader);
    }
    else if (reader.name() == QLatin1String("dTable"))
    {
        //!ToDo
        // dTable "CT_DTable"
        // reader.skipCurrentElement();
    }
    else if (reader.name() == QLatin1String("spPr"))
    {
        //!ToDo
        // spPr "a:CT_ShapeProperties"
        // reader.skipCurrentElement();
    }
    else if (reader.name() == QLatin1String("extLst"))
    {
        //!ToDo
        // extLst "CT_ExtensionList"
        // reader.skipCurrentElement();
    }

    return true;
}

bool ChartPrivate::loadXmlXxxChart(QXmlStreamReader &reader)
{
    QStringRef name = reader.name();

    if (name == QLatin1String("areaChart")) chartType = Chart::CT_AreaChart;
    else if (name == QLatin1String("area3DChart")) chartType = Chart::CT_Area3DChart;
    else if (name == QLatin1String("lineChart")) chartType = Chart::CT_LineChart;
    else if (name == QLatin1String("line3DChart")) chartType = Chart::CT_Line3DChart;
    else if (name == QLatin1String("stockChart")) chartType = Chart::CT_StockChart;
    else if (name == QLatin1String("radarChart")) chartType = Chart::CT_RadarChart;
    else if (name == QLatin1String("scatterChart")) chartType = Chart::CT_ScatterChart;
    else if (name == QLatin1String("pieChart")) chartType = Chart::CT_PieChart;
    else if (name == QLatin1String("pie3DChart")) chartType = Chart::CT_Pie3DChart;
    else if (name == QLatin1String("doughnutChart")) chartType = Chart::CT_DoughnutChart;
    else if (name == QLatin1String("barChart")) chartType = Chart::CT_BarChart;
    else if (name == QLatin1String("bar3DChart")) chartType = Chart::CT_Bar3DChart;
    else if (name == QLatin1String("ofPieChart")) chartType = Chart::CT_OfPieChart;
    else if (name == QLatin1String("surfaceChart")) chartType = Chart::CT_SurfaceChart;
    else if (name == QLatin1String("surface3DChart")) chartType = Chart::CT_Surface3DChart;
    else if (name == QLatin1String("bubbleChart")) chartType = Chart::CT_BubbleChart;
    else
    {
        qDebug() << "[undefined chart type] " << name;
        chartType = Chart::CT_NoStatementChart;
        return false;
    }

    while (!reader.atEnd())
    {
        reader.readNextStartElement();
        if (reader.tokenType() == QXmlStreamReader::StartElement)
        {
            if (reader.name() == QLatin1String("ser"))
            {
                loadXmlSer(reader);
            }
            else if (reader.name() == QLatin1String("axId"))
            {
                //!TODO
            }

        }
        else if (reader.tokenType() == QXmlStreamReader::EndElement &&
                   reader.name() == name)
        {
            break;
        }
    }

    return true;
}

bool ChartPrivate::loadXmlSer(QXmlStreamReader &reader)
{
    Q_ASSERT(reader.name() == QLatin1String("ser"));

    QSharedPointer<XlsxSeries> series = QSharedPointer<XlsxSeries>(new XlsxSeries);
    seriesList.append(series);

    while (!reader.atEnd() && !(reader.tokenType() == QXmlStreamReader::EndElement
                                && reader.name() == QLatin1String("ser")))
    {
        if (reader.readNextStartElement())
        {
            QStringRef name = reader.name();
            if (name == QLatin1String("cat") || name == QLatin1String("xVal"))
            {
                while (!reader.atEnd() && !(reader.tokenType() == QXmlStreamReader::EndElement
                                            && reader.name() == name))
                {
                    if (reader.readNextStartElement())
                    {
                        if (reader.name() == QLatin1String("numRef"))
                            series->axDataSource_numRef = loadXmlNumRef(reader);
                    }
                }
            }
            else if (name == QLatin1String("val") || name == QLatin1String("yVal"))
            {
                while (!reader.atEnd() && !(reader.tokenType() == QXmlStreamReader::EndElement
                                            && reader.name() == name))
                {
                    if (reader.readNextStartElement())
                    {
                        if (reader.name() == QLatin1String("numRef"))
                            series->numberDataSource_numRef = loadXmlNumRef(reader);
                    }
                }
            }
            else if (name == QLatin1String("extLst"))
            {
                while (!reader.atEnd() && !(reader.tokenType() == QXmlStreamReader::EndElement
                                            && reader.name() == name))
                {
                    reader.readNextStartElement();
                }
            }
        }
    }

    return true;
}


QString ChartPrivate::loadXmlNumRef(QXmlStreamReader &reader)
{
    Q_ASSERT(reader.name() == QLatin1String("numRef"));

    while (!reader.atEnd() && !(reader.tokenType() == QXmlStreamReader::EndElement
                                && reader.name() == QLatin1String("numRef")))
    {
        if (reader.readNextStartElement())
        {
            if (reader.name() == QLatin1String("f"))
                return reader.readElementText();
        }
    }

    return QString();
}

void ChartPrivate::saveXmlChart(QXmlStreamWriter &writer) const
{
    //----------------------------------------------------
    // c:chart
    writer.writeStartElement(QStringLiteral("c:chart"));

    //----------------------------------------------------
    // c:title

    saveXmlChartTitle(writer); // wrtie 'chart title'

    //----------------------------------------------------
    // c:plotArea

    writer.writeStartElement(QStringLiteral("c:plotArea"));

    switch (chartType)
    {
        case Chart::CT_AreaChart:       saveXmlAreaChart(writer); break;
        case Chart::CT_Area3DChart:     saveXmlAreaChart(writer); break;
        case Chart::CT_LineChart:       saveXmlLineChart(writer); break;
        case Chart::CT_Line3DChart:     saveXmlLineChart(writer); break;
        case Chart::CT_StockChart: break;
        case Chart::CT_RadarChart: break;
        case Chart::CT_ScatterChart:    saveXmlScatterChart(writer); break;
        case Chart::CT_PieChart:        saveXmlPieChart(writer); break;
        case Chart::CT_Pie3DChart:      saveXmlPieChart(writer); break;
        case Chart::CT_DoughnutChart:   saveXmlDoughnutChart(writer); break;
        case Chart::CT_BarChart:        saveXmlBarChart(writer); break;
        case Chart::CT_Bar3DChart:      saveXmlBarChart(writer); break;
        case Chart::CT_OfPieChart: break;
        case Chart::CT_SurfaceChart:  break;
        case Chart::CT_Surface3DChart: break;
        case Chart::CT_BubbleChart:  break;
        default:  break;
    }

    saveXmlAxis(writer); // c:catAx, c:valAx, c:serAx, c:dateAx (choose one)

    //!TODO: write element
    // c:dTable CT_DTable
    // c:spPr   CT_ShapeProperties
    // c:extLst CT_ExtensionList

    writer.writeEndElement(); // c:plotArea

    //!TODO:  save-legend // c:legend

    writer.writeEndElement(); // c:chart
}

bool ChartPrivate::loadXmlChartTitle(QXmlStreamReader &reader)
{
    //!TODO : load chart title

    Q_ASSERT(reader.name() == QLatin1String("title"));

    while (!reader.atEnd() && !(reader.tokenType() == QXmlStreamReader::EndElement
                                && reader.name() == QLatin1String("title")))
    {
        if (reader.readNextStartElement())
        {
            if (reader.name() == QLatin1String("tx")) // c:tx
                return loadXmlChartTitleTx(reader);
        }
    }

    return false;
}

bool ChartPrivate::loadXmlChartTitleTx(QXmlStreamReader &reader)
{
    Q_ASSERT(reader.name() == QLatin1String("tx"));

    while (!reader.atEnd() && !(reader.tokenType() == QXmlStreamReader::EndElement
                                && reader.name() == QLatin1String("tx")))
    {
        if (reader.readNextStartElement())
        {
            if (reader.name() == QLatin1String("rich")) // c:rich
                return loadXmlChartTitleTxRich(reader);
        }
    }

    return false;
}

bool ChartPrivate::loadXmlChartTitleTxRich(QXmlStreamReader &reader)
{
    Q_ASSERT(reader.name() == QLatin1String("rich"));

    while (!reader.atEnd() && !(reader.tokenType() == QXmlStreamReader::EndElement
                                && reader.name() == QLatin1String("rich")))
    {
        if (reader.readNextStartElement())
        {
            if (reader.name() == QLatin1String("p")) // a:p
                return loadXmlChartTitleTxRichP(reader);
        }
    }

    return false;
}

bool ChartPrivate::loadXmlChartTitleTxRichP(QXmlStreamReader &reader)
{
    Q_ASSERT(reader.name() == QLatin1String("p"));

    while (!reader.atEnd() && !(reader.tokenType() == QXmlStreamReader::EndElement
                                && reader.name() == QLatin1String("p")))
    {
        if (reader.readNextStartElement())
        {
            if (reader.name() == QLatin1String("r")) // a:r
                return loadXmlChartTitleTxRichP_R(reader);
        }
    }

    return false;
}

bool ChartPrivate::loadXmlChartTitleTxRichP_R(QXmlStreamReader &reader)
{
    Q_ASSERT(reader.name() == QLatin1String("r"));

    while (!reader.atEnd() && !(reader.tokenType() == QXmlStreamReader::EndElement
                                && reader.name() == QLatin1String("r")))
    {
        if (reader.readNextStartElement())
        {
            if (reader.name() == QLatin1String("t")) // a:t
            {
                QString textValue = reader.readElementText();
                this->chartTitle = textValue;
                return true;
            }
        }
    }

    return false;
}


// wrtie 'chart title'
void ChartPrivate::saveXmlChartTitle(QXmlStreamWriter &writer) const
{
    if ( chartTitle.isEmpty() )
        return;

    writer.writeStartElement(QStringLiteral("c:title"));
    /*
    <xsd:complexType name="CT_Title">
        <xsd:sequence>
            <xsd:element name="tx"      type="CT_Tx"                minOccurs="0" maxOccurs="1"/>
            <xsd:element name="layout"  type="CT_Layout"            minOccurs="0" maxOccurs="1"/>
            <xsd:element name="overlay" type="CT_Boolean"           minOccurs="0" maxOccurs="1"/>
            <xsd:element name="spPr"    type="a:CT_ShapeProperties" minOccurs="0" maxOccurs="1"/>
            <xsd:element name="txPr"    type="a:CT_TextBody"        minOccurs="0" maxOccurs="1"/>
            <xsd:element name="extLst"  type="CT_ExtensionList"     minOccurs="0" maxOccurs="1"/>
        </xsd:sequence>
    </xsd:complexType>
    */

        writer.writeStartElement(QStringLiteral("c:tx"));
        /*
        <xsd:complexType name="CT_Tx">
            <xsd:sequence>
                <xsd:choice     minOccurs="1"   maxOccurs="1">
                <xsd:element    name="strRef"   type="CT_StrRef"        minOccurs="1" maxOccurs="1"/>
                <xsd:element    name="rich"     type="a:CT_TextBody"    minOccurs="1" maxOccurs="1"/>
                </xsd:choice>
            </xsd:sequence>
        </xsd:complexType>
        */

            writer.writeStartElement(QStringLiteral("c:rich"));
            /*
            <xsd:complexType name="CT_TextBody">
                <xsd:sequence>
                    <xsd:element name="bodyPr"      type="CT_TextBodyProperties"    minOccurs=" 1"  maxOccurs="1"/>
                    <xsd:element name="lstStyle"    type="CT_TextListStyle"         minOccurs="0"   maxOccurs="1"/>
                    <xsd:element name="p"           type="CT_TextParagraph"         minOccurs="1"   maxOccurs="unbounded"/>
                </xsd:sequence>
            </xsd:complexType>
            */

                writer.writeEmptyElement(QStringLiteral("a:bodyPr")); // <a:bodyPr/>
                /*
                <xsd:complexType name="CT_TextBodyProperties">
                    <xsd:sequence>
                        <xsd:element name="prstTxWarp" type="CT_PresetTextShape" minOccurs="0" maxOccurs="1"/>
                        <xsd:group ref="EG_TextAutofit" minOccurs="0" maxOccurs="1"/>
                        <xsd:element name="scene3d" type="CT_Scene3D" minOccurs="0" maxOccurs="1"/>
                        <xsd:group ref="EG_Text3D" minOccurs="0" maxOccurs="1"/>
                        <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
                    </xsd:sequence>
                    <xsd:attribute name="rot" type="ST_Angle" use="optional"/>
                    <xsd:attribute name="spcFirstLastPara" type="xsd:boolean" use="optional"/>
                    <xsd:attribute name="vertOverflow" type="ST_TextVertOverflowType" use="optional"/>
                    <xsd:attribute name="horzOverflow" type="ST_TextHorzOverflowType" use="optional"/>
                    <xsd:attribute name="vert" type="ST_TextVerticalType" use="optional"/>
                    <xsd:attribute name="wrap" type="ST_TextWrappingType" use="optional"/>
                    <xsd:attribute name="lIns" type="ST_Coordinate32" use="optional"/>
                    <xsd:attribute name="tIns" type="ST_Coordinate32" use="optional"/>
                    <xsd:attribute name="rIns" type="ST_Coordinate32" use="optional"/>
                    <xsd:attribute name="bIns" type="ST_Coordinate32" use="optional"/>
                    <xsd:attribute name="numCol" type="ST_TextColumnCount" use="optional"/>
                    <xsd:attribute name="spcCol" type="ST_PositiveCoordinate32" use="optional"/>
                    <xsd:attribute name="rtlCol" type="xsd:boolean" use="optional"/>
                    <xsd:attribute name="fromWordArt" type="xsd:boolean" use="optional"/>
                    <xsd:attribute name="anchor" type="ST_TextAnchoringType" use="optional"/>
                    <xsd:attribute name="anchorCtr" type="xsd:boolean" use="optional"/>
                    <xsd:attribute name="forceAA" type="xsd:boolean" use="optional"/>
                    <xsd:attribute name="upright" type="xsd:boolean" use="optional" default="false"/>
                    <xsd:attribute name="compatLnSpc" type="xsd:boolean" use="optional"/>
                </xsd:complexType>
                 */

                writer.writeEmptyElement(QStringLiteral("a:lstStyle")); // <a:lstStyle/>

                writer.writeStartElement(QStringLiteral("a:p"));
                /*
                <xsd:complexType name="CT_TextParagraph">
                    <xsd:sequence>
                        <xsd:element    name="pPr"          type="CT_TextParagraphProperties" minOccurs="0" maxOccurs="1"/>
                        <xsd:group      ref="EG_TextRun"    minOccurs="0" maxOccurs="unbounded"/>
                        <xsd:element    name="endParaRPr"   type="CT_TextCharacterProperties" minOccurs="0"
                        maxOccurs="1"/>
                    </xsd:sequence>
                </xsd:complexType>
                 */

                    // <a:pPr lvl="0">
                    writer.writeStartElement(QStringLiteral("a:pPr"));

                        writer.writeAttribute(QStringLiteral("lvl"), QStringLiteral("0"));

                        // <a:defRPr b="0"/>
                        writer.writeStartElement(QStringLiteral("a:defRPr"));

                            writer.writeAttribute(QStringLiteral("b"), QStringLiteral("0"));

                        writer.writeEndElement();  // a:defRPr

                    writer.writeEndElement();  // a:pPr

                /*
                <xsd:group name="EG_TextRun">
                    <xsd:choice>
                        <xsd:element name="r"   type="CT_RegularTextRun"/>
                        <xsd:element name="br"  type="CT_TextLineBreak"/>
                        <xsd:element name="fld" type="CT_TextField"/>
                    </xsd:choice>
                </xsd:group>
                */

                writer.writeStartElement(QStringLiteral("a:r"));
                /*
                <xsd:complexType name="CT_RegularTextRun">
                    <xsd:sequence>
                        <xsd:element name="rPr" type="CT_TextCharacterProperties" minOccurs="0" maxOccurs="1"/>
                        <xsd:element name="t"   type="xsd:string" minOccurs="1" maxOccurs="1"/>
                    </xsd:sequence>
                </xsd:complexType>
                 */

                    // <a:t>chart name</a:t>
                    writer.writeTextElement(QStringLiteral("a:t"), chartTitle);

                writer.writeEndElement();  // a:r

                writer.writeEndElement();  // a:p

            writer.writeEndElement();  // c:rich

        writer.writeEndElement();  // c:tx

        // <c:overlay val="0"/>
        writer.writeStartElement(QStringLiteral("c:overlay"));
            writer.writeAttribute(QStringLiteral("val"), QStringLiteral("0"));
        writer.writeEndElement();  // c:overlay

    writer.writeEndElement();  // c:title
}
// }}

void ChartPrivate::saveXmlPieChart(QXmlStreamWriter &writer) const
{
    QString name = chartType == Chart::CT_PieChart ? QStringLiteral("c:pieChart") : QStringLiteral("c:pie3DChart");

    writer.writeStartElement(name);

    //Do the same behavior as Excel, Pie prefer varyColors
    writer.writeEmptyElement(QStringLiteral("c:varyColors"));
    writer.writeAttribute(QStringLiteral("val"), QStringLiteral("1"));

    for (int i=0; i<seriesList.size(); ++i)
        saveXmlSer(writer, seriesList[i].data(), i);

    writer.writeEndElement(); //pieChart, pie3DChart
}

void ChartPrivate::saveXmlBarChart(QXmlStreamWriter &writer) const
{
    QString name = chartType == Chart::CT_BarChart ? QStringLiteral("c:barChart") : QStringLiteral("c:bar3DChart");

    writer.writeStartElement(name);

    writer.writeEmptyElement(QStringLiteral("c:barDir"));
    writer.writeAttribute(QStringLiteral("val"), QStringLiteral("col"));

    for ( int i = 0 ; i < seriesList.size() ; ++i )
    {
        saveXmlSer(writer, seriesList[i].data(), i);
    }

    if ( axisList.isEmpty() )
    {
        const_cast<ChartPrivate*>(this)->axisList.append(
                    QSharedPointer<XlsxAxis>(
                        new XlsxAxis( XlsxAxis::T_Cat, XlsxAxis::Bottom, 0, 1, axisNames[XlsxAxis::Bottom] )));

        const_cast<ChartPrivate*>(this)->axisList.append(
                    QSharedPointer<XlsxAxis>(
                        new XlsxAxis( XlsxAxis::T_Val, XlsxAxis::Left, 1, 0, axisNames[XlsxAxis::Left] )));
    }

    //Note: Bar3D have 2~3 axes
    int axisListSize = axisList.size();
    //Q_ASSERT( axisListSize == 2 || ( axisListSize == 3 && chartType == Chart::CT_Bar3DChart ) );
    if ( axisListSize == 2 ||
         ( axisListSize == 3 && chartType == Chart::CT_Bar3DChart ) )
    {
    }
    else
    {
        int dp = 0; // FOR DEBUG
    }

    for ( int i = 0 ; i < axisList.size() ; ++i )
    {
        writer.writeEmptyElement(QStringLiteral("c:axId"));
        writer.writeAttribute(QStringLiteral("val"), QString::number(axisList[i]->axisId));
    }

    writer.writeEndElement(); //barChart, bar3DChart
}

void ChartPrivate::saveXmlLineChart(QXmlStreamWriter &writer) const
{
    QString name = chartType==Chart::CT_LineChart ? QStringLiteral("c:lineChart") : QStringLiteral("c:line3DChart");

    writer.writeStartElement(name);

    // writer.writeEmptyElement(QStringLiteral("grouping")); // dev22

    for (int i=0; i<seriesList.size(); ++i)
        saveXmlSer(writer, seriesList[i].data(), i);

    if (axisList.isEmpty())
    {
        const_cast<ChartPrivate*>(this)->axisList.append(QSharedPointer<XlsxAxis>(new XlsxAxis(XlsxAxis::T_Cat, XlsxAxis::Bottom, 0, 1, axisNames[XlsxAxis::Bottom] )));
        const_cast<ChartPrivate*>(this)->axisList.append(QSharedPointer<XlsxAxis>(new XlsxAxis(XlsxAxis::T_Val, XlsxAxis::Left, 1, 0, axisNames[XlsxAxis::Left] )));
        if (chartType==Chart::CT_Line3DChart)
            const_cast<ChartPrivate*>(this)->axisList.append(QSharedPointer<XlsxAxis>(new XlsxAxis(XlsxAxis::T_Ser, XlsxAxis::Bottom, 2, 0)));
    }

    Q_ASSERT((axisList.size()==2||chartType==Chart::CT_LineChart)|| (axisList.size()==3 && chartType==Chart::CT_Line3DChart));

    for (int i=0; i<axisList.size(); ++i) {
        writer.writeEmptyElement(QStringLiteral("c:axId"));
        writer.writeAttribute(QStringLiteral("val"), QString::number(axisList[i]->axisId));
    }

    writer.writeEndElement(); //lineChart, line3DChart
}

void ChartPrivate::saveXmlScatterChart(QXmlStreamWriter &writer) const
{
    const QString name = QStringLiteral("c:scatterChart");

    writer.writeStartElement(name);

    writer.writeEmptyElement(QStringLiteral("c:scatterStyle"));

    for (int i=0; i<seriesList.size(); ++i)
        saveXmlSer(writer, seriesList[i].data(), i);

    if (axisList.isEmpty())
    {
        const_cast<ChartPrivate*>(this)->axisList.append(
                    QSharedPointer<XlsxAxis>(
                        new XlsxAxis(XlsxAxis::T_Val, XlsxAxis::Bottom, 0, 1, axisNames[XlsxAxis::Bottom] )));
        const_cast<ChartPrivate*>(this)->axisList.append(
                    QSharedPointer<XlsxAxis>(
                        new XlsxAxis(XlsxAxis::T_Val, XlsxAxis::Left, 1, 0, axisNames[XlsxAxis::Left] )));
    }

    int axisListSize = axisList.size();
    Q_ASSERT(axisListSize == 2);

    for (int i=0; i<axisList.size(); ++i)
    {
        writer.writeEmptyElement(QStringLiteral("c:axId"));
        writer.writeAttribute(QStringLiteral("val"), QString::number(axisList[i]->axisId));
    }

    writer.writeEndElement(); //c:scatterChart
}

void ChartPrivate::saveXmlAreaChart(QXmlStreamWriter &writer) const
{
    QString name = chartType==Chart::CT_AreaChart ? QStringLiteral("c:areaChart") : QStringLiteral("c:area3DChart");

    writer.writeStartElement(name);

    // writer.writeEmptyElement(QStringLiteral("grouping")); // dev22

    for (int i=0; i<seriesList.size(); ++i)
        saveXmlSer(writer, seriesList[i].data(), i);

    if (axisList.isEmpty())
    {
        const_cast<ChartPrivate*>(this)->axisList.append(QSharedPointer<XlsxAxis>(new XlsxAxis(XlsxAxis::T_Cat, XlsxAxis::Bottom, 0, 1)));
        const_cast<ChartPrivate*>(this)->axisList.append(QSharedPointer<XlsxAxis>(new XlsxAxis(XlsxAxis::T_Val, XlsxAxis::Left, 1, 0)));
    }

    //Note: Area3D have 2~3 axes
    Q_ASSERT(axisList.size()==2 || (axisList.size()==3 && chartType==Chart::CT_Area3DChart));

    for (int i=0; i<axisList.size(); ++i)
    {
        writer.writeEmptyElement(QStringLiteral("c:axId"));
        writer.writeAttribute(QStringLiteral("val"), QString::number(axisList[i]->axisId));
    }

    writer.writeEndElement(); //lineChart, line3DChart
}

void ChartPrivate::saveXmlDoughnutChart(QXmlStreamWriter &writer) const
{
    QString name = QStringLiteral("c:doughnutChart");

    writer.writeStartElement(name);

    writer.writeEmptyElement(QStringLiteral("c:varyColors"));
    writer.writeAttribute(QStringLiteral("val"), QStringLiteral("1"));

    for (int i=0; i<seriesList.size(); ++i)
        saveXmlSer(writer, seriesList[i].data(), i);

    writer.writeStartElement(QStringLiteral("c:holeSize"));
    writer.writeAttribute(QStringLiteral("val"), QString::number(50));

    writer.writeEndElement();
}

void ChartPrivate::saveXmlSer(QXmlStreamWriter &writer, XlsxSeries *ser, int id) const
{
    writer.writeStartElement(QStringLiteral("c:ser"));
    writer.writeEmptyElement(QStringLiteral("c:idx"));
    writer.writeAttribute(QStringLiteral("val"), QString::number(id));
    writer.writeEmptyElement(QStringLiteral("c:order"));
    writer.writeAttribute(QStringLiteral("val"), QString::number(id));

    if (!ser->axDataSource_numRef.isEmpty()) {
        if (chartType == Chart::CT_ScatterChart || chartType == Chart::CT_BubbleChart)
            writer.writeStartElement(QStringLiteral("c:xVal"));
        else
            writer.writeStartElement(QStringLiteral("c:cat"));
        writer.writeStartElement(QStringLiteral("c:numRef"));
        writer.writeTextElement(QStringLiteral("c:f"), ser->axDataSource_numRef);
        writer.writeEndElement();//c:numRef
        writer.writeEndElement();//c:cat or c:xVal
    }

    if (!ser->numberDataSource_numRef.isEmpty()) {
        if (chartType == Chart::CT_ScatterChart || chartType == Chart::CT_BubbleChart)
            writer.writeStartElement(QStringLiteral("c:yVal"));
        else
            writer.writeStartElement(QStringLiteral("c:val"));
        writer.writeStartElement(QStringLiteral("c:numRef"));
        writer.writeTextElement(QStringLiteral("c:f"), ser->numberDataSource_numRef);
        writer.writeEndElement();//c:numRef
        writer.writeEndElement();//c:val or c:yVal
    }

    writer.writeEndElement();//c:ser
}

bool ChartPrivate::loadXmlAxisCatAx(QXmlStreamReader &reader)
{

    XlsxAxis* axis = new XlsxAxis();
    axis->type = XlsxAxis::T_Cat;
    axisList.append( QSharedPointer<XlsxAxis>(axis) );

    // load EG_AxShared
    if ( ! loadXmlAxisEG_AxShared( reader, axis ) )
    {
        qDebug() << "failed to load EG_AxShared";
        return false;
    }

    //!TODO: load element
    // auto
    // lblAlgn
    // lblOffset
    // tickLblSkip
    // tickMarkSkip
    // noMultiLvlLbl
    // extLst

    return true;
}

bool ChartPrivate::loadXmlAxisDateAx(QXmlStreamReader &reader)
{

    XlsxAxis* axis = new XlsxAxis();
    axis->type = XlsxAxis::T_Date;
    axisList.append( QSharedPointer<XlsxAxis>(axis) );

    // load EG_AxShared
    if ( ! loadXmlAxisEG_AxShared( reader, axis ) )
    {
        qDebug() << "failed to load EG_AxShared";
        return false;
    }

    //!TODO: load element
    // auto
    // lblOffset
    // baseTimeUnit
    // majorUnit
    // majorTimeUnit
    // minorUnit
    // minorTimeUnit
    // extLst

    return true;
}

bool ChartPrivate::loadXmlAxisSerAx(QXmlStreamReader &reader)
{

    XlsxAxis* axis = new XlsxAxis();
    axis->type = XlsxAxis::T_Ser;
    axisList.append( QSharedPointer<XlsxAxis>(axis) );

    // load EG_AxShared
    if ( ! loadXmlAxisEG_AxShared( reader, axis ) )
    {
        qDebug() << "failed to load EG_AxShared";
        return false;
    }

    //!TODO: load element
    // tickLblSkip
    // tickMarkSkip
    // extLst

    return true;
}

bool ChartPrivate::loadXmlAxisValAx(QXmlStreamReader &reader)
{
    Q_ASSERT(reader.name() == QLatin1String("valAx"));

    XlsxAxis* axis = new XlsxAxis();
    axis->type = XlsxAxis::T_Val;
    axisList.append( QSharedPointer<XlsxAxis>(axis) );

    if ( ! loadXmlAxisEG_AxShared( reader, axis ) )
    {
        qDebug() << "failed to load EG_AxShared";
        return false;
    }

    //!TODO: load element
    // crossBetween
    // majorUnit
    // minorUnit
    // dispUnits
    // extLst

    return true;
}

/*
<xsd:group name="EG_AxShared">
    <xsd:sequence>
        <xsd:element name="axId" type="CT_UnsignedInt" minOccurs="1" maxOccurs="1"/> (*)(M)
        <xsd:element name="scaling" type="CT_Scaling" minOccurs="1" maxOccurs="1"/> (*)(M)
        <xsd:element name="delete" type="CT_Boolean" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="axPos" type="CT_AxPos" minOccurs="1" maxOccurs="1"/> (*)(M)
        <xsd:element name="majorGridlines" type="CT_ChartLines" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="minorGridlines" type="CT_ChartLines" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="title" type="CT_Title" minOccurs="0" maxOccurs="1"/> (*)
        <xsd:element name="numFmt" type="CT_NumFmt" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="majorTickMark" type="CT_TickMark" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="minorTickMark" type="CT_TickMark" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="tickLblPos" type="CT_TickLblPos" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="spPr" type="a:CT_ShapeProperties" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="txPr" type="a:CT_TextBody" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="crossAx" type="CT_UnsignedInt" minOccurs="1" maxOccurs="1"/> (*)(M)
        <xsd:choice minOccurs="0" maxOccurs="1">
            <xsd:element name="crosses" type="CT_Crosses" minOccurs="1" maxOccurs="1"/>
            <xsd:element name="crossesAt" type="CT_Double" minOccurs="1" maxOccurs="1"/>
        </xsd:choice>
    </xsd:sequence>
</xsd:group>
*/
bool ChartPrivate::loadXmlAxisEG_AxShared(QXmlStreamReader &reader, XlsxAxis* axis)
{
    Q_ASSERT( NULL != axis );
    Q_ASSERT( reader.name().endsWith("Ax") );
    QString name = reader.name().toString(); //

    while ( !reader.atEnd() )
    {
        reader.readNextStartElement();

        if ( reader.tokenType() == QXmlStreamReader::StartElement )
        {
            if ( reader.name() == QLatin1String("axId") )
            {
                // mandatory element

                // CT_UnsignedInt
                /*
                    <xsd:complexType name="CT_UnsignedInt">
                        <xsd:attribute name="val" type="xsd:unsignedInt" use="required"/>
                    </xsd:complexType>
                */

                int axId = reader.attributes().value("val").toInt();
                axis->axisId = axId;
            }
            else if ( reader.name() == QLatin1String("scaling") )
            {
                // mandatory element
                // CT_Scaling

                loadXmlAxisEG_AxShared_Scaling(reader, axis);
            }
            else if ( reader.name() == QLatin1String("delete") )
            {
                //!TODO
                // CT_Boolean
                /*
                    <xsd:complexType name="CT_Boolean">
                     <xsd:attribute name="val" type="s:ST_OnOff" default="0"/>
                    </xsd:complexType>

                    <xsd:simpleType name="ST_OnOff">
                     <xs:union memberTypes="xsd:boolean"/>
                    </xsd:simpleType>
                */

                bool bDelete;
                QString strDelete = reader.attributes().value(QLatin1String("val")).toString();
                if ( strDelete == "true" || strDelete == "1" )
                {
                    bDelete = true;
                }
                else if ( strDelete == "false" || strDelete == "0" )
                {
                    bDelete = false;
                }
                else
                {
                    // invalid
                }

            }
            else if ( reader.name() == QLatin1String("axPos") )
            {
                // mandatory element

                // CT_AxPos
                /*
                    <xsd:complexType name="CT_AxPos">
                     <xsd:attribute name="val" type="ST_AxPos" use="required"/>
                    </xsd:complexType>

                    <xsd:simpleType name="ST_AxPos">
                     <xsd:restriction base="xsd:string">
                      <xsd:enumeration value="b"/>
                      <xsd:enumeration value="l"/>
                      <xsd:enumeration value="r"/>
                      <xsd:enumeration value="t"/>
                     </xsd:restriction>
                    </xsd:simpleType>
                */

                QString axPosVal = reader.attributes().value(QLatin1String("val")).toString();

                if ( axPosVal == "l" ) {
                    axis->axisPos = XlsxAxis::Left;
                }
                else if ( axPosVal == "r" ) {
                    axis->axisPos = XlsxAxis::Right;
                }
                else if ( axPosVal == "t" ) {
                    axis->axisPos = XlsxAxis::Top;
                }
                else if ( axPosVal == "b" ) {
                    axis->axisPos = XlsxAxis::Bottom;
                }
                else {
                    // invalid pos
                    return false;
                }
            }
            else if ( reader.name() == QLatin1String("majorGridlines") )
            {
                //!TODO
                // CT_ChartLines


            }
            else if ( reader.name() == QLatin1String("minorGridlines") )
            {
                //!TODO
                // CT_ChartLines

            }
            else if ( reader.name() == QLatin1String("title") )
            {
                // title
                // CT_Title

                if ( !loadXmlAxisEG_AxShared_Title(reader, axis) )
                {
                    qDebug() << "failed to load EG_AxShared title.";
                    return false;
                }
            }
            else if ( reader.name() == QLatin1String("numFmt") )
            {
                //!TODO
                //! CT_NumFmt

            }
            else if ( reader.name() == QLatin1String("majorTickMark") )
            {
                //!TODO
                //! CT_TickMark

            }
            else if ( reader.name() == QLatin1String("minorTickMark") )
            {
                //!TODO
                //! CT_TickMark

            }
            else if ( reader.name() == QLatin1String("tickLblPos") )
            {
                //!TODO
                //! CT_TickLblPos

                QString val = reader.attributes().value("val").toString();

            }
            else if ( reader.name() == QLatin1String("spPr") )
            {
                //!TODO
                //! a:CT_ShapeProperties

            }
            else if ( reader.name() == QLatin1String("txPr") )
            {
                //!TODO
                //! a:CT_TextBody

                if ( loadTxPr( reader, axis ) )
                {
                }
                else
                {
                }
            }
            else if ( reader.name() == QLatin1String("crossAx") )
            {
                // mandatory element
                // CT_UnsignedInt
                /*
                <xsd:complexType name="CT_UnsignedInt">
                    <xsd:attribute name="val" type="xsd:unsignedInt" use="required"/>
                </xsd:complexType>
                */

                if ( ! reader.attributes().value(QLatin1String("val")).isNull() )
                {
                    uint crossAx = reader.attributes().value(QLatin1String("val")).toUInt();
                    axis->crossAx = crossAx;
                }
                else
                {
                }
            }
            else if ( reader.name() == QLatin1String("crosses") )
            {
                //!TODO
                // CT_Crosses
                /*
                <xsd:complexType name="CT_Crosses">
                 <xsd:attribute name="val" type="ST_Crosses" use="required"/>
                </xsd:complexType>

                <xsd:simpleType name="ST_Crosses">
                 <xsd:restriction base="xsd:string">
                  <xsd:enumeration value="autoZero"/>
                  <xsd:enumeration value="max"/>
                  <xsd:enumeration value="min"/>
                 </xsd:restriction>
                </xsd:simpleType>
                */

                QString strCrosses = reader.attributes().value(QLatin1String("val")).toString();
                // autoZero, etc
                if ( strCrosses == "autoZero" )
                {
                    //!TODO: set value to autozero
                }
                else if ( strCrosses == "max" )
                {
                    //!TODO: set value to max
                }
                else if ( strCrosses == "min" )
                {
                    //!TODO: set value to min
                }
                else
                {
                    // invalid Crosses
                }
            }
            else if ( reader.name() == QLatin1String("crossesAt") )
            {
                //!TODO
                // CT_Double
                /*
                <xsd:complexType name="CT_Double">
                <xsd:attribute name="val" type="xsd:double" use="required"/>
                </xsd:complexType>
                */

                double dCrossesAt = reader.attributes().value(QLatin1String("val")).toDouble();
            }
            else
            {
                // undefined element
            }
        }
        else if ( reader.tokenType() == QXmlStreamReader::EndElement &&
                  reader.name().toString() == name )
        {
            break;
        }
    } // while ( !reader.atEnd() ) ...

    return true;
}

bool ChartPrivate::loadTxPr(QXmlStreamReader &reader, XlsxAxis* axis)
{
    Q_ASSERT(reader.name() == QLatin1String("txPr"));

    // CT_TextBody
    /*
    <xsd:complexType name="CT_TextBody">
     <xsd:sequence>
      <xsd:element name="bodyPr"   type="CT_TextBodyProperties" minOccurs="1" maxOccurs="1"/>
      <xsd:element name="lstStyle" type="CT_TextListStyle"      minOccurs="0" maxOccurs="1"/>
      <xsd:element name="p"        type="CT_TextParagraph"      minOccurs="1" maxOccurs="unbounded"/>
     </xsd:sequence>
    </xsd:complexType>
    */

    while ( !reader.atEnd() )
    {
        reader.readNextStartElement();
        if ( reader.tokenType() == QXmlStreamReader::StartElement )
        {
            if ( reader.name() == QLatin1String("bodyPr") )
            {
                // mandatory element
                //!TODO: load CT_TextBodyProperties

                loadTxPr_BodyPr(reader, axis);
            }
            else if ( reader.name() == QLatin1String("lstStyle") )
            {
                //!TODO: load CT_TextListStyle

                loadTxPr_LstStyle(reader, axis);
            }
            else if ( reader.name() == QLatin1String("p") )
            {
                // mandatory element
                //!TODO: load CT_TextParagraph

                loadTxPr_P(reader, axis);
            }
            else
            { // undefined element
            }
        }
        else if ( reader.tokenType() == QXmlStreamReader::EndElement &&
                  reader.name() == "txPr" )
        {
            break;
        }
    } // while ( !reader.atEnd() )

    return true;
}

bool ChartPrivate::loadTxPr_BodyPr(QXmlStreamReader &reader, XlsxAxis* axis)
{
    Q_ASSERT(reader.name() == QLatin1String("bodyPr"));

    // CT_TextBodyProperties

    while ( !reader.atEnd() )
    {
        reader.readNextStartElement();
        if ( reader.tokenType() == QXmlStreamReader::StartElement )
        {
            // if ( reader.name() == QLatin1String("") )
            {
            }
        }
        else if ( reader.tokenType() == QXmlStreamReader::EndElement &&
                  reader.name() == "bodyPr" )
        {
            break;
        }
    }

    return true;
}

bool ChartPrivate::loadTxPr_LstStyle(QXmlStreamReader &reader, XlsxAxis* axis)
{
    Q_ASSERT(reader.name() == QLatin1String("lstStyle"));

    // CT_TextParagraph

    while ( !reader.atEnd() )
    {
        reader.readNextStartElement();
        if ( reader.tokenType() == QXmlStreamReader::StartElement )
        {
            // if ( reader.name() == QLatin1String("") )
            {
            }
        }
        else if ( reader.tokenType() == QXmlStreamReader::EndElement &&
                  reader.name() == "lstStyle" )
        {
            break;
        }
    }

    return true;
}

bool ChartPrivate::loadTxPr_P(QXmlStreamReader &reader, XlsxAxis* axis)
{
    Q_ASSERT(reader.name() == QLatin1String("p"));

    // CT_TextParagraph

    while ( !reader.atEnd() )
    {
        reader.readNextStartElement();
        if ( reader.tokenType() == QXmlStreamReader::StartElement )
        {
            // if ( reader.name() == QLatin1String("") )
            {
            }
        }
        else if ( reader.tokenType() == QXmlStreamReader::EndElement &&
                  reader.name() == "p" )
        {
            break;
        }
    }

    return true;
}

bool ChartPrivate::loadXmlAxisEG_AxShared_Scaling(QXmlStreamReader &reader, XlsxAxis* axis)
{
    Q_ASSERT(reader.name() == QLatin1String("scaling"));

    // CT_Scaling
    /*
    <xsd:complexType name="CT_Scaling">
     <xsd:sequence>
      <xsd:element name="logBase"     type="CT_LogBase"       minOccurs="0" maxOccurs="1"/>
      <xsd:element name="orientation" type="CT_Orientation"   minOccurs="0" maxOccurs="1"/>
      <xsd:element name="max"         type="CT_Double"        minOccurs="0" maxOccurs="1"/>
      <xsd:element name="min"         type="CT_Double"        minOccurs="0" maxOccurs="1"/>
      <xsd:element name="extLst"      type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
     </xsd:sequence>
    </xsd:complexType>
    */

    while ( !reader.atEnd() )
    {
        reader.readNextStartElement();
        if ( reader.tokenType() == QXmlStreamReader::StartElement )
        {
            if ( reader.name() == QLatin1String("logBase") )
            {
                /*
                    <xsd:complexType name="CT_LogBase">
                     <xsd:attribute name="val" type="ST_LogBase" use="required"/>
                    </xsd:complexType>

                    <xsd:simpleType 1145 name="ST_LogBase">
                     <xsd:restriction base="xsd:double">
                      <xsd:minInclusive value="2"/>
                      <xsd:maxInclusive value="1000"/>
                     </xsd:restriction>
                    </xsd:simpleType>
                 */

                double dSTLogBase = reader.attributes().value(QLatin1String("val")).toDouble();
                // range : 2 ~ 1000

            }
            else if ( reader.name() == QLatin1String("orientation") )
            {
                /*
                    <xsd:complexType name="CT_Orientation">
                     <xsd:attribute name="val" type="ST_Orientation" default="minMax"/>
                    </xsd:complexType>

                    <xsd:simpleType name="ST_Orientation">
                     <xsd:restriction base="xsd:string">
                      <xsd:enumeration value="maxMin"/>
                      <xsd:enumeration value="minMax"/>
                     </xsd:restriction>
                    </xsd:simpleType>
                */

                QString strOrientation = reader.attributes().value(QLatin1String("val")).toString();
                // minMax, maxMin
            }
            else if ( reader.name() == QLatin1String("max") )
            {
                /*
                <xsd:complexType name="CT_Double">
                 <xsd:attribute name="val" type="xsd:double" use="required"/>
                </xsd:complexType>
                */

                double dMax = reader.attributes().value(QLatin1String("val")).toDouble();
            }
            else if ( reader.name() == QLatin1String("min") )
            {
                /*
                <xsd:complexType name="CT_Double">
                 <xsd:attribute name="val" type="xsd:double" use="required"/>
                </xsd:complexType>
                */

                double dMin = reader.attributes().value(QLatin1String("val")).toDouble();
            }
            else if ( reader.name() == QLatin1String("extLst") )
            {
                /*
                    <xsd:complexType name="CT_ExtensionList">
                     <xsd:sequence>
                      <xsd:element name="ext" type="CT_Extension" minOccurs="0" maxOccurs="unbounded"/>
                     </xsd:sequence>
                    </xsd:complexType>

                    <xsd:complexType name="CT_Extension">
                     <xsd:sequence>
                      <xsd:any processContents="lax"/>
                     </xsd:sequence>
                     <xsd:attribute name="uri" type="xsd:token"/>
                    </xsd:complexType>
                */

                loadExtList(reader, axis);
            }
            else
            {
            }
        }
        else if ( reader.tokenType() == QXmlStreamReader::EndElement &&
                  reader.name() == "scaling" )
        {
            break;
        }
    }

    return true;
}

bool ChartPrivate::loadExtList(QXmlStreamReader &reader, XlsxAxis* axis)
{
    Q_ASSERT(reader.name() == QLatin1String("extLst"));

    while ( !reader.atEnd() )
    {
        reader.readNextStartElement();
        if ( reader.tokenType() == QXmlStreamReader::StartElement )
        {
            if ( reader.name() == QLatin1String("ext") )
            {
                QString ctExtension = reader.attributes().value(QLatin1String("uri")).toString();
            }
        }
        else if ( reader.tokenType() == QXmlStreamReader::EndElement &&
                  reader.name() == "extLst" )
        {
            break;
        }
    }

    return true;
}

/*
  <xsd:complexType name="CT_Title">
      <xsd:sequence>
          <xsd:element name="tx" type="CT_Tx" minOccurs="0" maxOccurs="1"/>
          <xsd:element name="layout" type="CT_Layout" minOccurs="0" maxOccurs="1"/>
          <xsd:element name="overlay" type="CT_Boolean" minOccurs="0" maxOccurs="1"/>
          <xsd:element name="spPr" type="a:CT_ShapeProperties" minOccurs="0" maxOccurs="1"/>
          <xsd:element name="txPr" type="a:CT_TextBody" minOccurs="0" maxOccurs="1"/>
          <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
      </xsd:sequence>
  </xsd:complexType>

<xsd:complexType name="CT_Tx">
    <xsd:sequence>
        <xsd:choice minOccurs="1" maxOccurs="1">
            <xsd:element name="strRef" type="CT_StrRef" minOccurs="1" maxOccurs="1"/>
            <xsd:element name="rich" type="a:CT_TextBody" minOccurs="1" maxOccurs="1"/>
        </xsd:choice>
    </xsd:sequence>
</xsd:complexType>

<xsd:complexType name="CT_StrRef">
    <xsd:sequence>
        <xsd:element name="f" type="xsd:string" minOccurs="1" maxOccurs="1"/>
        <xsd:element name="strCache" type="CT_StrData" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
</xsd:complexType>

<xsd:complexType name="CT_TextBody">
    <xsd:sequence>
        <xsd:element name="bodyPr" type="CT_TextBodyProperties" minOccurs="1" maxOccurs="1"/>
        <xsd:element name="lstStyle" type="CT_TextListStyle" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="p" type="CT_TextParagraph" minOccurs="1" maxOccurs="unbounded"/>
    </xsd:sequence>
</xsd:complexType>
  */
bool ChartPrivate::loadXmlAxisEG_AxShared_Title(QXmlStreamReader &reader, XlsxAxis* axis)
{
    Q_ASSERT(reader.name() == QLatin1String("title"));

    while ( !reader.atEnd() )
    {
        reader.readNextStartElement();

        if ( reader.tokenType() == QXmlStreamReader::StartElement )
        {
            if ( reader.name() == QLatin1String("tx") )
            {
                loadXmlAxisEG_AxShared_Title_Tx(reader, axis);
            }
            else if ( reader.name() == QLatin1String("overlay") )
            {
                //!TODO: load overlay
                loadXmlAxisEG_AxShared_Title_Overlay(reader, axis);
            }
            else if ( reader.name() == QLatin1String("layout") )
            {
                //!TODO
            }
            else if ( reader.name() == QLatin1String("spPr") )
            {
                //!TODO
            }
            else if ( reader.name() == QLatin1String("txPr") )
            {
                //!TODO
            }
            else if ( reader.name() == QLatin1String("extLst") )
            {
                //!TODO
            }
            else
            {
            }
        }
        else if ( reader.tokenType() == QXmlStreamReader::EndElement &&
                  reader.name() == "title" )
        {
            break;
        }
    }

    return true;
}

bool ChartPrivate::loadXmlAxisEG_AxShared_Title_Overlay(QXmlStreamReader &reader, XlsxAxis* axis)
{
    Q_ASSERT(reader.name() == QLatin1String("overlay"));

    while ( !reader.atEnd() )
    {
        reader.readNextStartElement();
        if ( reader.tokenType() == QXmlStreamReader::StartElement )
        {
        }
        else if ( reader.tokenType() == QXmlStreamReader::EndElement &&
                  reader.name() == "overlay" )
        {
            break;
        }
    }

    return true;
}

bool ChartPrivate::loadXmlAxisEG_AxShared_Title_Tx(QXmlStreamReader &reader, XlsxAxis* axis)
{
    Q_ASSERT(reader.name() == QLatin1String("tx")); // c:tx

    while ( !reader.atEnd() )
    {
        reader.readNextStartElement();
        if ( reader.tokenType() == QXmlStreamReader::StartElement )
        {
            if ( reader.name() == QLatin1String("rich") )
            {
                loadXmlAxisEG_AxShared_Title_Tx_Rich(reader, axis);
            }
            else
            {
            }
        }
        else if ( reader.tokenType() == QXmlStreamReader::EndElement &&
                  reader.name() == "tx" )
        {
            break;
        }
    }

    return true;
}

bool ChartPrivate::loadXmlAxisEG_AxShared_Title_Tx_Rich(QXmlStreamReader &reader, XlsxAxis* axis)
{
    Q_ASSERT(reader.name() == QLatin1String("rich")); // c:rich

    while ( !reader.atEnd() )
    {
        reader.readNextStartElement();
        if ( reader.tokenType() == QXmlStreamReader::StartElement )
        {
            if ( reader.name() == QLatin1String("p") )
            {
                loadXmlAxisEG_AxShared_Title_Tx_Rich_P(reader, axis);
            }
            else if ( reader.name() == QLatin1String("bodyPr") )
            {
                //!TODO


            }
            else if ( reader.name() == QLatin1String("lstStyle") )
            {
                //!TODO
            }
            else
            {
            }
        }
        else if ( reader.tokenType() == QXmlStreamReader::EndElement &&
                  reader.name() == "rich" )
        {
            break;
        }
    }

    return true;
}

bool ChartPrivate::loadXmlAxisEG_AxShared_Title_Tx_Rich_P(QXmlStreamReader &reader, XlsxAxis* axis)
{
    Q_ASSERT(reader.name() == QLatin1String("p"));

    while ( !reader.atEnd() )
    {
        reader.readNextStartElement();
        if ( reader.tokenType() == QXmlStreamReader::StartElement )
        {
            if ( reader.name() == QLatin1String("r") )
            {
                loadXmlAxisEG_AxShared_Title_Tx_Rich_P_R(reader, axis);
            }
            else if ( reader.name() == QLatin1String("pPr") )
            {
                loadXmlAxisEG_AxShared_Title_Tx_Rich_P_pPr(reader, axis);
            }
            else
            {

            }
        }
        else if ( reader.tokenType() == QXmlStreamReader::EndElement &&
                  reader.name() == "p" )
        {
            break;
        }
    }

    return true;
}

bool ChartPrivate::loadXmlAxisEG_AxShared_Title_Tx_Rich_P_pPr(QXmlStreamReader &reader, XlsxAxis* axis)
{
    Q_ASSERT(reader.name() == QLatin1String("pPr"));

    while ( !reader.atEnd() )
    {
        reader.readNextStartElement();
        if ( reader.tokenType() == QXmlStreamReader::StartElement )
        {
            if ( reader.name() == QLatin1String("defRPr") )
            {
                QString strDefRPr = reader.readElementText();
                int debugLine = 0;
            }
            else
            {
            }
        }
        else if ( reader.tokenType() == QXmlStreamReader::EndElement &&
                  reader.name() == "pPr" )
        {
            break;
        }
    }

    return true;
}

bool ChartPrivate::loadXmlAxisEG_AxShared_Title_Tx_Rich_P_R(QXmlStreamReader &reader, XlsxAxis* axis)
{
    Q_ASSERT(reader.name() == QLatin1String("r"));

    while ( !reader.atEnd() )
    {
        reader.readNextStartElement();
        if ( reader.tokenType() == QXmlStreamReader::StartElement )
        {
            if ( reader.name() == QLatin1String("t") )
            {
                QString strAxisName = reader.readElementText();
                XlsxAxis::AxisPos axisPos = axis->axisPos;
                axis->axisNames[ axisPos ] = strAxisName;
            }
            else if ( reader.name() == QLatin1String("rPr") )
            {
                //

                //!TODO
                //! loadXmlAxisEG_AxShared_Title_Tx_Rich_P_R_rPr(reader, axis);

                // a:rPr
                //       a:solidFill
                //                    a:srgbClr
                //       a:latin


            }
            else
            {
            }
        }
        else if ( reader.tokenType() == QXmlStreamReader::EndElement &&
                  reader.name() == "r" )
        {
            break;
        }
    }

    return true;
}

/*
<xsd:complexType name="CT_PlotArea">
    <xsd:sequence>
        <xsd:element name="layout" type="CT_Layout" minOccurs="0" maxOccurs="1"/>
        <xsd:choice minOccurs="1" maxOccurs="unbounded">
            <xsd:element name="areaChart" type="CT_AreaChart" minOccurs="1" maxOccurs="1"/>
            <xsd:element name="area3DChart" type="CT_Area3DChart" minOccurs="1" maxOccurs="1"/>
            <xsd:element name="lineChart" type="CT_LineChart" minOccurs="1" maxOccurs="1"/>
            <xsd:element name="line3DChart" type="CT_Line3DChart" minOccurs="1" maxOccurs="1"/>
            <xsd:element name="stockChart" type="CT_StockChart" minOccurs="1" maxOccurs="1"/>
            <xsd:element name="radarChart" type="CT_RadarChart" minOccurs="1" maxOccurs="1"/>
            <xsd:element name="scatterChart" type="CT_ScatterChart" minOccurs="1" maxOccurs="1"/>
            <xsd:element name="pieChart" type="CT_PieChart" minOccurs="1" maxOccurs="1"/>
            <xsd:element name="pie3DChart" type="CT_Pie3DChart" minOccurs="1" maxOccurs="1"/>
            <xsd:element name="doughnutChart" type="CT_DoughnutChart" minOccurs="1" maxOccurs="1"/>
            <xsd:element name="barChart" type="CT_BarChart" minOccurs="1" maxOccurs="1"/>
            <xsd:element name="bar3DChart" type="CT_Bar3DChart" minOccurs="1" maxOccurs="1"/>
            <xsd:element name="ofPieChart" type="CT_OfPieChart" minOccurs="1" maxOccurs="1"/>
            <xsd:element name="surfaceChart" type="CT_SurfaceChart" minOccurs="1" maxOccurs="1"/>
            <xsd:element name="surface3DChart" type="CT_Surface3DChart" minOccurs="1" maxOccurs="1"/>
            <xsd:element name="bubbleChart" type="CT_BubbleChart" minOccurs="1" maxOccurs="1"/>
        </xsd:choice>
        <xsd:choice minOccurs="0" maxOccurs="unbounded">
            <xsd:element name="valAx" type="CT_ValAx" minOccurs="1" maxOccurs="1"/>
            <xsd:element name="catAx" type="CT_CatAx" minOccurs="1" maxOccurs="1"/>
            <xsd:element name="dateAx" type="CT_DateAx" minOccurs="1" maxOccurs="1"/>
            <xsd:element name="serAx" type="CT_SerAx" minOccurs="1" maxOccurs="1"/>
        </xsd:choice>
        <xsd:element name="dTable" type="CT_DTable" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="spPr" type="a:CT_ShapeProperties" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
</xsd:complexType>
*/

/*
<xsd:complexType name="CT_CatAx">
    <xsd:sequence>
        <xsd:group ref="EG_AxShared" minOccurs="1" maxOccurs="1"/>
        <xsd:element name="auto" type="CT_Boolean" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="lblAlgn" type="CT_LblAlgn" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="lblOffset" type="CT_LblOffset" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="tickLblSkip" type="CT_Skip" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="tickMarkSkip" type="CT_Skip" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="noMultiLvlLbl" type="CT_Boolean" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
</xsd:complexType>
<!----------------------------------------------------------------------------->
<xsd:complexType name="CT_DateAx">
    <xsd:sequence>
        <xsd:group ref="EG_AxShared" minOccurs="1" maxOccurs="1"/>
        <xsd:element name="auto" type="CT_Boolean" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="lblOffset" type="CT_LblOffset" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="baseTimeUnit" type="CT_TimeUnit" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="majorUnit" type="CT_AxisUnit" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="majorTimeUnit" type="CT_TimeUnit" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="minorUnit" type="CT_AxisUnit" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="minorTimeUnit" type="CT_TimeUnit" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
</xsd:complexType>
<!----------------------------------------------------------------------------->
<xsd:complexType name="CT_SerAx">
    <xsd:sequence>
    <xsd:group ref="EG_AxShared" minOccurs="1" maxOccurs="1"/>
        <xsd:element name="tickLblSkip" type="CT_Skip" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="tickMarkSkip" type="CT_Skip" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
</xsd:complexType>
<!----------------------------------------------------------------------------->
<xsd:complexType name="CT_ValAx">
    <xsd:sequence>
        <xsd:group ref="EG_AxShared" minOccurs="1" maxOccurs="1"/>
        <xsd:element name="crossBetween" type="CT_CrossBetween" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="majorUnit" type="CT_AxisUnit" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="minorUnit" type="CT_AxisUnit" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="dispUnits" type="CT_DispUnits" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
</xsd:complexType>
*/

void ChartPrivate::saveXmlAxis(QXmlStreamWriter &writer) const
{
    for ( int i = 0 ; i < axisList.size() ; ++i )
    {
        XlsxAxis* axis = axisList[i].data();
        if ( NULL == axis )
            continue;

        if ( axis->type == XlsxAxis::T_Cat  ) { saveXmlAxisCatAx( writer, axis ); }
        if ( axis->type == XlsxAxis::T_Val  ) { saveXmlAxisValAx( writer, axis ); }
        if ( axis->type == XlsxAxis::T_Ser  ) { saveXmlAxisSerAx( writer, axis ); }
        if ( axis->type == XlsxAxis::T_Date ) { saveXmlAxisDateAx( writer, axis ); }
    }

}

void ChartPrivate::saveXmlAxisCatAx(QXmlStreamWriter &writer, XlsxAxis* axis) const
{
/*
<xsd:complexType name="CT_CatAx">
    <xsd:sequence>
        <xsd:group ref="EG_AxShared" minOccurs="1" maxOccurs="1"/>
        <xsd:element name="auto" type="CT_Boolean" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="lblAlgn" type="CT_LblAlgn" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="lblOffset" type="CT_LblOffset" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="tickLblSkip" type="CT_Skip" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="tickMarkSkip" type="CT_Skip" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="noMultiLvlLbl" type="CT_Boolean" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
</xsd:complexType>
*/

    writer.writeStartElement("c:catAx");

    saveXmlAxisEG_AxShared(writer, axis); // EG_AxShared

    //!TODO: write element
    // auto
    // lblAlgn
    // lblOffset
    // tickLblSkip
    // tickMarkSkip
    // noMultiLvlLbl
    // extLst

    writer.writeEndElement(); // c:catAx
}

void ChartPrivate::saveXmlAxisDateAx(QXmlStreamWriter &writer, XlsxAxis* axis) const
{
/*
<xsd:complexType name="CT_DateAx">
    <xsd:sequence>
        <xsd:group ref="EG_AxShared" minOccurs="1" maxOccurs="1"/>
        <xsd:element name="auto" type="CT_Boolean" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="lblOffset" type="CT_LblOffset" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="baseTimeUnit" type="CT_TimeUnit" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="majorUnit" type="CT_AxisUnit" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="majorTimeUnit" type="CT_TimeUnit" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="minorUnit" type="CT_AxisUnit" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="minorTimeUnit" type="CT_TimeUnit" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
</xsd:complexType>
*/

    writer.writeStartElement("c:dateAx");

    saveXmlAxisEG_AxShared(writer, axis); // EG_AxShared

    //!TODO: write element
    // auto
    // lblOffset
    // baseTimeUnit
    // majorUnit
    // majorTimeUnit
    // minorUnit
    // minorTimeUnit
    // extLst

    writer.writeEndElement(); // c:dateAx
}

void ChartPrivate::saveXmlAxisSerAx(QXmlStreamWriter &writer, XlsxAxis* axis) const
{
/*
<xsd:complexType name="CT_SerAx">
    <xsd:sequence>
    <xsd:group ref="EG_AxShared" minOccurs="1" maxOccurs="1"/>
        <xsd:element name="tickLblSkip" type="CT_Skip" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="tickMarkSkip" type="CT_Skip" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
</xsd:complexType>
*/

    writer.writeStartElement("c:serAx");

    saveXmlAxisEG_AxShared(writer, axis); // EG_AxShared

    //!TODO: write element
    // tickLblSkip
    // tickMarkSkip
    // extLst

    writer.writeEndElement(); // c:serAx
}

void ChartPrivate::saveXmlAxisValAx(QXmlStreamWriter &writer, XlsxAxis* axis) const
{
/*
<xsd:complexType name="CT_ValAx">
    <xsd:sequence>
        <xsd:group ref="EG_AxShared" minOccurs="1" maxOccurs="1"/>
        <xsd:element name="crossBetween" type="CT_CrossBetween" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="majorUnit" type="CT_AxisUnit" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="minorUnit" type="CT_AxisUnit" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="dispUnits" type="CT_DispUnits" minOccurs="0" maxOccurs="1"/>
        <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
</xsd:complexType>
*/

    writer.writeStartElement("c:valAx");

    saveXmlAxisEG_AxShared(writer, axis); // EG_AxShared

    //!TODO: write element
    // crossBetween
    // majorUnit
    // minorUnit
    // dispUnits
    // extLst

    writer.writeEndElement(); // c:valAx
}

void ChartPrivate::saveXmlAxisEG_AxShared(QXmlStreamWriter &writer, XlsxAxis* axis) const
{
    /*
    <xsd:group name="EG_AxShared">
        <xsd:sequence>
            <xsd:element name="axId" type="CT_UnsignedInt" minOccurs="1" maxOccurs="1"/> (*)
            <xsd:element name="scaling" type="CT_Scaling" minOccurs="1" maxOccurs="1"/> (*)
            <xsd:element name="delete" type="CT_Boolean" minOccurs="0" maxOccurs="1"/>
            <xsd:element name="axPos" type="CT_AxPos" minOccurs="1" maxOccurs="1"/> (*)
            <xsd:element name="majorGridlines" type="CT_ChartLines" minOccurs="0" maxOccurs="1"/>
            <xsd:element name="minorGridlines" type="CT_ChartLines" minOccurs="0" maxOccurs="1"/>
            <xsd:element name="title" type="CT_Title" minOccurs="0" maxOccurs="1"/> (***********************)
            <xsd:element name="numFmt" type="CT_NumFmt" minOccurs="0" maxOccurs="1"/>
            <xsd:element name="majorTickMark" type="CT_TickMark" minOccurs="0" maxOccurs="1"/>
            <xsd:element name="minorTickMark" type="CT_TickMark" minOccurs="0" maxOccurs="1"/>
            <xsd:element name="tickLblPos" type="CT_TickLblPos" minOccurs="0" maxOccurs="1"/>
            <xsd:element name="spPr" type="a:CT_ShapeProperties" minOccurs="0" maxOccurs="1"/>
            <xsd:element name="txPr" type="a:CT_TextBody" minOccurs="0" maxOccurs="1"/>
            <xsd:element name="crossAx" type="CT_UnsignedInt" minOccurs="1" maxOccurs="1"/> (*)
            <xsd:choice minOccurs="0" maxOccurs="1">
                <xsd:element name="crosses" type="CT_Crosses" minOccurs="1" maxOccurs="1"/>
                <xsd:element name="crossesAt" type="CT_Double" minOccurs="1" maxOccurs="1"/>
            </xsd:choice>
        </xsd:sequence>
    </xsd:group>
    */

    writer.writeEmptyElement(QStringLiteral("c:axId")); // 21.2.2.9. axId (Axis ID) (mandatory value)
        writer.writeAttribute(QStringLiteral("val"), QString::number(axis->axisId));

    writer.writeStartElement(QStringLiteral("c:scaling")); // CT_Scaling (mandatory value)
        writer.writeEmptyElement(QStringLiteral("c:orientation")); // CT_Orientation
            writer.writeAttribute(QStringLiteral("val"), QStringLiteral("minMax")); // ST_Orientation
    writer.writeEndElement(); // c:scaling

    writer.writeEmptyElement(QStringLiteral("c:axPos")); // axPos CT_AxPos (mandatory value)
        QString pos = GetAxisPosString( axis->axisPos );
        if ( !pos.isEmpty() )
        {
            writer.writeAttribute(QStringLiteral("val"), pos); // ST_AxPos
        }

    saveXmlAxisEG_AxShared_Title(writer, axis); // "c:title" CT_Title

    writer.writeEmptyElement(QStringLiteral("c:crossAx")); // crossAx (mandatory value)
        writer.writeAttribute(QStringLiteral("val"), QString::number(axis->crossAx));

}

void ChartPrivate::saveXmlAxisEG_AxShared_Title(QXmlStreamWriter &writer, XlsxAxis* axis) const
{
    // CT_Title

    /*
    <xsd:complexType name="CT_Title">
        <xsd:sequence>
            <xsd:element name="tx" type="CT_Tx" minOccurs="0" maxOccurs="1"/>
            <xsd:element name="layout" type="CT_Layout" minOccurs="0" maxOccurs="1"/>
            <xsd:element name="overlay" type="CT_Boolean" minOccurs="0" maxOccurs="1"/>
            <xsd:element name="spPr" type="a:CT_ShapeProperties" minOccurs="0" maxOccurs="1"/>
            <xsd:element name="txPr" type="a:CT_TextBody" minOccurs="0" maxOccurs="1"/>
            <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
        </xsd:sequence>
    </xsd:complexType>
    */
    /*
    <xsd:complexType name="CT_Tx">
        <xsd:sequence>
            <xsd:choice minOccurs="1" maxOccurs="1">
                <xsd:element name="strRef" type="CT_StrRef" minOccurs="1" maxOccurs="1"/>
                <xsd:element name="rich" type="a:CT_TextBody" minOccurs="1" maxOccurs="1"/>
            </xsd:choice>
        </xsd:sequence>
    </xsd:complexType>
    */
    /*
    <xsd:complexType name="CT_StrRef">
        <xsd:sequence>
            <xsd:element name="f" type="xsd:string" minOccurs="1" maxOccurs="1"/>
            <xsd:element name="strCache" type="CT_StrData" minOccurs="0" maxOccurs="1"/>
            <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
        </xsd:sequence>
    </xsd:complexType>
    */
    /*
    <xsd:complexType name="CT_TextBody">
        <xsd:sequence>
            <xsd:element name="bodyPr" type="CT_TextBodyProperties" minOccurs="1" maxOccurs="1"/>
            <xsd:element name="lstStyle" type="CT_TextListStyle" minOccurs="0" maxOccurs="1"/>
            <xsd:element name="p" type="CT_TextParagraph" minOccurs="1" maxOccurs="unbounded"/>
        </xsd:sequence>
    </xsd:complexType>
    */

    writer.writeStartElement("c:title");

    // CT_Tx {{
     writer.writeStartElement("c:tx");

      writer.writeStartElement("c:rich"); // CT_TextBody

       writer.writeEmptyElement(QStringLiteral("a:bodyPr")); // CT_TextBodyProperties

       writer.writeEmptyElement(QStringLiteral("a:lstStyle")); // CT_TextListStyle

       writer.writeStartElement("a:p");

        writer.writeStartElement("a:pPr");
            writer.writeAttribute(QStringLiteral("lvl"), QString::number(0));

            writer.writeStartElement("a:defRPr");
            writer.writeAttribute(QStringLiteral("b"), QString::number(0));
            writer.writeEndElement(); // a:defRPr
        writer.writeEndElement(); // a:pPr

        writer.writeStartElement("a:r");
        QString strAxisName = GetAxisName(axis);
        writer.writeTextElement( QStringLiteral("a:t"), strAxisName );
        writer.writeEndElement(); // a:r

       writer.writeEndElement(); // a:p

      writer.writeEndElement(); // c:rich

     writer.writeEndElement(); // c:tx
     // CT_Tx }}

     writer.writeStartElement("c:overlay");
        writer.writeAttribute(QStringLiteral("val"), QString::number(0)); // CT_Boolean
     writer.writeEndElement(); // c:overlay

    writer.writeEndElement(); // c:title

}

QString ChartPrivate::GetAxisPosString( XlsxAxis::AxisPos axisPos ) const
{
    QString pos;
    switch ( axisPos )
    {
        case XlsxAxis::Top    : pos = QStringLiteral("t"); break;
        case XlsxAxis::Bottom : pos = QStringLiteral("b"); break;
        case XlsxAxis::Left   : pos = QStringLiteral("l"); break;
        case XlsxAxis::Right  : pos = QStringLiteral("r"); break;
        default: break; // ??
    }

    return pos;
}

QString ChartPrivate::GetAxisName(XlsxAxis* axis) const
{
    QString strAxisName;
    if ( NULL == axis )
        return strAxisName;

    QString pos = GetAxisPosString( axis->axisPos ); // l, t, r, b
    if ( pos.isEmpty() )
        return strAxisName;

    strAxisName = axis->axisNames[ axis->axisPos ];
    return strAxisName;
}


QT_END_NAMESPACE_XLSX
