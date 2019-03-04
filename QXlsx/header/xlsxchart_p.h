// xlsxchart_p.h

#ifndef QXLSX_CHART_P_H
#define QXLSX_CHART_P_H

#include <QtGlobal>
#include <QObject>
#include <QString>
#include <QSharedPointer>
#include <QVector>
#include <QMap>
#include <QList>
#include <QXmlStreamReader>
#include <QXmlStreamWriter>

#include "xlsxabstractooxmlfile_p.h"
#include "xlsxchart.h"

#include "xdattr.h"
#include "xdnode.h"
#include "xdxmldomreader.h"

QT_BEGIN_NAMESPACE_XLSX

class XlsxSeries
{
public:
    //At present, we care about number cell ranges only!
    QString numberDataSource_numRef; // yval, val
    QString axDataSource_numRef; // xval, cat
};

class XlsxAxis
{
public:
    enum Type { T_None = (-1), T_Cat, T_Val, T_Date, T_Ser };
    enum AxisPos { None = (-1), Left, Right, Top, Bottom };
public:
    XlsxAxis(){}

    XlsxAxis( Type axisType,
              XlsxAxis::AxisPos axisPos,
              int id,
              int crossId,
              QString axisTitle = QString("") )
    {
        type = axisType;
        axisPos = axisPos;
        axisId = id;
        crossAx = crossId;

        if ( !axisTitle.isEmpty() )
        {
            axisNames[ axisPos ] = axisTitle;
        }
    }

public:
    Type type;
    XlsxAxis::AxisPos axisPos;
    int axisId;
    uint crossAx;
    QMap< XlsxAxis::AxisPos, QString > axisNames;
};

class ChartPrivate : public AbstractOOXmlFilePrivate
{
    Q_DECLARE_PUBLIC(Chart)
public:
    ChartPrivate(Chart *q, Chart::CreateFlag flag);
    ~ChartPrivate();
public:
    bool loadXmlChart(QXmlStreamReader &reader);
protected:
    bool loadXmlPlotArea(QXmlStreamReader &reader);
    bool loadXmlPlotAreaElement(QXmlStreamReader &reader);
public:
    bool loadXmlXxxChart(QXmlStreamReader &reader);
    bool loadXmlSer(QXmlStreamReader &reader);
    QString loadXmlNumRef(QXmlStreamReader &reader);
    bool loadXmlChartTitle(QXmlStreamReader &reader);
protected:
    bool loadXmlChartTitleTx(QXmlStreamReader &reader);
    bool loadXmlChartTitleTxRich(QXmlStreamReader &reader);
    bool loadXmlChartTitleTxRichP(QXmlStreamReader &reader);
    bool loadXmlChartTitleTxRichP_R(QXmlStreamReader &reader);
protected:
    bool loadXmlAxisCatAx(QXmlStreamReader &reader);
    bool loadXmlAxisDateAx(QXmlStreamReader &reader);
    bool loadXmlAxisSerAx(QXmlStreamReader &reader);
    bool loadXmlAxisValAx(QXmlStreamReader &reader);
    bool loadXmlAxisEG_AxShared(QXmlStreamReader &reader, XlsxAxis* axis);
    bool loadTxPr(QXmlStreamReader &reader, XlsxAxis* axis);
    bool loadTxPr_BodyPr(QXmlStreamReader &reader, XlsxAxis* axis);
    bool loadTxPr_LstStyle(QXmlStreamReader &reader, XlsxAxis* axis);
    bool loadTxPr_P(QXmlStreamReader &reader, XlsxAxis* axis);
    bool loadXmlAxisEG_AxShared_Scaling(QXmlStreamReader &reader, XlsxAxis* axis);
    bool loadExtList(QXmlStreamReader &reader, XlsxAxis* axis);
    bool loadXmlAxisEG_AxShared_Title(QXmlStreamReader &reader, XlsxAxis* axis);
    bool loadXmlAxisEG_AxShared_Title_Overlay(QXmlStreamReader &reader, XlsxAxis* axis);
    bool loadXmlAxisEG_AxShared_Title_Tx(QXmlStreamReader &reader, XlsxAxis* axis);
    bool loadXmlAxisEG_AxShared_Title_Tx_Rich(QXmlStreamReader &reader, XlsxAxis* axis);
    bool loadXmlAxisEG_AxShared_Title_Tx_Rich_P(QXmlStreamReader &reader, XlsxAxis* axis);
    bool loadXmlAxisEG_AxShared_Title_Tx_Rich_P_pPr(QXmlStreamReader &reader, XlsxAxis* axis);
    bool loadXmlAxisEG_AxShared_Title_Tx_Rich_P_R(QXmlStreamReader &reader, XlsxAxis* axis);

public:
    bool loadFromXmlFile(QIODevice *device);
protected:
    bool load1Chart(XMLDOM::XMLDOMReader *pReader, XMLDOM::Node* ptrChart);
    bool load1Lang(XMLDOM::XMLDOMReader *pReader, XMLDOM::Node* ptrLang);
    bool load1PrinterSettings(XMLDOM::XMLDOMReader *pReader, XMLDOM::Node* ptrPrinterSettings);
protected:
    bool load2Title(XMLDOM::XMLDOMReader* pReader, XMLDOM::Node* ptrTitle );
    bool load2PlotArea(XMLDOM::XMLDOMReader* pReader, XMLDOM::Node* ptrPlotArea );
protected:
    bool load3AreaChart(XMLDOM::XMLDOMReader* pReader, XMLDOM::Node* ptrAreaChart );
    bool load3Aread3DChart(XMLDOM::XMLDOMReader* pReader, XMLDOM::Node* ptrArea3DChart );
    bool load3LineChart(XMLDOM::XMLDOMReader* pReader, XMLDOM::Node* ptrLineChart );
    bool load3Line3DChart(XMLDOM::XMLDOMReader* pReader, XMLDOM::Node* ptrLine3DChart );
    bool load3StockChart(XMLDOM::XMLDOMReader* pReader, XMLDOM::Node* ptrStockChart );
    bool load3RadarChart(XMLDOM::XMLDOMReader* pReader, XMLDOM::Node* ptrRadarChart );
    bool load3SactterChart(XMLDOM::XMLDOMReader* pReader, XMLDOM::Node* ptrScatterChart );
    bool load3PieChart(XMLDOM::XMLDOMReader* pReader, XMLDOM::Node* ptrPiChart );
    bool load3Pie3DChart(XMLDOM::XMLDOMReader* pReader, XMLDOM::Node* ptrPie3DChart );
    bool load3DoughnutChart(XMLDOM::XMLDOMReader* pReader, XMLDOM::Node* ptrDoughnutChart );

    bool load3BarChart(XMLDOM::XMLDOMReader* pReader, XMLDOM::Node* ptrBarChart );
    bool load3Bar3DChart(XMLDOM::XMLDOMReader* pReader, XMLDOM::Node* ptrBar3DChart );
    bool load3EG_BarChartShared(XMLDOM::XMLDOMReader* pReader, XMLDOM::Node* ptrBarChart);

    bool load3OfPieChart(XMLDOM::XMLDOMReader* pReader, XMLDOM::Node* ptrOfPieChart );
    bool load3SurfaceChart(XMLDOM::XMLDOMReader* pReader, XMLDOM::Node* ptrSurfaceChart );
    bool load3Surface3DChart(XMLDOM::XMLDOMReader* pReader, XMLDOM::Node*ptrSurface3DChart );
    bool load3BubbleChart(XMLDOM::XMLDOMReader* pReader, XMLDOM::Node* ptrBubbleChart );
protected:
    bool load2Legend(XMLDOM::XMLDOMReader* pReader, XMLDOM::Node* ptrLegend );
    bool load2PlotVisOnly(XMLDOM::XMLDOMReader* pReader, XMLDOM::Node* ptrPlotVisOnly );

public:
    void saveXmlChart(QXmlStreamWriter &writer) const;
    void saveXmlChartTitle(QXmlStreamWriter &writer) const;
    void saveXmlPieChart(QXmlStreamWriter &writer) const;
    void saveXmlBarChart(QXmlStreamWriter &writer) const;
    void saveXmlLineChart(QXmlStreamWriter &writer) const;
    void saveXmlScatterChart(QXmlStreamWriter &writer) const;
    void saveXmlAreaChart(QXmlStreamWriter &writer) const;
    void saveXmlDoughnutChart(QXmlStreamWriter &writer) const;
    void saveXmlSer(QXmlStreamWriter &writer, XlsxSeries *ser, int id) const;
    void saveXmlAxis(QXmlStreamWriter &writer) const;
protected:
    void saveXmlAxisCatAx(QXmlStreamWriter &writer, XlsxAxis* axis) const;
    void saveXmlAxisDateAx(QXmlStreamWriter &writer, XlsxAxis* axis) const;
    void saveXmlAxisSerAx(QXmlStreamWriter &writer, XlsxAxis* axis) const;
    void saveXmlAxisValAx(QXmlStreamWriter &writer, XlsxAxis* axis) const;
protected:
    void saveXmlAxisEG_AxShared(QXmlStreamWriter &writer, XlsxAxis* axis) const;
    void saveXmlAxisEG_AxShared_Title(QXmlStreamWriter &writer, XlsxAxis* axis) const;
    QString GetAxisPosString( XlsxAxis::AxisPos axisPos ) const;
    QString GetAxisName(XlsxAxis* ptrXlsxAxis) const;
public:
    Chart::ChartType chartType;
    QList< QSharedPointer<XlsxSeries> > seriesList;
    QList< QSharedPointer<XlsxAxis> > axisList;
    QMap< XlsxAxis::AxisPos, QString > axisNames;
    QString chartTitle;
    AbstractSheet* sheet;
};

QT_END_NAMESPACE_XLSX

#endif // QXLSX_CHART_P_H
