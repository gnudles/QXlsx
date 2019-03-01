// main.cpp

#include <QtGlobal>
#include <QCoreApplication>
#include <QtCore>
#include <QVector>
#include <QVariant>
#include <QDebug> 

#include <iostream>
using namespace std;

#include "xlsxdocument.h"
#include "xlsxchartsheet.h"
#include "xlsxcellrange.h"
#include "xlsxchart.h"
#include "xlsxrichstring.h"
#include "xlsxworkbook.h"
using namespace QXlsx;

extern int test1(QVector<QVariant> params);
extern int test2(QVector<QVariant> params);

int main(int argc, char *argv[])
{
	QCoreApplication app(argc, argv);

    // QVector<QVariant> testParams1;
    // int ret = test1(testParams);
    // qDebug() << "test return value : " << ret;

    QVector<QVariant> testParams2;
    int ret = test2(testParams2);

	return 0; 
}

