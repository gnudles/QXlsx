// main.cpp

#include <QtGlobal>
#include <QObject>
#include <QString>
#include <QVector>
#include <QList>
#include <QVariant>
#include <QDir>

#include <QDebug>

#include <QCoreApplication>

#include <iostream>
using namespace std;

#include "xlsxdocument.h"
#include "xlsxchartsheet.h"
#include "xlsxcellrange.h"
#include "xlsxchart.h"
#include "xlsxrichstring.h"
#include "xlsxworkbook.h"
using namespace QXlsx;

bool loadAndSaveSlxsx(QString testDir);

int main(int argc, char *argv[])
{
    QCoreApplication app(argc, argv);

    qDebug() << "[debug] current path : " << QDir::currentPath();

    // Fix testDir for your own test environment;
    QString testDir
    testDir =  QDir::currentPath() + QString("../xlsx_files/");

    loadAndSaveSlxsx(testDir);

    return 0;
}

bool loadAndSaveSlxsx(QString testDir)
{




}
