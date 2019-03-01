// main.cpp

#include <QtGlobal>
#include <QObject>
#include <QString>
#include <QVector>
#include <QList>
#include <QVariant>
#include <QDir>
#include <QStringList>

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

bool loadAndSaveSlxsx(QString srcDir, QString destDir);

int main(int argc, char *argv[])
{
    QCoreApplication app(argc, argv);

    qDebug() << "[debug] current path : " << QDir::currentPath();

    // Fix directories for your own test environment;
    QString srcDir =  QDir::currentPath() + QString("/../xlsx_files/");
    QString dstDir =  QDir::currentPath() + QString("/../xlsx_files2/");

    loadAndSaveSlxsx( srcDir, dstDir );

    return 0;
}

bool loadAndSaveSlxsx(QString srcDir, QString destDir)
{
    QDir dir( srcDir );
    // QString s = dir.absoluteFilePath( testDir );
    QStringList el = dir.entryList(QStringList() << "*.*", QDir::Files);

    Q_FOREACH(QString entry, el)
    {
        using namespace QXlsx;

        QString srcFilePath = srcDir + entry;
        QString dstFilePath = destDir + entry;

        Document doc( srcFilePath );
        if ( !doc.load() )
        {
            qDebug() << "[debug] failed to load : " << srcFilePath;
            continue;
        }

        if ( !doc.saveAs( dstFilePath ) )
        {
            qDebug() << "[debug] failed to save : " << dstFilePath;
            continue;
        }

        qDebug() << dstFilePath;
    }

}
