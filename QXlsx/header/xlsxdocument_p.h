// xlsxdocument_p.h
// QXlsx https://github.com/j2doll/QXlsx
// QtXlsx https://github.com/dbzhang800/QtXlsxWriter

#ifndef XLSXDOCUMENT_P_H
#define XLSXDOCUMENT_P_H

#include <QtGlobal>
#include <QObject>
#include <QString>
#include <QIODevice>
#include <QSharedPointer>
#include <QMap>

#include "xlsxdocument.h"
#include "xlsxworkbook.h"
#include "xlsxcontenttypes_p.h"

QT_BEGIN_NAMESPACE_XLSX

class DocumentPrivate
{
    Q_DECLARE_PUBLIC(Document)
public:
    DocumentPrivate(Document *p);
    void init();

    bool loadPackage(QIODevice *device);
    bool savePackage(QIODevice *device) const;

    Document *q_ptr;
    const QString defaultPackageName; //default name when package name not specified
    QString packageName; //name of the .xlsx file

    QMap<QString, QString> documentProperties; //core, app and custom properties
    QSharedPointer<Workbook> workbook;
    QSharedPointer<ContentTypes> contentTypes;
	bool isLoad; 
};

QT_END_NAMESPACE_XLSX

#endif // XLSXDOCUMENT_P_H
