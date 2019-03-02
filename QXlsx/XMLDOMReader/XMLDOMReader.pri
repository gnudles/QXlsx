#
# XMLDOMReader.pri
#
# https://github.com/j2doll/XMLDOMReader

# You may fix environment value for your own project
isEmpty(XMLDOMREADER_PARENTPATH) {
    XMLDOMREADER_PARENTPATH = ../QXlsx/XMLDOMReader/
    # XMLDOMREADER_PARENTPATH = ../XMLDOMReader/
}

QT += xml

######################################################################
# source code 

INCLUDEPATH += $${XMLDOMREADER_PARENTPATH}

HEADERS += \
$${XMLDOMREADER_PARENTPATH}Attr.h \
$${XMLDOMREADER_PARENTPATH}Node.h \
$${XMLDOMREADER_PARENTPATH}XMLDOMReader.h

SOURCES += \
$${XMLDOMREADER_PARENTPATH}Attr.cpp \
$${XMLDOMREADER_PARENTPATH}Node.cpp \
$${XMLDOMREADER_PARENTPATH}XMLDOMReader.cpp




