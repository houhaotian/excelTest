#include "AxManager.h"
#include <QAxObject>
#include <QDebug>

#pragma execution_character_set("utf-8")

AxManager::AxManager(QObject *parent)
    : QObject(parent)
{
}

AxManager::~AxManager()
{
}

int AxManager::openExcelFile(const QString &path)
{
    excel = new QAxObject("Excel.Application");
    excel->dynamicCall("SetVisible(bool)", false); //true 表示操作文件时可见，false表示为不可见
    workbooks = excel->querySubObject("WorkBooks");

    //————————————————按文件路径打开文件————————————————————
    workbook = workbooks->querySubObject("Open(QString&)", path);


    std::string str = path.toStdString();
    str.append(":file do not exist!");
    const char* info = str.c_str();

    Q_ASSERT_X(workbook!=nullptr,"openExcelFile",info);
    // 获取打开的excel文件中所有的工作sheet
    worksheets = workbook->querySubObject("WorkSheets");

    //—————————————————Excel文件中表的个数:——————————————————
    int iWorkSheet = worksheets->property("Count").toInt();
    qDebug() << QString("Excel文件中表的个数: %1").arg(QString::number(iWorkSheet));

    return iWorkSheet;
}

void AxManager::setSheetIndex(int index)
{
    // ————————————————获取第n个工作表 querySubObject("Item(int)", n);——————————
    worksheet = worksheets->querySubObject("Item(int)", index);//本例获取第一个，最后参数填1

      //—————————获取该sheet的数据范围（可以理解为有数据的矩形区域）————
    usedrange = worksheet->querySubObject("UsedRange");
}

bool AxManager::closeExcelFile()
{
    workbook->dynamicCall("Close()");
    excel->dynamicCall("Quit()");
    if (excel)
    {
        delete excel;
        excel = NULL;
        qDebug() << "close?";
    }
    qDebug() << "close!";
    return true;
}

void AxManager::loadData()
{
    //———————————————————获取行数———————————————
    QAxObject * rows = usedrange->querySubObject("Rows");
    int iRows = rows->property("Count").toInt();
    qDebug() << QString("行数为: %1").arg(QString::number(iRows));

    /**************************************************************/
       //先读第一个部门的数据，即C列
    for (char x('C'); x <= 'F'; ++x) {
        auto hash = getHash(x);
        for (int row = 1; row <= iRows; ++row) {
            //——————————————读出数据—————————————
            QAxObject *range = worksheet->querySubObject("Range(QString)", QString("%1%2").arg(x).arg(row));
            QString value = range->property("Value").toString();

            //如果目标值小于等于0（可能不是数字）就不管
            bool ret;
            if (qFuzzyIsNull(value.toDouble(&ret))) {
                if (ret == false)
                    continue;
                continue;
            }
            QAxObject *range1 = worksheet->querySubObject("Range(QString)", QString("A%1").arg(row));
            QString menu = range1->property("Value").toString();
            qDebug() << QString("第A%1列为：").arg(row) + menu;
            qDebug() << QString("第%1%2列为：").arg(x).arg(row) + value;
            hash->insert(menu, value);
        }
    }
}

void AxManager::writeData(int index)
{
    //———————————————————获取行数———————————————
    QAxObject * rows = usedrange->querySubObject("Rows");
    int iRows = rows->property("Count").toInt();
    qDebug() << QString("行数为: %1").arg(QString::number(iRows));
  
    auto h = getHash(index);
    auto hash = *h;
    for (int row = 1; row <= iRows; ++row) {
        //遍历保存的key也就是总表的编码
        //读出目标文件的A列即编码value
        QAxObject *range2 = worksheet->querySubObject("Range(QString)", QString("%1%2").arg('A').arg(row));
        QString value = range2->property("Value").toString();
        for (auto key : hash.keys())
        {
            if (value == key)
            {
                //—————————————写入数据—————————————
                QAxObject *range2 = worksheet->querySubObject("Range(QString)", QString("D%1").arg(row));
                //写入数据, 第6行，第6列
                range2->setProperty("Value", QString(hash.value(key)));
                QString newStr = range2->property("Value").toString();
                qDebug() << "写入数据后数据为：" + newStr;
            }
        }
    }
}

QHash<QString, QString> * AxManager::getHash(char x)
{
    Q_ASSERT(x >= 'C'&&x <= 'F');
    switch (x)
    {
    case 'C':
        return &m_hash1;
    case 'D':
        return &m_hash2;
    case 'E':
        return &m_hash3;
    case 'F':
        return &m_hash4;
    }
}

QHash<QString, QString> * AxManager::getHash(int index)
{
    Q_ASSERT(index >= 1 && index <= 4);

    switch (index)
    {
    case 1:
        return &m_hash1;
    case 2:
        return &m_hash2;
    case 3:
        return &m_hash3;
    case 4:
        return &m_hash4;
    }
}
