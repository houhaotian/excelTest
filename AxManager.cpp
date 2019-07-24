﻿#include "AxManager.h"
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


bool AxManager::Test(QString &path)
{
    //QAxObject *excel = NULL;    //本例中，excel设定为Excel文件的操作对象
    //QAxObject *workbooks = NULL;
    //QAxObject *workbook = NULL;  //Excel操作对象
    //excel = new QAxObject("Excel.Application");
    //excel->dynamicCall("SetVisible(bool)", false); //true 表示操作文件时可见，false表示为不可见
    //workbooks = excel->querySubObject("WorkBooks");


    ////————————————————按文件路径打开文件————————————————————
    //workbook = workbooks->querySubObject("Open(QString&)", path);
    //// 获取打开的excel文件中所有的工作sheet
    //QAxObject * worksheets = workbook->querySubObject("WorkSheets");


    ////—————————————————Excel文件中表的个数:——————————————————
    //int iWorkSheet = worksheets->property("Count").toInt();
    //qDebug() << QString("Excel文件中表的个数: %1").arg(QString::number(iWorkSheet));


    //// ————————————————获取第n个工作表 querySubObject("Item(int)", n);——————————
    //QAxObject * worksheet = worksheets->querySubObject("Item(int)", 1);//本例获取第一个，最后参数填1


    ////—————————获取该sheet的数据范围（可以理解为有数据的矩形区域）————
    //QAxObject * usedrange = worksheet->querySubObject("UsedRange");

    ////———————————————————获取行数———————————————
    //QAxObject * rows = usedrange->querySubObject("Rows");
    //int iRows = rows->property("Count").toInt();
    //qDebug() << QString("行数为: %1").arg(QString::number(iRows));

    ////————————————获取列数—————————
    //QAxObject * columns = usedrange->querySubObject("Columns");
    //int iColumns = columns->property("Count").toInt();
    //qDebug() << QString("列数为: %1").arg(QString::number(iColumns));

    ////————————数据的起始行———
    //int iStartRow = rows->property("Row").toInt();
    //qDebug() << QString("起始行为: %1").arg(QString::number(iStartRow));

    ////————————数据的起始列————————————
    //int iColumn = columns->property("Column").toInt();
    //qDebug() << QString("起始列为: %1").arg(QString::number(iColumn));

    ///**************************************************************/
    ////先读第一个部门的数据，即C列
    //for (int row = 1; row < iRows; ++row) {
    //    //——————————————读出数据—————————————
    //    QAxObject *range = worksheet->querySubObject("Range(QString)", QString("C%1").arg(row));
    //    QString value = range->property("Value").toString();
    //    

    //    qDebug() << QString("第C%1列为：").arg(row) + value;
    //    //如果目标值小于等于0（可能不是数字）就不管
    //    bool ret;
    //    qDebug() << value.toUInt(&ret);
    //    if (value.toDouble(&ret) == 0) {
    //        if (ret == false)
    //            continue;
    //    }
    //    QAxObject *range1 = worksheet->querySubObject("Range(QString)", QString("A%1").arg(row));
    //    QString menu = range1->property("Value").toString();
    //    qDebug() << QString("第A%1列为：").arg(row) + menu;
    //    m_hash.insert(menu, value);
    //}


//—————————————写入数据—————————————
//获取F6的位置
    QAxObject *range2 = worksheet->querySubObject("Range(QString)", "F6");
    //写入数据, 第6行，第6列
    range2->setProperty("Value", "中共十九大");
    QString newStr = "";
    newStr = range2->property("Value").toString();
    qDebug() << "写入数据后，第6行，第6列的数据为：" + newStr;

    //!!!!!!!一定要记得close，不然系统进程里会出现n个EXCEL.EXE进程
    workbook->dynamicCall("Close()");
    excel->dynamicCall("Quit()");
    if (excel)
    {
        delete excel;
        excel = NULL;
    }

    return true;
}

int AxManager::openExcelFile(const QString &path)
{
    excel = new QAxObject("Excel.Application");
    excel->dynamicCall("SetVisible(bool)", false); //true 表示操作文件时可见，false表示为不可见
    workbooks = excel->querySubObject("WorkBooks");

    //————————————————按文件路径打开文件————————————————————
    workbook = workbooks->querySubObject("Open(QString&)", path);
    // 获取打开的excel文件中所有的工作sheet
    QAxObject * worksheets = workbook->querySubObject("WorkSheets");

    //—————————————————Excel文件中表的个数:——————————————————
    int iWorkSheet = worksheets->property("Count").toInt();
    qDebug() << QString("Excel文件中表的个数: %1").arg(QString::number(iWorkSheet));

    // ————————————————获取第n个工作表 querySubObject("Item(int)", n);——————————
    worksheet = worksheets->querySubObject("Item(int)", 1);//本例获取第一个，最后参数填1

      //—————————获取该sheet的数据范围（可以理解为有数据的矩形区域）————
    usedrange = worksheet->querySubObject("UsedRange");

    return iWorkSheet;
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
    for (char x('C'); x < 'F'; ++x) {
        for (int row = 1; row < iRows; ++row) {
            //——————————————读出数据—————————————
            QAxObject *range = worksheet->querySubObject("Range(QString)", QString("%1%2").arg(x).arg(row));
            QString value = range->property("Value").toString();

            //如果目标值小于等于0（可能不是数字）就不管
            bool ret;
            if (value.toDouble(&ret) == 0) {
                if (ret == false)
                    continue;
            }
            QAxObject *range1 = worksheet->querySubObject("Range(QString)", QString("A%1").arg(row));
            QString menu = range1->property("Value").toString();
            qDebug() << QString("第A%1列为：").arg(row) + menu;
            qDebug() << QString("第C%1列为：").arg(row) + value;
            m_hash.insert(menu, value);
        }
    }
   
}