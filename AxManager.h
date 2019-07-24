#pragma once

#include <QObject>
#include <QHash>

class QAxObject;

class AxManager : public QObject
{
    Q_OBJECT

public:

    AxManager(QObject *parent = 0);
    ~AxManager();
    bool Test(QString &path);


    QString ConvertFromNumber(int number)
    {
        QString resultStr = "";
        while (number > 0)
        {
            int k = number % 26;
            if (k == 0)
                k = 26;
            resultStr.push_front(QChar(k + 64));
            number = (number - k) / 26;
        }

        return resultStr;
    }
    //打开excel文件并返回表的个数，如果返回-1则打开失败
    int openExcelFile(const QString &path);
    void setSheetIndex(int index);
    bool closeExcelFile();
    void loadData();
    void writeData(int index);

private:
    QHash<QString, QString> *getHash(char x);
    QHash<QString, QString> *getHash(int index);
    char getAimColumn(int index);
private:
    //存编号，目标值
    QHash<QString, QString> m_hash1;//对应C列数据
    QHash<QString, QString> m_hash2;//对应D列数据
    QHash<QString, QString> m_hash3;//对应E列数据
    QHash<QString, QString> m_hash4;//对应F列数据
    
    QAxObject *excel;    //本例中，excel设定为Excel文件的操作对象
    QAxObject *workbooks;
    QAxObject *workbook;  //Excel操作对象
    QAxObject *worksheets;//打开的excel的所有sheet
    QAxObject *worksheet;//工作表
    QAxObject *usedrange;//有效的sheet工作范围
};
