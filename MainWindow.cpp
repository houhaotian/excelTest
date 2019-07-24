#include "MainWindow.h"
#include "AxManager.h"


MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
{
    ui.setupUi(this);

    AxManager *axm = new AxManager(this);

    //这一套可以把所有东西放到内存
    axm->openExcelFile(QString("c:/Users/123/Desktop/excTest/sum.xls"));
    axm->setSheetIndex(1);
    axm->loadData();
    axm->closeExcelFile();

    axm->openExcelFile(QString("c:/Users/123/Desktop/excTest/wanted.xlsx"));
    for (int i = 1; i <= 4; i++) {
        axm->setSheetIndex(i);
        axm->writeData(i);
    }
    axm->closeExcelFile();


}

