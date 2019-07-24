#include "MainWindow.h"
#include "AxManager.h"


MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
{
    ui.setupUi(this);

    AxManager *axm = new AxManager(this);

    axm->openExcelFile(QString("c:/Users/123/Desktop/excTest/sum.xls"));
    axm->loadData();
    axm->closeExcelFile();
}

