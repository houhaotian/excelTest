#include <QtWidgets/QApplication>
#include "AxManager.h"
#include "SplashScreen.h"

#pragma execution_character_set("utf-8")

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);

    SplashScreen splash;
    splash.show();
    AxManager *axm = new AxManager;
    splash.setProgress(1);
    splash.setText("申请资源");
    //这一套可以把所有东西放到内存
    axm->openExcelFile(QString("c:/Users/houha/Desktop/test/sum.xls"));
    splash.setProgress(10);
    splash.setText("打开文件");
    axm->setSheetIndex(1);
    axm->loadData();
    splash.setProgress(30);
    splash.setText("全部数据读取完毕");

    axm->closeExcelFile();
    splash.setProgress(35);
    splash.setText("关闭源文件");

    axm->openExcelFile(QString("c:/Users/houha/Desktop/test/wanted.xlsx"));
    splash.setProgress(40);
    splash.setText("打开目标文件");
    for (int i = 1; i <= 4; ++i) {
        axm->setSheetIndex(i);
        axm->writeData(i);
        int process = 40 + 10 * i;
        splash.setProgress(process);
        splash.setText(QString("写入第%1个文件").arg(i));
        if (i == 4) {
            splash.setProgress(99);
            splash.setText(QString("执行完毕，请点击确定"));
        }
    }
   
    axm->closeExcelFile();

    return 0;
}
