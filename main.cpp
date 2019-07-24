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
    splash.setText("������Դ");
    //��һ�׿��԰����ж����ŵ��ڴ�
    axm->openExcelFile(QString("c:/Users/houha/Desktop/test/sum.xls"));
    splash.setProgress(10);
    splash.setText("���ļ�");
    axm->setSheetIndex(1);
    axm->loadData();
    splash.setProgress(30);
    splash.setText("ȫ�����ݶ�ȡ���");

    axm->closeExcelFile();
    splash.setProgress(35);
    splash.setText("�ر�Դ�ļ�");

    axm->openExcelFile(QString("c:/Users/houha/Desktop/test/wanted.xlsx"));
    splash.setProgress(40);
    splash.setText("��Ŀ���ļ�");
    for (int i = 1; i <= 4; ++i) {
        axm->setSheetIndex(i);
        axm->writeData(i);
        int process = 40 + 10 * i;
        splash.setProgress(process);
        splash.setText(QString("д���%1���ļ�").arg(i));
        if (i == 4) {
            splash.setProgress(99);
            splash.setText(QString("ִ����ϣ�����ȷ��"));
        }
    }
   
    axm->closeExcelFile();

    return 0;
}
