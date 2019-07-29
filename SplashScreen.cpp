#include "SplashScreen.h"

SplashScreen::SplashScreen(QWidget *parent /*= 0*/)
{
    QPixmap map(":/MainWindow/resource/splash.jpg");
    setPixmap(map);
    m_bar = new QProgressBar(this);
    m_bar->setRange(0, 100);
    m_bar->setGeometry(0, height() - 20, width(), 20);

  /*  m_label = new QLabel(this);
    m_label->setStyleSheet(QString("font:12ps"));
    m_label->setGeometry(250, height() - 50, 200, 20);*/
}

SplashScreen::~SplashScreen()
{

}

void SplashScreen::setText(const QString &text)
{
   /* m_label->setText(text);
    m_label->repaint();*/
    showMessage(text, Qt::AlignCenter);
}

void SplashScreen::setProgress(int p)
{
    m_bar->setValue(p);
}

