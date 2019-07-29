#pragma once

#include <QObject>
#include <QSplashScreen>
#include <QProgressBar>
#include <QLabel>

class SplashScreen :public QSplashScreen
{
public:
    SplashScreen(QWidget *parent = 0);
    ~SplashScreen();

    void setText(const QString &text);
    void setProgress(int p);
private:
    QProgressBar *m_bar;
    QLabel *m_label;
};