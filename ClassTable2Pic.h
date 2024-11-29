#pragma once

#include <QtWidgets/QMainWindow>
#include "ui_ClassTable2Pic.h"

class ClassTable2Pic : public QMainWindow
{
    Q_OBJECT

public:
    ClassTable2Pic(QWidget *parent = nullptr);
    ~ClassTable2Pic();
public slots:
    void solve();
private:
    Ui::ClassTable2PicClass ui;
};
