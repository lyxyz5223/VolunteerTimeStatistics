#include "ClassTable2Pic.h"
#include <QtWidgets/QApplication>
#include <atlbase.h>
#include <atlcom.h>

int main(int argc, char *argv[])
{
    //HRESULT hr = CoInitializeEx(0, COINIT_MULTITHREADED);
    QApplication a(argc, argv);
    ClassTable2Pic w;
    w.show();
    return a.exec();
}
