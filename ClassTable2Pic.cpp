#include "ClassTable2Pic.h"
#include "win32Header.h"
//mine
#include "stringProcess.h"
#include <thread>
#include <qfiledialog.h>
#include <qmessagebox.h>
#include <sstream>
#include<qinputdialog.h>
vector<Excel::_ApplicationPtr> appList;
struct dt {
	string name;
	string id;
	int time;//志愿时
};
struct index {
	int row;
	int column;
};
struct dataStruct {
	index name;
	index id;
	index time;
};
struct headerStruct {
	string h1;
	string h2;
	string h3;
};
using namespace std;
bool isDigits(string str)
{
	for (char ch : str)
		if (ch > '9' || ch < '0')
			return false;
	return true;
}
void ClassTable2Pic::solve()
{
	setAttribute(Qt::WA_AlwaysStackOnTop, true);
	//窗口置顶
	::SetWindowPos((HWND)this->winId(), HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
	using namespace Excel;
	using namespace _com_util;
	//数据输入
	string DocumentPathAndName = "C:\\Users\\lyxyz5223\\Downloads\\工作簿1 - 副本.xlsx";
	QFileDialog* fd = new QFileDialog(this);
	DocumentPathAndName = fd->getOpenFileName().toStdString();
	if (DocumentPathAndName == "")
		return;
	//创建/打开Excel应用实例
	Excel::_ApplicationPtr app;
	app.CreateInstance(__uuidof(Application));
	appList.push_back(app);
	app->Workbooks->Open(ConvertStringToBSTR(UTF8ToANSI(DocumentPathAndName).c_str()), 1, 0);//打开xls
	app->PutVisible(0, xlVisible);
	SetActiveWindow((HWND)app->ActiveWindow->GetHwnd());
	if (ui.selectHeader->isChecked())
	{
		string h1, h2, h3;//三个表头
		QMessageBox::information(this, "请选择", "请在Excel表格中选择表头，\n如选中“姓名 学号 志愿时长”这三项，\n只能选择三项，计算时：第一项用于比较合并，第三项用于求和，第二项取内容最长的一个单元格\n选择完成请点确定，然后稍等\n切记：稍等，勿因为程序阻塞就以为程序没在做事情");
		RangePtr selectionRange = app->ActiveWindow->GetRangeSelection();
		h1 = ConvertBSTRToString(_bstr_t(selectionRange->GetItem(1, 1).GetVARIANT()));
		h2 = ConvertBSTRToString(_bstr_t(selectionRange->GetItem(1, 2).GetVARIANT()));
		h3 = ConvertBSTRToString(_bstr_t(selectionRange->GetItem(1, 3).GetVARIANT()));
		Excel::_WorksheetPtr ws;//WorkSheet
		Excel::_WorkbookPtr wb;//WorkBook
		wb = app->ActiveWorkbook;
		ws = app->ActiveWorkbook->ActiveSheet;
		Excel::RangePtr pRange = ws->GetUsedRange();
		long ColumnCount = pRange->Columns->GetCount();
		long RowCount = pRange->Rows->GetCount();
		cout << "ColumnCount:" << ColumnCount << endl;
		cout << "RowCount:" << RowCount << endl;
		vector<dt> data;
		vector<dataStruct> ind;
		for (long i = 0; i < RowCount; i++)
		{
			for (long j = 0; j < ColumnCount; j++)
			{
				string str = ConvertBSTRToString(_bstr_t(pRange->GetItem(i + 1, j + 1).GetVARIANT()));
				if (str == h1)
					ind.push_back({ { i + 1,j + 1 }, {0,0}, {0,0} });
				else if (str == h2)
					ind[ind.size() - 1].id = { i + 1,j + 1 };
				else if (str == h3)
					ind[ind.size() - 1].time = { i + 1,j + 1 };
				cout << str << "\t";
			}
			cout << endl;
		}
		for (auto i = ind.begin(); i != ind.end(); i++)
		{
			for (long row = (*i).name.row + 1; row < RowCount; row++)
			{
				if (i->name.column == 0 || i->name.row == 0
					|| i->id.row == 0 || i->id.column == 0
					|| i->time.row == 0 || i->time.column == 0)
					break;
				string name = ConvertBSTRToString(_bstr_t(pRange->GetItem(row, i->name.column).GetVARIANT()));
				if (name == "")//如果名字为空
					break;//结束该区块的搜索，切换到下一数据区
				string id = ConvertBSTRToString(_bstr_t(pRange->GetItem(row, i->id.column).GetVARIANT()));
				string time = ConvertBSTRToString(_bstr_t(pRange->GetItem(row, i->time.column).GetVARIANT()));
				bool exist = false;
				auto t = data.begin();
				for (; t != data.end(); t++)
				{
					if (t->name == name)
					{
						exist = true;
						break;
					}
				}
				if (!exist)
					data.push_back({ name,id, stoi(time) });
				else
				{
					if (t->id.size() < id.size())
						t->id = id;
					try {
						t->time += stoi(time);
					}
					catch (const std::invalid_argument& e)
					{
						cerr << e.what() << endl;
						QMessageBox::critical(this, "err,数据存在错误！", ANSIToUTF8(e.what()).c_str());
						assert(e.what());
					}
				}
			}
		}
		QMessageBox::information(this, "确认", "确认后继续");
		_WorksheetPtr newWorkSheet = wb->Worksheets->Add(_variant_t(wb->Worksheets->GetItem(1), true), vtMissing, _variant_t(1), xlWorksheet);
#ifdef _DEBUG
		system("cls");
#endif //_DEBUG
		cout << h1 << "\t" << h2 << "\t" << h3 << endl;
		QString textDelimiter = ui.delimiter->text();
		QString resultText = QString::fromStdString(ANSIToUTF8(h1)) + textDelimiter;
		resultText += QString::fromStdString(ANSIToUTF8(h2)) + textDelimiter;
		resultText += QString::fromStdString(ANSIToUTF8(h3)) + "\n";
		newWorkSheet->Cells->Item[1, 1] = ConvertStringToBSTR(h1.c_str());//name
		newWorkSheet->Cells->Item[1, 2] = ConvertStringToBSTR(h2.c_str());//id
		newWorkSheet->Cells->Item[1, 3] = ConvertStringToBSTR(h3.c_str());//time
		for (size_t i = 0; i < data.size(); i++)
		{
			newWorkSheet->Cells->Item[i + 2, 1] = ConvertStringToBSTR(data[i].name.c_str());//name
			newWorkSheet->Cells->Item[i + 2, 2] = ConvertStringToBSTR(data[i].id.c_str());//id
			newWorkSheet->Cells->Item[i + 2, 3] = ConvertStringToBSTR(to_string(data[i].time).c_str());//time
			resultText += ANSIToUTF8(data[i].name);
			resultText += textDelimiter;
			resultText += ANSIToUTF8(data[i].id);
			resultText += textDelimiter;
			resultText += ANSIToUTF8(to_string(data[i].time) + "\n");
			cout << data[i].name << "\t" << data[i].id << "\t" << data[i].time << endl;
		}
		ui.log->setText(resultText);
	}
	else if (ui.selectRange->isChecked())
	{
		//QMessageBox::information(this, "注意", "请在弹出的Excel窗口中选择需要处理的Sheet\n完成后点击确定");
		QMessageBox::information(this, "注意", "请在弹出的Excel窗口中选择需要处理的数据范围\n(可含表头，但若表头第三项为数字请不要包含表头，否则表头将按数据处理)，完成后点击确定\n注：按住Ctrl键可以多选", QMessageBox::NoButton);
		_WorkbookPtr wb = app->ActiveWorkbook;
		_WorksheetPtr ws = wb->ActiveSheet;
		Excel::RangePtr UsedRange = ws->GetUsedRange();
		long ColumnCount = UsedRange->Columns->GetCount();
		long RowCount = UsedRange->Rows->GetCount();
		RangePtr sr = app->ActiveWindow->GetRangeSelection();
		vector<headerStruct> header;
		vector<dt> data;
		long areaCount = sr->GetAreas()->GetCount();
		for (int currentArea = 1; currentArea <= areaCount; currentArea++)
		{
			AreasPtr sra = sr->GetAreas();
			RangePtr procRange = sra->GetItem(currentArea);
			bool haveHeader = false;
			string headerTextJudge = ConvertBSTRToString(_bstr_t(procRange->GetItem(1, 3).GetVARIANT()));
			if (!isDigits(headerTextJudge))
			{
				haveHeader = true;
				header.push_back({
					ConvertBSTRToString(_bstr_t(procRange->GetItem(1, 1).GetVARIANT())),
					ConvertBSTRToString(_bstr_t(procRange->GetItem(1, 2).GetVARIANT())),
					ConvertBSTRToString(_bstr_t(procRange->GetItem(1, 3).GetVARIANT()))
					});
			}
			for (int i = (haveHeader ? 2 : 1); i <= RowCount; i++)
			{
				string name = ConvertBSTRToString(_bstr_t(procRange->GetItem(i, 1).GetVARIANT()));
				string id = ConvertBSTRToString(_bstr_t(procRange->GetItem(i, 2).GetVARIANT()));
				string time = ConvertBSTRToString(_bstr_t(procRange->GetItem(i, 3).GetVARIANT()));
				if ((name == "" && id == "") || !isDigits(time))
					break;
				bool exist = false; 
				auto iter = data.begin();
				for (iter; iter != data.end(); iter++)
				{
					if (name == iter->name)
					{
						exist = true;
						break;
					}
				}
				if (!exist)
					data.push_back({ name,id,stoi(time) });
				else
					iter->time += stoi(time);
			}
		}
		QMessageBox::information(this, "确认", "确认后继续");
		_WorksheetPtr newWorkSheet = wb->Worksheets->Add(_variant_t(wb->Worksheets->GetItem(1), true), vtMissing, _variant_t(1), xlWorksheet);
		QString textDelimiter = ui.delimiter->text();
		QString resultText = QString::fromStdString(ANSIToUTF8(header[0].h1)) + textDelimiter;
		resultText += QString::fromStdString(ANSIToUTF8(header[0].h2)) + textDelimiter;
		resultText += QString::fromStdString(ANSIToUTF8(header[0].h3)) + "\n";
		newWorkSheet->Cells->Item[1, 1] = ConvertStringToBSTR(header[0].h1.c_str());//name
		newWorkSheet->Cells->Item[1, 2] = ConvertStringToBSTR(header[0].h2.c_str());//id
		newWorkSheet->Cells->Item[1, 3] = ConvertStringToBSTR(header[0].h3.c_str());//time
		for (size_t i = 0; i < data.size(); i++)
		{
			newWorkSheet->Cells->Item[i + 2, 1] = ConvertStringToBSTR(data[i].name.c_str());//name
			newWorkSheet->Cells->Item[i + 2, 2] = ConvertStringToBSTR(data[i].id.c_str());//id
			newWorkSheet->Cells->Item[i + 2, 3] = ConvertStringToBSTR(to_string(data[i].time).c_str());//time
			resultText += ANSIToUTF8(data[i].name);
			resultText += textDelimiter;
			resultText += ANSIToUTF8(data[i].id);
			resultText += textDelimiter;
			resultText += ANSIToUTF8(to_string(data[i].time) + "\n");
			cout << data[i].name << "\t" << data[i].id << "\t" << data[i].time << endl;
		}
		ui.log->setText(resultText);
	}
	//窗口不置顶
	::SetWindowPos((HWND)this->winId(), HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);	using namespace Excel;
	setAttribute(Qt::WA_AlwaysStackOnTop, false);
}
ClassTable2Pic::ClassTable2Pic(QWidget *parent)
	: QMainWindow(parent)
{
	ui.setupUi(this);
	show();
}

ClassTable2Pic::~ClassTable2Pic()
{
	for (auto i = appList.begin(); i != appList.end(); i++)
	{
		if (*i != 0)
		{
			try {
				(*i)->Quit();
				(*i).Release();
			}
			catch (...)
			{

			}
		}
	}
}
