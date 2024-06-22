#include <afxwin.h>
#include "NopEx.h"
class MyApp : public CWinApp
{
	virtual BOOL InitInstance()override
	{
// TODO: 调用 AfxInitRichEdit2() 以初始化 richedit2 库。\n"		CWinApp::InitInstance();
		AfxInitRichEdit2();
		NopEx app;
		m_pMainWnd = &app;
		app.DoModal();
		return TRUE;
	}
};

MyApp theApp;