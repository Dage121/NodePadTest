#include <afxwin.h>
#include "NopEx.h"
class MyApp : public CWinApp
{
	virtual BOOL InitInstance()override
	{
// TODO: ���� AfxInitRichEdit2() �Գ�ʼ�� richedit2 �⡣\n"		CWinApp::InitInstance();
		AfxInitRichEdit2();
		NopEx app;
		m_pMainWnd = &app;
		app.DoModal();
		return TRUE;
	}
};

MyApp theApp;