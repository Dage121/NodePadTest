// NopEx.cpp: 实现文件
//

#include "NopEx.h"
#include <RichOle.h>
#include <afxole.h>
#include <chrono>

// NopEx 对话框

IMPLEMENT_DYNAMIC(NopEx, CDialogEx)

NopEx::NopEx(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_DIALOG1, pParent)
{

}

NopEx::~NopEx()
{
}

void NopEx::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_CONTENT, m_Content);
}


BEGIN_MESSAGE_MAP(NopEx, CDialogEx)
	ON_WM_SIZE()
END_MESSAGE_MAP()


// NopEx 消息处理程序


BOOL NopEx::PreTranslateMessage(MSG* pMsg)
{
	// TODO: 在此添加专用代码和/或调用基类
	if (pMsg->message == WM_KEYDOWN)
	{
		if (GetKeyState(VK_CONTROL) && pMsg->wParam == 0x56) { GetClipBmp(); }
		if (GetKeyState(VK_CONTROL) && pMsg->wParam == 0x54) {
			::SetWindowPos(this->m_hWnd, isTopmost ? HWND_TOPMOST : HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
			isTopmost = !isTopmost; // 不置顶
		}
		//	MessageBox(0, 0, 0);
		switch (pMsg->wParam)
		{
			case VK_RETURN:
				//m_Content.SetSel(-1, -1);
				m_Content.ReplaceSel(L"\n");
				return TRUE;
			case VK_TAB:
				m_Content.ReplaceSel(L"\t");
			case VK_ESCAPE:
				return TRUE;
		}
	}
	return CDialogEx::PreTranslateMessage(pMsg);
}


BOOL NopEx::OnInitDialog()
{
	CDialogEx::OnInitDialog();
	HICON mhIcon = AfxGetApp()->LoadIconW(IDI_ICON1);
	SetIcon(mhIcon, TRUE);

	SetClipboardViewer();
	CRect rect;
	GetClientRect(&rect);
	m_Content.MoveWindow(&rect);
	// 设置 RichEdit2 控件的默认字体
	CHARFORMAT cf;
	cf.cbSize = sizeof(CHARFORMAT);
	cf.dwMask = CFM_FACE | CFM_SIZE | CFM_BOLD | CFM_COLOR;
	cf.dwEffects = 0; // 清除所有效果
	cf.yHeight = 400; // 字体高度（单位是 1/20 点）
	cf.crTextColor = RGB(0, 255, 255); // 字体颜色
	wcscpy_s(cf.szFaceName, _T("Arial")); // 字体名称
	m_Content.SetDefaultCharFormat(cf);
	m_Content.SetBackgroundColor(FALSE, RGB(39, 39, 39));
	// 设置为 0 启用自动换行
	m_Content.SendMessage(EM_SETTARGETDEVICE, 0, 0);
	// TODO:  在此添加额外的初始化
	SetBackgroundColor(RGB(0, 0, 0));
	::SetWindowPos(this->m_hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
	isTopmost = true; // 不置顶

	return TRUE;  // return TRUE unless you set the focus to a control
				  // 异常: OCX 属性页应返回 FALSE
}

void NopEx::GetClipBmp()
{
	if (OpenClipboard())
	{
		if (IsClipboardFormatAvailable(CF_BITMAP))
		{
			HBITMAP hBitmap = (HBITMAP)GetClipboardData(CF_BITMAP);
			InsertPicture(hBitmap);
		}
		CloseClipboard();
	}
}

void NopEx::OnSize(UINT nType, int cx, int cy)
{
	CDialogEx::OnSize(nType, cx, cy);
	//m_Content.MoveWindow(10, 40, cx - 20, cy - 50);
	m_Content.MoveWindow(0, 0, cx, cy);
	// TODO: 在此处添加消息处理程序代码
}

void NopEx::InsertPicture(HBITMAP hBmp)
{
	STGMEDIUM stgm;
	stgm.tymed = TYMED_GDI;
	stgm.hBitmap = hBmp;
	stgm.pUnkForRelease = NULL;

	FORMATETC fm;
	fm.cfFormat = CF_BITMAP;
	fm.ptd = NULL;
	fm.dwAspect = DVASPECT_CONTENT;
	fm.lindex = -1;
	fm.tymed = TYMED_GDI;

	COleDataSource oleDataSource;
	oleDataSource.CacheData(CF_BITMAP, &stgm);
	LPDATAOBJECT dataObject = (LPDATAOBJECT)oleDataSource.GetInterface(&IID_IDataObject);

	if (OleQueryCreateFromData(dataObject) != OLE_S_STATIC)
		return;

	LPOLECLIENTSITE oleClientSite;
	if (S_OK != m_Content.GetIRichEditOle()->GetClientSite(&oleClientSite))
		return;

	// 内存分配
	LPLOCKBYTES lockBytes = NULL;
	if (S_OK == CreateILockBytesOnHGlobal(NULL, TRUE, &lockBytes) && lockBytes)
	{
		IStorage* storage = NULL;
		if (S_OK == StgCreateDocfileOnILockBytes(lockBytes, STGM_SHARE_EXCLUSIVE | STGM_CREATE | STGM_READWRITE, 0, &storage) && storage)
		{
			IOleObject* oleObject = NULL;
			if (S_OK == OleCreateStaticFromData(dataObject, IID_IOleObject, OLERENDER_FORMAT, &fm, oleClientSite, storage, (void**)&oleObject) && oleObject)
			{
				CLSID clsid;
				if (S_OK == oleObject->GetUserClassID(&clsid))
				{
					REOBJECT reobject = { sizeof(REOBJECT) };
					reobject.clsid = clsid;
					reobject.cp = REO_CP_SELECTION;
					reobject.dvaspect = DVASPECT_CONTENT;
					reobject.poleobj = oleObject;
					reobject.polesite = oleClientSite;
					reobject.pstg = storage;
					// 插入OLE对象
					m_Content.GetIRichEditOle()->InsertObject(&reobject);
				}
				oleObject->Release();
			}
			storage->Release();
		}
		lockBytes->Release();
	}
	oleClientSite->Release();
}
