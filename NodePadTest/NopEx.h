#pragma once

#include <afxdialogex.h>
#include "resource.h"

// NopEx 对话框

class NopEx : public CDialogEx
{
	DECLARE_DYNAMIC(NopEx)

public:
	NopEx(CWnd* pParent = nullptr);   // 标准构造函数
	virtual ~NopEx();

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_DIALOG1 };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持
		
	DECLARE_MESSAGE_MAP()
public:
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	virtual BOOL OnInitDialog();
	void InsertPicture(HBITMAP hBmp);
	void GetClipBmp();
	bool isTopmost;
	CRichEditCtrl m_Content;
	afx_msg void OnSize(UINT nType, int cx, int cy);
};
