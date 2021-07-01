
// IEViewDlg.h : 头文件
//

#pragma once


// CIEViewDlg 对话框
class CIEViewDlg : public CDHtmlDialog
{
// 构造
public:
	CIEViewDlg(CWnd* pParent = NULL);	// 标准构造函数

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_IEVIEW_DIALOG, IDH = IDR_HTML_IEVIEW_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持

	BOOL CallJScript(const CString strFunc, const CStringArray& paramArray);
	HRESULT OnButtonOK(IHTMLElement *pElement);
	HRESULT OnButtonCancel(IHTMLElement *pElement);
	virtual BOOL CanAccessExternal();
	void OnNavigateComplete(LPDISPATCH pDisp, LPCTSTR szUrl);
	void SuppressScriptError();
	void ExecuteScript(CString &strScript, CString &strLanguage);
	// 实现
protected:
	HICON m_hIcon;


	// 生成的消息映射函数
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
	DECLARE_DHTML_EVENT_MAP()
public:
	virtual void OnDocumentComplete(LPDISPATCH pDisp, LPCTSTR szUrl);
};
