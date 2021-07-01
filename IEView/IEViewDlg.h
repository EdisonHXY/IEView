
// IEViewDlg.h : ͷ�ļ�
//

#pragma once


// CIEViewDlg �Ի���
class CIEViewDlg : public CDHtmlDialog
{
// ����
public:
	CIEViewDlg(CWnd* pParent = NULL);	// ��׼���캯��

// �Ի�������
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_IEVIEW_DIALOG, IDH = IDR_HTML_IEVIEW_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV ֧��

	BOOL CallJScript(const CString strFunc, const CStringArray& paramArray);
	HRESULT OnButtonOK(IHTMLElement *pElement);
	HRESULT OnButtonCancel(IHTMLElement *pElement);
	virtual BOOL CanAccessExternal();
	void OnNavigateComplete(LPDISPATCH pDisp, LPCTSTR szUrl);
	void SuppressScriptError();
	void ExecuteScript(CString &strScript, CString &strLanguage);
	// ʵ��
protected:
	HICON m_hIcon;


	// ���ɵ���Ϣӳ�亯��
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
	DECLARE_DHTML_EVENT_MAP()
public:
	virtual void OnDocumentComplete(LPDISPATCH pDisp, LPCTSTR szUrl);
};
