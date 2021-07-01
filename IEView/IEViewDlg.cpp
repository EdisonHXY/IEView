
// IEViewDlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "IEView.h"
#include "IEViewDlg.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// ����Ӧ�ó��򡰹��ڡ��˵���� CAboutDlg �Ի���

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// �Ի�������
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

// ʵ��
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(IDD_ABOUTBOX)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CIEViewDlg �Ի���

BEGIN_DHTML_EVENT_MAP(CIEViewDlg)
	DHTML_EVENT_ONCLICK(_T("ButtonOK"), OnButtonOK)
	DHTML_EVENT_ONCLICK(_T("ButtonCancel"), OnButtonCancel)
END_DHTML_EVENT_MAP()


CIEViewDlg::CIEViewDlg(CWnd* pParent /*=NULL*/)
	: CDHtmlDialog(IDD_IEVIEW_DIALOG, IDR_HTML_IEVIEW_DIALOG, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CIEViewDlg::DoDataExchange(CDataExchange* pDX)
{
	CDHtmlDialog::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CIEViewDlg, CDHtmlDialog)
	ON_WM_SYSCOMMAND()
END_MESSAGE_MAP()


BOOL CIEViewDlg::CallJScript(const CString strFunc, const CStringArray& paramArray)
{
	CComPtr<IDispatch> script;
	m_spHtmlDoc->get_Script(&script); //��ȡ�ű�ָ��
	DISPID id;

	script->GetIDsOfNames(IID_NULL, &CComBSTR(strFunc), 1, LOCALE_USER_DEFAULT, &id); // ���ݽű��нӿ�����ȡ��Ӧ��ID
	if (id < 0)
	{
		return FALSE;
	}

	COleVariant ret;
	COleVariant *args = new COleVariant[paramArray.GetSize()];

	for (int i = 0 ;i < paramArray.GetSize();++i)
	{
		args[i] = paramArray[i];
	}
	
	DISPPARAMS params = { 0 };
	params.cArgs = paramArray.GetSize();
	params.rgvarg = args;
	script->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, &ret, NULL, NULL);

	return TRUE;
}



// CIEViewDlg ��Ϣ�������

BOOL CIEViewDlg::OnInitDialog()
{
	SetHostFlags(DOCHOSTUIFLAG_FLAT_SCROLLBAR | DOCHOSTUIFLAG_NO3DBORDER);
	CDHtmlDialog::OnInitDialog();

	// ��������...���˵�����ӵ�ϵͳ�˵��С�

	// IDM_ABOUTBOX ������ϵͳ���Χ�ڡ�
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// ���ô˶Ի����ͼ�ꡣ  ��Ӧ�ó��������ڲ��ǶԻ���ʱ����ܽ��Զ�
	//  ִ�д˲���
	SetIcon(m_hIcon, TRUE);			// ���ô�ͼ��
	SetIcon(m_hIcon, FALSE);		// ����Сͼ��

	// TODO: �ڴ���Ӷ���ĳ�ʼ������
	int cx = GetSystemMetrics(SM_CXSCREEN);
	int cy = GetSystemMetrics(SM_CYSCREEN);
	MoveWindow(0, 0, cx, cy);
	
	//���δ���
	m_pBrowserApp->put_Silent(VARIANT_TRUE);


	//��ȡ·��
	TCHAR szFilePath[MAX_PATH + 1];
	GetModuleFileName(NULL, szFilePath, MAX_PATH);

	CString tmpStr = _tcsrchr(szFilePath, '\\');
	(_tcsrchr(szFilePath, '\\'))[0] = 0;
	CString dir = szFilePath;
	int iIndex = tmpStr.ReverseFind('.');
	CString fileName = tmpStr.Mid(1, iIndex - 1);
	
	//��ȡ����
	CString cfgPath = dir + "\\url.ini";
	CFileFind finder;   //�����Ƿ����ini�ļ����������ڣ�������һ���µ�Ĭ�����õ�ini�ļ�
	BOOL ifFind = finder.FindFile(cfgPath);
	CString urlPath = "www.baidu.com";
	if (!ifFind)
	{
		::WritePrivateProfileString("SYS", "url", urlPath, cfgPath);
	}
	else
	{
		::GetPrivateProfileString("SYS","url",urlPath,urlPath.GetBuffer(MAX_PATH),MAX_PATH,cfgPath);
	}

	Navigate(urlPath);

	SetExternalDispatch(GetIDispatch(TRUE));
	
	return TRUE;  // ���ǽ��������õ��ؼ������򷵻� TRUE
}

void CIEViewDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDHtmlDialog::OnSysCommand(nID, lParam);
	}
}

// �����Ի��������С����ť������Ҫ����Ĵ���
//  �����Ƹ�ͼ�ꡣ  ����ʹ���ĵ�/��ͼģ�͵� MFC Ӧ�ó���
//  �⽫�ɿ���Զ���ɡ�

void CIEViewDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // ���ڻ��Ƶ��豸������

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// ʹͼ���ڹ����������о���
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// ����ͼ��
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDHtmlDialog::OnPaint();
	}
}

//���û��϶���С������ʱϵͳ���ô˺���ȡ�ù��
//��ʾ��
HCURSOR CIEViewDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

HRESULT CIEViewDlg::OnButtonOK(IHTMLElement* /*pElement*/)
{
	OnOK();
	return S_OK;
}

HRESULT CIEViewDlg::OnButtonCancel(IHTMLElement* /*pElement*/)
{
	OnCancel();
	return S_OK;
}


BOOL CIEViewDlg::CanAccessExternal()
{
	return TRUE;
}

void CIEViewDlg::OnNavigateComplete(LPDISPATCH pDisp, LPCTSTR szUrl)
{
	CDHtmlDialog::OnNavigateComplete(pDisp, szUrl);

	SuppressScriptError(); // ���νű�����
}

void CIEViewDlg::SuppressScriptError()
{
	// Ҫִ�е����α���ű�
	CString strScript = _T("window.onerror=function myonerror(){return true}");
	CString strLanguage("JavaScript");
	ExecuteScript(strScript, strLanguage);
}

void CIEViewDlg::ExecuteScript(CString &strScript, CString &strLanguage)
{
	IHTMLDocument2* pIHtmlDoc = NULL;
	GetDHtmlDocument(&pIHtmlDoc);
	if (!pIHtmlDoc) return;

	IHTMLWindow2* pIhtmlwindow2 = NULL;
	pIHtmlDoc->get_parentWindow(&pIhtmlwindow2);
	if (!pIhtmlwindow2) return;

	BSTR bstrScript = strScript.AllocSysString();
	BSTR bstrLanguage = strLanguage.AllocSysString();
	VARIANT pRet;

	// ע��ű�����ǰҳ��
	pIhtmlwindow2->execScript(bstrScript, bstrLanguage, &pRet);
	::SysFreeString(bstrScript);
	::SysFreeString(bstrLanguage);
	pIhtmlwindow2->Release();
}


void CIEViewDlg::OnDocumentComplete(LPDISPATCH pDisp, LPCTSTR szUrl)
{
	CDHtmlDialog::OnDocumentComplete(pDisp, szUrl);

	// TODO: �ڴ����ר�ô����/����û���
	
}
