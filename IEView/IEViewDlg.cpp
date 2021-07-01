
// IEViewDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "IEView.h"
#include "IEViewDlg.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// 用于应用程序“关于”菜单项的 CAboutDlg 对话框

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

// 实现
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


// CIEViewDlg 对话框

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
	m_spHtmlDoc->get_Script(&script); //获取脚本指针
	DISPID id;

	script->GetIDsOfNames(IID_NULL, &CComBSTR(strFunc), 1, LOCALE_USER_DEFAULT, &id); // 根据脚本中接口名获取对应的ID
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



// CIEViewDlg 消息处理程序

BOOL CIEViewDlg::OnInitDialog()
{
	SetHostFlags(DOCHOSTUIFLAG_FLAT_SCROLLBAR | DOCHOSTUIFLAG_NO3DBORDER);
	CDHtmlDialog::OnInitDialog();

	// 将“关于...”菜单项添加到系统菜单中。

	// IDM_ABOUTBOX 必须在系统命令范围内。
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

	// 设置此对话框的图标。  当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	// TODO: 在此添加额外的初始化代码
	int cx = GetSystemMetrics(SM_CXSCREEN);
	int cy = GetSystemMetrics(SM_CYSCREEN);
	MoveWindow(0, 0, cx, cy);
	
	//屏蔽错误
	m_pBrowserApp->put_Silent(VARIANT_TRUE);


	//获取路径
	TCHAR szFilePath[MAX_PATH + 1];
	GetModuleFileName(NULL, szFilePath, MAX_PATH);

	CString tmpStr = _tcsrchr(szFilePath, '\\');
	(_tcsrchr(szFilePath, '\\'))[0] = 0;
	CString dir = szFilePath;
	int iIndex = tmpStr.ReverseFind('.');
	CString fileName = tmpStr.Mid(1, iIndex - 1);
	
	//读取配置
	CString cfgPath = dir + "\\url.ini";
	CFileFind finder;   //查找是否存在ini文件，若不存在，则生成一个新的默认设置的ini文件
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
	
	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
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

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。  对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CIEViewDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作区矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDHtmlDialog::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
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

	SuppressScriptError(); // 屏蔽脚本报错
}

void CIEViewDlg::SuppressScriptError()
{
	// 要执行的屏蔽报错脚本
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

	// 注入脚本到当前页面
	pIhtmlwindow2->execScript(bstrScript, bstrLanguage, &pRet);
	::SysFreeString(bstrScript);
	::SysFreeString(bstrLanguage);
	pIhtmlwindow2->Release();
}


void CIEViewDlg::OnDocumentComplete(LPDISPATCH pDisp, LPCTSTR szUrl)
{
	CDHtmlDialog::OnDocumentComplete(pDisp, szUrl);

	// TODO: 在此添加专用代码和/或调用基类
	
}
