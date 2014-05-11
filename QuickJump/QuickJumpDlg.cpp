// QuickJumpDlg.cpp : implementation file
//

#include "stdafx.h"
#include "QuickJump.h"
#include "QuickJumpDlg.h"
#include "ShellApiEx.h"

#include <Shlwapi.h>

#include "StatLink.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CAboutDlg dialog used for App About

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// Dialog Data
	enum { IDD = IDD_ABOUTBOX };

protected:
    virtual BOOL OnInitDialog();
	virtual void DoDataExchange(CDataExchange* pDX);

protected:
    CStaticLink m_MailUrl;

// Implementation
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD), m_MailUrl(_T("https://github.com/hufuman"))
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}

BOOL CAboutDlg::OnInitDialog()
{
    CDialog::OnInitDialog();

    m_MailUrl.SubclassDlgItem(IDC_LINK_MAIL, this);
    return TRUE;
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CQuickJumpDlg dialog

CQuickJumpDlg::CQuickJumpDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CQuickJumpDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
    m_hLoadedShlwapi = NULL;
}

void CQuickJumpDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CQuickJumpDlg, CDialog)


	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
    ON_WM_QUERYDRAGICON()
    ON_WM_NCDESTROY()
    ON_WM_TIMER()

	ON_EN_SETFOCUS(IDC_EDIT_PATH, OnSetfocusEditPath)
	ON_BN_CLICKED(IDC_RADIO_FILE, OnRadioFile)
	ON_BN_CLICKED(IDC_RADIO_REG, OnRadioReg)
	ON_BN_CLICKED(IDC_CHECK_EXECUTE, OnCheckExecute)


    ON_MESSAGE(WM_HOTKEY, OnHotKey)
    ON_MESSAGE(QUICK_JUMP_NOTIFY_MESSAGE, OnQJNotifyMessage)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CQuickJumpDlg message handlers

BOOL CQuickJumpDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// Add "About..." menu item to system menu.

	// IDM_ABOUTBOX must be in the system command range.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		CString strAboutMenu;
		strAboutMenu.LoadString(IDS_ABOUTBOX);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon
	
	// TODO: Add extra initialization here
    SetWindowPos(&wndTopMost, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);

    CheckDlgButton(IDC_RADIO_REG, BST_CHECKED);

    HMODULE hModule = GetModuleHandle(_T("shlwapi.dll"));
    if(hModule == NULL)
    {
        m_hLoadedShlwapi = LoadLibrary(_T("shlwapi.dll"));
        hModule = m_hLoadedShlwapi;
    }
    if(hModule != NULL)
    {
        HWND hEditWnd = ::GetDlgItem(m_hWnd, IDC_EDIT_PATH);
        HRESULT hResult = SHAutoComplete(hEditWnd, SHACF_DEFAULT);
    }

    SyncButtonState();

    RegisterHotKey(m_hWnd, 100, MOD_CONTROL, _T('1'));

    CenterWindow();
    SwitchWindowState(TRUE);

	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CQuickJumpDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialog::OnSysCommand(nID, lParam);
	}
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CQuickJumpDlg::OnPaint() 
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, (WPARAM) dc.GetSafeHdc(), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

// The system calls this to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CQuickJumpDlg::OnQueryDragIcon()
{
	return (HCURSOR) m_hIcon;
}

void CQuickJumpDlg::OnOK() 
{
    CString strPath;

    GetDlgItemText(IDC_EDIT_PATH, strPath);
    BOOL bFile = IsDlgButtonChecked(IDC_RADIO_FILE);
    if(strPath.IsEmpty())
    {
        return;
    }
    if(bFile)
	{
		JumpToFilePath(strPath);
    }
    else
	{
		CString strValue;
		GetDlgItemText(IDC_EDIT_VALUE, strValue);
		JumpToRegistry(strPath, strValue);
    }
}

BOOL CQuickJumpDlg::JumpToFilePath(LPCTSTR szFilePath)
{
    BOOL bExecuteChecked = IsDlgButtonChecked(IDC_CHECK_EXECUTE);
    if(bExecuteChecked)
    {
        ShellExecute(m_hWnd, _T("open"), szFilePath, NULL, NULL, SW_SHOW);
    }
    else
    {
	    if(!exSHOpenFolderAndSelectItem(szFilePath))
        {
            CString strMsg;
            strMsg.Format(_T("Failed to Open File: %s"), szFilePath);
            MessageBox(strMsg, NULL, MB_OK | MB_ICONWARNING);
        }
    }
    return TRUE;
}

BOOL CQuickJumpDlg::JumpToRegistry(LPCTSTR szRegPath, LPCTSTR szRegValue)
{
	if(!exSHOpenRegKey(szRegPath, szRegValue))
    {
        CString strMsg;
        strMsg.Format(_T("Failed to Open Registry: %s"), szRegPath);
        MessageBox(strMsg, NULL, MB_ICONWARNING | MB_OK);
    }
    return TRUE;
}

void CQuickJumpDlg::OnSetfocusEditPath() 
{
    ::PostMessage(::GetDlgItem(m_hWnd, IDC_EDIT_PATH), EM_SETSEL, 0, -1);
}

void CQuickJumpDlg::OnNcDestroy() 
{
	CDialog::OnNcDestroy();
	if(m_hLoadedShlwapi != NULL)
    {
        FreeLibrary(m_hLoadedShlwapi);
        m_hLoadedShlwapi = NULL;
    }
    ModifyNotifyIcon(FALSE);
}

void CQuickJumpDlg::OnRadioFile() 
{
    SyncButtonState();
}

void CQuickJumpDlg::OnRadioReg() 
{
    SyncButtonState();
}

void CQuickJumpDlg::OnCheckExecute() 
{
    SyncButtonState();
}

BOOL CQuickJumpDlg::SyncButtonState()
{
    BOOL bFileChecked = IsDlgButtonChecked(IDC_RADIO_FILE);
    BOOL bExecuteChecked = IsDlgButtonChecked(IDC_CHECK_EXECUTE);
    GetDlgItem(IDC_CHECK_EXECUTE)->ShowWindow(bFileChecked ? SW_SHOW : SW_HIDE);
    SetDlgItemText(IDOK, (bExecuteChecked && bFileChecked) ? _T("&Open") : _T("&Jump"));

    HWND hEditWnd = ::GetDlgItem(m_hWnd, IDC_EDIT_PATH);
    if(bFileChecked)
    {
        HRESULT hResult = SHAutoComplete(hEditWnd, SHACF_DEFAULT);
    }
    else
    {
        HRESULT hResult = SHAutoComplete(hEditWnd, SHACF_AUTOSUGGEST_FORCE_OFF);
    }
    return TRUE;
}

LRESULT CQuickJumpDlg::OnHotKey(WPARAM wParam, LPARAM lParam)
{
    if(LOWORD(lParam) == MOD_CONTROL && HIWORD(lParam) == _T('1'))
    {
        SwitchWindowState(!IsWindowVisible());
    }
    return 0;
}

BOOL CQuickJumpDlg::ModifyNotifyIcon(BOOL bShowIcon)
{
    NOTIFYICONDATA Data = { 0 };
    Data.cbSize = sizeof(Data);
    Data.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
    Data.hIcon = m_hIcon;
    Data.hWnd = m_hWnd;
    Data.uCallbackMessage = QUICK_JUMP_NOTIFY_MESSAGE;
    lstrcat(Data.szTip, _T("Show QuickJump"));
    Data.uID = 0;
    Shell_NotifyIcon(bShowIcon ? NIM_ADD : NIM_DELETE, &Data);
    return TRUE;
}

BOOL CQuickJumpDlg::SwitchWindowState(BOOL bShow)
{
    if(bShow)
    {
        ModifyNotifyIcon(FALSE);
        AnimateWindow(200, AW_ACTIVATE | AW_CENTER);
    }
    else
    {
        ModifyNotifyIcon(TRUE);
        AnimateWindow(200, AW_HIDE | AW_CENTER);
    }
    return TRUE;
}

LRESULT CQuickJumpDlg::OnQJNotifyMessage(WPARAM wParam, LPARAM lParam)
{
    if(lParam == WM_LBUTTONDBLCLK)
    {
        SwitchWindowState(!IsWindowVisible());
    }
    return 0;
}
