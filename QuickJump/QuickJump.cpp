// QuickJump.cpp : Defines the class behaviors for the application.
//

#include "stdafx.h"
#include "QuickJump.h"
#include "QuickJumpDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

BEGIN_MESSAGE_MAP(CQuickJumpApp, CWinApp)
END_MESSAGE_MAP()


CQuickJumpApp::CQuickJumpApp()
{
}


CQuickJumpApp theApp;



BOOL CQuickJumpApp::InitInstance()
{
	AfxEnableControlContainer();

    CoInitialize(NULL);

	CQuickJumpDlg dlg;
	m_pMainWnd = &dlg;
	int nResponse = dlg.DoModal();

	if (nResponse == IDOK)
	{
	}
	else if (nResponse == IDCANCEL)
	{
	}
    CoUninitialize();

	// Since the dialog has been closed, return FALSE so that we exit the
	//  application, rather than start the application's message pump.
	return FALSE;
}
