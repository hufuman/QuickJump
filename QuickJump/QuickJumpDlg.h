#pragma once

#define QUICK_JUMP_NOTIFY_MESSAGE       (WM_USER + 1001)

class CQuickJumpDlg : public CDialog
{
// Construction
public:
	CQuickJumpDlg(CWnd* pParent = NULL);

	enum { IDD = IDD_QuickJump_DIALOG };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);

// Implementation
protected:
	HICON       m_hIcon;
    HMODULE     m_hLoadedShlwapi;

protected:
    BOOL JumpToFilePath(LPCTSTR szFilePath);
    BOOL JumpToRegistry(LPCTSTR szRegPath, LPCTSTR szRegValue);
    BOOL SyncButtonState();
    BOOL SwitchWindowState(BOOL bShow);
    BOOL ModifyNotifyIcon(BOOL bShowIcon);

	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	afx_msg void OnOK();

	afx_msg void OnSetfocusEditPath();
	afx_msg void OnNcDestroy();
	afx_msg void OnRadioFile();
	afx_msg void OnRadioReg();
	afx_msg void OnCheckExecute();

    afx_msg LRESULT OnHotKey(WPARAM wParam, LPARAM lParam);
    afx_msg LRESULT OnQJNotifyMessage(WPARAM wParam, LPARAM lParam);
	DECLARE_MESSAGE_MAP()
};
