#include "stdafx.h"
#include "ShellApiEx.h"
#include <shlobj.h>
#include <algorithm>

#define MSG_BOX(str) MessageBox(NULL, str, _T("QuickJump"), MB_OK);

BOOL RemoveDoubleSlash(tstring& strPath)
{
    tstring::size_type pos = strPath.find(_T("\\\\"));
    while(pos != tstring::npos)
    {
        strPath.erase(pos, 1);
        pos = strPath.find(_T("\\\\"), pos);
    }
    int nLen = strPath.length();
    if(nLen > 0 && strPath[nLen - 1] == _T('\\'))
    {
        strPath.resize(nLen - 1);
    }
    return TRUE;
}

/**
 * 
 * Open the directory that contains the file object and select it.
 * Refer to: http://www.vckbase.net/document/viewdoc/?id=1595
 *
 * 
 * 
 */

HRESULT GetShellFolderViewDual(ITEMIDLIST* pidl, IShellFolderViewDual** ppIShellFolderViewDual)
{
	IWebBrowserApp* pIWebBrowserApp;
	IDispatch* pDoc;
	HWND hWnd;
	HRESULT hr;
	HINSTANCE ghSHDOCVW;

	HRESULT (WINAPI*gpfSHGetIDispatchForFolder)(ITEMIDLIST*, IWebBrowserApp**);

	*ppIShellFolderViewDual = NULL;

	ghSHDOCVW = LoadLibrary(_T("SHDOCVW.DLL"));
	if(ghSHDOCVW == NULL)
		return FALSE;

	pIWebBrowserApp=NULL;
	gpfSHGetIDispatchForFolder = (HRESULT (WINAPI*)(ITEMIDLIST*, IWebBrowserApp**)) GetProcAddress(ghSHDOCVW, "SHGetIDispatchForFolder");
	if(gpfSHGetIDispatchForFolder == NULL)
    {
        FreeLibrary(ghSHDOCVW);
		return FALSE;
    }

	if(SUCCEEDED(gpfSHGetIDispatchForFolder(pidl, &pIWebBrowserApp))) 
	{
		if(SUCCEEDED(pIWebBrowserApp->get_HWND((long*)&hWnd))) 
		{
			SetForegroundWindow(hWnd);
			ShowWindow(hWnd, SW_SHOWNORMAL);
		}

		if(SUCCEEDED(hr = pIWebBrowserApp->get_Document(&pDoc)))
		{
			pDoc->QueryInterface(IID_IShellFolderViewDual, (void**)ppIShellFolderViewDual);
			pDoc->Release();
		}
		pIWebBrowserApp->Release();
	}
	FreeLibrary(ghSHDOCVW);

	return (ppIShellFolderViewDual != NULL);
}

BOOL exSHGetIDLFromPath(LPCTSTR szFolderPath, LPITEMIDLIST* ppidl)
{
    IShellFolder* pShellFolder = NULL;
    HRESULT hResult = SHGetDesktopFolder(&pShellFolder);
    if(FAILED(hResult))
    {
        return FALSE;
    }
    hResult = pShellFolder->ParseDisplayName(NULL, NULL, (LPOLESTR)szFolderPath, NULL, ppidl, 0);
    pShellFolder->Release();
    return SUCCEEDED(hResult);
}

BOOL exSHOpenFolderAndSelectItem(ITEMIDLIST *pidlFolder)
{
	ITEMIDLIST *pidl, *pidl2;
	ITEMIDLIST* pIdlFile;
	USHORT cb;
    BOOL bResult = FALSE;
	IShellFolderViewDual* pIShellFolderViewDual;

	HRESULT (WINAPI *gpfSHOpenFolderAndSelectItems)(LPCITEMIDLIST*, UINT, LPCITEMIDLIST*, DWORD);

    HINSTANCE ghShell32;

	ghShell32 = LoadLibrary(_T("Shell32.DLL"));
	if(ghShell32 == NULL)
		return FALSE;

	gpfSHOpenFolderAndSelectItems = (HRESULT (WINAPI*)(LPCITEMIDLIST*, UINT, LPCITEMIDLIST*, DWORD))GetProcAddress(ghShell32, "SHOpenFolderAndSelectItems");
	if(gpfSHOpenFolderAndSelectItems != NULL)
	{
        bResult = SUCCEEDED(gpfSHOpenFolderAndSelectItems((LPCITEMIDLIST*)pidlFolder, 0, (LPCITEMIDLIST*)NULL, 0));
		FreeLibrary(ghShell32);
		return bResult;
	}
	FreeLibrary(ghShell32);

	/// If the OS doesn't support SHOpenFolderAndSelectItems(), we make it ourself.
	pidl = pidlFolder;
	pIdlFile = pidl;
	// Find the offset
	while (cb = pIdlFile->mkid.cb) 
	{
		pidl2 = pIdlFile;
		pIdlFile = (ITEMIDLIST*)((BYTE*)pIdlFile + cb);
	}

	cb = pidl2->mkid.cb;
	pidl2->mkid.cb = 0;

    bResult = FALSE;
	// Open the directory
	if(SUCCEEDED(GetShellFolderViewDual(pidl, &pIShellFolderViewDual))) 
	{
		pidl2->mkid.cb = cb;
		// 0 Deselect the item.  
		// 1 Select the item.  
		// 3 Put the item in edit mode.  
		// 4 Deselect all but the specified item.  
		// 8 Ensure the item is displayed in the view.  
		// 0x10 Give the item the focus.  
		COleVariant bszFile(pidl2);

		if(pIShellFolderViewDual != NULL)
		{
			// Select the file object.
			pIShellFolderViewDual->SelectItem(bszFile, 0x1d);
			pIShellFolderViewDual->Release();
		}
		bResult = TRUE;
	}

	return bResult;
}

BOOL exSHOpenFolderAndSelectItem(LPCTSTR szFilePath)
{
    if(szFilePath == NULL || szFilePath[0] == 0)
    {
        return FALSE;
    }

    tstring strFilePath(szFilePath);
    RemoveDoubleSlash(strFilePath);

    if(GetFileAttributes(strFilePath.c_str()) == (DWORD)-1)
    {
        return FALSE;
    }
    HMODULE hModule = LoadLibrary(_T("shell32.dll"));
    if(hModule == NULL)
    {
        return FALSE;
    }

	LPITEMIDLIST pIDL = NULL;
	BOOL bResult = exSHGetIDLFromPath(strFilePath.c_str(), &pIDL);
	if(bResult)
	{
		bResult = exSHOpenFolderAndSelectItem(pIDL);
		CoTaskMemFree(pIDL);
	}
    if(!bResult)
    {
        ASSERT(false);
        bResult = (ShellExecute(NULL, _T("explore"), strFilePath.c_str(), 0, 0, SW_SHOWNORMAL) > (HINSTANCE)32);
    }
    FreeLibrary(hModule);
	return bResult;
}

HKEY GetKeyFromName(LPCTSTR szName)
{
	static struct
	{
		LPCTSTR szName;
		HKEY	hRootKey;
	} _Data[] =
	{
		_T("HKEY_CLASSES_ROOT"), HKEY_CLASSES_ROOT,
		_T("HKEY_CURRENT_USER"), HKEY_CURRENT_USER,
		_T("HKEY_LOCAL_MACHINE"), HKEY_LOCAL_MACHINE,
		_T("HKEY_USERS"), HKEY_USERS,
		_T("HKEY_CURRENT_CONFIG"), HKEY_CURRENT_CONFIG,

		_T("HKCR"), HKEY_CLASSES_ROOT,
		_T("HKCU"), HKEY_CURRENT_USER,
		_T("HKLM"), HKEY_LOCAL_MACHINE,
		_T("HKU"), HKEY_USERS,
	};
	for(int i=0; i<sizeof(_Data) / sizeof(_Data[0]); i++)
	{
		if(_tcsicmp(szName, _Data[i].szName) == 0)
		{
			return _Data[i].hRootKey;
		}
	}
	return NULL;
}

BOOL IsRegKeyExists(LPCTSTR szRegPath, LPCTSTR szValue)
{
	LPCTSTR pName = _tcschr(szRegPath, _T('\\'));
	if(pName == NULL)
	{
		return (GetKeyFromName(szRegPath) != NULL);
	}
	tstring strRegPath = szRegPath;
	tstring strRootRegName = strRegPath.substr(0, pName - szRegPath);
	strRegPath = pName + 1;

	HKEY hRootKey = GetKeyFromName(strRootRegName.c_str());
	if(hRootKey == NULL) return FALSE;

	HKEY hKey = NULL;
	BOOL bResult = FALSE;
	if(RegOpenKeyEx(hRootKey, strRegPath.c_str(), 0, KEY_QUERY_VALUE, &hKey) == ERROR_SUCCESS)
	{
		if(szValue != NULL && szValue[0] != 0)
		{
			BYTE byTemp;
			DWORD dwLen = sizeof(BYTE);
			LONG lResult = RegQueryValueEx(hKey, szValue, 0, 0, &byTemp, &dwLen);
			if(lResult == ERROR_SUCCESS || lResult == ERROR_MORE_DATA)
			{
				bResult = TRUE;
			}
		}
		else
		{
			bResult = TRUE;
		}
		RegCloseKey(hKey);
	}
	return bResult;
}

BOOL FindRegeditWnd(HWND& hRegeditWnd, HWND& hTreeViewWnd, HWND& hListViewWnd)
{
	hRegeditWnd = FindWindow(_T("RegEdit_RegEdit"), NULL );
	if(hRegeditWnd == NULL)
	{
		SHELLEXECUTEINFO info = { 0 };
		info.cbSize = sizeof(info);
		info.fMask	= SEE_MASK_NOCLOSEPROCESS;
		info.lpVerb	= _T("open");
		info.lpFile	= _T("regedit.exe");
		info.nShow	= SW_SHOWNORMAL;
		if(!ShellExecuteEx( &info ))
            return FALSE;
		WaitForInputIdle( info.hProcess, INFINITE );
        for(int i=0; i<10; ++ i)
        {
            hRegeditWnd = FindWindow(_T("RegEdit_RegEdit"), NULL );
            if(hRegeditWnd != NULL)
                break;
            ::Sleep(500);
        }
        ::CloseHandle(info.hProcess);
	}
	if(hRegeditWnd != NULL)
	{
		hTreeViewWnd = FindWindowEx(hRegeditWnd, NULL, _T("SysTreeView32"), NULL );
		hListViewWnd = FindWindowEx(hRegeditWnd, NULL, _T("SysListView32"), NULL );
	}
	return (hRegeditWnd != NULL) && (hTreeViewWnd != NULL);
}

BOOL TranslateRegRoot(tstring& strRegPath)
{
    TCHAR szTranslateMap[][2][30] = 
    {
        _T("HKCR\\"), _T("HKEY_CLASSES_ROOT\\"),
        _T("HCR\\"), _T("HKEY_CLASSES_ROOT\\"),
        _T("HKLM\\"), _T("HKEY_LOCAL_MACHINE\\"),
        _T("HKCU\\"), _T("HKEY_CURRENT_USER\\"),
        _T("HKU\\"), _T("HKEY_USERS\\"),
        _T("HKCC\\"), _T("HKEY_CURRENT_CONFIG\\"),
    };
    BOOL bResult = FALSE;
    int nRootLen = 0;
    for(int i=0; i<sizeof(szTranslateMap) / sizeof(szTranslateMap[0]); i++)
    {
        nRootLen = _tcslen(szTranslateMap[i][0]);
        if(_tcsnicmp(szTranslateMap[i][0], strRegPath.c_str(), nRootLen) == 0)
        {
            tstring strTemp(szTranslateMap[i][1]);
            strTemp += strRegPath.substr(nRootLen);
            strRegPath = strTemp;
            bResult = TRUE;
            break;
        }
    }
    return FALSE;
}

BOOL exSHOpenRegKey(LPCTSTR szRegPath, LPCTSTR szRegKeyName)
{
    if(szRegPath == NULL || szRegPath[0] == _T('\0'))
    {
        return FALSE;
    }
    tstring strRegPath(szRegPath);
    TranslateRegRoot(strRegPath);

    RemoveDoubleSlash(strRegPath);

	if(!IsRegKeyExists(strRegPath.c_str(), NULL))
	{
		return FALSE;
	}

	HWND hTreeViewWnd = NULL;
	HWND hRegeditWnd = NULL;
	HWND hListViewWnd = NULL;

	if(!FindRegeditWnd(hRegeditWnd, hTreeViewWnd, hListViewWnd))
	{
		return FALSE;
	}
	SetForegroundWindow(hRegeditWnd);
	SetFocus(hRegeditWnd);

	ShowWindow(hTreeViewWnd, SW_SHOWNORMAL);
	SetForegroundWindow(hTreeViewWnd);

	int i = 0;
	// Close it up
	for(i=0; i<30; ++i)
	{
		SendMessage(hTreeViewWnd, WM_KEYDOWN, VK_LEFT, 0 );
	}

	SendMessage(hTreeViewWnd, WM_NULL, 0, 0);

	SendMessage(hTreeViewWnd, WM_KEYDOWN, VK_RIGHT, 0 );
	int nLen = strRegPath.length();
	for (i=0; i<nLen; ++i)
	{
		if( strRegPath[i] == '\\' )
		{
			SendMessage(hTreeViewWnd, WM_KEYDOWN, VK_RIGHT, 0 );
			SendMessage(hTreeViewWnd, WM_NULL, 0, 0);
		}
		else
		{
			SendMessage(hTreeViewWnd, WM_CHAR, toupper(strRegPath[i]), 0 );
		}
	}

	if(szRegKeyName != NULL && szRegKeyName[0] != 0)
	{
	    SetForegroundWindow(hListViewWnd);
	    SetFocus(hListViewWnd);

		int nLen = _tcslen(szRegKeyName);
		for (i=0; i<nLen; ++i)
		{
			SendMessage(hListViewWnd, WM_CHAR, toupper(szRegKeyName[i]), 0 );
		}
	}

	SetForegroundWindow(hRegeditWnd);
	SetFocus(hRegeditWnd);
	return TRUE;
}