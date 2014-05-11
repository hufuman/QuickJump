#pragma once

BOOL exSHGetIDLFromPath(LPCTSTR szFolderPath, LPITEMIDLIST* ppidl);

BOOL exSHOpenFolderAndSelectItem(LPCTSTR szFilePath);
BOOL exSHOpenFolderAndSelectItem(ITEMIDLIST *pidlFolder);

BOOL exSHOpenRegKey(LPCTSTR szRegPath, LPCTSTR szRegKeyName);