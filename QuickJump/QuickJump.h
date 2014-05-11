#pragma once

#include "resource.h"		// main symbols

class CQuickJumpApp : public CWinApp
{
public:
	CQuickJumpApp();

// Overrides
public:
	virtual BOOL InitInstance();

// Implementation
	DECLARE_MESSAGE_MAP()
};
