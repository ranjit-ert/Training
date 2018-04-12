// pop3_imap4_simple.vcNative.h : main header file for the PROJECT_NAME application
//

#pragma once

#ifndef __AFXWIN_H__
	#error "include 'stdafx.h' before including this file for PCH"
#endif

#include "resource.h"		// main symbols


// Cpop3_imap4_simplevcNativeApp:
// See pop3_imap4_simple.vcNative.cpp for the implementation of this class
//

class Cpop3_imap4_simplevcNativeApp : public CWinApp
{
public:
	Cpop3_imap4_simplevcNativeApp();

// Overrides
	public:
	virtual BOOL InitInstance();

// Implementation

	DECLARE_MESSAGE_MAP()
};

extern Cpop3_imap4_simplevcNativeApp theApp;