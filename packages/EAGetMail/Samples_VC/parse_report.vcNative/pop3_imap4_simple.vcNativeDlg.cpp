// pop3_imap4_simple.vcNativeDlg.cpp : implementation file
//
//  ===============================================================================
// |    THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF      |
// |    ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO    |
// |    THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A         |
// |    PARTICULAR PURPOSE.                                                    |
// |    Copyright (c)2010-2012  ADMINSYSTEM SOFTWARE LIMITED                         |
// |
// |    Project: It demonstrates how to use EAGetMail to receive/parse email.
// |        
// |
// |    File: Form1 : implementation file
// |
// |    Author: Ivan Lui ( ivan@emailarchitect.net )
//  ===============================================================================
#include "stdafx.h"
#include "pop3_imap4_simple.vcNative.h"
#include "pop3_imap4_simple.vcNativeDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

// CAboutDlg dialog used for App About
class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// Dialog Data
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Implementation
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
END_MESSAGE_MAP()


// Cpop3_imap4_simplevcNativeDlg dialog




Cpop3_imap4_simplevcNativeDlg::Cpop3_imap4_simplevcNativeDlg(CWnd* pParent /*=NULL*/)
	: CDialog(Cpop3_imap4_simplevcNativeDlg::IDD, pParent)
	, textServer(_T(""))
	, textUser(_T(""))
	, textPassword(_T(""))
	, chkSSL(FALSE)
	, chkLeaveCopy(FALSE)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	m_bCancel = VARIANT_FALSE;
	m_uidlfile = _T("uidl.txt");

	TCHAR szFile[MAX_PATH+1];
	memset( szFile, 0, sizeof(szFile));
	::GetModuleFileName( NULL, szFile, MAX_PATH );
	LPCTSTR pszBuf = _tcsrchr( szFile, _T('\\') );
	if( pszBuf != NULL )
	{
		m_curpath.Append( szFile, pszBuf - szFile );
	}
	else
	{/*impossible*/
		m_curpath.Append( szFile );
	}

}

void Cpop3_imap4_simplevcNativeDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_EDIT_SERVER, textServer);
	DDV_MaxChars(pDX, textServer, 256);
	DDX_Text(pDX, IDC_EDIT_USER, textUser);
	DDV_MaxChars(pDX, textUser, 256);
	DDX_Text(pDX, IDC_EDIT_PASSWORD, textPassword);
	DDV_MaxChars(pDX, textPassword, 256);
	DDX_Check(pDX, IDC_CHECK_SSL, chkSSL);
	DDX_Control(pDX, IDC_COMBO_AUTHTYPE, lstAuthType);
	DDX_Control(pDX, IDC_COMBO_PROTOCOL, lstProtocol);
	DDX_Check(pDX, IDC_CHECK_LEAVECOPY, chkLeaveCopy);
	DDX_Control(pDX, IDC_PROGRESS1, pgBar);
	DDX_Control(pDX, IDC_LIST_MAIL, lstMail);
	DDX_Control(pDX, IDC_EDIT1, m_textReport);
}

BEGIN_MESSAGE_MAP(Cpop3_imap4_simplevcNativeDlg, CDialog)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	//}}AFX_MSG_MAP
	ON_BN_CLICKED(IDC_BUTTON_START, &Cpop3_imap4_simplevcNativeDlg::OnBnClickedButtonStart)
	ON_BN_CLICKED(IDC_BUTTON_CANCEL, &Cpop3_imap4_simplevcNativeDlg::OnBnClickedButtonCancel)
	ON_NOTIFY(NM_CLICK, IDC_LIST_MAIL, &Cpop3_imap4_simplevcNativeDlg::OnNMClickListMail)
	ON_BN_CLICKED(IDC_BUTTON_DEL, &Cpop3_imap4_simplevcNativeDlg::OnBnClickedButtonDel)
	ON_CBN_SELCHANGE(IDC_COMBO_PROTOCOL, &Cpop3_imap4_simplevcNativeDlg::OnCbnSelchangeComboProtocol)
END_MESSAGE_MAP()


// Cpop3_imap4_simplevcNativeDlg message handlers

BOOL Cpop3_imap4_simplevcNativeDlg::OnInitDialog()
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

	::CoInitialize( NULL );
	// TODO: Add extra initialization here
		
	lstAuthType.AddString(_T("USER/LOGIN"));
	lstAuthType.AddString(_T("APOP(CRAM-MD5)"));
	lstAuthType.AddString(_T("NTLM"));

	lstAuthType.SetCurSel( 0 );

	lstProtocol.AddString(_T("POP3"));
	lstProtocol.AddString(_T("IMAP4"));
	lstProtocol.AddString(_T("Exchange Web Service - 2007/2010"));
	lstProtocol.AddString(_T("Exchange WebDAV - 2000/2003"));

	lstProtocol.SetCurSel(0);

	pgBar.SetRange( 0, 100 );
	pgBar.SetPos( 0 );

	static TCHAR*	headers[4]	= { _T("From"), _T("Subject"), _T("Date"), _T("File") };
	static INT		cxs[4] = {150, 200, 150, 1 };
	DWORD dwStyle	=  lstMail.GetExtendedStyle();
	dwStyle			|= LVS_EX_FULLROWSELECT | LVS_EX_FLATSB;
	lstMail.SetExtendedStyle( dwStyle );


	
	INT			i = 0; 
	INT			nCount = 4;
	LVCOLUMN	column;
	
	column.mask	= LVCF_TEXT|LVCF_WIDTH;
	//column.cx	=  100;
	for( i = 0; i < nCount; i++ )
	{
		column.cx = cxs[i];
		column.pszText = headers[i];
		lstMail.InsertColumn( i, &column );
	}

	LoadMails();

	return TRUE;  // return TRUE  unless you set the focus to a control
}

void Cpop3_imap4_simplevcNativeDlg::OnSysCommand(UINT nID, LPARAM lParam)
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

void Cpop3_imap4_simplevcNativeDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

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

// The system calls this function to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR Cpop3_imap4_simplevcNativeDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

//============================================
// EAGetMail Event Handler
//============================================
//Catching the following events is not necessary, 
//just make the application more user friendly.
//If you use the object in asp.net/windows service or non-gui application, 
//You need not to catch the following events.
//To learn more detail, please refer to the code in EAGetMail EventHandler region
HRESULT __stdcall 
Cpop3_imap4_simplevcNativeDlg::OnIdleHandler(
		IDispatch * oSender,
        VARIANT_BOOL * Cancel)
{

	DoEvents();
	*Cancel = m_bCancel;

	return S_OK;
}

HRESULT __stdcall 
Cpop3_imap4_simplevcNativeDlg::OnConnectedHandler(
		IDispatch * oSender,
        VARIANT_BOOL * Cancel)
{
	CWnd* pStatus = GetDlgItem(IDC_STATIC_STATUS);
	pStatus->SetWindowText( _T("Connected"));
	DoEvents();
	*Cancel = m_bCancel;

	return S_OK;	
}

HRESULT __stdcall 
Cpop3_imap4_simplevcNativeDlg::OnQuitHandler(
		IDispatch * oSender,
        VARIANT_BOOL * Cancel)
{
	CWnd* pStatus = GetDlgItem(IDC_STATIC_STATUS);
	pStatus->SetWindowText( _T("Quiting..."));
	return S_OK;		
}

HRESULT __stdcall 
Cpop3_imap4_simplevcNativeDlg::OnSendCommandHandler(
		IDispatch * oSender,
        VARIANT data,
        VARIANT_BOOL IsDataStream,
        VARIANT_BOOL * Cancel)
{
	return S_OK;	
}

HRESULT __stdcall 
Cpop3_imap4_simplevcNativeDlg::OnReceiveResponseHandler (
        IDispatch * oSender,
        VARIANT data,
        VARIANT_BOOL IsDataStream,
        VARIANT_BOOL * Cancel )
{
	return S_OK;	
}

HRESULT __stdcall 
Cpop3_imap4_simplevcNativeDlg::OnSecuringHandler (
        IDispatch * oSender,
        VARIANT_BOOL * Cancel )
{
	CWnd* pStatus = GetDlgItem(IDC_STATIC_STATUS);
	pStatus->SetWindowText( _T("Securing..."));

	DoEvents();
	*Cancel = m_bCancel;
	return S_OK;	
}

HRESULT __stdcall 
Cpop3_imap4_simplevcNativeDlg::OnAuthorizedHandler (
        IDispatch * oSender,
        VARIANT_BOOL * Cancel )
{
	CWnd* pStatus = GetDlgItem(IDC_STATIC_STATUS);
	pStatus->SetWindowText( _T("Authorized"));

	DoEvents();
	*Cancel = m_bCancel;
	return S_OK;	
}

HRESULT __stdcall 
Cpop3_imap4_simplevcNativeDlg::OnSendingDataStreamHandler (
        IDispatch * oSender,
        long Sent,
        long Total,
        VARIANT_BOOL * Cancel )
{

	return S_OK;	
}

HRESULT __stdcall 
Cpop3_imap4_simplevcNativeDlg::OnReceivingDataStreamHandler (
        IDispatch * oSender,
        IDispatch * oInfo,
        long Received,
        long Total,
        VARIANT_BOOL * Cancel )
{
	DoEvents();
	*Cancel = m_bCancel;
	pgBar.SetRange32( 0, Total );
	pgBar.SetPos( Received );
	return S_OK;	
}

void Cpop3_imap4_simplevcNativeDlg::DoEvents()
{
	MSG msg; 
	while(PeekMessage(&msg,NULL,0,0,PM_REMOVE))
	{ 
		if(msg.message == WM_QUIT) 
			return; 

		TranslateMessage(&msg); 
		DispatchMessage(&msg); 
	}  
}
//============================================
// end EAGetMail Event Handler
//============================================

void Cpop3_imap4_simplevcNativeDlg::OnBnClickedButtonStart()
{
	UpdateData( TRUE );

	if( textServer.GetLength() == 0 ||
		textUser.GetLength() == 0 ||
		textPassword.GetLength() == 0 )
	{
		MessageBox( _T("Please input server, user and password."), NULL, MB_OK );
		return;
	}

	//::CoInitialize( NULL );

	CWnd* pStatus = GetDlgItem(IDC_STATIC_STATUS);
	CButton* pbtnStart = (CButton*)GetDlgItem(IDC_BUTTON_START);
	CButton* pbtnCancel = (CButton*)GetDlgItem(IDC_BUTTON_CANCEL);


	const int MailServerPop3 = 0;
	const int MailServerImap4 = 1;
	const int MailServerEWS = 2;
	const int MailServerDAV = 3;

	IMailClientPtr oClient;
	oClient.CreateInstance( "EAGetMailObj.MailClient" );

	IMailServerPtr oServer;
	oServer.CreateInstance( "EAGetMailObj.MailServer" );
	DispEventAdvise(oClient.GetInterfacePtr());

	// UIDL is the identifier of every email on POP3/IMAP4/Exchange server, to avoid retrieve
    // the same email from server more than once, we record the email UIDL retrieved every time
    // if you delete the email from server every time and not to leave a copy of email on
    // the server, then please remove all the function about uidl.
    // UIDLManager wraps the function to write/read uidl record from a text file.
	IUIDLManagerPtr oUIDLManager;
	oUIDLManager.CreateInstance("EAGetMailObj.UIDLManager");

	try
	{
        // For evaluation usage, please use "TryIt" as the license code, otherwise the
        // "invalid license code" exception will be thrown. However, the object will expire in 1-2 months, then
        // "trial version expired" exception will be thrown.
        oClient->LicenseCode = _T("TryIt");

		oServer->put_Server((BSTR)textServer.GetString());
		oServer->put_User((BSTR)textUser.GetString());
		oServer->put_Password((BSTR)textPassword.GetString());
		oServer->put_AuthType( lstAuthType.GetCurSel());
		INT nProtocol = lstProtocol.GetCurSel();

		oServer->put_Protocol( nProtocol);
		oServer->put_SSLConnection( (chkSSL)?VARIANT_TRUE:VARIANT_FALSE);
		if( nProtocol == MailServerPop3 )
		{
			if( chkSSL )
			{
				oServer->put_Port( 995 );
			}
			else
			{
				oServer->put_Port( 110 );
			}
		}
		else
		{
			if( chkSSL )
			{
				oServer->put_Port( 993 );
				
			}
			else
			{
				oServer->put_Port( 143 );
			}
		}

		CString mailFolder;
		mailFolder.Format( _T("%s\\inbox"), m_curpath.GetString());
		::CreateDirectory( mailFolder.GetString(), NULL );

		// Load existed UIDL records to UIDLManager
		CString file;
		file.Format( _T("%s\\%s"), m_curpath.GetString(), m_uidlfile.GetString() );
		oUIDLManager->Load((const TCHAR*)file);

		m_bCancel = VARIANT_FALSE;
		pbtnStart->EnableWindow( FALSE );
		pbtnCancel->EnableWindow( TRUE );

		pStatus->SetWindowText( _T("Connecting server..." ));
		oClient->Connect( oServer );

		_variant_t arInfo = oClient->GetMailInfos();
		
		// Remove the local uidl which is not existed on the server.
        oUIDLManager->SyncUIDL(oServer, arInfo);
        oUIDLManager->Update();

		SAFEARRAY *psa = arInfo.parray;

		long LBound = 0, UBound = 0;
		SafeArrayGetLBound( psa, 1, &LBound );
		SafeArrayGetUBound( psa, 1, &UBound );
	
		INT count = UBound-LBound+1;
		TCHAR szBuf[200];
		memset(szBuf, 0, sizeof(szBuf));
		::wsprintf( szBuf, _T("Total %d email(s)"), count );
		pStatus->SetWindowText( szBuf );

		for( long i = LBound; i <= UBound; i++ )
		{
			_variant_t vtInfo;
			SafeArrayGetElement( psa, &i, &vtInfo );
			
			IMailInfoPtr pInfo;
			vtInfo.pdispVal->QueryInterface(__uuidof(IMailInfo), (void**)&pInfo);
			
			IUIDLItemPtr oUIDLItem = oUIDLManager->FindUIDL( oServer, pInfo->UIDL );
			if(oUIDLItem != NULL )
			{
				// this email has been downloaded before.
				continue;
			}

			CString fileName;
			SYSTEMTIME curtm;
			::GetLocalTime( &curtm );
			fileName.Format( _T("%04d%02d%02d%02d%02d%02d%03d%d.eml"),
				curtm.wYear,
				curtm.wMonth,
				curtm.wDay,
				curtm.wHour,
				curtm.wMinute,
				curtm.wSecond,
				curtm.wMilliseconds,
				i );

			memset(szBuf, 0, sizeof(szBuf));
			::wsprintf( szBuf, _T("Receiving %d/%d mail header ... "), i+1, count );
			pStatus->SetWindowText( szBuf );

			IMailPtr oMail;
			oMail.CreateInstance("EAGetMailObj.Mail");
			oMail->LicenseCode = _T("TryIt");

			oMail->Load(oClient->GetMailHeader(pInfo));

			if (oMail->IsReport == VARIANT_FALSE )
			{
				// Not a report, continue
				// Add the email uidl to uidl file to avoid we check it next time. 
				oUIDLManager->AddUIDL(oServer, pInfo->UIDL, (const TCHAR*)fileName);
				continue;
			}

            // This is a report, get the entire email.
			oMail = oClient->GetMail(pInfo);
	

			CString full = mailFolder;
			full.Append( _T("\\"));
			full.Append( fileName );
			oMail->SaveAs( full.GetString(), VARIANT_TRUE );

			CString from = (TCHAR*)oMail->From->Name;
			from += _T("<");
			from += (TCHAR*)oMail->From->Address;
			from += _T(">");

			CString subject = (TCHAR*)oMail->Subject;

			DATE dt = oMail->ReceivedDate;
			SYSTEMTIME recvtm;
			::VariantTimeToSystemTime( dt, &recvtm );
			CString sdate;
			sdate.Format( _T("%04d-%02d-%02d %02d:%02d:%02d"),
				recvtm.wYear, recvtm.wMonth, recvtm.wDay, recvtm.wHour, recvtm.wMinute, recvtm.wSecond );


			INT nIndex = lstMail.InsertItem( 0, _T("temporary value"));
			lstMail.SetItemText( nIndex, 0, from.GetString());
			lstMail.SetItemText( nIndex, 1, subject.GetString());
			lstMail.SetItemText( nIndex, 2, sdate.GetString());
			lstMail.SetItemText( nIndex, 3, full.GetString());
			DWORD_PTR dtSort = (DWORD_PTR)(dt * 1000);
			lstMail.SetItemData( nIndex, dtSort ); 
			
			oMail.Release();

			if( chkLeaveCopy )
			{
				// Add the email uidl to uidl file to avoid we retrieve it next time. 
				oUIDLManager->AddUIDL( oServer, pInfo->UIDL, (const TCHAR*)fileName);
			}
		}

		if( !chkLeaveCopy )
		{
			pStatus->SetWindowText( _T("Deleting ..."));
			for( long i = LBound; i <= UBound; i++)
			{
				_variant_t vtInfo;
				SafeArrayGetElement( psa, &i, &vtInfo );
			
				IMailInfoPtr pInfo;
				vtInfo.pdispVal->QueryInterface(__uuidof(IMailInfo), (void**)&pInfo);
				oClient->Delete( pInfo );
				
				// Remove UIDL from local uidl file.
                oUIDLManager->RemoveUIDL( oServer, pInfo->UIDL );
				
			}
		}
		
		// Delete method just mark the email as deleted, 
		// Quit method pure the emails from server exactly.
		arInfo.Clear();
		oClient->Quit();
		pStatus->SetWindowText( _T("Completed"));
	}
	catch( _com_error &ep )
	{
		MessageBox((const TCHAR*)ep.Description(), _T("Error"), MB_OK );
		pStatus->SetWindowText( (const TCHAR*)ep.Description());
	}
	
	// Update the uidl list to local uidl file and then we can load it next time.
	oUIDLManager->Update();

	DispEventUnadvise(oClient.GetInterfacePtr());
	oClient.Release();
	oServer.Release();

	lstMail.SortItems( MyCompareProc,  0 );
	
	pbtnStart->EnableWindow( TRUE );
	pbtnCancel->EnableWindow( FALSE );
}

void Cpop3_imap4_simplevcNativeDlg::OnBnClickedButtonCancel()
{
	m_bCancel = VARIANT_TRUE;
	CButton* pBtn = (CButton*)GetDlgItem( IDC_BUTTON_CANCEL );
	pBtn->EnableWindow( FALSE );
}


//#region Parse and Display Mails
void
Cpop3_imap4_simplevcNativeDlg::LoadMails()
{
	lstMail.DeleteAllItems();

	CString mailFolder =  m_curpath;
	mailFolder.Append( _T("\\inbox" ));
	
	::CreateDirectory( mailFolder.GetString(), NULL );
	

	CString find = mailFolder;
	find.Append( _T("\\*.eml" ));

	WIN32_FIND_DATA findData;
	HANDLE hFind = ::FindFirstFile( find.GetString(), &findData );

	if( hFind == INVALID_HANDLE_VALUE )
		return;

	do
	{
		if((findData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY ) != FILE_ATTRIBUTE_DIRECTORY )
		{
			CString file = mailFolder;
			file.Append( _T("\\" ));
			file.Append( findData.cFileName );
			
			IMailPtr oMail;
			oMail.CreateInstance( "EAGetMailObj.Mail" );
			//For evaluation usage, please use "TryIt" as the license code, otherwise the 
			//"invalid license code" exception will be thrown. However, the object will expire in 1-2 months, then
			//"trial version expired" exception will be thrown.
			oMail->LicenseCode = _T("TryIt");

			// LoadFile( file, true ) only load the email header to Mail object to save the CPU and memory
			// the Mail object will load the whole email file later automatically if bodytext or attachment is required..
			oMail->LoadFile( file.GetString(), VARIANT_TRUE );

			
			CString from = (TCHAR*)oMail->From->Name;
			from += _T("<");
			from += (TCHAR*)oMail->From->Address;
			from += _T(">");

			CString subject = (TCHAR*)oMail->Subject;

			DATE dt = oMail->ReceivedDate;
			SYSTEMTIME recvtm;
			::VariantTimeToSystemTime( dt, &recvtm );
			CString sdate;
			sdate.Format( _T("%04d-%02d-%02d %02d:%02d:%02d"),
				recvtm.wYear, recvtm.wMonth, recvtm.wDay, recvtm.wHour, recvtm.wMinute, recvtm.wSecond );


			INT nIndex = lstMail.InsertItem( 0, _T("temporary value"));
			lstMail.SetItemText( nIndex, 0, from.GetString());
			lstMail.SetItemText( nIndex, 1, subject.GetString());
			lstMail.SetItemText( nIndex, 2, sdate.GetString());
			lstMail.SetItemText( nIndex, 3, file.GetString());
			DWORD_PTR dtSort = (DWORD_PTR)(dt * 1000);
			lstMail.SetItemData( nIndex, dtSort ); 
			
			oMail->Clear();
			oMail.Release();
		}
		
	}while( ::FindNextFile( hFind, &findData ));

	lstMail.SortItems( MyCompareProc,  0 );
}

void 
Cpop3_imap4_simplevcNativeDlg::ShowReport( LPCTSTR lpszFile )
{
	try
	{
		IMailPtr oMail;
		oMail.CreateInstance("EAGetMailObj.Mail");
		oMail->LicenseCode = _T("TryIt");
		oMail->LoadFile(lpszFile, VARIANT_FALSE );

		if(oMail->IsReport == VARIANT_FALSE )
		{
			MessageBox( _T("This is not a report" ));
			return;
		}

		const int FailureReport = 0;
		const int DeliveryReceipt = 1;
		const int ReadReceipt = 2;
		CString s;
		IMailReportPtr oReport = oMail->GetReport();
		switch( oReport->ReportType )
		{
		case DeliveryReceipt:
			s = _T("This is a deliver receipt.\r\n\r\n");
			break;
		case ReadReceipt:
			s = _T("This is a read receipt.\r\n\r\n");
			break;
		default:
			s = _T("This is a failure report.\r\n\r\n");
			break;
		}

		CString t;
		t.Format( _T("Original Sender: %s\r\n"), (const TCHAR*)oReport->OriginalSender );
		s += t;

		t.Format( _T("Original Recipient: %s\r\n"), (const TCHAR*)oReport->OriginalRecipient );
		s += t;

		t.Format( _T("Original Message-ID: %s\r\n\r\n"), (const TCHAR*)oReport->OriginalMessageID );
		s += t;

		if( oReport->ReportType == FailureReport )
		{
			t.Format( _T("Original Subject: %s\r\n"), (const TCHAR*)oReport->OriginalSubject );
			s += t;

			t.Format( _T("Report MTA: %s\r\n"), (const TCHAR*)oReport->ReportMTA );
			s += t;

			t.Format( _T("Error Code: %s\r\n"), (const TCHAR*)oReport->ErrCode );
			s += t;

			t.Format( _T("Error Description: %s\r\n\r\n"), (const TCHAR*)oReport->ErrDescription );
			s += t;
			
			s += _T("---- Original Message Header ----\r\n\r\n");
			IHeaderCollectionPtr oHeaders = oReport->OriginalHeaders;
			
			int count = oHeaders->Count;
			for( int i = 0; i < count; i++ )
			{
				IHeaderItemPtr oHeader = oHeaders->Item(i);
				t.Format( _T("%s: %s\r\n"), (const TCHAR*)oHeader->HeaderKey, (const TCHAR*)oHeader->HeaderValue );
				s += t;
			}
		}
		m_textReport.SetWindowTextW( s );
		
	}
	catch(  _com_error &ep )
	{
		MessageBox( (const TCHAR*)ep.Description(), NULL, MB_OK );
	}
}


void Cpop3_imap4_simplevcNativeDlg::OnBnClickedButtonDel()
{
	INT nIndex =  lstMail.GetSelectionMark();
	if( nIndex == -1 )
		return;

	CString fileName = lstMail.GetItemText( nIndex, 3 );
	if( MessageBox( _T("Are you sure to delete selected email?"), NULL, MB_OKCANCEL ) != IDOK )
		return;

	int pos = fileName.ReverseFind(_T('.'));
	CString mainName = fileName.Mid(0, pos);
	CString htmlName = mainName + _T(".htm");

	CString tempFolder = mainName;
	
	::DeleteFile( htmlName );
	CString find = tempFolder;
	find.Append( _T("\\*"));
	WIN32_FIND_DATA findData;
	HANDLE hFind = ::FindFirstFile( find, &findData );
	if( hFind != INVALID_HANDLE_VALUE )
	{
		do
		{
			if((findData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != FILE_ATTRIBUTE_DIRECTORY)
			{
				CString tname = tempFolder;
				tname.Append( _T("\\"));
				tname.Append( findData.cFileName );
				::DeleteFile( tname.GetString());
			}
		
		}while( ::FindNextFile( hFind, &findData));

		::FindClose( hFind );
	}

	::RemoveDirectory( tempFolder.GetString());
	lstMail.DeleteItem( nIndex );

}

int CALLBACK Cpop3_imap4_simplevcNativeDlg::MyCompareProc(LPARAM lParam1, LPARAM lParam2, 
			LPARAM lParamSort)
{
	return (int)(lParam2 -lParam1);
}


void Cpop3_imap4_simplevcNativeDlg::OnNMClickListMail(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMITEMACTIVATE pNMItemActivate = reinterpret_cast<NMITEMACTIVATE*>(pNMHDR);
	// TODO: Add your control notification handler code here
	*pResult = 0;

	if( pNMItemActivate->iItem == -1 )
		return;

	CString file = lstMail.GetItemText( pNMItemActivate->iItem, 3 );
	if( file.GetLength() > 0 )
		ShowReport( file.GetString());
}

void Cpop3_imap4_simplevcNativeDlg::OnCbnSelchangeComboProtocol()
{
	 // By default, Exchange Web Service requires SSL connection.
     // For other protocol, please set SSL connection based on your server setting manually
           
	//const int MailServerEWS = 2;
	int selected = lstProtocol.GetCurSel();
	UpdateData( TRUE );
	if( selected == 2 )
	{
		chkSSL = TRUE;
		UpdateData( FALSE );
	}
	
}
