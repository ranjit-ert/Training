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

#include "AddFolderDlg.h"
#include "FolderDlg.h"

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
	DDX_Control(pDX, IDC_PROGRESS1, pgBar);
	DDX_Control(pDX, IDC_LIST_MAIL, lstMail);
	DDX_Control(pDX, IDC_EXPLORER_MAIL, webMail);
	DDX_Control(pDX, IDC_TREE1, trFolders);
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
	ON_NOTIFY(NM_RCLICK, IDC_TREE1, &Cpop3_imap4_simplevcNativeDlg::OnNMRClickTree1)
	ON_COMMAND(ID_MAILMENU_REFRESHFOLDERS, &Cpop3_imap4_simplevcNativeDlg::OnMailmenuRefreshfolders)
	ON_NOTIFY(TVN_SELCHANGED, IDC_TREE1, &Cpop3_imap4_simplevcNativeDlg::OnTvnSelchangedTree1)
	ON_COMMAND(ID_MAILMENU_REFRESHMAILS, &Cpop3_imap4_simplevcNativeDlg::OnMailmenuRefreshmails)
	ON_COMMAND(ID_MAILMENU_ADDFOLDER, &Cpop3_imap4_simplevcNativeDlg::OnMailmenuAddfolder)
	ON_COMMAND(ID_MAILMENU_DELETEFOLDER, &Cpop3_imap4_simplevcNativeDlg::OnMailmenuDeletefolder)
	ON_COMMAND(ID_MAILMENU_RENAMEFOLDER, &Cpop3_imap4_simplevcNativeDlg::OnMailmenuRenamefolder)
	ON_NOTIFY(TVN_ENDLABELEDIT, IDC_TREE1, &Cpop3_imap4_simplevcNativeDlg::OnTvnEndlabeleditTree1)
	ON_BN_CLICKED(IDC_BUTTON_UNDELETE, &Cpop3_imap4_simplevcNativeDlg::OnBnClickedButtonUndelete)
	ON_BN_CLICKED(IDC_BUTTON_UNREAD, &Cpop3_imap4_simplevcNativeDlg::OnBnClickedButtonUnread)
	ON_BN_CLICKED(IDC_BUTTON_PURE, &Cpop3_imap4_simplevcNativeDlg::OnBnClickedButtonPure)
	ON_BN_CLICKED(IDC_BUTTON_COPY, &Cpop3_imap4_simplevcNativeDlg::OnBnClickedButtonCopy)
	ON_BN_CLICKED(IDC_BUTTON_MOVE, &Cpop3_imap4_simplevcNativeDlg::OnBnClickedButtonMove)
	ON_BN_CLICKED(IDC_BUTTON_UPLOAD, &Cpop3_imap4_simplevcNativeDlg::OnBnClickedButtonUpload)
	ON_BN_CLICKED(IDC_BUTTON_QUIT, &Cpop3_imap4_simplevcNativeDlg::OnBnClickedButtonQuit)
	ON_NOTIFY(NM_CUSTOMDRAW, IDC_LIST_MAIL, &Cpop3_imap4_simplevcNativeDlg::OnNMCustomdrawListMail)
	ON_WM_CLOSE()
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

	webMail.Navigate( _T("about:blank"), NULL, NULL, NULL, NULL );
	
	EnableIdle( FALSE );
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
	DoEvents();
	*Cancel = m_bCancel;
	pgBar.SetRange32( 0, Total );
	pgBar.SetPos( Sent );
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

	EnableIdle(FALSE);
	m_bCancel = FALSE;

    CWnd* pBtnCancel = GetDlgItem( IDC_BUTTON_CANCEL );
	pBtnCancel->EnableWindow( FALSE );

	CWnd* pBtnStart = GetDlgItem( IDC_BUTTON_START );

	if( textServer.GetLength() == 0 ||
		textUser.GetLength() == 0 ||
		textPassword.GetLength() == 0 )
	{
		MessageBox( _T("Please input server, user and password."), NULL, MB_OK );
		return;
	}

	pBtnStart->EnableWindow( FALSE );

	lstMail.DeleteAllItems();
	_ClearMailItems();

	trFolders.DeleteAllItems();
	_ClearFolders();

	//::CoInitialize( NULL );

	CWnd* pStatus = GetDlgItem(IDC_STATIC_STATUS);
	CButton* pbtnStart = (CButton*)GetDlgItem(IDC_BUTTON_START);
	CButton* pbtnCancel = (CButton*)GetDlgItem(IDC_BUTTON_CANCEL);

	
	const int MailServerPop3 = 0;
	const int MailServerImap4 = 1;
	const int MailServerEWS = 2;
	const int MailServerDAV = 3;

	if( oClient != NULL )
	{
		DispEventUnadvise(oClient.GetInterfacePtr());
		oClient.Release();
	}

	if( oCurServer != NULL )
		oCurServer.Release();

	if( oUIDLManager != NULL )
		oUIDLManager.Release();
	
	oClient.CreateInstance("EAGetMailObj.MailClient");
	oClient->LicenseCode = _T("TryIt");
	DispEventAdvise(oClient.GetInterfacePtr());	

	oCurServer.CreateInstance( "EAGetMailObj.MailServer" );
	oUIDLManager.CreateInstance( "EAGetMailObj.UIDLManager");

	oCurServer->Server = textServer.GetString();
	oCurServer->User = textUser.GetString();
	oCurServer->Password = textPassword.GetString();
	oCurServer->AuthType = lstAuthType.GetCurSel();

	INT nProtocol = lstProtocol.GetCurSel()+1;

	oCurServer->Protocol = nProtocol;

	oCurServer->SSLConnection = (chkSSL)?VARIANT_TRUE:VARIANT_FALSE;

	if( nProtocol == MailServerImap4 )
	{
		if( chkSSL )
		{
			oCurServer->Port = 993;
		}
		else
		{
			oCurServer->Port = 143;
		}
	}

	try
	{
		// Enable log file 
        // oClient->LogFileName = _T("d:\\imap.txt");
		
		pBtnCancel->EnableWindow( TRUE );
		pStatus->SetWindowText(_T("Connecting ... "));
        oClient->Connect(oCurServer);
		CString curserver;
		curserver.Format( _T("%s\\%s"), (const TCHAR*)oCurServer->Server, (const TCHAR*)oCurServer->User );
		curserver.MakeLower();

		HTREEITEM hItem = trFolders.InsertItem( curserver );
		trFolders.SetItemData( hItem, (DWORD_PTR)oCurServer.GetInterfacePtr());

		trFolders.SelectItem( hItem );
		ShowNode();

		EnableIdle( TRUE );
	}
	catch( _com_error &ep )
	{
		MessageBox((const TCHAR*)ep.Description(), _T("Error"), MB_OK );
		pStatus->SetWindowText((const TCHAR*)ep.Description());
		oCurServer.Release();

		pBtnStart->EnableWindow( TRUE );
		pBtnCancel->EnableWindow( FALSE );
	}
}
// Cancel current operation
void Cpop3_imap4_simplevcNativeDlg::OnBnClickedButtonCancel()
{
	m_bCancel = VARIANT_TRUE;
	CButton* pBtn = (CButton*)GetDlgItem( IDC_BUTTON_CANCEL );
	pBtn->EnableWindow( FALSE );
}


//#region Parse and Display Mails
void 
Cpop3_imap4_simplevcNativeDlg::ShowMail( LPCTSTR lpszFile )
{
	try
	{
		CString fileName = lpszFile;

		int pos = fileName.ReverseFind(_T('.'));
		CString mainName = fileName.Mid(0, pos);
		CString htmlName = mainName + _T(".htm");

		CString tempFolder = mainName;
		WIN32_FIND_DATA findData;
		HANDLE hFind = ::FindFirstFile( htmlName.GetString(), &findData );
		BOOL bExist = FALSE;

		if( hFind != INVALID_HANDLE_VALUE )
		{
			::FindClose( hFind );
			bExist = TRUE;
		}

		if( !bExist )
		{	//we haven't generate the html for this email, generate it now.
			_GenerateHtmlForEmail( htmlName, fileName, tempFolder );
		}

		
		webMail.Navigate( htmlName.GetString(), NULL, NULL, NULL, NULL );
	}
	catch(  _com_error &ep )
	{
		MessageBox((const TCHAR*)ep.Description(), NULL, MB_OK );
	}
}

//we generate a html + attachment folder for every email, once the html is create,
// next time we don't need to parse the email again.
void 
Cpop3_imap4_simplevcNativeDlg::_GenerateHtmlForEmail( CString &htmlName, CString &emlFile, CString& tempFolder )
{
	//For evaluation usage, please use "TryIt" as the license code, otherwise the 
	//"invalid license code" exception will be thrown. However, the object will expire in 1-2 months, then
	//"trial version expired" exception will be thrown.
	IMailPtr oMail;
	oMail.CreateInstance( "EAGetMailObj.Mail");
	oMail->LicenseCode = _T("TryIt");
	oMail->LoadFile( emlFile.GetString(), VARIANT_FALSE );

	if( oMail->IsEncrypted == VARIANT_TRUE )
	{
		try
		{
			//this email is encrypted, we decrypt it by user default certificate.
			// you can also use specified certificate like this
			// oCert = new Certificate();
			// oCert.Load("c:\\test.pfx", "pfxpassword", Certificate.CertificateKeyLocation.CRYPT_USER_KEYSET)
			// oMail = oMail.Decrypt( oCert );
			oMail = oMail->Decrypt( NULL );
		}
		catch( _com_error &ep )
		{
			MessageBox((const TCHAR*)ep.Description(), NULL, MB_OK );
			oMail->LoadFile( emlFile.GetString(), VARIANT_FALSE );
		}	
	}

	if( oMail->IsSigned == VARIANT_TRUE )
	{
		try
		{
			//this email is digital signed.
			ICertificatePtr oCert = oMail->VerifySignature();
			MessageBox( _T("This email contains a valid digital signature."), NULL, MB_OK );
			//you can add the certificate to your certificate storage like this
			//cert.AddToStore( Certificate.CertificateStoreLocation.CERT_SYSTEM_STORE_CURRENT_USER,
			//	"addressbook" );
			// then you can use send the encrypted email back to this sender.
		}
		catch(_com_error &ep )
		{
			MessageBox((const TCHAR*)ep.Description(), NULL, MB_OK );
		}
	}

	CString html = (TCHAR*)oMail->HtmlBody;
	CString hdr;
	hdr.Preallocate( 1024 * 5 );

	hdr.Append( _T("<font face=\"Courier New,Arial\" size=2>"));
	hdr.Append( _T("<b>From:</b> "));

	CString tp = (TCHAR*)oMail->From->Name;
	tp += _T("<");
	tp += (TCHAR*)oMail->From->Address;
	tp += _T(">");
	hdr.Append( _FormatHtmlTag(tp.GetString()));
	hdr.Append( _T("<br>"));

	long LBound = 0, UBound = 0;
	SAFEARRAY *psa = NULL;

	_variant_t arAddr = oMail->To;
	
	psa = arAddr.parray;
	SafeArrayGetLBound( psa, 1, &LBound );
	SafeArrayGetUBound( psa, 1, &UBound );
	
	INT count = UBound-LBound+1;
	if( count > 0 )
	{
		hdr.Append( _T("<b>To:</b> "));
		for( long i = LBound; i <= UBound; i++ )
		{
			_variant_t vtAddr;
			SafeArrayGetElement( psa, &i, &vtAddr );
			
			IMailAddressPtr pAddr;
			vtAddr.pdispVal->QueryInterface( __uuidof(IMailAddress), (void**)&pAddr );
			
			tp = (TCHAR*)pAddr->Name;
			tp += _T("<");
			tp += (TCHAR*)pAddr->Address;
			tp += _T(">");

			hdr.Append( _FormatHtmlTag( tp.GetString()));
			if( i < UBound )
			{
				hdr.Append( _T(";"));
			}
		}
		hdr.Append( _T("<br>"));
	}
	
	arAddr.Clear();
	arAddr = oMail->Cc;
	
	psa = arAddr.parray;
	SafeArrayGetLBound( psa, 1, &LBound );
	SafeArrayGetUBound( psa, 1, &UBound );
	
	count = UBound-LBound+1;
	if( count > 0 )
	{
		hdr.Append( _T("<b>Cc:</b> "));
		for( long i = LBound; i <= UBound; i++ )
		{
			_variant_t vtAddr;
			SafeArrayGetElement( psa, &i, &vtAddr );
			
			IMailAddressPtr pAddr;
			vtAddr.pdispVal->QueryInterface( __uuidof(IMailAddress), (void**)&pAddr );

			tp = (TCHAR*)pAddr->Name;
			tp += _T("<");
			tp += (TCHAR*)pAddr->Address;
			tp += _T(">");

			hdr.Append( _FormatHtmlTag( tp.GetString()));
			if( i < UBound )
			{
				hdr.Append( _T(";"));
			}
		}
		hdr.Append( _T("<br>"));
	}
	
	hdr.Append( _T( "<b>Subject:</b>" ));
	hdr.Append( _FormatHtmlTag((TCHAR*)oMail->Subject));
	hdr.Append( _T("<br>"));

	_variant_t atts = oMail->Attachments;
	psa = atts.parray;
	SafeArrayGetLBound( psa, 1, &LBound );
	SafeArrayGetUBound( psa, 1, &UBound );
	
	count = UBound-LBound+1;
	if( count > 0 )
	{
		::CreateDirectory( tempFolder.GetString(), NULL );
		hdr.Append( _T("<b>Attachments:</b>"));
		for( long i = LBound; i <= UBound; i++ )
		{
			_variant_t vtAtt;
			SafeArrayGetElement( psa, &i, &vtAtt );
			
			IAttachmentPtr pAtt;
			vtAtt.pdispVal->QueryInterface( __uuidof(IAttachment), (void**)&pAtt );

			CString name = (TCHAR*)pAtt->Name;
			if( name.CompareNoCase( _T("winmail.dat")) == 0 )
			{
				//this attachment is in OUTLOOK RTF format, decode it here.
				_variant_t tatts;
				try
				{
					tatts = oMail->ParseTNEF( pAtt->Content, VARIANT_TRUE );
				}
				catch(_com_error &ep )
				{
					MessageBox((const TCHAR*)ep.Description(), NULL, MB_OK );
					continue;
				}
	
				long XLBound = 0, XUBound = 0;

				SAFEARRAY* tpsa = tatts.parray;
				SafeArrayGetLBound( tpsa, 1, &XLBound );
				SafeArrayGetUBound( tpsa, 1, &XUBound );
				for( long x = XLBound; x <= XUBound; x++ )
				{
					_variant_t vttAtt;
					SafeArrayGetElement( tpsa, &x, &vttAtt );
					IAttachmentPtr ptAtt;
					vttAtt.pdispVal->QueryInterface( __uuidof(IAttachment), (void**)&ptAtt );

					CString tattname = tempFolder;
					tattname.Append( _T("\\"));
					tattname.Append((TCHAR*)ptAtt->Name );
					ptAtt->SaveAs( tattname.GetString(), VARIANT_TRUE );
					
					hdr.Append( _T("<a href=\""));
					hdr.Append(tattname);
					hdr.Append( _T("\" target=\"_blank\">"));
					hdr.Append((TCHAR*)ptAtt->Name);
					hdr.Append(_T("</a> "));
				}
				continue;
			}
			CString attname = tempFolder;
			attname.Append(_T("\\"));
			attname.Append((TCHAR*)pAtt->Name);
			pAtt->SaveAs( attname.GetString(), VARIANT_TRUE );

			hdr.Append( _T("<a href=\""));
			hdr.Append(attname);
			hdr.Append( _T("\" target=\"_blank\">"));
			hdr.Append((TCHAR*)pAtt->Name);
			hdr.Append(_T("</a> "));

			CString contentID = (TCHAR*)pAtt->ContentID;
			CString contentType = (TCHAR*)pAtt->ContentType;
			if( contentID.GetLength() > 0 )
			{
				CString find = _T("cid:");
				find.Append( contentID );
				//show embedded image.
				html.Replace( find, attname );
			}
			else if( _tcsnicmp( contentType.GetString(), _T("image/"), _tcslen(_T("image/"))) == 0 )
			{
				//show attached image
				html.Append( _T("<hr><img src=\""));
				html.Append( attname );
				html.Append( _T("\">"));
			}
		}
	}
	
	html = hdr + "<hr>" + html;
	_ReplaceHtmlCharset( html );

	IToolsPtr oTools;
	oTools.CreateInstance("EAGetMailObj.Tools");
	oTools->WriteTextFile( htmlName.GetString(), html.GetString(), CP_UTF8 );

	oMail->Clear();
	oMail.Release();
}
CString 
Cpop3_imap4_simplevcNativeDlg::_ReplaceHtmlCharset( CString &s )
{
	int start = 0;
	int pos = 0;
	BOOL bCharset = FALSE;
	while((pos = s.Find( _T("<"), start )) != -1 )
	{
		LPCTSTR pT = (LPCTSTR)s + pos;
		pos++;
		if( pos >= s.GetLength())
			break;

		start = pos;
		// header end
		if( _tcsnicmp( pT, _T("</head"), _tcslen(_T("</head"))) == 0 )
			break;

		if( _tcsnicmp( pT, _T("<body"), _tcslen(_T("<body"))) == 0 )
			break;
		
		if( _tcsnicmp( pT, _T("<meta"), _tcslen(_T("<meta"))) == 0 )
		{
			pos = s.Find( _T(">"), start );
			if( pos == -1 )
				break;

		
			CString tmeta = s.Mid( start, pos - start );
			int buf = pos-start;
			// because CString.MakeLower has a crash problem, we have to use our makelower function
			LPTSTR pszMeta = new TCHAR[buf+1];
			memset( pszMeta, 0, (buf+1)*sizeof(TCHAR));
			memcpy( pszMeta, (LPCTSTR)tmeta, buf*sizeof(TCHAR));
			::_tcslwr_s( pszMeta, buf+1 );
			
			if(_tcsstr( pszMeta, _T("charset")) != NULL )
				bCharset = TRUE;

			delete []pszMeta;
			if( bCharset )
			{
				s.Replace( (LPCTSTR)tmeta, _T("meta HTTP-EQUIV=\"Content-Type\" Content=\"text/html; charset=utf-8\""));
				break;
			}

			start++;
			if( start >= s.GetLength())
				break;
		}
	}


	if( !bCharset )
	{
		s.Insert( 0, _T("<meta HTTP-EQUIV=\"Content-Type\" Content=\"text/html; charset=utf-8\">"));
	}

	return s;
}
CString
Cpop3_imap4_simplevcNativeDlg::_FormatHtmlTag( LPCTSTR lpszSrc )
{
	CString src = lpszSrc;

	src.Replace( _T(">"), _T("&gt;"));
	src.Replace( _T("<"), _T("&lt;"));
	return src;
}

// Delete email
void Cpop3_imap4_simplevcNativeDlg::OnBnClickedButtonDel()
{
	if( lstMail.GetSelectedCount() == 0 )
		return;

	HTREEITEM hNode = trFolders.GetSelectedItem();
	if( hNode == NULL )
		return;

	const int MailServerPop3 = 0;
	const int MailServerImap4 = 1;
	const int MailServerEWS = 2;
	const int MailServerDAV = 3;	
	CWnd* pStatus = GetDlgItem(IDC_STATIC_STATUS);

	EnableIdle( FALSE );
	try
	{
		INT nIndex = -1;
		ConnectServer(hNode);

		while((nIndex = lstMail.GetNextItem( nIndex, LVNI_SELECTED  )) != -1)
		{
			LPMAILINFOITEM pMailItem = (LPMAILINFOITEM)lstMail.GetItemData( nIndex );
			IMailInfo *info = pMailItem->oInfo;
			
			oClient->Delete( info );
			if( oCurServer->Protocol == MailServerImap4 )
			{
				lstMail.Update( nIndex );
			}
			else
			{
				lstMail.DeleteItem( nIndex );
				nIndex = -1;
			}
		}
	}
	catch( _com_error &ep )
	{
		MessageBox((const TCHAR*)ep.Description(), _T("Error"), MB_OK );
		pStatus->SetWindowText((const TCHAR*)ep.Description());
		oClient->Close();
	}

	EnableIdle( TRUE );
}

int CALLBACK Cpop3_imap4_simplevcNativeDlg::MyCompareProc(LPARAM lParam1, LPARAM lParam2, 
			LPARAM lParamSort)
{
	LPMAILINFOITEM p1 = (LPMAILINFOITEM)lParam1;
	LPMAILINFOITEM p2 = (LPMAILINFOITEM)lParam2;
	return (int)(p2->sortDate -p1->sortDate);
}

// show current email
void Cpop3_imap4_simplevcNativeDlg::OnNMClickListMail(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMITEMACTIVATE pNMItemActivate = reinterpret_cast<NMITEMACTIVATE*>(pNMHDR);
	// TODO: Add your control notification handler code here
	*pResult = 0;

	if( lstMail.GetSelectedCount() != 1 )
		return;

	INT index = lstMail.GetNextItem( -1, LVNI_SELECTED );
	if( index == -1 )
		return;

	LPMAILINFOITEM pMailItem = (LPMAILINFOITEM)lstMail.GetItemData(index);
	CWnd* pStatus = GetDlgItem(IDC_STATIC_STATUS);

	EnableIdle(FALSE);

	try
	{
		IMailInfo* pInfo = pMailItem->oInfo;
		IToolsPtr oTool;
		oTool.CreateInstance("EAGetMailObj.Tools");

		HTREEITEM hNode = trFolders.GetSelectedItem();
		CString mailbox = GetFolderByNode(hNode);
		_CreateFullFolder(mailbox);

		// Find current email record in UIDL file.
		IUIDLItemPtr oUIDL = oUIDLManager->FindUIDL(oCurServer, pInfo->UIDL);
		if (oUIDL == NULL)
		{
			// show never happen except you delete the file from the folder manually.
			MessageBox(_T("No email file found!"));
			EnableIdle(TRUE);
			return;
		}

		// Get the  local file name for this email UIDL
		CString emlFile;
		emlFile.Format( _T("%s\\%s"), (const TCHAR*)mailbox, (const TCHAR*)oUIDL->FileName);

		int pos = emlFile.ReverseFind(_T('.'));
		CString mainName = emlFile.Mid(0, pos);
		CString htmlName = mainName;
		htmlName += _T(".htm");

		// only mail header is retrieved, now retrieve full content of mail.
		if (oTool->ExistFile((const TCHAR*)htmlName) == VARIANT_FALSE )
		{
			IMailPtr oMail = oClient->GetMail(pInfo);
			oMail->SaveAs((const TCHAR*)emlFile, VARIANT_TRUE );
			pgBar.SetRange( 0, 100 );
			pgBar.SetPos( 0 );
		}

		if (pInfo->Read == VARIANT_FALSE )
		{
			oClient->MarkAsRead(pInfo, VARIANT_TRUE);
		}

		ShowMail(emlFile);

		lstMail.Update( index );
	}
	catch( _com_error &ep )
	{
		MessageBox((const TCHAR*)ep.Description(), _T("Error"), MB_OK );
		pStatus->SetWindowText((const TCHAR*)ep.Description());
		oClient->Close();
	}

	EnableIdle(TRUE);
	
}
void Cpop3_imap4_simplevcNativeDlg::OnCbnSelchangeComboProtocol()
{
	 // By default, Exchange Web Service requires SSL connection.
     // For other protocol, please set SSL connection based on your server setting manually
           
	//const int MailServerEWS = 2;
	int selected = lstProtocol.GetCurSel() + 1;
	UpdateData( TRUE );
	if( selected == 2 )
	{
		chkSSL = TRUE;
		UpdateData( FALSE );
	}
	
}

// show mail list for current folder
void Cpop3_imap4_simplevcNativeDlg::ShowNode()
{
	HTREEITEM hItem = trFolders.GetSelectedItem();
	if( hItem == NULL )
	{
		lstMail.DeleteAllItems();
		_ClearMailItems();
		return;
	}

	CWnd* pStatus = GetDlgItem(IDC_STATIC_STATUS);

	try
	{
		HTREEITEM hItemParent = trFolders.GetParentItem( hItem );
		if (hItemParent == NULL )
		{
			// Current node is root node, 
			// So we get all folders from server and display it
			lstMail.DeleteAllItems();
			_ClearMailItems();

			ConnectServer(hItem);

			pStatus->SetWindowText( _T("Refreshing Folders ..."));
			_variant_t fds = oClient->Imap4Folders;
			_ClearFolders();

			ExpandFolders(fds, hItem);

			trFolders.Expand( hItem, TVE_EXPAND );
			pStatus->SetWindowText(_T(""));
		}
		else
		{
			IImap4Folder* fd = (IImap4Folder*)trFolders.GetItemData(hItem);
			if (fd->NoSelect == VARIANT_FALSE )
			{
				//Display emails list in current folder
				LoadServerMails(hItem, fd);
			}
			else
			{
				pStatus->SetWindowText(_T(""));
				lstMail.DeleteAllItems();// This is a folder without email storage
				_ClearMailItems();

			}
		}
	}
	catch( _com_error &ep )
	{
		MessageBox((const TCHAR*)ep.Description(), _T("Error"), MB_OK );
		pStatus->SetWindowText((const TCHAR*)ep.Description());
		oClient->Close();
	}
}

// Connect server indeed
void Cpop3_imap4_simplevcNativeDlg::ConnectServer(HTREEITEM hNode)
{
	CWnd* pStatus = GetDlgItem(IDC_STATIC_STATUS);
	if (oClient->Connected == VARIANT_FALSE)
	{
		pStatus->SetWindowText(_T("Connecting ... "));

		m_bCancel = FALSE;
		oClient->Connect(oCurServer);
		if (trFolders.GetParentItem( hNode ) != NULL )
		{
			IImap4Folder *pFolder = (IImap4Folder*)trFolders.GetItemData( hNode );
			oClient->SelectFolder( pFolder );
		}
	}
}

// show server folders
void Cpop3_imap4_simplevcNativeDlg::ExpandFolders(_variant_t &fds, HTREEITEM hNode )
{
	_DeleteSubNode( hNode );
	SAFEARRAY *psa = fds.parray;
	long LBound = 0, UBound = 0;
	SafeArrayGetLBound( psa, 1, &LBound );
	SafeArrayGetUBound( psa, 1, &UBound );

	INT count = UBound-LBound+1;
	for( long i = LBound; i <= UBound; i++ )
	{
		_variant_t vtFolder;
		SafeArrayGetElement( psa, &i, &vtFolder );

		IImap4FolderPtr oFolder = NULL;
		vtFolder.pdispVal->QueryInterface(__uuidof(IImap4Folder), (void**)&oFolder);
		
		HTREEITEM hSub = trFolders.InsertItem( oFolder->Name, hNode );
		trFolders.SetItemData( hSub, (DWORD_PTR)oFolder.GetInterfacePtr());
		m_arFolder.Add(oFolder.GetInterfacePtr());
		
		ExpandFolders( oFolder->SubFolders, hSub );

		trFolders.Expand( hSub, TVE_EXPAND );
		oFolder.Detach();
	}
}
void Cpop3_imap4_simplevcNativeDlg::OnNMRClickTree1(NMHDR *pNMHDR, LRESULT *pResult)
{
	if( trFolders.GetSelectedItem() == NULL )
		return;

	*pResult = 0;
	CMenu mnuPopupSubmit;
	BOOL b = mnuPopupSubmit.LoadMenu(IDR_MENU1);

	POINT point;
	::GetCursorPos( &point );
	CMenu *mnuPopupMenu = mnuPopupSubmit.GetSubMenu(0);
	mnuPopupMenu->TrackPopupMenu(TPM_LEFTALIGN | TPM_RIGHTBUTTON, point.x, point.y, this);
}
// refresh folder
void Cpop3_imap4_simplevcNativeDlg::OnMailmenuRefreshfolders()
{
	HTREEITEM hItem = trFolders.GetSelectedItem();
	if (hItem ==  NULL)
		return;

	CWnd* pStatus = GetDlgItem(IDC_STATIC_STATUS);

	EnableIdle(FALSE);
	try
	{
		HTREEITEM hNode = hItem;
		while( trFolders.GetParentItem( hNode ) != NULL )
		{
			hNode = trFolders.GetParentItem( hNode );
		}

		trFolders.SelectItem( NULL );
		lstMail.DeleteAllItems();
		_ClearMailItems();

		ConnectServer(hNode);
		oClient->RefreshFolders();

		pStatus->SetWindowText( _T("Refreshing Folders ..."));
		_variant_t fds = oClient->Imap4Folders;
		
		_ClearFolders();
		ExpandFolders(fds, hNode);
		trFolders.Expand( hNode, TVE_EXPAND );
		pStatus->SetWindowText(_T(""));
	}
	catch( _com_error &ep )
	{
		MessageBox((const TCHAR*)ep.Description(), _T("Error"), MB_OK );
		pStatus->SetWindowText((const TCHAR*)ep.Description());
		oClient->Close();
	}

	EnableIdle(TRUE);
}

 // Get email list from server and diplay it in listview.
void 
Cpop3_imap4_simplevcNativeDlg::LoadServerMails( HTREEITEM hNode, IImap4Folder* fd )
{
	CWnd* pStatus = GetDlgItem(IDC_STATIC_STATUS);
	lstMail.DeleteAllItems();
	_ClearMailItems();

	try
	{
		IToolsPtr oTools;
		oTools.CreateInstance("EAGetMailObj.Tools");

		ConnectServer(hNode);
		CString localfolder = GetFolderByNode(hNode);
		_CreateFullFolder(localfolder);

		pStatus->SetWindowText(_T("Refreshing email(s) ..."));
		oClient->SelectFolder(fd);
		_variant_t infos = oClient->GetMailInfos();

		// UIDL is the identifier of every email on POP3/IMAP4/Exchange server, to avoid retrieve
		// the same email from server more than once, we record the email UIDL retrieved every time
		// UIDLManager wraps the function to write/read uidl record from a text file.
		// Load existed uidl records to UIDLManager
		CString uidlfile;
		uidlfile.Format( _T("%s\\uidl.txt"), (const TCHAR*)localfolder);
		oUIDLManager->Load((const TCHAR*)uidlfile);

		// Remove the local uidl which is not existed on the server.
		oUIDLManager->SyncUIDL(oCurServer, infos);
		oUIDLManager->Update();

		// Remove the email file on local disk that is not existed on server
		_ClearLocalMails( infos, localfolder);
		
		SAFEARRAY *psa = infos.parray;

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

			CString localfile = localfolder;
			localfile.Append( _T("\\"));
			localfile.Append( fileName );

			IUIDLItemPtr uidl_item = oUIDLManager->FindUIDL( oCurServer, pInfo->UIDL );
			if(uidl_item != NULL )
			{
				localfile.Format( _T("%s\\%s"), (const TCHAR*)localfolder, (const TCHAR*)uidl_item->FileName );
			}
			
			memset(szBuf, 0, sizeof(szBuf));
			::wsprintf( szBuf, _T("Checking %d/%d email ..."), i+1, count );
			pStatus->SetWindowText( szBuf );

			IMailPtr oMail;
			oMail.CreateInstance("EAGetMailObj.Mail");
			oMail->LicenseCode = _T("TryIt");

			if (oTools->ExistFile((const TCHAR*)localfile))
			{
				oMail->LoadFile((const TCHAR*)localfile, VARIANT_TRUE );
				// This mail has been downloaded from server.
			}
			else
			{
				// Get the mail header from server.
				oMail->Load(oClient->GetMailHeader(pInfo));
				oMail->SaveAs((const TCHAR*)localfile, VARIANT_TRUE );
			}

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
			lstMail.SetItemText( nIndex, 3, localfile.GetString());
			
			LPMAILINFOITEM pMailItem = new MAILINFOITEM;
			pMailItem->sortDate = (DWORD)(dt * 1000);
			pMailItem->oInfo = pInfo.GetInterfacePtr();
			pMailItem->oInfo->AddRef();

			lstMail.SetItemData( nIndex, (DWORD_PTR)pMailItem); 
			lstMail.Update(nIndex);

			m_arMailItem.Add( pMailItem ); // add to cache
			
			oMail.Release();

			if (uidl_item == NULL )
			{
				// Add the email UIDL and local file name to uidl file to avoid we retrieve it again. 
				oUIDLManager->AddUIDL(oCurServer, pInfo->UIDL, (const TCHAR*)fileName);
			}
		}

		lstMail.SortItems( MyCompareProc,  0 );
		// Update the uidl list to local uidl file and then we can load it next time.
		oUIDLManager->Update();

		CString tcount;
		tcount.Format( _T("Total %d email(s)"), count );
		pStatus->SetWindowText((const TCHAR*)tcount);
	}
	catch (_com_error &ep)
	{
		// Update the uidl list to local uidl file and then we can load it next time.
		oUIDLManager->Update();
		pStatus->SetWindowText((const TCHAR*)ep.Description());
		throw ep;
	}
}

// convert server path to local path
CString	
Cpop3_imap4_simplevcNativeDlg::GetFolderByNode( HTREEITEM hNode )
{
	CString folder = _T("");
	if( hNode == NULL )
		return folder;

	IImap4Folder* fd = (IImap4Folder*)trFolders.GetItemData( hNode );
	while( trFolders.GetParentItem( hNode ) != NULL )
	{
		hNode = trFolders.GetParentItem( hNode );
	}
	CString root = trFolders.GetItemText(hNode);

	folder.Format( _T("%s\\%s\\%s"), (const TCHAR*)m_curpath, (const TCHAR*)root, (const TCHAR*)fd->LocalPath);
	return folder;
}

// Create local folder
void 
Cpop3_imap4_simplevcNativeDlg::_CreateFullFolder( CString &folder )
{
	IToolsPtr oTools;
	oTools.CreateInstance("EAGetMailObj.Tools");
	if(oTools->ExistFile( (const TCHAR*)folder ) == VARIANT_TRUE )
	{
		return;
	}

	int pos = 0;
	while ((pos = folder.Find('\\', pos)) != -1)
	{
		if (pos > 2)
		{
			CString s = folder.Mid(0, pos);
			if (oTools->ExistFile((const TCHAR*)s) == VARIANT_FALSE )
			{
				oTools->CreateFolder((const TCHAR*)s);
			}
		}
		pos++;
	}

	if (oTools->ExistFile((const TCHAR*)folder) == VARIANT_FALSE )
	{
		oTools->CreateFolder((const TCHAR*)folder);
	}
}

 // Clear local file that is not existed on server.
void 
Cpop3_imap4_simplevcNativeDlg::_ClearLocalMails( _variant_t &infos, CString &folder )
{
	CString find = folder;
	find += _T("\\*.eml");
	IToolsPtr oTools;
	oTools.CreateInstance("EAGetMailObj.Tools");
	_variant_t files = oTools->GetFiles((const TCHAR*)find);

	SAFEARRAY *psa = files.parray;

	long LBound = 0, UBound = 0;
	SafeArrayGetLBound( psa, 1, &LBound );
	SafeArrayGetUBound( psa, 1, &UBound );

	INT count = UBound-LBound+1;
	for( long i = LBound; i <= UBound; i++ )
	{
		_variant_t vtFile;
		SafeArrayGetElement( psa, &i, &vtFile );
		CString s = vtFile.bstrVal;
		CString fileName = s;

		int pos = s.ReverseFind( _T('\\'));
		if (pos != -1)
			s = s.Mid(pos + 1);

		BOOL bfind = FALSE;

		IUIDLItemPtr item = oUIDLManager->FindLocalFile((const TCHAR*)s);
		if (item != NULL)
		{
			bfind = TRUE;
		}

		if (!bfind)
		{
			::DeleteFile((const TCHAR*)fileName);
			int p = fileName.ReverseFind(_T('.'));
			CString tempFolder = fileName.Mid(0, p);
			CString htmlName = tempFolder;
			htmlName += _T(".htm");

			if ( oTools->ExistFile((const TCHAR*)htmlName)== VARIANT_TRUE )
				::DeleteFile((const TCHAR*)htmlName);

			if (oTools->ExistFile((const TCHAR*)tempFolder)== VARIANT_TRUE)
			{
				oTools->RemoveFolder((const TCHAR*)tempFolder, VARIANT_TRUE );
			}
		}
	}	
}

// show mail list for specified folder.
void Cpop3_imap4_simplevcNativeDlg::OnTvnSelchangedTree1(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMTREEVIEW pNMTreeView = reinterpret_cast<LPNMTREEVIEW>(pNMHDR);
	// TODO: Add your control notification handler code here
	*pResult = 0;
	EnableIdle(FALSE);
	ShowNode();
	EnableIdle(TRUE);
}

void
Cpop3_imap4_simplevcNativeDlg::_DeleteSubNode( HTREEITEM hNode )
{
	// Delete all of the children of hmyItem.
	if (trFolders.ItemHasChildren(hNode))
	{
		HTREEITEM hNextItem;
		HTREEITEM hChildItem = trFolders.GetChildItem(hNode);

		while (hChildItem != NULL)
		{
			hNextItem = trFolders.GetNextItem(hChildItem, TVGN_NEXT);
			trFolders.DeleteItem(hChildItem);
			hChildItem = hNextItem;
		}
	}
}

// Refresh mails
void Cpop3_imap4_simplevcNativeDlg::OnMailmenuRefreshmails()
{

	HTREEITEM hNode = trFolders.GetSelectedItem();
	if (hNode == NULL)
		return;
	CWnd* pStatus = GetDlgItem(IDC_STATIC_STATUS);

	EnableIdle(FALSE);
	try
	{

		lstMail.DeleteAllItems();
		_ClearMailItems();

		ConnectServer(hNode);
		oClient->RefreshMailInfos();
		pStatus->SetWindowText( _T("Refreshing Mails ..."));

		ShowNode();
	}
	catch( _com_error &ep )
	{
		MessageBox((const TCHAR*)ep.Description(), _T("Error"), MB_OK );
		pStatus->SetWindowText((const TCHAR*)ep.Description());
		oClient->Close();
	}

	EnableIdle(TRUE);

}
// Add folder
void Cpop3_imap4_simplevcNativeDlg::OnMailmenuAddfolder()
{
	CAddFolderDlg *pDlg = new CAddFolderDlg();
	
	if( pDlg->DoModal() == IDOK )
	{
		EnableIdle(FALSE);
		try
		{
			CString folder = pDlg->folderName;
			HTREEITEM hNode = trFolders.GetSelectedItem();
			ConnectServer( hNode );
			IImap4Folder* pFolder = NULL;
			if( trFolders.GetParentItem( hNode ) != NULL )
			{
				pFolder = (IImap4Folder*)trFolders.GetItemData( hNode );
			}

			IImap4FolderPtr pNew = oClient->CreateFolder( pFolder, (const TCHAR*)folder );
			HTREEITEM hNew = trFolders.InsertItem( folder, hNode );
			
			trFolders.SetItemData( hNew , (DWORD_PTR)pNew.GetInterfacePtr());
			m_arFolder.Add(pNew.GetInterfacePtr());

			trFolders.Expand( hNode, TVE_EXPAND );
			pNew.Detach();

		}catch( _com_error &ep )
		{
			MessageBox((const TCHAR*)ep.Description(), _T("Error"), MB_OK );
			oClient->Close();
		}

		EnableIdle(TRUE);
	}

	delete pDlg;
}

// Delete folder
void Cpop3_imap4_simplevcNativeDlg::OnMailmenuDeletefolder()
{
	HTREEITEM hNode = trFolders.GetSelectedItem();
	if( hNode == NULL )
		return;

	if( trFolders.GetParentItem( hNode ) == NULL )
		return;

	EnableIdle(FALSE);
	try
	{
		ConnectServer( hNode );
		IImap4Folder* oFolder = (IImap4Folder*)trFolders.GetItemData( hNode );
		oClient->DeleteFolder( oFolder );

		trFolders.SelectItem( NULL );
		trFolders.DeleteItem( hNode );
		
		lstMail.DeleteAllItems();
		_ClearMailItems();


	}catch( _com_error &ep )
	{
		MessageBox((const TCHAR*)ep.Description(), _T("Error"), MB_OK );
		oClient->Close();
	}

	EnableIdle(TRUE);
}
// Rename folder
void Cpop3_imap4_simplevcNativeDlg::OnMailmenuRenamefolder()
{
	HTREEITEM hNode = trFolders.GetSelectedItem();
	if( hNode == NULL )
		return;

	if( trFolders.GetParentItem( hNode ) == NULL )
		return;

	//trFolders.SetFocus();
	trFolders.EditLabel( hNode );
}

void Cpop3_imap4_simplevcNativeDlg::OnTvnEndlabeleditTree1(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMTVDISPINFO pTVDispInfo = reinterpret_cast<LPNMTVDISPINFO>(pNMHDR);
	// TODO: Add your control notification handler code here
	*pResult = 0;
	
	EnableIdle(FALSE);
	try
	{
		HTREEITEM hNode = pTVDispInfo->item.hItem;
		if (hNode != NULL ) // rename folder
		{
			ConnectServer(hNode);
			CString cur_localpath = GetFolderByNode(hNode);
			IImap4Folder *pFolder = (IImap4Folder*)trFolders.GetItemData( hNode );

			oClient->RenameFolder(pFolder, pTVDispInfo->item.pszText );
			trFolders.EndEditLabelNow( FALSE );

			trFolders.SetItemText( hNode, pTVDispInfo->item.pszText );
			CString new_localpath = GetFolderByNode(hNode);

			// Try to rename local folder as well.
			::MoveFile( (const TCHAR*)cur_localpath, (const TCHAR*)new_localpath );

		}
	}
	catch( _com_error &ep )
	{
		trFolders.EndEditLabelNow( TRUE );
		MessageBox((const TCHAR*)ep.Description(), _T("Error"), MB_OK );
		oClient->Close();
	}

	EnableIdle(TRUE);

}
// Undelete deleted email on IMAP4 server
void Cpop3_imap4_simplevcNativeDlg::OnBnClickedButtonUndelete()
{
	if( lstMail.GetSelectedCount() == 0 )
		return;

	HTREEITEM hNode = trFolders.GetSelectedItem();
	if( hNode == NULL )
		return;

	const int MailServerPop3 = 0;
	const int MailServerImap4 = 1;
	const int MailServerEWS = 2;
	const int MailServerDAV = 3;	
	CWnd* pStatus = GetDlgItem(IDC_STATIC_STATUS);

	if (oCurServer->Protocol == MailServerEWS ||
		oCurServer->Protocol == MailServerDAV)
	{
		// EWS and WebDAV doesn't support this operating
		return;
	}

	EnableIdle( FALSE );
	try
	{
		INT nIndex = -1;
		ConnectServer(hNode);

		while((nIndex = lstMail.GetNextItem( nIndex, LVNI_SELECTED  )) != -1)
		{
			LPMAILINFOITEM pMailItem = (LPMAILINFOITEM)lstMail.GetItemData( nIndex );
			IMailInfo *info = pMailItem->oInfo;

			if( info->Deleted == VARIANT_TRUE )
			{
				oClient->Undelete( info );
				lstMail.Update( nIndex );
			}
		}
	}
	catch( _com_error &ep )
	{
		MessageBox((const TCHAR*)ep.Description(), _T("Error"), MB_OK );
		pStatus->SetWindowText((const TCHAR*)ep.Description());
		oClient->Close();
	}

	EnableIdle( TRUE );
}

// Mark email unread
void Cpop3_imap4_simplevcNativeDlg::OnBnClickedButtonUnread()
{
	if( lstMail.GetSelectedCount() == 0 )
		return;

	HTREEITEM hNode = trFolders.GetSelectedItem();
	if( hNode == NULL )
		return;

	const int MailServerPop3 = 0;
	const int MailServerImap4 = 1;
	const int MailServerEWS = 2;
	const int MailServerDAV = 3;	
	CWnd* pStatus = GetDlgItem(IDC_STATIC_STATUS);

	EnableIdle( FALSE );
	try
	{
		INT nIndex = -1;
		ConnectServer(hNode);

		while((nIndex = lstMail.GetNextItem( nIndex, LVNI_SELECTED  )) != -1)
		{
			LPMAILINFOITEM pMailItem = (LPMAILINFOITEM)lstMail.GetItemData( nIndex );
			IMailInfo *info = pMailItem->oInfo;
			if( info->Read == VARIANT_TRUE )
			{
				oClient->MarkAsRead( info, VARIANT_FALSE );
				lstMail.Update( nIndex );
			}
		}
	}
	catch( _com_error &ep )
	{
		MessageBox((const TCHAR*)ep.Description(), _T("Error"), MB_OK );
		pStatus->SetWindowText((const TCHAR*)ep.Description());
		oClient->Close();
	}

	EnableIdle( TRUE );
}
// Pure deleted emails
void Cpop3_imap4_simplevcNativeDlg::OnBnClickedButtonPure()
{
	HTREEITEM hNode = trFolders.GetSelectedItem();
	if( hNode == NULL )
		return;

	const int MailServerPop3 = 0;
	const int MailServerImap4 = 1;
	const int MailServerEWS = 2;
	const int MailServerDAV = 3;	
	CWnd* pStatus = GetDlgItem(IDC_STATIC_STATUS);

	if (oCurServer->Protocol == MailServerEWS ||
		oCurServer->Protocol == MailServerDAV)
	{
		// EWS and WebDAV doesn't support this operating
		return;
	}

	if( trFolders.GetParentItem( hNode ) == NULL )
	{
		return;
	}

	EnableIdle( FALSE );
	try
	{
		
		ConnectServer(hNode);
		IImap4Folder* fd = (IImap4Folder*)trFolders.GetItemData( hNode );

		oClient->Expunge();
		LoadServerMails( hNode, fd );
		
	}
	catch( _com_error &ep )
	{
		MessageBox((const TCHAR*)ep.Description(), _T("Error"), MB_OK );
		pStatus->SetWindowText((const TCHAR*)ep.Description());
		oClient->Close();
	}
	EnableIdle( TRUE );
}

// Copy emails
void Cpop3_imap4_simplevcNativeDlg::OnBnClickedButtonCopy()
{
	if( lstMail.GetSelectedCount() == 0 )
		return;

	HTREEITEM hNode = trFolders.GetSelectedItem();
	if( hNode == NULL )
		return;

	const int MailServerPop3 = 0;
	const int MailServerImap4 = 1;
	const int MailServerEWS = 2;
	const int MailServerDAV = 3;	
	CWnd* pStatus = GetDlgItem(IDC_STATIC_STATUS);


	EnableIdle( FALSE );
	try
	{
		INT nIndex = -1;
		ConnectServer(hNode);
		IImap4Folder* fd = (IImap4Folder*)trFolders.GetItemData( hNode );
		
		CFolderDlg *pDlg = new CFolderDlg();
		pDlg->folders = oClient->Imap4Folders;
		
		if( pDlg->DoModal() != IDOK )
		{
			delete pDlg;
			EnableIdle( TRUE );
			return;
		}
		
		IImap4Folder* dest = pDlg->m_oFolder;
		dest->AddRef();
		delete pDlg;

		if( _tcsicmp((const TCHAR*)fd->FullPath, (const TCHAR*)dest->FullPath ) == 0 )
		{
			dest->Release();
			EnableIdle( TRUE );
			return;
		}

		while((nIndex = lstMail.GetNextItem( nIndex, LVNI_SELECTED  )) != -1)
		{
			LPMAILINFOITEM pMailItem = (LPMAILINFOITEM)lstMail.GetItemData( nIndex );
			IMailInfo *info = pMailItem->oInfo;
			oClient->Copy( info, dest );
		}

		dest->Release();
	}
	catch( _com_error &ep )
	{
		MessageBox((const TCHAR*)ep.Description(), _T("Error"), MB_OK );
		pStatus->SetWindowText((const TCHAR*)ep.Description());
		oClient->Close();
	}

	EnableIdle( TRUE );
}

// Move email
void Cpop3_imap4_simplevcNativeDlg::OnBnClickedButtonMove()
{
	if( lstMail.GetSelectedCount() == 0 )
		return;

	HTREEITEM hNode = trFolders.GetSelectedItem();
	if( hNode == NULL )
		return;

	const int MailServerPop3 = 0;
	const int MailServerImap4 = 1;
	const int MailServerEWS = 2;
	const int MailServerDAV = 3;	
	CWnd* pStatus = GetDlgItem(IDC_STATIC_STATUS);

	EnableIdle( FALSE );
	try
	{
		INT nIndex = -1;
		ConnectServer(hNode);
		IImap4Folder* fd = (IImap4Folder*)trFolders.GetItemData( hNode );
		
		CFolderDlg *pDlg = new CFolderDlg();
		pDlg->folders = oClient->Imap4Folders;
		
		if( pDlg->DoModal() != IDOK )
		{
			delete pDlg;
			EnableIdle(TRUE);
			return;
		}
		
		IImap4Folder* dest = pDlg->m_oFolder;
		dest->AddRef();
		delete pDlg;

		if( _tcsicmp((const TCHAR*)fd->FullPath, (const TCHAR*)dest->FullPath ) == 0 )
		{
			dest->Release();
			EnableIdle(TRUE);
			return;
		}

		while((nIndex = lstMail.GetNextItem( nIndex, LVNI_SELECTED  )) != -1)
		{
			LPMAILINFOITEM pMailItem = (LPMAILINFOITEM)lstMail.GetItemData( nIndex );
			IMailInfo *info = pMailItem->oInfo;
			oClient->Move( info, dest );
			if (oCurServer->Protocol == MailServerEWS ||
				oCurServer->Protocol == MailServerDAV)
			{
				lstMail.DeleteItem( nIndex );
				nIndex = -1;
			}
			else
			{
				lstMail.Update( nIndex );
			}
		}

		dest->Release();
	}
	catch( _com_error &ep )
	{
		MessageBox((const TCHAR*)ep.Description(), _T("Error"), MB_OK );
		pStatus->SetWindowText((const TCHAR*)ep.Description());
		oClient->Close();
	}

	EnableIdle( TRUE );
}
// Upload EML file to folder
void Cpop3_imap4_simplevcNativeDlg::OnBnClickedButtonUpload()
{

	CWnd* pStatus = GetDlgItem(IDC_STATIC_STATUS);

	HTREEITEM hNode = trFolders.GetSelectedItem();
	if (hNode == NULL)
		return;

	if (trFolders.GetParentItem( hNode ) == NULL )
		return;

	CFileDialog *pFileDlg = new CFileDialog( TRUE, _T("*.EML"), NULL, 4|2, _T("Email File (*.EML)|*.EML||"));
	pFileDlg->m_ofn.Flags;
	pFileDlg->m_ofn.Flags |= OFN_FILEMUSTEXIST;
	
	if( pFileDlg->DoModal() != IDOK )
		return;

	EnableIdle( FALSE );
	try
	{
		CString fileName = pFileDlg->GetFileName();
		CString path = pFileDlg->GetFolderPath();
		CString fullpath;
		fullpath.Format( _T("%s\\%s"), path, fileName );
		IMailPtr oMail;
		oMail.CreateInstance( "EAGetMailObj.Mail");
		oMail->LicenseCode = _T("TryIt");
		oMail->LoadFile( (const TCHAR*)fullpath, VARIANT_FALSE );

		ConnectServer(hNode);
		IImap4Folder* fd = (IImap4Folder*)trFolders.GetItemData( hNode );
		oClient->Append( fd, oMail->Content );
		oClient->RefreshMailInfos();
		ShowNode();
	}
	catch( _com_error &ep )
	{
		MessageBox((const TCHAR*)ep.Description(), _T("Error"), MB_OK );
		pStatus->SetWindowText((const TCHAR*)ep.Description());
		oClient->Close();
	}
	EnableIdle( TRUE );
}

// Enable/disable control 
void Cpop3_imap4_simplevcNativeDlg::EnableIdle(BOOL bIdle) 
{
	CWnd* pBtnDelete = GetDlgItem(IDC_BUTTON_DEL);
	CWnd* pBtnUndelete = GetDlgItem(IDC_BUTTON_UNDELETE);
	CWnd* pBtnUnread = GetDlgItem(IDC_BUTTON_UNREAD);
	CWnd* pBtnPure = GetDlgItem(IDC_BUTTON_PURE);
	CWnd* pBtnMove = GetDlgItem(IDC_BUTTON_MOVE);
	CWnd* pBtnCopy = GetDlgItem(IDC_BUTTON_COPY);
	CWnd* pBtnUpload = GetDlgItem(IDC_BUTTON_UPLOAD);
	

	CWnd* pBtnCancel = GetDlgItem(IDC_BUTTON_CANCEL);
	CWnd* pBtnQuit = GetDlgItem(IDC_BUTTON_QUIT);
	CWnd* pBtnStart = GetDlgItem(IDC_BUTTON_START);

	pBtnDelete->EnableWindow( bIdle );
	pBtnUndelete->EnableWindow( bIdle );
	pBtnUnread->EnableWindow( bIdle );
	pBtnPure->EnableWindow( bIdle );
	pBtnMove->EnableWindow( bIdle );
	pBtnCopy->EnableWindow( bIdle );
	pBtnUpload->EnableWindow( bIdle );

	if( lstMail.GetSelectedCount() == 0 )
	{
		pBtnDelete->EnableWindow( FALSE );
		pBtnUndelete->EnableWindow( FALSE );
		pBtnUnread->EnableWindow( FALSE );
		pBtnCopy->EnableWindow( FALSE );
		pBtnMove->EnableWindow( FALSE );
		//pBtnPure->EnableWindow( FALSE );
	}

	HTREEITEM hNode = trFolders.GetSelectedItem();
	if( hNode == NULL )
	{
		pBtnUpload->EnableWindow( FALSE );
	}
	else
	{
		if( trFolders.GetParentItem( hNode ) == NULL )
		{
			pBtnUpload->EnableWindow( FALSE );
		}
	}

    pBtnCancel->EnableWindow(!bIdle);

	if( pBtnStart->IsWindowEnabled())
		pBtnCancel->EnableWindow( FALSE );

	pBtnQuit->EnableWindow(bIdle);
	trFolders.EnableWindow( bIdle );
	lstMail.EnableWindow( bIdle );

	const int MailServerPop3 = 0;
	const int MailServerImap4 = 1;
	const int MailServerEWS = 2;
	const int MailServerDAV = 3;

	if (oCurServer != NULL )
	{
		if (oCurServer->Protocol == MailServerEWS ||
			oCurServer->Protocol == MailServerDAV )
		{// Exchange WebDAV and EWS doesn't support this operating
			pBtnUndelete->EnableWindow( FALSE );
			pBtnPure->EnableWindow(FALSE );
		}
	}
}

// Disconnect connections
void Cpop3_imap4_simplevcNativeDlg::OnBnClickedButtonQuit()
{
	CWnd* pStatus = GetDlgItem(IDC_STATIC_STATUS);
	try
	{
		CWnd* pBtnStart = GetDlgItem(IDC_BUTTON_START);

		pBtnStart->EnableWindow(TRUE);
		trFolders.DeleteAllItems();
		_ClearFolders();

		lstMail.DeleteAllItems();
		_ClearMailItems();

		webMail.Navigate( _T("about:blank"), NULL, NULL, NULL, NULL );

		oClient->Logout();
		oClient->Close();

		pStatus->SetWindowText(_T("Discconnected"));
		
	}
	catch( ... )
	{
	}

	EnableIdle(FALSE);
}

// Set font based on Mail flags.
void Cpop3_imap4_simplevcNativeDlg::OnNMCustomdrawListMail(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMCUSTOMDRAW pNMCD = reinterpret_cast<LPNMCUSTOMDRAW>(pNMHDR);
	// TODO: Add your control notification handler code here
	*pResult = 0;

	NMLVCUSTOMDRAW* customDrawItem =  reinterpret_cast<NMLVCUSTOMDRAW*>(pNMHDR); 
	// Initialize to default processing: 
	*pResult = CDRF_DODEFAULT; 
	if( CDDS_PREPAINT  == pNMCD->dwDrawStage )
	{
		*pResult = CDRF_NOTIFYITEMDRAW;
	}
	else if (CDDS_ITEMPREPAINT == pNMCD->dwDrawStage) 
	{ 
		LPMAILINFOITEM pMailItem =  (LPMAILINFOITEM)pNMCD->lItemlParam;
		if( pMailItem == NULL )
			return;

		HFONT realFont = HFONT(::SelectObject(pNMCD->hdc, HFONT(::GetStockObject(SYSTEM_FONT)))); 
		LOGFONT logFont; 
		::GetObject(realFont, sizeof(LOGFONT), &logFont);


		if( pMailItem->oInfo->Read == VARIANT_TRUE )
		{
			logFont.lfWeight = FW_NORMAL;
		}
		else
		{
			logFont.lfWeight = FW_BOLD;
		}

		if( pMailItem->oInfo->Deleted == VARIANT_TRUE )	
		{
			logFont.lfStrikeOut = TRUE;
		}
		else
		{
			logFont.lfStrikeOut = FALSE;
		}

		HFONT newFont = ::CreateFontIndirect(&logFont); 
		::SelectObject(pNMCD->hdc, newFont);
		::DeleteObject(newFont);

		*pResult = CDRF_NEWFONT; 
	} 

}

// Clear all cached LPMAILINFOITEM object
void 
Cpop3_imap4_simplevcNativeDlg::_ClearMailItems()
{
	INT count = m_arMailItem.GetSize();
	for( int i = 0; i < count; i++ )
	{
		LPMAILINFOITEM pItem = m_arMailItem[i];
		pItem->oInfo->Release();
		delete pItem;
	}
	m_arMailItem.RemoveAll();
}
// Clear all cached IImap4Folder object
void 
Cpop3_imap4_simplevcNativeDlg::_ClearFolders()
{
	INT count = m_arFolder.GetSize();
	for( int i = 0; i < count; i++ )
	{
		IImap4Folder* fd = m_arFolder[i];
		fd->Release();
	}
	m_arFolder.RemoveAll();
}	
void Cpop3_imap4_simplevcNativeDlg::OnClose()
{
	// TODO: Add your message handler code here and/or call default
	_ClearMailItems();
	_ClearFolders();
	__super::OnClose();
}
