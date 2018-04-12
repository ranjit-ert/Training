// FolderDlg.cpp : implementation file
//

#include "stdafx.h"
#include "pop3_imap4_simple.vcNative.h"
#include "FolderDlg.h"


// CFolderDlg dialog

IMPLEMENT_DYNAMIC(CFolderDlg, CDialog)

CFolderDlg::CFolderDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CFolderDlg::IDD, pParent)
{

}

CFolderDlg::~CFolderDlg()
{
}

void CFolderDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_TREE1, trFolders);
}


BEGIN_MESSAGE_MAP(CFolderDlg, CDialog)
	ON_BN_CLICKED(IDOK, &CFolderDlg::OnBnClickedOk)
END_MESSAGE_MAP()


// CFolderDlg message handlers
void CFolderDlg::ExpandFolders(_variant_t &fds, HTREEITEM hNode )
{
	//_DeleteSubNode( hNode );
	SAFEARRAY *psa = fds.parray;
	long LBound = 0, UBound = 0;
	SafeArrayGetLBound( psa, 1, &LBound );
	SafeArrayGetUBound( psa, 1, &UBound );

	INT count = UBound-LBound+1;
	for( long i = LBound; i <= UBound; i++ )
	{
		_variant_t vtFolder;
		SafeArrayGetElement( psa, &i, &vtFolder );

		IImap4FolderPtr pFolder = NULL;
		vtFolder.pdispVal->QueryInterface(__uuidof(IImap4Folder), (void**)&pFolder);
		
		HTREEITEM hSub = trFolders.InsertItem( pFolder->Name, hNode );
		trFolders.SetItemData( hSub, (DWORD_PTR)pFolder.GetInterfacePtr());

		ExpandFolders( pFolder->SubFolders, hSub );

		trFolders.Expand( hSub, TVE_EXPAND );
		pFolder.Detach();
	}
}
BOOL CFolderDlg::OnInitDialog()
{
	CDialog::OnInitDialog();
	HTREEITEM hRoot = trFolders.InsertItem( _T("Root Folder"));
	// TODO:  Add extra initialization here
	ExpandFolders( folders, hRoot );
	trFolders.Expand( hRoot, TVE_EXPAND );
	return TRUE;  // return TRUE unless you set the focus to a control
	// EXCEPTION: OCX Property Pages should return FALSE
}

void CFolderDlg::OnBnClickedOk()
{
	HTREEITEM hNode = trFolders.GetSelectedItem();
	if( hNode == NULL )
	{
		MessageBox( _T("Please select a folder!"));
		return;
	}
	
	if( trFolders.GetParentItem( hNode ) == NULL )
	{
		MessageBox( _T("Please select a folder!"));
		return;
	}

	m_oFolder = (IImap4Folder*)trFolders.GetItemData( hNode );
	
	OnOK();
}
