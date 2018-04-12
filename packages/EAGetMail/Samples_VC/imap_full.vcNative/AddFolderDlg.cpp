// AddFolderDlg.cpp : implementation file
//

#include "stdafx.h"
#include "pop3_imap4_simple.vcNative.h"
#include "AddFolderDlg.h"


// CAddFolderDlg dialog

IMPLEMENT_DYNAMIC(CAddFolderDlg, CDialog)

CAddFolderDlg::CAddFolderDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CAddFolderDlg::IDD, pParent)
	, folderName(_T(""))
{

}

CAddFolderDlg::~CAddFolderDlg()
{
}

void CAddFolderDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_EDIT1, folderName);
	DDV_MaxChars(pDX, folderName, 200);
}


BEGIN_MESSAGE_MAP(CAddFolderDlg, CDialog)
	ON_BN_CLICKED(IDOK, &CAddFolderDlg::OnBnClickedOk)
END_MESSAGE_MAP()


// CAddFolderDlg message handlers

void CAddFolderDlg::OnBnClickedOk()
{
	UpdateData( TRUE );
	if( folderName.GetLength() == 0 )
	{
		MessageBox( _T("Please input folder name!"));
		return;
	}
	OnOK();
}
