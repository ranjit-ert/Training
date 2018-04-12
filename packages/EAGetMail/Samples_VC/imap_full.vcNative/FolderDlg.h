#pragma once
#include "afxcmn.h"


// CFolderDlg dialog

class CFolderDlg : public CDialog
{
	DECLARE_DYNAMIC(CFolderDlg)

public:
	CFolderDlg(CWnd* pParent = NULL);   // standard constructor
	virtual ~CFolderDlg();

// Dialog Data
	enum { IDD = IDD_DIALOG1 };

	void ExpandFolders(_variant_t &fds, HTREEITEM hNode );

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	CTreeCtrl trFolders;
	_variant_t folders;
	IImap4Folder* m_oFolder;
	virtual BOOL OnInitDialog();
	afx_msg void OnBnClickedOk();
};
