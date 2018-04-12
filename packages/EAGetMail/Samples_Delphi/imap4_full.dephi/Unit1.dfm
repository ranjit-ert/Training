object Form1: TForm1
  Left = 2
  Top = 219
  Width = 949
  Height = 500
  Caption = 'Full Example for IMAP4 and Exchange Web Service/WebDAV'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  OnResize = FormResize
  PixelsPerInch = 96
  TextHeight = 13
  object lblStatus: TLabel
    Left = 8
    Top = 424
    Width = 31
    Height = 13
    Caption = 'Ready'
  end
  object GroupBox1: TGroupBox
    Left = 8
    Top = 8
    Width = 265
    Height = 353
    Caption = 'Accout Information'
    TabOrder = 0
    object Label1: TLabel
      Left = 8
      Top = 26
      Width = 31
      Height = 13
      Caption = 'Server'
    end
    object Label2: TLabel
      Left = 8
      Top = 54
      Width = 22
      Height = 13
      Caption = 'User'
    end
    object Label3: TLabel
      Left = 8
      Top = 79
      Width = 46
      Height = 13
      Caption = 'Password'
    end
    object Label4: TLabel
      Left = 8
      Top = 147
      Width = 49
      Height = 13
      Caption = 'Auth Type'
    end
    object Label5: TLabel
      Left = 8
      Top = 173
      Width = 39
      Height = 13
      Caption = 'Protocol'
    end
    object textServer: TEdit
      Left = 74
      Top = 23
      Width = 177
      Height = 21
      TabOrder = 0
    end
    object textUser: TEdit
      Left = 74
      Top = 49
      Width = 177
      Height = 21
      TabOrder = 1
    end
    object textPassword: TEdit
      Left = 74
      Top = 76
      Width = 177
      Height = 21
      PasswordChar = '*'
      TabOrder = 2
    end
    object chkSSL: TCheckBox
      Left = 8
      Top = 106
      Width = 241
      Height = 25
      Caption = 'My server requires SSL connection'
      TabOrder = 3
    end
    object lstAuthType: TComboBox
      Left = 75
      Top = 142
      Width = 177
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 4
    end
    object lstProtocol: TComboBox
      Left = 75
      Top = 169
      Width = 178
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 5
      OnChange = lstProtocolChange
    end
    object btnStart: TButton
      Left = 8
      Top = 248
      Width = 241
      Height = 25
      Caption = 'Start'
      TabOrder = 6
      OnClick = btnStartClick
    end
    object btnCancel: TButton
      Left = 8
      Top = 312
      Width = 241
      Height = 25
      Caption = 'Cancel'
      Enabled = False
      TabOrder = 7
      OnClick = btnCancelClick
    end
    object btnQuit: TButton
      Left = 8
      Top = 279
      Width = 241
      Height = 25
      Caption = 'Quit'
      Enabled = False
      TabOrder = 8
      OnClick = btnQuitClick
    end
  end
  object lstMail: TListView
    Left = 448
    Top = 13
    Width = 481
    Height = 153
    Columns = <
      item
        Caption = 'From'
        Width = 150
      end
      item
        Caption = 'Subject'
        Width = 190
      end
      item
        Caption = 'Date'
        Width = 150
      end
      item
        Width = 0
      end>
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    GridLines = True
    HideSelection = False
    MultiSelect = True
    RowSelect = True
    ParentFont = False
    SortType = stText
    TabOrder = 1
    ViewStyle = vsReport
    OnAdvancedCustomDrawItem = lstMailAdvancedCustomDrawItem
    OnCompare = lstMailCompare
    OnSelectItem = lstMailSelectItem
  end
  object btnDel: TButton
    Left = 448
    Top = 169
    Width = 64
    Height = 25
    Caption = 'Delete'
    Enabled = False
    TabOrder = 2
    OnClick = btnDelClick
  end
  object webMail: TWebBrowser
    Left = 448
    Top = 200
    Width = 481
    Height = 233
    TabOrder = 3
    ControlData = {
      4C000000B6310000151800000000000000000000000000000000000000000000
      000000004C000000000000000000000001000000E0D057007335CF11AE690800
      2B2E126208000000000000004C0000000114020000000000C000000000000046
      8000000000000000000000000000000000000000000000000000000000000000
      00000000000000000100000000000000000000000000000000000000}
  end
  object trFolders: TTreeView
    Left = 280
    Top = 14
    Width = 161
    Height = 435
    HideSelection = False
    Indent = 19
    PopupMenu = PopupMenu1
    TabOrder = 4
    OnChange = trFoldersChange
    OnEdited = trFoldersEdited
  end
  object pgBar: TProgressBar
    Left = 448
    Top = 440
    Width = 481
    Height = 9
    TabOrder = 5
  end
  object btnUndelete: TButton
    Left = 516
    Top = 169
    Width = 63
    Height = 25
    Caption = 'Undelete'
    Enabled = False
    TabOrder = 6
    OnClick = btnUndeleteClick
  end
  object btnUnread: TButton
    Left = 583
    Top = 169
    Width = 64
    Height = 25
    Caption = 'Unread'
    Enabled = False
    TabOrder = 7
    OnClick = btnUnreadClick
  end
  object btnPure: TButton
    Left = 651
    Top = 169
    Width = 64
    Height = 25
    Caption = 'Pure'
    Enabled = False
    TabOrder = 8
    OnClick = btnPureClick
  end
  object btnCopy: TButton
    Left = 719
    Top = 169
    Width = 64
    Height = 25
    Caption = 'Copy'
    Enabled = False
    TabOrder = 9
    OnClick = btnCopyClick
  end
  object btnMove: TButton
    Left = 787
    Top = 169
    Width = 64
    Height = 25
    Caption = 'Move'
    Enabled = False
    TabOrder = 10
    OnClick = btnMoveClick
  end
  object btnUpload: TButton
    Left = 855
    Top = 169
    Width = 64
    Height = 25
    Caption = 'Upload'
    Enabled = False
    TabOrder = 11
    OnClick = btnUploadClick
  end
  object PopupMenu1: TPopupMenu
    Left = 120
    Top = 400
    object RefreshFolders1: TMenuItem
      Caption = 'Refresh Folders'
      OnClick = RefreshFolders1Click
    end
    object RefreshMails1: TMenuItem
      Caption = 'Refresh Mails'
      OnClick = RefreshMails1Click
    end
    object AddFolder1: TMenuItem
      Caption = 'Add Folder'
      OnClick = AddFolder1Click
    end
    object DeleteFolder1: TMenuItem
      Caption = 'Delete Folder'
      OnClick = DeleteFolder1Click
    end
    object RenameFolder1: TMenuItem
      Caption = 'Rename Folder'
      OnClick = RenameFolder1Click
    end
  end
end
