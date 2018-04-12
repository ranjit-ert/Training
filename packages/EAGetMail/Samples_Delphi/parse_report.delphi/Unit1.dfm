object Form1: TForm1
  Left = 204
  Top = 192
  Width = 786
  Height = 484
  Caption = 'Retrieve and Parse Report'
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
  object Label6: TLabel
    Left = 8
    Top = 328
    Width = 265
    Height = 49
    AutoSize = False
    Caption = 
      'Warning: if "leave a copy of message on server" is not checked, ' +
      ' the emails on the server will be deleted !'
    Color = clBtnFace
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clNavy
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentColor = False
    ParentFont = False
    WordWrap = True
  end
  object lblTotal: TLabel
    Left = 280
    Top = 176
    Width = 24
    Height = 13
    Caption = 'Total'
  end
  object GroupBox1: TGroupBox
    Left = 8
    Top = 8
    Width = 265
    Height = 313
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
      Top = 138
      Width = 49
      Height = 13
      Caption = 'Auth Type'
    end
    object Label5: TLabel
      Left = 8
      Top = 164
      Width = 39
      Height = 13
      Caption = 'Protocol'
    end
    object lblStatus: TLabel
      Left = 8
      Top = 250
      Width = 31
      Height = 13
      Caption = 'Ready'
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
      Top = 104
      Width = 241
      Height = 25
      Caption = 'My server requires SSL connection'
      TabOrder = 3
    end
    object lstAuthType: TComboBox
      Left = 75
      Top = 133
      Width = 177
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 4
    end
    object lstProtocol: TComboBox
      Left = 75
      Top = 160
      Width = 177
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 5
      OnChange = lstProtocolChange
    end
    object chkLeaveCopy: TCheckBox
      Left = 8
      Top = 192
      Width = 249
      Height = 17
      Caption = 'Leave a copy of message on server'
      TabOrder = 6
    end
    object btnStart: TButton
      Left = 8
      Top = 216
      Width = 113
      Height = 25
      Caption = 'Start'
      TabOrder = 7
      OnClick = btnStartClick
    end
    object btnCancel: TButton
      Left = 137
      Top = 216
      Width = 113
      Height = 25
      Caption = 'Cancel'
      Enabled = False
      TabOrder = 8
      OnClick = btnCancelClick
    end
    object pgBar: TProgressBar
      Left = 8
      Top = 289
      Width = 243
      Height = 17
      TabOrder = 9
    end
  end
  object lstMail: TListView
    Left = 280
    Top = 13
    Width = 489
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
    RowSelect = True
    ParentFont = False
    SortType = stText
    TabOrder = 1
    ViewStyle = vsReport
    OnCompare = lstMailCompare
    OnSelectItem = lstMailSelectItem
  end
  object btnDel: TButton
    Left = 704
    Top = 171
    Width = 65
    Height = 22
    Caption = 'Delete'
    TabOrder = 2
    OnClick = btnDelClick
  end
  object textReport: TMemo
    Left = 280
    Top = 200
    Width = 489
    Height = 241
    ScrollBars = ssBoth
    TabOrder = 3
  end
end
