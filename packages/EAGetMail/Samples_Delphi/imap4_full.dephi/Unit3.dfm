object Form3: TForm3
  Left = 298
  Top = 213
  BorderStyle = bsDialog
  Caption = 'Select Folder'
  ClientHeight = 449
  ClientWidth = 321
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object trFolders: TTreeView
    Left = 8
    Top = 8
    Width = 305
    Height = 393
    HideSelection = False
    Indent = 19
    ReadOnly = True
    TabOrder = 0
  end
  object btnOK: TButton
    Left = 160
    Top = 416
    Width = 73
    Height = 25
    Caption = 'OK'
    Default = True
    TabOrder = 1
    OnClick = btnOKClick
  end
  object btnCancel: TButton
    Left = 240
    Top = 416
    Width = 73
    Height = 25
    Caption = 'Cancel'
    TabOrder = 2
    OnClick = btnCancelClick
  end
end
