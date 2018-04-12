object Form2: TForm2
  Left = 199
  Top = 260
  BorderStyle = bsDialog
  Caption = 'Add Folder'
  ClientHeight = 74
  ClientWidth = 449
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 8
    Top = 16
    Width = 63
    Height = 13
    Caption = 'Folder Name:'
  end
  object textFolder: TEdit
    Left = 88
    Top = 8
    Width = 353
    Height = 21
    TabOrder = 0
  end
  object btnOK: TButton
    Left = 248
    Top = 40
    Width = 97
    Height = 25
    Caption = 'OK'
    TabOrder = 1
    OnClick = btnOKClick
  end
  object btnCancel: TButton
    Left = 352
    Top = 40
    Width = 89
    Height = 25
    Caption = 'Cancel'
    TabOrder = 2
    OnClick = btnCancelClick
  end
end
