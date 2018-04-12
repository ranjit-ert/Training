unit Unit2;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls;

type
  TForm2 = class(TForm)
    Label1: TLabel;
    textFolder: TEdit;
    btnOK: TButton;
    btnCancel: TButton;
    procedure btnCancelClick(Sender: TObject);
    procedure btnOKClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form2: TForm2;

implementation

{$R *.dfm}

procedure TForm2.btnCancelClick(Sender: TObject);
begin
  
  ModalResult := mrCancel;
end;

procedure TForm2.btnOKClick(Sender: TObject);
begin
  if textFolder.Text = '' then
  begin
    ShowMessage( 'Please input folder name!' );
    exit;
  end;

  ModalResult := mrOK;
end;

end.
