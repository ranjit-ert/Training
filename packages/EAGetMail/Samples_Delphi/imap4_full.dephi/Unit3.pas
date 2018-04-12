unit Unit3;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls;

type
  TForm3 = class(TForm)
    trFolders: TTreeView;
    btnOK: TButton;
    btnCancel: TButton;
    procedure btnOKClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form3: TForm3;

implementation

{$R *.dfm}

procedure TForm3.btnOKClick(Sender: TObject);
var
  oNode: TTreeNode;
begin
  oNode := trFolders.Selected;
  if oNode = nil then
  begin
    ShowMessage( 'Please select a folder!');
    exit;
  end;

  if oNode.Parent = nil then
  begin
    ShowMessage( 'Please select a folder!');
    exit;
  end;

  ModalResult := mrOK;
end;

procedure TForm3.btnCancelClick(Sender: TObject);
begin
  ModalResult := mrCancel;
end;

end.
