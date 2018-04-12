unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, OleCtrls, SHDocVw_TLB, ComCtrls, StdCtrls, StrUtils, EAGetMailObjLib_TLB,
  SHDocVw, Menus, Unit2, Unit3;

type
  TForm1 = class(TForm)
    GroupBox1: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    textServer: TEdit;
    textUser: TEdit;
    textPassword: TEdit;
    chkSSL: TCheckBox;
    Label4: TLabel;
    Label5: TLabel;
    lstAuthType: TComboBox;
    lstProtocol: TComboBox;
    btnStart: TButton;
    btnCancel: TButton;
    lstMail: TListView;
    btnDel: TButton;
    lblStatus: TLabel;
    webMail: TWebBrowser;
    trFolders: TTreeView;
    pgBar: TProgressBar;
    btnQuit: TButton;
    btnUndelete: TButton;
    btnUnread: TButton;
    btnPure: TButton;
    btnCopy: TButton;
    btnMove: TButton;
    btnUpload: TButton;
    PopupMenu1: TPopupMenu;
    RefreshFolders1: TMenuItem;
    RefreshMails1: TMenuItem;
    AddFolder1: TMenuItem;
    DeleteFolder1: TMenuItem;
    RenameFolder1: TMenuItem;
    procedure FormCreate(Sender: TObject);
    procedure btnStartClick(Sender: TObject);

    function WideStringToString(const ws: WideString): AnsiString;
    function StringToWideString(const s: AnsiString): WideString;

    procedure EnableButton( enabled: Bool);
    procedure ConnectServer();

    // Folder management
    procedure ShowFolders();
    procedure ExpandFolders( oFolders: OleVariant; oNode: TTreeNode);
    procedure ExpandFoldersEx( oFolders: OleVariant; oNode: TTreeNode);
    procedure CreateFullFolder( folder: WideString );

    // Display email functions
    procedure LoadMails();
    procedure LoadServerMails( oNode: TTreeNode; oFolder: IImap4Folder);
    procedure ClearLocalMails( infos: OleVariant; folder: WideString);
    procedure ShowMail( fileName: WideString );
    procedure GenerateHtmlForEmail(htmlName: WideString; emlFile: WideString; tempFolder: WideString);


    // EAGetMail event handler
    procedure OnIdle(ASender: TObject; const oSender: IDispatch; var Cancel: WordBool);
    procedure OnConnected(ASender: TObject; const oSender: IDispatch; var Cancel: WordBool);
    procedure OnQuit(ASender: TObject; const oSender: IDispatch; var Cancel: WordBool);
    procedure OnSecuring(ASender: TObject; const oSender: IDispatch; var Cancel: WordBool);
    procedure OnAuthorized(ASender: TObject; const oSender: IDispatch;
      var Cancel: WordBool);
    procedure OnSendingDataStream(ASender: TObject; const oSender: IDispatch;
      Sent: Integer; Total: Integer;
      var Cancel: WordBool);
    procedure OnReceivingDataStream(ASender: TObject; const oSender: IDispatch;
      const oInfo: IDispatch;
      Received: Integer; Total: Integer;
      var Cancel: WordBool);
    procedure btnCancelClick(Sender: TObject);
    procedure lstMailSelectItem(Sender: TObject; Item: TListItem;
      Selected: Boolean);
    procedure btnDelClick(Sender: TObject);
    procedure lstMailCompare(Sender: TObject; Item1, Item2: TListItem;
      Data: Integer; var Compare: Integer);
    procedure FormResize(Sender: TObject);
    procedure lstProtocolChange(Sender: TObject);
    procedure trFoldersChange(Sender: TObject; Node: TTreeNode);
    procedure RefreshFolders1Click(Sender: TObject);
    procedure RefreshMails1Click(Sender: TObject);
    procedure AddFolder1Click(Sender: TObject);
    procedure DeleteFolder1Click(Sender: TObject);
    procedure RenameFolder1Click(Sender: TObject);
    procedure trFoldersEdited(Sender: TObject; Node: TTreeNode;
      var S: String);
    procedure btnQuitClick(Sender: TObject);
    procedure lstMailAdvancedCustomDrawItem(Sender: TCustomListView;
      Item: TListItem; State: TCustomDrawState; Stage: TCustomDrawStage;
      var DefaultDraw: Boolean);
    procedure btnUndeleteClick(Sender: TObject);
    procedure btnUnreadClick(Sender: TObject);
    procedure btnPureClick(Sender: TObject);
    procedure btnCopyClick(Sender: TObject);
    procedure btnMoveClick(Sender: TObject);
    procedure btnUploadClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;
const
  MailServerConnectType_ConnectSSLAuto = 0;
  MailServerConnectType_ConnectSSL = 1;
  MailServerConnectType_ConnectTLS = 2;

  ProxyType_Socks4 = 0;
  ProxyType_Socks5 = 1;
  ProxyType_Http = 2;

  MailServerAuthLogin = 0;
  MailServerAuthCRAM5 = 1;
  MailServerAuthNTLM = 2;

  MailServerPop3 = 0;
  MailServerImap4 = 1;
  MailServerEWS = 2;   // For Exchange Web Service - Exchange 2007/2010
  MailServerDAV = 3; // For Exchange Web DAV - Exchange 2000/2003

var
  Form1: TForm1;
  m_bCancel: WordBool;
  m_curpath: WideString;
  m_uidlfile: WideString;
  oClient: TMailClient;
  oCurServer: TMailServer;
  oUIDLManager: TUIDLManager;
  oTools: TTools;

implementation

{$R *.dfm}

// EAGetMail event handler
procedure TForm1.OnIdle(ASender: TObject; const oSender: IDispatch; var Cancel: WordBool);
begin
  Application.ProcessMessages();
  Cancel := m_bCancel;
end;

procedure TForm1.OnConnected(ASender: TObject; const oSender: IDispatch; var Cancel: WordBool);
begin
  lblStatus.Caption := 'Connected';
  Cancel := m_bCancel;
end;

procedure TForm1.OnQuit(ASender: TObject; const oSender: IDispatch; var Cancel: WordBool);
begin
  lblStatus.Caption := 'Disconnecting ...';
end;

procedure TForm1.OnSecuring(ASender: TObject; const oSender: IDispatch; var Cancel: WordBool);
begin
  lblStatus.Caption := 'Securing ...';
  Cancel := m_bCancel;
end;

procedure TForm1.OnAuthorized(ASender: TObject; const oSender: IDispatch;
      var Cancel: WordBool);
begin
  lblStatus.Caption := 'Authorized';
  Cancel := m_bCancel;
end;

procedure TForm1.OnSendingDataStream(ASender: TObject; const oSender: IDispatch;
      Sent: Integer; Total: Integer;
      var Cancel: WordBool);
begin
  pgBar.Max := Total;
  pgBar.Min := 0;
  pgBar.Position := Sent;
  Cancel := m_bCancel;
end;

procedure TForm1.OnReceivingDataStream(ASender: TObject; const oSender: IDispatch;
      const oInfo: IDispatch;
      Received: Integer; Total: Integer;
      var Cancel: WordBool);
begin
  pgBar.Max := Total;
  pgBar.Min := 0;
  pgBar.Position := Received;
  Cancel := m_bCancel;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  lstAuthType.AddItem('USER/LOGIN', nil);
  lstAuthType.AddItem('APOP/AUTH CRAM-MD5', nil);
  lstAuthType.AddItem('NTLM', nil);
  lstAuthType.ItemIndex := 0;

  lstProtocol.AddItem('IMAP4', nil);
  lstProtocol.AddItem('Exchange Web Service - 2007/2010', nil);
  lstProtocol.AddItem('Exchange WebDAV - 2000/2003', nil);
  lstProtocol.ItemIndex := 0;

  webMail.Navigate('about:blank');
  m_curpath := GetCurrentDir();
  m_uidlfile := m_curpath + '\uidl.txt';

  m_bCancel := false;
  oTools := TTools.Create(Application);
  oUIDLManager := TUIDLManager.Create(Application);
  oCurServer := nil;

  lstMail.Items.Clear();
  trFolders.Items.Clear();

end;

procedure TForm1.btnStartClick(Sender: TObject);
var
  oServer: TMailServer;
begin

  lstMail.Items.Clear();
  trFolders.Items.Clear();

  EnableButton(False);
  m_bCancel := False;
  btnCancel.Enabled := False;

  if trim(textServer.Text) = '' then
  begin
    ShowMessage( 'Plese input server address!' );
    textServer.SetFocus();
    exit;
  end;

  if(trim(textUser.Text) = '' )  then
  begin
    ShowMessage( 'Please input user!');
    textUser.SetFocus();
    exit;
  end;

  if(trim(textPassword.Text) = '') then
  begin
    ShowMessage( 'Please input password!');
    textPassword.SetFocus();
    exit;
  end;



  oServer := TMailServer.Create(Application);
  oServer.Server := trim(textServer.Text);
  oServer.User := trim(textUser.Text);
  oServer.Password := trim(textPassword.Text);
  oServer.AuthType := lstAuthType.ItemIndex;
  oServer.Protocol := lstProtocol.ItemIndex + 1;
  oServer.SSLConnection := chkSSL.Checked;

  if oServer.Protocol = MailServerPop3 then
    if oServer.SSLConnection then
      oServer.Port := 995
    else
      oServer.Port := 110
  else
    if oServer.SSLConnection then
      oServer.Port :=993
    else
      oServer.Port := 143;

  m_bCancel := false;
  lblStatus.Caption := 'Connecting server ...';
  btnStart.Enabled := false;
  btnCancel.Enabled := true;
  pgBar.Position := 0;

  try
    oClient := TMailClient.Create(Application);
    oClient.LicenseCode := 'TryIt';

    // Add eagetmail event handler
    // You don't have to use EAGetMail Event,
    // but using Event make your application more user friendly
    oClient.OnIdle := OnIdle;
    oClient.OnConnected := OnConnected;
    oClient.OnQuit := OnQuit;
    oClient.OnSecuring := OnSecuring;
    oClient.OnAuthorized := OnAuthorized;
    oClient.OnSendingDataStream := OnSendingDataStream;
    oClient.OnReceivingDataStream := OnReceivingDataStream;

    oCurServer := oServer;
    btnCancel.Enabled := true;
    m_bCancel := False;
    
    oClient.Connect1( oCurServer.DefaultInterface );

    ShowFolders();
    EnableButton(True);
  except
    on ep:Exception do
    begin
        if oClient <>nil then
          oClient.Close();

        trFolders.Items.Clear();
        ShowMessage(ep.Message);
        lblStatus.Caption := ep.Message;
        btnStart.Enabled := true;
        btnCancel.Enabled := false;
    end;
    
  end;
end;

procedure TForm1.ShowFolders();
var
  oNode: TTreeNode;
begin
  try
    trFolders.Items.Clear();

    oNode := trFolders.Items.AddChild(nil,
        oCurServer.Server + '\' + oCurServer.User);

    trFolders.Selected := oNode;
    ConnectServer(); //connect server on demand

    ExpandFolders( oClient.Imap4Folders, oNode);
    oNode.Expand(true);
    // EnableButton True
  except
    on ep:Exception do
    begin
        oClient.Close();
        ShowMessage(ep.Message);
        // EnableButton True
    end;
  end;

end;

procedure TForm1.ConnectServer();
var
oNode: TTreeNode;
begin
   // EnableButton False
    
    if oClient.Connected then
        Exit;

    lblStatus.Caption := 'Connecting server ...';
    m_bCancel := False;
    oClient.Connect1(oCurServer.DefaultInterface);
    lblStatus.Caption := 'Online';
    
    oNode := trFolders.Selected;
   
    if oNode =nil then
      Exit;

    if oNode.Parent = nil then
      Exit;

     // select current folder
     oClient.SelectFolder(IImap4Folder(oNode.Data));
end;

procedure TForm1.ExpandFolders(oFolders: OleVariant; oNode: TTreeNode);
var
  oSubNode: TTreeNode;
  oFolder: IImap4Folder;
  i, UBound: integer;
begin
  oNode.DeleteChildren();
  UBound := VarArrayHighBound( oFolders, 1 );
  for i:= 0 to UBound do
  begin
       oFolder := IDispatch(VarArrayGet(oFolders, i)) as IImap4Folder;
       oSubNode := trFolders.Items.AddChild( oNode, WideStringToString(oFolder.Name));
       oSubNode.Data := Pointer(oFolder);
       ExpandFolders( oFolder.SubFolders, oSubNode );
  end;
end;

procedure TForm1.ExpandFoldersEx(oFolders: OleVariant; oNode: TTreeNode);
var
  oSubNode: TTreeNode;
  oFolder: IImap4Folder;
  i, UBound: integer;
begin
  oNode.DeleteChildren();
  UBound := VarArrayHighBound( oFolders, 1 );
  for i:= 0 to UBound do
  begin
       oFolder := IDispatch(VarArrayGet(oFolders, i)) as IImap4Folder;
       oSubNode := Form3.trFolders.Items.AddChild( oNode, WideStringToString(oFolder.Name));
       oSubNode.Data := Pointer(oFolder);
       ExpandFolders( oFolder.SubFolders, oSubNode );
  end;
end;

procedure TForm1.btnCancelClick(Sender: TObject);
begin
  m_bCancel := true;
  btnCancel.Enabled := false;
end;

// Display mails
procedure TForm1.LoadMails();
var
  oMail: TMail;
  oTools: TTools;
  i, UBound: integer;
  find, mailFolder: WideString;
  files: OleVariant;
  fileName: WideString;
  item: TListItem;
begin
  mailFolder := m_curpath + '\inbox';
  find := mailFolder + '\*.eml';

  oTools := TTools.Create(Application);
  oTools.CreateFolder(mailFolder);

  files := oTools.GetFiles(find);
  UBound := VarArrayHighBound( files, 1 );
  for i := 0 to UBound do
  begin
    fileName := VarArrayGet(files, i );
    oMail := TMail.Create(Application);
    oMail.LicenseCode := 'TryIt';

    // Only load the email header temporially
    oMail.LoadFile(fileName, true);

    item := TListItem.Create(lstMail.Items);
    item.SubItems.Add(oMail.Subject);
    item.SubItems.Add(FormatDateTime('yyyy-MM-dd hh:mm:ss', oMail.ReceivedDate));
    item.SubItems.Add(fileName);
    lstMail.Items.AddItem(item);
    item.Caption := oMail.From.Address;
  end;

end;

procedure TForm1.lstMailSelectItem(Sender: TObject; Item: TListItem;
  Selected: Boolean);
var
  fileName, folder, tempFolder, htmlName: WideString;
  oNode: TTreeNode;
  oFolder: IImap4Folder;
  oInfo: IMailInfo;
  uidl_item: IUIDLItem;
  oMail: IMail;
begin
  oNode := trFolders.Selected;
  if oNode = nil then
    exit;

  if oNode.Parent = nil then
    exit;
    
  if not Selected then
    exit;

  if lstMail.SelCount <> 1 then
    exit;

  EnableButton( False );
  try
    oFolder := IImap4Folder(oNode.Data);
    oInfo := IMailInfo(Item.Data);

    ConnectServer();

    folder := m_curpath + '\' + oCurServer.Server + '\' + oCurServer.User +
    '\' + oFolder.LocalPath;

    CreateFullFolder( folder );

    uidl_item := oUIDLManager.FindUIDL(oCurServer.DefaultInterface, oInfo.UIDL);
    if uidl_item = nil then
    begin
      // should never happen except you delete the file from the folder manually.
      ShowMessage('No email file found');
      exit;
    end;

    // Get the  local file name for this email UIDL
    fileName := folder + '\' + uidl_item.fileName;

    tempFolder := LeftStr( fileName, length(fileName)-4 );
    htmlName := tempFolder + '.htm';
    // only mail header is retrieved, now retrieve full content of mail.
    if not oTools.ExistFile(htmlName) then
    begin
      lblStatus.Caption := 'Downloading email ...';
      oMail := oClient.GetMail(oInfo);
      oMail.SaveAs( fileName, true );
    end;

    if not oInfo.Read then
      oClient.MarkAsRead(oInfo, true);


    item.Update();
    ShowMail( fileName );

    lblStatus.Caption := Format( 'Total %d email(s)', [lstMail.Items.Count]);
  except
    on ep:Exception do
    begin
      oClient.Close();
      ShowMessage(ep.Message);
    end;
  end;
  EnableButton( True );
end;

procedure TForm1.ShowMail(fileName: WideString);
var
  mainName, htmlName, tempFolder: WideString;
  oTools: TTools;
begin
// remove .eml extension
  mainName := LeftStr( fileName, length(fileName)-4 );
  htmlName := mainName + '.htm';
  tempFolder := mainName;

  oTools := TTools.Create(Application);
  if not oTools.ExistFile(htmlName) then
  begin
    // we haven't generate the html for this email, generate it now.
    GenerateHtmlForEmail(htmlName, fileName, tempFolder );
  end;

  webMail.Navigate(htmlName);
end;

procedure TForm1.GenerateHtmlForEmail(htmlName: WideString;
emlFile: WideString; tempFolder: WideString);
var
  oMail: TMail;
  oTools: TTools;
  i, x, UBound, XBound: integer;
  att, tatt: IAttachment;
  addr: IMailAddress;
  addrs, atts, tatts: OleVariant;
  html, hdr, attname, tattname: WideString;
begin
  oTools := TTools.Create(Application);
  oMail := TMail.Create(Application);
  oMail.LicenseCode := 'TryIt';

  oMail.LoadFile( emlFile, false );

  try
    if oMail.IsEncrypted then
      oMail.Load(oMail.Decrypt(nil).Content);
  except
    on ep:Exception do
      ShowMessage( 'Decrypt Error: ' + ep.Message );
  end;

  try
    if oMail.IsSigned then
      oMail.VerifySignature();
  except
    on ep:Exception do
      ShowMessage( 'Verify Digital Signature Error: ' + ep.Message );
  end;

  html := oMail.HtmlBody;
  hdr := hdr + '<font face="Courier New,Arial" size="2">';
  hdr := hdr + '<b>From:</b> ' + oMail.From.Name + '&lt;' + oMail.From.Address + '&gt;' + '<br>';

  addrs := oMail.ToAddr;
  UBound := VarArrayHighBound( addrs, 1 );
  if( UBound >= 0) then
  begin
    hdr := hdr + '<b>To:</b>';
    for i:= 0 to UBound do
    begin
      addr := IDispatch(VarArrayGet(addrs, i)) as IMailAddress;
      hdr := hdr + addr.Name + '&lt;' + addr.Address + '&gt';
      if( i < UBound ) then
        hdr := hdr + ';';
    end;
    hdr := hdr + '<br>' + #13#10;
  end;

  addrs := oMail.Cc;
  UBound := VarArrayHighBound( addrs, 1 );
  if( UBound >= 0) then
  begin
    hdr := hdr + '<b>Cc:</b>';
    for i:= 0 to UBound do
    begin
      addr := IDispatch(VarArrayGet(addrs, i)) as IMailAddress;
      hdr := hdr + addr.Name + '&lt;' + addr.Address + '&gt';
      if( i < UBound ) then
        hdr := hdr + ';';
    end;
    hdr := hdr + '<br>' + #13#10;
  end;

  hdr := hdr + '<b>Subject:</b>' + oMail.Subject + '<br>' + #13#10;


  // Parse attachment
  atts := oMail.Attachments;
  UBound := VarArrayHighBound( atts, 1 );
  if( UBound >= 0 ) then
  begin
    // create a temporal folder to store attachments.
    if not oTools.ExistFile(tempFolder) then
      oTools.CreateFolder(tempFolder);

    hdr := hdr + '<b>Attachments:</b>';
    for i:= 0 to UBound do
    begin
      att := IDispatch(VarArrayGet(atts, i)) as IAttachment;
      if LowerCase(att.Name) = 'winmail.dat' then
      begin
      // this is outlook rtf (TNEF) attachment, decode it here
        tatts := oMail.ParseTNEF(att.Content, true );
        XBound := VarArrayHighBound(tatts, 1 );
        for x:= 0 to XBound do
        begin
          tatt := IDispatch(VarArrayGet(tatts,x)) as IAttachment;
          tattname := tempFolder + '\' + tatt.Name;
          tatt.SaveAs(tattname, true);
          hdr := hdr + '<a href="' + tattname + '" target="_blank">' + tatt.Name + '</a> ';
        end;
      end
      else
      begin
        attname := tempFolder + '\' + att.Name;
        att.SaveAs(attname, true);
        hdr := hdr + '<a href="' + attname + '" target="_blank">' + att.Name + '</a> ';
        // show embedded images
        if att.ContentID <> '' then
        begin
          // StringReplace doesn't support some non-ascii characters very well.
          html := StringReplace( html, 'cid:' + att.ContentID, attname, [rfReplaceAll, rfIgnoreCase]);
        end
        else if Pos('image/', att.ContentType ) = 1 then
        begin
          html := html + '<hr><img src="' + attname + '">';
        end;

      end;
    end;
  end;

  hdr := '<meta http-equiv="Content-Type" content="text-html; charset=utf-8">' + hdr;
  html := hdr + '<hr>' + html;
  oTools.WriteTextFile( htmlName, html, 65001 );

end;

procedure TForm1.btnDelClick(Sender: TObject);
var
  oInfo: IMailInfo;
  item: TListItem;
  oNode: TTreeNode;
  i: integer;
begin
  oNode := trFolders.Selected;
  
  if oNode = nil then
    exit;

  if oNode.Parent = nil then
    exit;

  if lstMail.SelCount < 1 then
    exit;

  EnableButton( False );
  try
    ConnectServer();
    i := 0;
    while i < lstMail.Items.Count do
    begin
      item := lstMail.Items[i];
      if not item.Selected then
      begin
        i := i+1;
        continue;
      end;

      oInfo := IMailInfo(item.Data);
      oClient.Delete(oInfo);
      if oCurServer.Protocol <> MailServerImap4 then
        lstMail.Items.Delete(i)
      else
      begin
        i := i+1;
        item.Update();
      end;
    end;

  except
    on ep:Exception do
    begin
      oClient.Close();
      ShowMessage( ep.Message );
    end;
  end;
  
  EnableButton(True);

end;

procedure TForm1.lstMailCompare(Sender: TObject; Item1, Item2: TListItem;
  Data: Integer; var Compare: Integer);
var
  s1, s2: string;
begin
    s1:= Item1.SubItems[1];;
    s2 := Item2.SubItems[1];
    Compare := -CompareText(s1,s2);
end;

procedure TForm1.FormResize(Sender: TObject);
begin
  if Form1.Width < 945 then
    Form1.Width := 945;

  if Form1.Height < 484 then
    Form1.Height := 484;

  lstMail.Width := Form1.Width - 465;
  webMail.Width := lstMail.Width;
  webMail.Height := Form1.Height - 260;
  pgBar.Top :=  Form1.Height - 55;
  trFolders.Height := Form1.Height - 60;
end;

procedure TForm1.lstProtocolChange(Sender: TObject);
begin
// by default, Exchange Web Service requires SSL connection
// for other protocols, please set SSL connection based on your server setting.
  if lstProtocol.ItemIndex + 1 = MailServerEWS then
    chkSSL.Checked := true;
end;

// because delphi doesn't support unicode very well in VCL, so
// we have to convert the ansistring to unicode by current default codepage.
function TForm1.WideStringToString(const ws: WideString): AnsiString;
var
  l: integer;
begin
  if ws = '' then
    Result := ''
  else
  begin
    l := WideCharToMultiByte(GetACP(), 0,
      @ws[1], - 1, nil, 0, nil, nil);
    SetLength(Result, l - 1);
    if l > 1 then
      WideCharToMultiByte(GetACP(), 0,
        @ws[1], - 1, @Result[1], l - 1, nil, nil);
  end;
end;

function TForm1.StringToWideString(const s: AnsiString): WideString;
var
  l: integer;
begin
  if s = '' then
    Result := ''
  else 
  begin
    l := MultiByteToWideChar(GetACP(), 0, PAnsiChar(@s[1]), - 1, nil, 0);
    SetLength(Result, l - 1);
    if l > 1 then
      MultiByteToWideChar(GetACP(), 0, PAnsiChar(@s[1]),
        - 1, PWideChar(@Result[1]), l - 1);
  end;
end; 
procedure TForm1.trFoldersChange(Sender: TObject; Node: TTreeNode);
begin
  lstMail.Items.Clear();
  if Node.Parent = nil then
  begin
  EnableButton(True);
      Exit;
  end;

  EnableButton(False);
  try
     LoadServerMails( Node, IImap4Folder(Node.Data));
  except
    on ep:Exception do
    begin
      oClient.Close();
      ShowMessage( ep.Message );
    end;
  end;
  EnableButton(True);
end;
procedure TForm1.LoadServerMails( oNode: TTreeNode; oFolder: IImap4Folder);
var
  folder: WideString;
  infos: OleVariant;
  uidl_item: IUIDLItem;
  i, UBound: integer;
  oInfo: IMailInfo;
  fileName, localFile: WideString;
  oMail: TMail;
  item: TListItem;
begin
  lstMail.Items.Clear();
  if oFolder.NoSelect then
  begin
    oNode.Expand(true);
    Exit;
  end;

   //EnableButton False
  try
    folder := m_curpath + '\' + oCurServer.Server + '\' + oCurServer.User +
    '\' + oFolder.LocalPath;

    CreateFullFolder( folder );
    ConnectServer();

    lblStatus.Caption := 'Refreshing emails ...';
    oClient.SelectFolder(oFolder);
    oUIDLManager.Load( folder + '\uidl.txt');

    infos := oClient.GetMailInfos();

    // Remove the local uidl which is not existed on the server.
    oUIDLManager.SyncUIDL(oCurServer.DefaultInterface, infos);
    oUIDLManager.Update();

    // Remove the email on local disk which is not existed on server
    ClearLocalMails(infos, folder);

    UBound := VarArrayHighBound( infos, 1 );
    for i:= 0 to UBound do
    begin
      lblStatus.Caption := Format( 'Retrieve summary %d/%d ...', [i+1, UBound+1]);
      oInfo := IDispatch(VarArrayGet(infos, i)) as IMailInfo;

      fileName := oTools.GenFileName(i) + '.eml';
      localFile := folder + '\' + fileName;

      uidl_item := oUIDLManager.FindUIDL(oCurServer.DefaultInterface, oInfo.UIDL);
      if uidl_item <> nil then
        localFile := folder + '\' + uidl_item.FileName;

      oMail := TMail.Create(Application);
      oMail.LicenseCode := 'TryIt';

      if oTools.ExistFile(localFile) then
        oMail.LoadFile(localFile, false)
      else
      begin
        oMail.Load(oClient.GetMailHeader(oInfo));
        oMail.SaveAs(localFile, true);
      end;

      item := TListItem.Create(lstMail.Items);

      item.SubItems.Add(WideStringToString(oMail.Subject));
      item.SubItems.Add(FormatDateTime('yyyy-MM-dd hh:mm:ss', oMail.ReceivedDate));
      item.SubItems.Add(fileName);
      item.Data := Pointer(oInfo);
      lstMail.Items.AddItem(item);
      item.Caption := oMail.From.Address;

      if uidl_item = nil then
        oUIDLManager.AddUIDL( oCurServer.DefaultInterface,
        oInfo.UIDL, fileName );

      item.Update();

      oUIDLManager.Update();

    end;
    lblStatus.Caption := Format( 'Total %d email(s)', [lstMail.Items.Count]);

  except
    on ep:Exception do
    begin
      oUIDLManager.Update();
      Raise;
    end;
  end;

  //EnableButton True;
  
end;

procedure TForm1.CreateFullFolder( folder: WideString );
var
  s: WideString;
  npos: integer;
begin
    if oTools.ExistFile(folder) then
        Exit;

    npos := 1;
    while true do
    begin
      npos := PosEx('\', folder, npos );
      if npos > 3 then
      begin
        s := MidStr( folder, 1, npos - 1 );
        if Not oTools.ExistFile(s) then
          oTools.CreateFolder(s);

      end
      else if npos < 1 then
        Break;

      npos := npos + 1;
    end;

    if Not oTools.ExistFile(folder) then
      oTools.CreateFolder(folder);  
end;
// refresh folders
procedure TForm1.RefreshFolders1Click(Sender: TObject);
begin

  EnableButton( False );
  try
    lstMail.Items.Clear();
    oClient.RefreshFolders();
    ShowFolders();

  except
    on ep:Exception do
    begin
      oClient.Close();
      ShowMessage( ep.Message );

    end;
  end;

  EnableButton( True );

end;

procedure TForm1.RefreshMails1Click(Sender: TObject);
var
  oNode: TTreeNode;
  oFolder: IImap4Folder;
begin
  oNode := trFolders.Selected;
  if oNode = nil then
    exit;

  if oNode.Parent = nil then
    exit;


  EnableButton( False );
  try
    oFolder := IImap4Folder(oNode.Data);
    ConnectServer();
    oClient.SelectFolder(oFolder);
    oClient.RefreshMailInfos();

    LoadServerMails( oNode, oFolder );

  except
    on ep:Exception do
    begin
      oClient.Close();
      ShowMessage( ep.Message );

    end;
  end;

  EnableButton( True );
end;

procedure TForm1.AddFolder1Click(Sender: TObject);
var
  oDlg: TForm2;
  oNode, oNewNode: TTreeNode;
  folderName: WideString;
  oParent, oFolder: IImap4Folder;
begin
  oDlg := TForm2.Create(Application);
  if oDlg.ShowModal() <> mrOK then
    exit;

 EnableButton( False );
 try
    ConnectServer();
    oNode := trFolders.Selected;
    oParent := nil;
    if oNode <> nil then
      if oNode.Parent <> nil then
        oParent := IImap4Folder(oNode.Data);


    folderName := StringToWideString(oDlg.textFolder.Text);
    oFolder := oClient.CreateFolder(oParent, folderName);

    oNewNode := trFolders.Items.AddChild(oNode, oDlg.textFolder.Text);
    oNewNode.Data := Pointer(oFolder);

    if oNode <> nil then
      oNode.Expand(true);

  except
    on ep:Exception do
    begin
      oClient.Close();
      ShowMessage( ep.Message );

    end;
  end;

  EnableButton( True );
  
end;

procedure TForm1.DeleteFolder1Click(Sender: TObject);
var
  oNode: TTreeNode;
  oFolder: IImap4Folder;
begin
  oNode := trFolders.Selected;
  if oNode = nil then
    exit;

  if oNode.Parent = nil then
    exit;


 EnableButton( False );
 try
    ConnectServer();
    oFolder := IImap4Folder(oNode.Data);
    oClient.DeleteFolder(oFolder);
    lstMail.Items.Clear();
    trFolders.Items.Delete(oNode);

  except
    on ep:Exception do
    begin
      oClient.Close();
      ShowMessage( ep.Message );
    end;
  end;

  EnableButton( True );

end;

procedure TForm1.RenameFolder1Click(Sender: TObject);
var
  oNode: TTreeNode;
begin
  oNode := trFolders.Selected;
  if oNode = nil then
    exit;

  if oNode.Parent = nil then
    exit;

  oNode.EditText();

end;

procedure TForm1.trFoldersEdited(Sender: TObject; Node: TTreeNode;
  var S: String);
var
  oFolder: IImap4Folder;
begin
  EnableButton( False );
  try
    ConnectServer();
    oFolder := IImap4Folder(Node.Data);
    oClient.RenameFolder(oFolder, StringToWideString(S));
    Node.Text := S;
  except
    on ep:Exception do
    begin
      Node.EndEdit(true);
      S := Node.Text;
      oClient.Close();
      ShowMessage( ep.Message );
    end;
  end;
  EnableButton( True );
end;


procedure TForm1.EnableButton( enabled: Bool);
var
  oNode: TTreeNode;
begin
//
    btnDel.Enabled := enabled;
    btnUndelete.Enabled := enabled;
    btnUnread.Enabled := enabled;
    btnPure.Enabled := enabled;
    btnMove.Enabled := enabled;
    btnCopy.Enabled := enabled;
    btnUpload.Enabled := enabled;

    if lstMail.SelCount = 0 then
    begin;
      btnDel.Enabled := False;
      btnUndelete.Enabled := False;
      btnUnread.Enabled := False;
      btnCopy.Enabled := False;
      btnMove.Enabled := False;
    end;

    oNode := trFolders.Selected;
    if oNode = nil then
    begin
      btnUpload.Enabled := False;
      btnPure.Enabled := False;
    end
    else if oNode.Parent = nil then
    begin
      btnUpLoad.Enabled := False;
      btnPure.Enabled := False;
    end;

    btnCancel.Enabled := (Not enabled);

    if btnStart.Enabled then
        btnCancel.Enabled := False;

    btnQuit.Enabled := enabled;
    trFolders.Enabled := enabled;
    lstMail.Enabled := enabled;

    if  (oCurServer <> nil ) then
        if (oCurServer.Protocol = MailServerEWS) Or
            (oCurServer.Protocol = MailServerDAV) then
        begin
        // Exchange WebDAV and EWS doesn't support this operating
            btnUndelete.Enabled := False;
            btnPure.Enabled := False;
        end;

end;

procedure TForm1.btnQuitClick(Sender: TObject);
begin
  try
    oClient.Logout();

    webMail.Navigate ('about:blank');
  except
    on ep:Exception do
    begin

    end;
  end;
  lstMail.Items.Clear();
  trFolders.Items.Clear();
  btnStart.Enabled := True;
  oClient.Close();
  EnableButton( False );
end;

procedure TForm1.lstMailAdvancedCustomDrawItem(Sender: TCustomListView;
  Item: TListItem; State: TCustomDrawState; Stage: TCustomDrawStage;
  var DefaultDraw: Boolean);
var
  oInfo: IMailInfo;
begin
  if item.Data = nil then
    exit;

  oInfo := IMailInfo(item.Data);
  if not oInfo.Read then
    Sender.Canvas.Font.Style := Sender.Canvas.Font.Style + [fsBold]
  else
    Sender.Canvas.Font.Style := Sender.Canvas.Font.Style - [fsBold];

  if oInfo.Deleted then
    Sender.Canvas.Font.Style := Sender.Canvas.Font.Style +  [fsStrikeOut]
  else
    Sender.Canvas.Font.Style := Sender.Canvas.Font.Style -  [fsStrikeOut];

end;

procedure TForm1.btnUndeleteClick(Sender: TObject);
var
  oInfo: IMailInfo;
  item: TListItem;
  oNode: TTreeNode;
  i: integer;
begin
  oNode := trFolders.Selected;
  
  if oNode = nil then
    exit;

  if oNode.Parent = nil then
    exit;

  if lstMail.SelCount < 1 then
    exit;

  if (oCurServer.Protocol = MailServerEWS) Or
    (oCurServer.Protocol = MailServerDAV) then
    exit;

  EnableButton( False );
  try
    ConnectServer();
    i := 0;
    while i < lstMail.Items.Count do
    begin
      item := lstMail.Items[i];
      if not item.Selected then
      begin
        i := i+1;
        continue;
      end;

      oInfo := IMailInfo(item.Data);
      if oInfo.Deleted then
        oClient.Undelete(oInfo);

      i := i+1;
      item.Update();

    end;

  except
    on ep:Exception do
    begin
      oClient.Close();
      ShowMessage( ep.Message );
    end;
  end;

  EnableButton(True);

end;


procedure TForm1.btnUnreadClick(Sender: TObject);
var
  oInfo: IMailInfo;
  item: TListItem;
  oNode: TTreeNode;
  i: integer;
begin
  oNode := trFolders.Selected;
  
  if oNode = nil then
    exit;

  if oNode.Parent = nil then
    exit;

  if lstMail.SelCount < 1 then
    exit;

  EnableButton( False );
  try
    ConnectServer();
    i := 0;
    while i < lstMail.Items.Count do
    begin
      item := lstMail.Items[i];
      if not item.Selected then
      begin
        i := i+1;
        continue;
      end;

      oInfo := IMailInfo(item.Data);
      if oInfo.Read  then
        oClient.MarkAsRead(oInfo, False);

      i := i+1;
      item.Update();

    end;

  except
    on ep:Exception do
    begin
      oClient.Close();
      ShowMessage( ep.Message );
    end;
  end;
  
  EnableButton(True);

end;

procedure TForm1.btnPureClick(Sender: TObject);
var
  oFolder: IImap4Folder;
  oNode: TTreeNode;
begin
  oNode := trFolders.Selected;
  
  if oNode = nil then
    exit;

  if oNode.Parent = nil then
    exit;


  if (oCurServer.Protocol = MailServerEWS) Or
    (oCurServer.Protocol = MailServerDAV) then
    exit;
        
  EnableButton( False );
  try
    oFolder := IImap4Folder(oNode.Data);
    ConnectServer();
    oClient.SelectFolder(oFolder);
    oClient.Expunge();

    LoadServerMails( oNode, oFolder );

  except
    on ep:Exception do
    begin
      oClient.Close();
      ShowMessage( ep.Message );
    end;
  end;
  
  EnableButton(True);

end;

procedure TForm1.btnCopyClick(Sender: TObject);
var
  oInfo: IMailInfo;
  oDest, oFolder: IImap4Folder;
  item: TListItem;
  oNode, oRoot: TTreeNode;
  i: integer;
begin
  oNode := trFolders.Selected;
  
  if oNode = nil then
    exit;

  if oNode.Parent = nil then
    exit;

  if lstMail.SelCount < 1 then
    exit;

  EnableButton( False );
  try
    ConnectServer();

    Form3.trFolders.Items.Clear();
    oRoot := Form3.trFolders.Items.AddChildFirst(nil, 'Root Folder');
    ExpandFoldersEx( oClient.Imap4Folders, oRoot );
    oRoot.Expand(true);

    if Form3.ShowModal() <> mrOK then
    begin
      EnableButton(True);
      exit;
    end;

    oDest := IImap4Folder(Form3.trFolders.Selected.Data);
    oFolder := IImap4Folder(oNode.Data);

    if WideCompareText(oDest.FullPath, oFolder.FullPath ) = 0 then
    begin
      EnableButton(True);
      exit;
    end;

    i := 0;
    while i < lstMail.Items.Count do
    begin
      item := lstMail.Items[i];
      if not item.Selected then
      begin
        i := i+1;
        continue;
      end;

      oInfo := IMailInfo(item.Data);
      oClient.Copy( oInfo, oDest);
      i := i+1;

    end;

  except
    on ep:Exception do
    begin
      oClient.Close();
      ShowMessage( ep.Message );
    end;
  end;

  EnableButton(True);

end;

procedure TForm1.btnMoveClick(Sender: TObject);
var
  oInfo: IMailInfo;
  oDest, oFolder: IImap4Folder;
  item: TListItem;
  oNode, oRoot: TTreeNode;
  i: integer;
begin
  oNode := trFolders.Selected;

  if oNode = nil then
    exit;

  if oNode.Parent = nil then
    exit;

  if lstMail.SelCount < 1 then
    exit;

  EnableButton( False );
  try
    ConnectServer();

    Form3.trFolders.Items.Clear();
    oRoot := Form3.trFolders.Items.AddChildFirst(nil, 'Root Folder');
    ExpandFoldersEx( oClient.Imap4Folders, oRoot );
    oRoot.Expand(true);


    if Form3.ShowModal() <> mrOK then
    begin
      EnableButton(True);
      exit;
    end;

    oDest := IImap4Folder(Form3.trFolders.Selected.Data);
    oFolder := IImap4Folder(oNode.Data);

    if WideCompareText(oDest.FullPath, oFolder.FullPath ) = 0 then
    begin
      EnableButton(True);
      exit;
    end;

    i := 0;
    while i < lstMail.Items.Count do
    begin
      item := lstMail.Items[i];
      if not item.Selected then
      begin
        i := i+1;
        continue;
      end;

      oInfo := IMailInfo(item.Data);
      oClient.Move( oInfo, oDest);

      if oCurServer.Protocol <> MailServerImap4 then
        lstMail.Items.Delete(i)
      else
      begin
        i := i+1;
        item.Update();
      end;
    end;

  except
    on ep:Exception do
    begin
      oClient.Close();
      ShowMessage( ep.Message );
    end;
  end;

  EnableButton(True);

end;

procedure TForm1.btnUploadClick(Sender: TObject);
var
  pFileDlg : TOpenDialog;
  fileName : string;
  oNode: TTreeNode;
  oMail: TMail;
  oFolder: IImap4Folder;
begin
  oNode := trFolders.Selected;

  if oNode = nil then
    exit;

  if oNode.Parent = nil then
    exit;

  pFileDlg := TOpenDialog.Create(Form1);
  pFileDlg.Filter := 'Email File (*.eml)|*.EML';
  if not pFileDlg.Execute() then
  begin
    pFileDlg.Destroy();
    Exit;
  end;
  fileName := pFileDlg.FileName;
  pFileDlg.Destroy();

  EnableButton(False);
  try
    ConnectServer();
    oMail := TMail.Create(Application);
    oMail.LicenseCode := 'TryIt';
    oMail.LoadFile(StringToWideString(fileName), False);

    oFolder := IImap4Folder(oNode.Data);
    oClient.Append(oFolder, oMail.Content);
    oClient.RefreshMailInfos();
    LoadServerMails( oNode, oFolder );

   except
    on ep:Exception do
    begin
      oClient.Close();
      ShowMessage( ep.Message );
    end;
  end;

  EnableButton(True);
end;

procedure TForm1.ClearLocalMails( infos: OleVariant; folder: WideString );
var
  files: OleVariant;
  i, UBound: integer;
  fileName, shortName: WideString;
  nPos: integer;
  bFind: bool;
  htmlName, tempFolder: WideString;
begin
  files := oTools.GetFiles( folder + '\*.eml');
  UBound := VarArrayHighBound( files, 1 );
  for i:= 0 to UBound do
  begin
    fileName := WideString(VarArrayGet(files, i));
    shortName := fileName;
    while True do
    begin
      nPos := Pos(  '\', shortName );
      if nPos < 1 then
        break;

      shortName := MidStr( shortName, nPos+1, length(shortName) - nPos);
    end;

    bFind := False;
    if oUIDLManager.FindLocalFile(shortName) <> nil then
      bFind := True;

    if Not bFind then
    begin
      oTools.RemoveFile(fileName);
      tempFolder := LeftStr( fileName, length(fileName)-4 );
      htmlName := tempFolder + '.htm';

      if oTools.ExistFile(htmlName) then
        oTools.RemoveFile( htmlName );

      if oTools.ExistFile(tempFolder) then
        oTools.RemoveFolder(tempFolder, True);

      end;
  end;
end;

end.
