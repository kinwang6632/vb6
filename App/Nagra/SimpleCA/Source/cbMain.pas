unit cbMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, IniFiles, Contnrs,
  Encryption_TLB,
  cxEdit, cxControls, cxContainer, cxTextEdit, DB, ADODB, cxMaskEdit,
  cxDropDownEdit, cxImageComboBox, cxMemo, cxGroupBox, cxRadioGroup,
  cxGridTableView, ImgList, Menus, cxGraphics;

const
  DBINIFILE = 'CONFIG.INI';
  DBSECTION = 'DBINFO';


type

  TCASimple = record
    UserId: String;
    UserPass: String;
    UserName: String;
    UserLock: Boolean;
    CompStr: String;
    ExpireDate: String;
    ChLimitCount: Integer;
    ChLimitDay: Integer;
    Admin: Boolean;
  end;

  PChannel = ^TChannel;
  TChannel = record
    CodeNo: String;
    Description: String;
    ChannelId: String;
  end;

  TSendNagra = record
    Actioin: String;
    HighlevelCmdId: string;
    IccNo: String;
    StbNo: String;
    SubscriptionBeginDate: String;
    SubscriptionEndDate: String;
    Notes: String;
    CmdStatus: String;
    Operator: String;
    Seqno: String;
    CompCode: String;
    MisIrdCmdId: String;
    MisIrdCmdData: String;
    StbFlag: String;
    ProductText: String;
  end;

  TLoader = class
  private
    FIniObj: TIniFile;
    FEncrypObj: _Password;
    FEncrypKey: WideString;
    function ProduceDecryptFile(aSourceFileName: String): String;
  protected
    procedure LoadFromFile(aFullFileName: String); virtual; abstract;
  public
    constructor Create;
    destructor Destroy; override;
  end;

  TSoLoader = class(TLoader)
  private
    FSoIndex: Integer;
    FSoList: TObjectList;
    function GetCompCode: String;
    function GetCompName: String;
    function GetSid: String;
    function GetOwner: String;
    function GetPassword: String;
    function GetEnabled: Boolean;
    property SoIndex: Integer read FSoIndex write FSoIndex;
  public
    procedure LoadFromFile(aFullFileName: String); override;
    property LoaderResult: TObjectList read FSoList;
    constructor Create;
    destructor Destroy; override;
  end;

  TSoInfo = class
  private
    FCompCode: String;
    FPassword: String;
    FCompName: String;
    FSid: String;
    FOwner: String;
    FEnable: Boolean;
  public
    property CompCode: String read FCompCode;
    property CompName: String read FCompName;
    property Sid: String read FSid;
    property Owner: String read FOwner;
    property Password: String read FPassword;
    property Enable: Boolean read FEnable;
    constructor Create(aLoader: TLoader; const aSoIndex: Integer);
    destructor Destroy; override;
  end;

  TCom = record
    Owner: String;
    Pass: String;
    Sid: String;
  end;


  TfmMain = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    Button1: TButton;
    Bevel1: TBevel;
    Label1: TLabel;
    Label2: TLabel;
    txtUserId: TcxTextEdit;
    txtPassword: TcxTextEdit;
    btnLogin: TButton;
    Bevel2: TBevel;
    ComConnection: TADOConnection;
    cxEditStyleController1: TcxEditStyleController;
    Label4: TLabel;
    txtIcc: TcxTextEdit;
    Label5: TLabel;
    txtStb: TcxTextEdit;
    btnAddProd: TButton;
    Label6: TLabel;
    txtChannel: TcxMemo;
    cxGroupBox1: TcxGroupBox;
    rdoADB: TcxRadioButton;
    rdoHDT: TcxRadioButton;
    btnCmdA1: TButton;
    btnCmdA2: TButton;
    btnCmdE1: TButton;
    btnCmdB1: TButton;
    btnCmdB2: TButton;
    ComReader: TADOQuery;
    SoConnection: TADOConnection;
    SoDataWriter: TADOQuery;
    Panel4: TPanel;
    lblMsg: TLabel;
    Image1: TImage;
    Panel5: TPanel;
    Image2: TImage;
    Label3: TLabel;
    cmbSO: TcxImageComboBox;
    SoDbImageOk: TImage;
    SoDbImageNone: TImage;
    lblStep: TLabel;
    ImageList1: TImageList;
    ComWriter: TADOQuery;
    btnUser: TButton;
    SoReader: TADOQuery;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormDestroy(Sender: TObject);
    procedure btnLoginClick(Sender: TObject);
    procedure cmbSOPropertiesChange(Sender: TObject);
    procedure btnAddProdClick(Sender: TObject);
    procedure btnCmdA1Click(Sender: TObject);
    procedure btnCmdA2Click(Sender: TObject);
    procedure btnCmdE1Click(Sender: TObject);
    procedure btnCmdB1Click(Sender: TObject);
    procedure btnCmdB2Click(Sender: TObject);
    procedure btnUserClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
    FCASimple: TCASimple;
    FSoList: TObjectList;
    FChList: TList;
    FPinCode: String;
    FOpenChannelDay: Integer;
    FCom: TCom;
    function DoDbLogin: Boolean;
    function DoLogout: Boolean;
    function DoUserLogin: Boolean;
    function DoChangeSoDB: Boolean;
    procedure ChangeUIState(aLogin, aSoLogin: Boolean);
    procedure ReleaseSoList;
    procedure CreateSoList;
    procedure ReleaseChList;
    procedure SoListBindToComboBox;
    procedure WriteToSendNagra(aCmd: String);
    function CalcSubBeginDate: String;
    function CalcSubEndDate: String;
    function GetActionTex(aCmd: String): String;
    function GetChId(aCmd: String): String;
    function GetChText: String;
    function GetSendNagraSeq: String;
    function VdDoActionCount(aCmd: String; var aCount: Integer): Boolean;
    function VdExpire: Boolean;
    function CheckInput(aCmd: String): Boolean;
    procedure ReloadCASimple;
    procedure DataSetToCASimple;
  public
    { Public declarations }
    property SoList: TObjectList read FSoList;
  end;

var
  fmMain: TfmMain;

implementation

{$R *.dfm}

uses cbUtilis, cbChannel, cbUser;

{ TLoader }

constructor TLoader.Create;
begin
  FEncrypObj := CoPassword.Create;
  FEncrypKey := 'CS';
end;

{ ---------------------------------------------------------------------------- }

destructor TLoader.Destroy;
begin
  FEncrypObj := nil;
  inherited;
end;

{ ---------------------------------------------------------------------------- }

function TLoader.ProduceDecryptFile(aSourceFileName: String): String;

     { ---------------------------------------------------- }

     function GetTempDir: String;
     var
       aBuffer: array [0..MAX_PATH] of Char;
     begin
       ZeroMemory( @aBuffer[0], SizeOf( aBuffer ) );
       GetTempPath( SizeOf( aBuffer ), @aBuffer[0] );
       Result := aBuffer;
     end;

     { ---------------------------------------------------- }

var
  aIndex: Integer;
  aStrList: TStringList;
  aPrefix, aLineStr, aTempDir: String;
begin
  Result := EmptyStr;
  if not FileExists( aSourceFileName ) then Exit;
  aStrList := TStringList.Create;
  try
    aStrList.LoadFromFile( aSourceFileName );
    for aIndex := 0 to aStrList.Count - 1 do
    begin
      aLineStr := aStrList[aIndex];
      aPrefix := Copy( aLineStr, 1, 2 );
      if ( aPrefix = '//' ) or ( aPrefix = EmptyStr ) then Continue;
      if Pos( 'COMPCODE', aLineStr ) <= 0 then
        aStrList[aIndex] := FEncrypObj.Decrypt( aLineStr, FEncrypKey )
      else
        aStrList[aIndex] := aLineStr;
    end;
    aTempDir := GetTempDir;
    Result := IncludeTrailingPathDelimiter( aTempDir ) + ExtractFileName(
      aSourceFileName );
    aStrList.SaveToFile( Result );  
  finally
     aStrList.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

{ TSoLoader }

constructor TSoLoader.Create;
begin
  inherited Create;
  FSoList := nil;
  FSoIndex := -1;
end;

{ ---------------------------------------------------------------------------- }

destructor TSoLoader.Destroy;
begin
  FSoList := nil;
  inherited;
end;

{ ---------------------------------------------------------------------------- }

function TSoLoader.GetCompCode: String;
begin
  Result := FIniObj.ReadString( DBSECTION, Format( 'COMPCODE_%d', [FSoIndex+1] ),
    EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

function TSoLoader.GetCompName: String;
begin
  Result := FIniObj.ReadString( DBSECTION, Format( 'COMPNAME_%d', [FSoIndex+1] ),
    EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

function TSoLoader.GetOwner: String;
begin
  Result := FIniObj.ReadString( DBSECTION, Format( 'USERID_%d', [FSoIndex+1] ),
    EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

function TSoLoader.GetPassword: String;
begin
  Result := FIniObj.ReadString( DBSECTION, Format( 'PASSWORD_%d', [FSoIndex+1] ),
    EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

function TSoLoader.GetSid: String;
begin
  Result := FIniObj.ReadString( DBSECTION, Format( 'ALIAS_%d', [FSoIndex+1] ),
    EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

function TSoLoader.GetEnabled: Boolean;
var
  aStr: String;
begin
  aStr := FIniObj.ReadString( DBSECTION, Format( 'ENABLE_%d', [FSoIndex+1] ),
    'N' );
  Result := ( aStr = 'Y' );
end;

{ ---------------------------------------------------------------------------- }

procedure TSoLoader.LoadFromFile(aFullFileName: String);
var
  aIndex, aCount: Integer;
begin
  FIniObj := TIniFile.Create( ProduceDecryptFile( aFullFileName ) );
  try
    FSoList := nil;
    aCount := FIniObj.ReadInteger( DBSECTION, 'DB_COUNT', 0 );
    if aCount <= 0 then Exit;
    FSoList := TObjectList.Create( True );
    for aIndex := 0 to aCount - 1 do
      FSoList.Add( TSoInfo.Create( Self, aIndex ) );
    DeleteFile( FIniObj.FileName );
  finally
    FIniObj.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

{ TSoInfo }

constructor TSoInfo.Create(aLoader: TLoader; const aSoIndex: Integer);
begin
  TSoLoader( aLoader ).SoIndex := aSoIndex;
  FCompCode := TSoLoader( aLoader ).GetCompCode;
  FCompName := TSoLoader( aLoader ).GetCompName;
  FSid := TSoLoader( aLoader ).GetSid;
  FOwner := TSoLoader( aLoader ).GetOwner;
  FPassword := TSoLoader( aLoader ).GetPassword;
  FEnable := TSoLoader( aLoader ).GetEnabled;
end;

{ ---------------------------------------------------------------------------- }

destructor TSoInfo.Destroy;
begin
  inherited;
end;

{ ---------------------------------------------------------------------------- }

{ TForm1 }

procedure TfmMain.FormCreate(Sender: TObject);
var
  aFileName, aFileName2: String;
  aIni: TIniFile;
  aLoader: TLoader;
begin
  aFileName := IncludeTrailingPathDelimiter( ExtractFilePath(
    ParamStr( 0 ) ) ) + DBINIFILE;
  if not FileExists( aFileName ) then
  begin
    WarningMsg( '系統檔案不存在: ' + DBINIFILE + ', 無法執行此模組。' );
    Application.Terminate;
  end;
  aLoader := TLoader.Create;
  try
    aFileName2 := aLoader.ProduceDecryptFile( aFileName );
    aIni := TIniFile.Create( aFileName2 );
    try
      FCom.Sid := aIni.ReadString( 'COM', 'ALIAS', 'CATVN' );
      FCom.Owner := aIni.ReadString( 'COM', 'USERID', 'COM' );
      FCom.Pass := aIni.ReadString( 'COM', 'PASSWORD', 'COM' );
    finally
      aIni.Free;
    end;
    DeleteFile( aFileName2 );
  finally
    aLoader.Free;
  end;
  FChList := TList.Create;
  ChangeUIState( False, False );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.FormShow(Sender: TObject);
begin
  if txtUserId.CanFocusEx then txtUserId.SetFocus;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin

end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.FormDestroy(Sender: TObject);
begin
  try
    ReleaseChList;
    FChList.Free;
    ReleaseSoList;
  except
    { ... }
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.ChangeUIState(aLogin, aSoLogin: Boolean);
begin
  if not aLogin then
  begin
    lblMsg.Caption := '尚未登入。';
    lblStep.Caption := EmptyStr;
    txtUserId.Text := EmptyStr;
    txtUserId.Properties.ReadOnly := False;
    txtPassword.Enabled := True;
    txtPassword.Text := EmptyStr;
    btnLogin.Enabled := True;
    cmbSO.Enabled := False;
    cmbSO.Properties.OnChange := nil;
    try
      cmbSO.ItemIndex := -1;
    finally
      cmbSO.Properties.OnChange := cmbSOPropertiesChange;
    end;
    txtIcc.Enabled := False;
    txtStb.Enabled := False;
    rdoADB.Enabled := False;
    rdoHDT.Enabled := False;
    btnAddProd.Enabled := False;
    txtChannel.Enabled := False;
    btnCmdA1.Enabled := False;
    btnCmdA2.Enabled := False;
    btnCmdB1.Enabled := False;
    btnCmdB2.Enabled := False;
    btnCmdE1.Enabled := False;
    SoDbImageOk.Visible := False;
    SoDbImageNone.Visible := True;
    btnUser.Visible := False;
    btnLogin.Caption := '登入';
    btnLogin.Enabled := True;
  end else
  begin
    lblMsg.Caption := Format( '已登入為:%s。', [txtUserId.Text] );
    if cmbSO.ItemIndex < 0 then
      lblStep.Caption := '請選擇系統台。'
    else
      lblStep.Caption := '請輸入STB或ICC。';
    txtUserId.Properties.ReadOnly := True;
    txtPassword.Enabled := False;
    txtPassword.Text := EmptyStr;
    cmbSO.Enabled := True;
    txtIcc.Enabled := True;
    txtStb.Enabled := True;
    rdoADB.Enabled := True;
    rdoHDT.Enabled := True;
    rdoADB.Checked := True;
    btnAddProd.Enabled := True;
    txtChannel.Enabled := True;
    btnUser.Visible := FCASimple.Admin;
    btnLogin.Caption := '登出';
    btnLogin.Enabled := True;
  end;
  if ( aLogin and aSoLogin ) then
  begin
    btnCmdA1.Enabled := True;
    btnCmdA2.Enabled := True;
    btnCmdB1.Enabled := True;
    btnCmdB2.Enabled := True;
    btnCmdE1.Enabled := True;
    SoDbImageOk.Visible := True;
    SoDbImageNone.Visible := False;
  end else
  begin
    btnCmdA1.Enabled := False;
    btnCmdA2.Enabled := False;
    btnCmdB1.Enabled := False;
    btnCmdB2.Enabled := False;
    btnCmdE1.Enabled := False;
    SoDbImageOk.Visible := False;
    SoDbImageNone.Visible := True;
    if cmbSO.ItemIndex >= 0 then
      lblStep.Caption := '所選擇的系統台無法作業。';
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.DoDbLogin: Boolean;
begin
  ComConnection.Close;
  ComConnection.ConnectionString := Format(
   'Provider=MSDAORA.1;Password=%s;User ID=%s;Data Source=%s;Persist Security Info=True',
    [FCom.Pass, FCom.Owner, FCom.Sid] );
  try
    ComConnection.Open;
  except
    on E: Exception do
    begin
      ErrorMsg( Format( '資料庫連結有誤, 原因:%s。', [E.Message] ) );
      ErrorMsg( Format( '資料庫連結字元:%s。', [ComConnection.ConnectionString] ) );
    end;
  end;
  Result := ComConnection.Connected;
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.DoLogout: Boolean;
begin
  if ComConnection.Connected then
    try
      ComConnection.Connected := False;
    except
      { ... }
    end;
  if SoConnection.Connected then
    try
      SoConnection.Connected := False;
    except
      { ... }
    end;
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.DoUserLogin: Boolean;
var
  aPassObj: _Password;
  aKey: WideString;
  aPassword: String;
begin
  Result := False;
  if Trim( txtUserId.Text ) = EmptyStr then
  begin
    WarningMsg( '請輸入使用者帳號。' );
    if txtUserId.CanFocusEx then txtUserId.SetFocus;
    Exit;
  end;
  if Trim( txtPassword.Text ) = EmptyStr then
  begin
    WarningMsg( '請輸入使用者密碼。' );
    if txtPassword.CanFocusEx then txtPassword.SetFocus;
    Exit;
  end;
  aPassObj := CoPassword.Create;
  aKey := 'CS';
  aPassword := aPassObj.Encrypt( txtPassword.Text, aKey );
  txtPassword.Text := aPassword;
  ComReader.Close;
  ComReader.SQL.Text :=
    ' SELECT * FROM CASIMPLE   ' +
    '  WHERE USERID = :ID      ' +
    '    AND PASSWORD = :PASS  ';
  ComReader.Parameters.ParamByName( 'ID' ).Value := txtUserId.Text;
  ComReader.Parameters.ParamByName( 'PASS' ).Value := txtPassword.Text;
  try
    ComReader.Open;
    DataSetToCASimple;
    Result := not ComReader.IsEmpty;
    ComReader.Close;
    if not Result then
    begin
      WarningMsg( '無此帳號或密碼/帳號錯誤, 請重新登入。' );
      if txtUserId.CanFocusEx then txtUserId.SetFocus;
      txtPassword.Text := EmptyStr;
      Exit;
    end;
    Result := not FCASimple.UserLock;
    if not Result then
    begin
      WarningMsg( '此帳號已被鎖定, 請重新登入。' );
      if txtUserId.CanFocusEx then txtUserId.SetFocus;
      txtPassword.Text := EmptyStr;
      Exit;
    end;
    Result := VdExpire;
    if not Result then
    begin
      WarningMsg( '此帳號已過使用期限, 請重新登入。' );
      if txtUserId.CanFocusEx then txtUserId.SetFocus;
      txtPassword.Text := EmptyStr;
      Exit;
    end;
  except
    on E: Exception do
    begin
      ErrorMsg( '檢核使用者有誤, 原因:%s。' );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.btnLoginClick(Sender: TObject);
begin
  Screen.Cursor := crSQLWait;
  try
    if ( btnLogin.Caption = '登入' ) then
    begin
      if not DoDbLogin then Exit;
      if not DoUserLogin then Exit;
      CreateSoList;
      SoListBindToComboBox;
      ChangeUIState( True, False );
    end else
    begin
      DoLogout;
      ChangeUIState( False, False );
    end;  
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.ReleaseSoList;
begin
  if not Assigned( FSoList ) then Exit;
  while FSoList.Count > 0 do
    FSoList.Delete( 0 );
  FreeAndNil( FSoList );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.ReleaseChList;
var
  aIndex: Integer;
begin
  for aIndex := 0 to FChList.Count - 1 do
  begin
    if Assigned( FCHList[aIndex] ) then
      Dispose( PChannel( FChList[aIndex] ) );  
  end;
  FChList.Clear;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.CreateSoList;
var
  aLoader: TSoLoader;
  aFileName: String;
begin
  ReleaseSoList;
  aLoader := TSoLoader.Create;
  try
    aFileName := IncludeTrailingPathDelimiter( ExtractFilePath(
      ParamStr( 0 ) ) ) + DBINIFILE;
    aLoader.LoadFromFile( aFileName );
    FSoList := aLoader.LoaderResult;
  finally
    aLoader.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.SoListBindToComboBox;

  { ---------------------------------- }

  function GetSoObject(aCompCode: String): Pointer;
  var
    aIndex: Integer;
  begin
    Result := nil;
    for aIndex := 0 to FSoList.Count - 1 do
    begin
      if ( TSoInfo( FSoList[aIndex] ).CompCode = aCompCode ) then
      begin
        Result := FSoList[aIndex];
        Break;
      end;
    end;

  end;

  { ---------------------------------- }

var
  aTemp, aCompCode: String;
  aItem: TcxImageComboBoxItem;
  aSo: TSoInfo;
begin
  cmbSO.Properties.OnChange := nil;
  try
    cmbSO.Properties.Items.Clear;
    aTemp := FCASimple.CompStr;
    repeat
      aCompCode := ExtractValue( aTemp );
      if ( aCompCode <> EmptyStr ) then
      begin
        aSo := GetSoObject( aCompCode );
        if Assigned( aSo ) then
        begin
          aItem := cmbSO.Properties.Items.Add;
          aItem.Value := aSo.CompCode;
          aItem.Tag := Integer( aSo );
          aItem.Description := aSo.CompName;
          aItem.ImageIndex := 0;
        end;
      end;
    until ( aTemp = EmptyStr );
  finally
    cmbSO.Properties.OnChange := cmbSOPropertiesChange;
  end;
  //if ( cmbSO.ItemIndex < 0 ) and ( cmbSO.Properties.Items.Count = 1 ) then
  //  cmbSO.ItemIndex := 0;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.cmbSOPropertiesChange(Sender: TObject);
begin
  Screen.Cursor := crSQLWait;
  try
    if DoChangeSoDB then
      ChangeUIState( True, True )
    else
      ChangeUIState( True, False )
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.DoChangeSoDB: Boolean;
var
  aSo:TSoInfo;
begin
  aSo := TSoInfo( cmbSo.Properties.Items[cmbSO.ItemIndex].Tag );
  SoConnection.Close;
  SoConnection.ConnectionString := Format(
   'Provider=MSDAORA.1;Password=%s;User ID=%s;Data Source=%s;Persist Security Info=True',
    [aSo.Password, aSo.FOwner, aSo.Sid] );
  try
    SoConnection.Open;
  except
    on E: Exception do
    begin
      ErrorMsg( Format( '切換系統台有誤, 原因:%s。', [E.Message] ) );
    end;
  end;
  Result := SoConnection.Connected;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.btnAddProdClick(Sender: TObject);
var
  aIndex: Integer;
  aController: TcxGridTableController;
  aChannel: PChannel;
begin
  fmChannel := TfmChannel.Create( nil );
  try
    if ( fmChannel.ShowModal = mrOk ) then
    begin
      aController := fmChannel.ChGridView.Controller;
      if ( aController.SelectedRowCount > 0 ) then
      begin
        ReleaseChList;
        txtChannel.Lines.Clear;
      end;
      for aIndex := 0 to aController.SelectedRowCount - 1 do
      begin
        New( aChannel );
        aChannel.CodeNo := VarToStrDef(
          aController.SelectedRows[aIndex].Values[fmChannel.ViewCODENO.Index], EmptyStr );
        aChannel.Description := VarToStrDef(
          aController.SelectedRows[aIndex].Values[fmChannel.ViewDESCRIPTION.Index], EmptyStr );
        aChannel.ChannelId := VarToStrDef(
          aController.SelectedRows[aIndex].Values[fmChannel.ViewCHANNELID.Index], EmptyStr );
        FChList.Add( aChannel );
        txtChannel.Lines.Add( aChannel.Description );
      end;
    end;
  finally
    fmChannel.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.WriteToSendNagra(aCmd: String);
var
  aSo:TSoInfo;
  aSendNagra: TSendNagra;
  aStbFlag, aSendNagraSeq, aActionText: String;
begin
  aSo := TSoInfo( cmbSo.Properties.Items[cmbSO.ItemIndex].Tag );
  aStbFlag := '0';
  if rdoHDT.Checked then aStbFlag := '1';
  aSendNagraSeq := GetSendNagraSeq;
  if ( aSendNagraSeq = EmptyStr ) then
  begin
    ErrorMsg( Format( '發送指令有誤, 原因:%s。', ['產生指令序號有誤。'] ) );
    Exit;
  end;
  aActionText := GetActionTex( aCmd );
  if ( aCmd = 'A1' ) then
  begin
    aSendNagra.Actioin := aActionText;
    aSendNagra.HighlevelCmdId := 'A1';
    aSendNagra.IccNo := txtIcc.Text;
    aSendNagra.StbNo := txtStb.Text;
    aSendNagra.SubscriptionBeginDate := EmptyStr;
    aSendNagra.SubscriptionEndDate := EmptyStr;
    aSendNagra.Notes := EmptyStr;
    aSendNagra.CmdStatus := 'W';
    aSendNagra.Operator := FCASimple.UserName;
    aSendNagra.Seqno := aSendNagraSeq;
    aSendNagra.CompCode := aSo.CompCode;
    aSendNagra.MisIrdCmdId := '4';
    aSendNagra.MisIrdCmdData := EmptyStr;
    aSendNagra.StbFlag := aStbFlag;
    aSendNagra.ProductText := EmptyStr;
  end else
  if ( aCmd = 'A2' ) then
  begin
    aSendNagra.Actioin := aActionText;
    aSendNagra.HighlevelCmdId := 'A2';
    aSendNagra.IccNo := txtIcc.Text;
    aSendNagra.StbNo := EmptyStr;
    aSendNagra.SubscriptionBeginDate := EmptyStr;
    aSendNagra.SubscriptionEndDate := EmptyStr;
    aSendNagra.Notes := EmptyStr;
    aSendNagra.CmdStatus := 'W';
    aSendNagra.Operator := FCASimple.UserName;
    aSendNagra.Seqno := aSendNagraSeq;
    aSendNagra.CompCode := aSo.CompCode;
    aSendNagra.MisIrdCmdId := EmptyStr;
    aSendNagra.MisIrdCmdData := EmptyStr;
    aSendNagra.StbFlag := aStbFlag;
    aSendNagra.ProductText := EmptyStr;
  end else
  if ( aCmd = 'E1' ) then
  begin
    aSendNagra.Actioin := aActionText;
    aSendNagra.HighlevelCmdId := 'E1';
    aSendNagra.IccNo := txtIcc.Text;
    aSendNagra.StbNo := txtStb.Text;
    aSendNagra.SubscriptionBeginDate := EmptyStr;
    aSendNagra.SubscriptionEndDate := EmptyStr;
    aSendNagra.Notes := EmptyStr;
    aSendNagra.CmdStatus := 'W';
    aSendNagra.Operator := FCASimple.UserName;
    aSendNagra.Seqno := aSendNagraSeq;
    aSendNagra.CompCode := aSo.CompCode;
    aSendNagra.MisIrdCmdId := '1';
    aSendNagra.MisIrdCmdData := Copy( Nvl( FPinCode, '1234' ), 1, 4 );
    aSendNagra.StbFlag := aStbFlag;
    aSendNagra.ProductText := EmptyStr;
  end else
  if ( aCmd = 'B1' ) then
  begin
    aSendNagra.Actioin := aActionText;
    aSendNagra.HighlevelCmdId := 'B1';
    aSendNagra.IccNo := txtIcc.Text;
    aSendNagra.StbNo := txtStb.Text;
    aSendNagra.SubscriptionBeginDate := CalcSubBeginDate;
    aSendNagra.SubscriptionEndDate := CalcSubEndDate;
    aSendNagra.Notes := GetChId( aCmd );
    aSendNagra.CmdStatus := 'W';
    aSendNagra.Operator := FCASimple.UserName;
    aSendNagra.Seqno := aSendNagraSeq;
    aSendNagra.CompCode := aSo.CompCode;
    aSendNagra.MisIrdCmdId := EmptyStr;
    aSendNagra.MisIrdCmdData := EmptyStr;
    aSendNagra.StbFlag := aStbFlag;
    aSendNagra.ProductText := Copy( GetChText, 1, 1024 );
  end else
  if ( aCmd = 'B2' ) then
  begin
    aSendNagra.Actioin := aActionText;
    aSendNagra.HighlevelCmdId := 'B2';
    aSendNagra.IccNo := txtIcc.Text;
    aSendNagra.StbNo := txtStb.Text;
    aSendNagra.SubscriptionBeginDate := EmptyStr;
    aSendNagra.SubscriptionEndDate := EmptyStr;
    aSendNagra.Notes := GetChId( aCmd );
    aSendNagra.CmdStatus := 'W';
    aSendNagra.Operator := FCASimple.UserName;
    aSendNagra.Seqno := aSendNagraSeq;
    aSendNagra.CompCode := aSo.CompCode;
    aSendNagra.MisIrdCmdId := EmptyStr;
    aSendNagra.MisIrdCmdData := EmptyStr;
    aSendNagra.StbFlag := aStbFlag;
    aSendNagra.ProductText := Copy( GetChText, 1, 1024 );
  end;
  { 寫入 Log }
  ComWriter.Close;
  ComWriter.SQL.Text :=
    ' INSERT INTO CAWAREHOUSELOG( ICC_NO, STB_NO, PRODUCTID, ACTION,      ' +
    '       AUTHORSTARTDATE, AUTHORSTOPDATE, OPERATOR, UPDTIME, PINCODE ) ' +
    ' VALUES ( :1, :2, :3, :4, TO_DATE( :5, ''YYYY/MM/DD'' ),             ' +
    '          TO_DATE( :6, ''YYYY/MM/DD'' ), :7, SYSDATE, :9 )           ';
  ComWriter.Parameters.ParamByName( '1' ).Value := aSendNagra.IccNo;
  ComWriter.Parameters.ParamByName( '2' ).Value := aSendNagra.StbNo;
  ComWriter.Parameters.ParamByName( '3' ).Value := aSendNagra.ProductText;
  ComWriter.Parameters.ParamByName( '4' ).Value := aSendNagra.Actioin;
  ComWriter.Parameters.ParamByName( '5' ).Value := aSendNagra.SubscriptionBeginDate;
  ComWriter.Parameters.ParamByName( '6' ).Value := aSendNagra.SubscriptionEndDate;
  ComWriter.Parameters.ParamByName( '7' ).Value := aSendNagra.Operator;
  if aCmd = 'E1' then
    ComWriter.Parameters.ParamByName( '9' ).Value := Copy( Nvl( FPinCode, '1234' ), 1, 4 )
  else
    ComWriter.Parameters.ParamByName( '9' ).Value := EmptyStr;
  try
    ComWriter.ExecSQL;
  except
    on E: Exception do
    begin
      ErrorMsg( Format( '發送指令有誤, 原因:%s。', [E.Message] ) );
      Exit;
    end;
  end;
  { 寫入 Send_Nagra 傳送指令 }
  SoDataWriter.Close;
  SoDataWriter.SQL.Text :=
    ' INSERT INTO COM.SEND_NAGRA( HIGH_LEVEL_CMD_ID, ICC_NO, STB_NO,   ' +
    '    SUBSCRIPTION_BEGIN_DATE, SUBSCRIPTION_END_DATE, NOTES,        ' +
    '    CMD_STATUS, OPERATOR, SEQNO, COMPCODE, MIS_IRD_CMD_ID,        ' +
    '    MIS_IRD_CMD_DATA, STB_FLAG  )                                 ' +
    ' VALUES ( :1, :2, :3, TO_DATE( :4, ''YYYY/MM/DD'' ),              ' +
    '          TO_DATE( :5, ''YYYY/MM/DD'' ), :6, :7, :8, :9,          ' +
    '          :10, :11, :12, :13 )                                    ';
  SoDataWriter.Parameters.ParamByName( '1' ).Value := aSendNagra.HighlevelCmdId;
  SoDataWriter.Parameters.ParamByName( '2' ).Value := aSendNagra.IccNo;
  SoDataWriter.Parameters.ParamByName( '3' ).Value := aSendNagra.StbNo;
  SoDataWriter.Parameters.ParamByName( '4' ).Value := aSendNagra.SubscriptionBeginDate;
  SoDataWriter.Parameters.ParamByName( '5' ).Value := aSendNagra.SubscriptionEndDate;
  SoDataWriter.Parameters.ParamByName( '6' ).Value := aSendNagra.Notes;
  SoDataWriter.Parameters.ParamByName( '7' ).Value := aSendNagra.CmdStatus;
  SoDataWriter.Parameters.ParamByName( '8' ).Value := aSendNagra.Operator;
  SoDataWriter.Parameters.ParamByName( '9' ).Value := aSendNagra.Seqno;
  SoDataWriter.Parameters.ParamByName( '10' ).Value := aSendNagra.CompCode;
  SoDataWriter.Parameters.ParamByName( '11' ).Value := aSendNagra.MisIrdCmdId;
  SoDataWriter.Parameters.ParamByName( '12' ).Value := aSendNagra.MisIrdCmdData;
  SoDataWriter.Parameters.ParamByName( '13' ).Value := aSendNagra.StbFlag;
  try
    SoDataWriter.ExecSQL;
    InfoMsg( '動作執行中, 3分鐘內應可執行完畢。' );
  except
    on E: Exception do
    begin
      ErrorMsg( Format( '發送指令有誤, 原因:%s。', [E.Message] ) );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.CalcSubBeginDate: String;
begin
  Result := EmptyStr;
  ComReader.Close;
  ComReader.SQL.Text :=
    ' SELECT TO_CHAR( SYSDATE, ''YYYY/MM/DD'' ) FROM DUAL ';
  try
    ComReader.Open;
    Result := ComReader.Fields[0].AsString;
    ComReader.Close;
  except
    Result := FormatDateTime( 'yyyy/mm/dd', Now );
  end;
end;


{ ---------------------------------------------------------------------------- }

function TfmMain.CalcSubEndDate: String;
begin
  Result := EmptyStr;
  ComReader.Close;
  ComReader.SQL.Text := Format(
    ' SELECT TO_CHAR( SYSDATE + %d, ''YYYY/MM/DD'' ) FROM DUAL ', [FOpenChannelDay] );
  try
    ComReader.Open;
    Result := ComReader.Fields[0].AsString;
    ComReader.Close;
  except
    Result := FormatDateTime( 'yyyy/mm/dd', Now + FOpenChannelDay );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.GetActionTex(aCmd: String): String;
begin
  if ( aCmd = 'A1' ) then
    Result := '開機'
  else
  if ( aCmd = 'A2' ) then
    Result := '關機'
  else
  if ( aCmd = 'B1' ) then
    Result := '開頻道'
  else
  if ( aCmd = 'B2' ) then
    Result := '關頻道'
  else
  if ( aCmd = 'E1' ) then
    Result := '設定親子密碼';
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.GetChId(aCmd: String): String;
var
  aIndex: Integer;
  aChannel: PChannel;
begin
  Result := EmptyStr;
  for aIndex := 0 to FChList.Count - 1 do
  begin
    aChannel := PChannel( FChList[aIndex]);
    if ( aCmd = 'B1' ) then
      Result := Result + Format( 'A~%s,', [aChannel.ChannelId] )
    else
    if ( aCmd = 'B2' ) then
      Result := Result + Format( '%s,', [aChannel.ChannelId] );
  end;
  if IsDelimiter( ',', Result, Length( Result ) ) then
    Delete( Result, Length( Result ), 1 );
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.GetChText: String;
var
  aIndex: Integer;
  aChannel: PChannel;
begin
  Result := EmptyStr;
  for aIndex := 0 to FChList.Count - 1 do
  begin
    aChannel := PChannel( FChList[aIndex]);
    Result := Result + Format( '%s,', [aChannel.Description] );
  end;
  if IsDelimiter( ',', Result, Length( Result ) ) then
    Delete( Result, Length( Result ), 1 );
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.GetSendNagraSeq: String;
begin
  Result := EmptyStr;
  SoReader.Close;
  SoReader.SQL.Text := ' SELECT COM.S_SENDNAGRA_SEQNO.NEXTVAL FROM DUAL ';
  try
    SoReader.Open;
    Result := SoReader.Fields[0].AsString;
    SoReader.Close;
  except
    {...} 
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.VdDoActionCount(aCmd: String; var aCount: Integer): Boolean;
var
  aActionText: String;
begin
  aActionText := GetActionTex( aCmd );
  ComReader.Close;
  ComReader.SQL.Text := Format(
    ' SELECT COUNT(1) FROM CAWAREHOUSELOG A ' +
    '  WHERE A.STB_NO = ''%s''              ' +
    '    AND A.ACTION = ''%s''              ',
    [txtStb.Text, aActionText] );
  try
    ComReader.Open;
    aCount := ComReader.Fields[0].AsInteger;
    ComReader.Close;
  except
    aCount := FCASimple.ChLimitCount + 1;
  end;
  Result := ( aCount < FCASimple.ChLimitCount );
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.VdExpire: Boolean;
var
  aDate: String;
begin
  Result := True;
  if ( FCASimple.ExpireDate = EmptyStr ) then
    Exit;
  ComReader.Close;
  ComReader.SQL.Text :=
    ' SELECT TO_CHAR( SYSDATE, ''YYYY/MM/DD'' ) FROM DUAL ';
  try
    ComReader.Open;
    aDate := ComReader.Fields[0].AsString;
    ComReader.Close;
  except
    aDate := FormatDateTime( 'yyyy/mm/dd', Now );
  end;
  Result := StrToDate( aDate ) <= StrToDate( FCASimple.ExpireDate );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.btnCmdA1Click(Sender: TObject);
begin
  if not CheckInput( 'A1' ) then Exit;
  WriteToSendNagra( 'A1' );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.btnCmdA2Click(Sender: TObject);
begin
  if not CheckInput( 'A2' ) then Exit;
  WriteToSendNagra( 'A2' );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.btnCmdE1Click(Sender: TObject);
begin
  if not CheckInput( 'E1' ) then Exit;
  FPinCode := '1234';
  if InputQuery( '輸入', '請輸入親子密碼', FPinCode ) then
  begin
    if Trim( FPinCode ) <> EmptyStr then
      WriteToSendNagra( 'E1' );
  end;    
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.CheckInput(aCmd: String): Boolean;
begin
  Result := False;
  ReloadCASimple;
  if ( FCASimple.UserLock ) then
  begin
    WarningMsg( '此帳號已被鎖定, 請重新登入。' );
    ChangeUIState( False, False );
    Exit;
  end;
  if not VdExpire then
  begin
    WarningMsg( '此帳號已過使用期限, 請重新登入。' );
    ChangeUIState( False, False );
    Exit;
  end;
  if ( txtIcc.Text = EmptyStr ) then
  begin
    WarningMsg( '請輸入ICC NO' );
    if txtIcc.CanFocusEx then txtIcc.SetFocus;
    Exit;
  end;
  if ( txtStb.Text = EmptyStr ) then
  begin
    WarningMsg( '請輸入STB NO' );
    if txtStb.CanFocusEx then txtStb.SetFocus;
    Exit;
  end;
  if ( aCmd = 'B1' ) or ( aCmd = 'B2' ) then
  begin
    if FChList.Count = 0 then
    begin
      WarningMsg( '請點選頻道' );
      Exit;
    end;
  end;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.btnCmdB1Click(Sender: TObject);
var
  aCount: Integer;
  aCountText: String;
begin
  if not CheckInput( 'B1' ) then Exit;
  if FCASimple.ChLimitCount > 0 then
  begin
    if not VdDoActionCount( 'B1', aCount ) then
    begin
      WarningMsg( Format( '此 STB 已執行過 %d 次開頻道動作, 無法再次開頻道。', [aCount] ) );
      if txtStb.CanFocusEx then txtStb.SetFocus;
      Exit;
    end;
  end;
  if FCASimple.ChLimitDay <= 0 then
  begin
    if InputQuery( '輸入', '請輸入開頻道天數', aCountText ) then
    begin
      if not TryStrToInt( aCountText, FOpenChannelDay ) then
      begin
        WarningMsg( '所輸入的開頻道天數應為數字' );
        Exit;
      end;
      if FOpenChannelDay <= 0 then
      begin
        WarningMsg( '所輸入的開頻道天數應該大於零' );
        Exit;
      end;
      WriteToSendNagra( 'B1' );
    end;
  end else
  begin
    FOpenChannelDay := FCASimple.ChLimitDay;
    WriteToSendNagra( 'B1' );
  end;

end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.btnCmdB2Click(Sender: TObject);
begin
  if not CheckInput( 'B2' ) then Exit;
  WriteToSendNagra( 'B2' );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.ReloadCASimple;
begin
  ComReader.Close;
  ComReader.SQL.Text :=
    ' SELECT * FROM CASIMPLE   ' +
    '  WHERE USERID = :ID      ' +
    '    AND PASSWORD = :PASS  ';
  ComReader.Parameters.ParamByName( 'ID' ).Value := FCASimple.UserId;
  ComReader.Parameters.ParamByName( 'PASS' ).Value := FCASimple.UserPass;
  try
    ComReader.Open;
    DataSetToCASimple;
    ComReader.Close;
  except
    on E: Exception do
    begin
      ErrorMsg( Format( '重新載入模組參數有誤, 原因:%s。', [E.Message] ) );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.DataSetToCASimple;
begin
  FCASimple.UserId := txtUserId.Text;
  if txtPassword.Text = EmptyStr then
    FCASimple.UserPass := FCASimple.UserPass
  else
    FCASimple.UserPass := txtPassword.Text;
  FCASimple.UserName := ComReader.FieldByName( 'USERNAME' ).AsString;
  FCASimple.UserLock := ( ComReader.FieldByName( 'USERLOCK' ).AsString <> 'N' );
  FCASimple.CompStr := ComReader.FieldByName( 'COMPSTR' ).AsString;
  FCASimple.ExpireDate := ComReader.FieldByName( 'EXPIREDATE' ).AsString;
  FCASimple.ChLimitCount :=
    Nvl( ComReader.FieldByName( 'CHLIMITCOUNT' ).AsString, 2 );
  FCASimple.ChLimitDay :=
    Nvl( ComReader.FieldByName( 'CHLIMITDAY' ).AsString, 7 );
  if FCASimple.ChLimitDay > 0 then
    btnCmdB1.Caption := Format( '開頻道(%d天)', [FCASimple.ChLimitDay] );
  FCASimple.Admin := ( ComReader.FieldByName( 'ADMIN' ).AsString = 'Y' );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.btnUserClick(Sender: TObject);
begin
  fmUser := TfmUser.Create( nil );
  try
    fmUser.ShowModal;
  finally
    fmUser.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.Button1Click(Sender: TObject);
begin
  Self.Close;
end;

{ ---------------------------------------------------------------------------- }

end.

