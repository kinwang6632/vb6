unit cbMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Menus, ExtCtrls, DB,
  cxLookAndFeelPainters, dxSkinsCore, dxSkinsDefaultPainters,
  cxLookAndFeels, cxTextEdit, StdCtrls, cxButtons, cxControls,
  cxContainer, cxEdit, cxGroupBox, cxGraphics, cxMaskEdit, cxDropDownEdit,
  cxImageComboBox, dxSkinscxPCPainter, cxPC, cxStyles,
  VirtualTable, DBAccess, Ora, OracleCI, MemDS;

type

  TSoInfo = class(TObject)
  private
    FCompCode: string;
    FDbAliase: String;
    FDbAccount: String;
    FDbPassword: string;
    FConnectionString: String;
  public
    property CompCode: String read FCompCode write FCompCode;
    property DbAliase: String read FDbAliase write FDbAliase;
    property DbAccount: String read FDbAccount write FDbAccount;
    property DbPassword: String read FDbPassword write FDbPassword;
    property ConnectionString: String read FConnectionString write FConnectionString; 
  end;

  TfmMain = class(TForm)
    pnTitle: TPanel;
    Panel2: TPanel;
    cxGroupBox2: TcxGroupBox;
    lblConnectStatus: TLabel;
    btnConnect: TcxButton;
    cxGroupBox5: TcxGroupBox;
    txtNewSmartCard: TcxTextEdit;
    txtNewStepBox: TcxTextEdit;
    cxLookAndFeelController1: TcxLookAndFeelController;
    Label1: TLabel;
    Label9: TLabel;
    pnClient: TPanel;
    CommandPage: TcxPageControl;
    A1: TcxTabSheet;
    A2: TcxTabSheet;
    A3: TcxTabSheet;
    A4: TcxTabSheet;
    Label2: TLabel;
    txtOldStepBox: TcxTextEdit;
    Label3: TLabel;
    txtOldSmartCard: TcxTextEdit;
    A6: TcxTabSheet;
    A8: TcxTabSheet;
    A9: TcxTabSheet;
    B1: TcxTabSheet;
    B2: TcxTabSheet;
    B3: TcxTabSheet;
    B4: TcxTabSheet;
    B5: TcxTabSheet;
    B6: TcxTabSheet;
    B7: TcxTabSheet;
    B9: TcxTabSheet;
    C4: TcxTabSheet;
    E1: TcxTabSheet;
    E2: TcxTabSheet;
    E3: TcxTabSheet;
    E9: TcxTabSheet;
    Label13: TLabel;
    txtA1ZipCode: TcxTextEdit;
    Label14: TLabel;
    txtA1BatId: TcxTextEdit;
    lbNoParamNeeded: TLabel;
    Label4: TLabel;
    txtA6ZipCode: TcxTextEdit;
    Label5: TLabel;
    txtA6BatId: TcxTextEdit;
    Label6: TLabel;
    txtA6ProductId: TcxTextEdit;
    Label7: TLabel;
    Label8: TLabel;
    Label10: TLabel;
    txtA6StartDate: TcxMaskEdit;
    txtA6EndDate: TcxMaskEdit;
    Label11: TLabel;
    txtA8BatId: TcxTextEdit;
    Label12: TLabel;
    txtA9ZipCode: TcxTextEdit;
    Label15: TLabel;
    txtA9ProductId: TcxTextEdit;
    Label16: TLabel;
    Label17: TLabel;
    txtA9StartDate: TcxMaskEdit;
    Label18: TLabel;
    txtA9EndDate: TcxMaskEdit;
    Label19: TLabel;
    txtB1ProductId: TcxTextEdit;
    Label20: TLabel;
    Label21: TLabel;
    txtB1StartDate: TcxMaskEdit;
    Label22: TLabel;
    txtB1EndDate: TcxMaskEdit;
    Label23: TLabel;
    txtB2ProductId: TcxTextEdit;
    Label24: TLabel;
    Label25: TLabel;
    txtB4ProductId: TcxTextEdit;
    Label26: TLabel;
    Label27: TLabel;
    txtB5ProductId: TcxTextEdit;
    Label28: TLabel;
    Label29: TLabel;
    txtB6ProductId: TcxTextEdit;
    Label30: TLabel;
    Label32: TLabel;
    txtB6EndDate: TcxMaskEdit;
    Label31: TLabel;
    txtB7ProductId: TcxTextEdit;
    Label33: TLabel;
    Label34: TLabel;
    txtB7StartDate: TcxMaskEdit;
    Label35: TLabel;
    txtB7EndDate: TcxMaskEdit;
    Label36: TLabel;
    txtB7IccNo: TcxTextEdit;
    Label37: TLabel;
    Label38: TLabel;
    txtB9ZipCode: TcxTextEdit;
    Label39: TLabel;
    txtB9BatId: TcxTextEdit;
    Label40: TLabel;
    txtB9ProductId: TcxTextEdit;
    Label41: TLabel;
    Label42: TLabel;
    txtB9StartDate: TcxMaskEdit;
    Label43: TLabel;
    txtB9EndDate: TcxMaskEdit;
    Label44: TLabel;
    txtC4ZipCode: TcxTextEdit;
    Label45: TLabel;
    txtE1PinCode: TcxTextEdit;
    Label46: TLabel;
    txtE2Message: TcxTextEdit;
    Label47: TLabel;
    txtE2Priority: TcxComboBox;
    Label48: TLabel;
    Label49: TLabel;
    Label50: TLabel;
    txtE3NetworkId: TcxTextEdit;
    txtE3TransportId: TcxTextEdit;
    txtE3ServiceId: TcxTextEdit;
    Label51: TLabel;
    txtE9BatId: TcxTextEdit;
    OraSession: TOraSession;
    Label52: TLabel;
    Label53: TLabel;
    txtDbUserId: TcxTextEdit;
    Label54: TLabel;
    txtDbPassword: TcxTextEdit;
    cmbTNS: TcxComboBox;
    pnHor: TPanel;
    btnWriteSendNagra: TcxButton;
    OraWriter: TOraSQL;
    OraReader: TOraQuery;
    Label55: TLabel;
    Label56: TLabel;
    lbSendResult: TLabel;
    Label57: TLabel;
    lbErrCode: TLabel;
    Label58: TLabel;
    lbErrMsg: TLabel;
    Label59: TLabel;
    txtB5StartDate: TcxMaskEdit;
    Label60: TLabel;
    txtB5EndDate: TcxMaskEdit;
    procedure FormCreate(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormDestroy(Sender: TObject);
    procedure btnConnectClick(Sender: TObject);
    procedure CommandPageChange(Sender: TObject);
    procedure btnWriteSendNagraClick(Sender: TObject);
  private
    { Private declarations }
    FSoInfo: TSoInfo;
    FSubBeginDateList: TList;
    FSubEndDateList: TList;
    FSubProductList: TList;
    FZipCodeList: TList;
    FBatIdList: TList;
    FCurrentActiveSeqNo: String;
    procedure BuildTNS;
    procedure InitGUI;
    procedure FormControlAddToList;
    procedure DisconnectFromOracle;
    procedure MakeSubBeginDateSameValue(ADate: String);
    procedure MakeSubEndDateSameValue(ADate: String);
    function VdDbUserIdAndPassword: Boolean;
    function VdSmartCardAndStb(const ACmdText: String): Boolean;
    function VdParameter(const ACmdText: String): Boolean;
    function ConnectToOracle: Boolean;
    function IsNotNeedParameterCommand(const ACmdText: String): Boolean;
    function GetCompCode: String;
    function GetSeqNo: String;
    {}
    procedure WriteCommand(const ACmdText: String);
    procedure WriteA1;
    procedure WriteA2;
    procedure WriteA3;
    procedure WriteA4;
    procedure WriteA6;
    procedure WriteA8;
    procedure WriteA9;
    procedure WriteB1;
    procedure WriteB2;
    procedure WriteB3;
    procedure WriteB4;
    procedure WriteB5;
    procedure WriteB6;
    procedure WriteB7;
    procedure WriteB9;
    procedure WriteC4;
    procedure WriteE1;
    procedure WriteE2;
    procedure WriteE3;
    procedure WriteE9;
  public
    { Public declarations }
  end;

var
  fmMain: TfmMain;

implementation

uses cbUtilis, DateUtils, ComObj;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TfmMain.FormCreate(Sender: TObject);
begin
  FSoInfo := TSoInfo.Create;
  FSubBeginDateList := TList.Create;
  FSubEndDateList := TList.Create;
  FSubProductList := TList.Create;
  FZipCodeList := TList.Create;
  FBatIdList := TList.Create;
  BuildTNS;
  InitGUI;
  FormControlAddToList;
  MakeSubBeginDateSameValue( EmptyStr );
  MakeSubEndDateSameValue( EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin

end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.FormDestroy(Sender: TObject);
begin
  FSoInfo.Free;
  FSubBeginDateList.Free;
  FSubEndDateList.Free;
  FSubProductList.Free;
  FZipCodeList.Free;
  FBatIdList.Free;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.BuildTNS;
begin
  cmbTNS.Properties.Items.Clear;
  cmbTNS.Properties.Items.AddStrings( OracleAliasList );
  if ( cmbTNS.Properties.Items.Count > 0 ) then cmbTNS.ItemIndex := 0;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.InitGUI;
begin
  CommandPage.ActivePageIndex := 0;
  btnWriteSendNagra.Enabled := False;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.FormControlAddToList;
begin
  FSubBeginDateList.Add( txtA6StartDate );
  FSubBeginDateList.Add( txtA9StartDate );
  FSubBeginDateList.Add( txtB1StartDate );
  FSubBeginDateList.Add( txtB5StartDate );
  FSubBeginDateList.Add( txtB7StartDate );
  FSubBeginDateList.Add( txtB9StartDate );
  {}
  FSubEndDateList.Add( txtA6EndDate );
  FSubEndDateList.Add( txtA9EndDate );
  FSubEndDateList.Add( txtB1EndDate );
  FSubEndDateList.Add( txtB5EndDate );
  FSubEndDateList.Add( txtB6EndDate );
  FSubEndDateList.Add( txtB7EndDate );
  FSubEndDateList.Add( txtB9EndDate );
  {}
  FSubProductList.Add( txtA6ProductId );
  FSubProductList.Add( txtA9ProductId );
  FSubProductList.Add( txtB1ProductId );
  FSubProductList.Add( txtB2ProductId );
  FSubProductList.Add( txtB4ProductId );
  FSubProductList.Add( txtB5ProductId );
  FSubProductList.Add( txtB6ProductId );
  FSubProductList.Add( txtB7ProductId );
  FSubProductList.Add( txtB9ProductId );
  {}
  FZipCodeList.Add( txtA1ZipCode );
  FZipCodeList.Add( txtA6ZipCode );
  FZipCodeList.Add( txtA9ZipCode );
  FZipCodeList.Add( txtB9ZipCode );
  FZipCodeList.Add( txtC4ZipCode );
  {}
  FBatIdList.Add( txtA1BatId );
  FBatIdList.Add( txtA6BatId );
  FBatIdList.Add( txtA8BatId );
  FBatIdList.Add( txtB9BatId );
  FBatIdList.Add( txtE9BatId );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.DisconnectFromOracle;
begin
  if OraSession.Connected then OraSession.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.MakeSubBeginDateSameValue(ADate: String);
var
  aIndex: Integer;
begin
  if ( ADate = EmptyStr ) then
    ADate := FormatDateTime( 'yyyy/mm/dd', Date );
  for aIndex := 0 to FSubBeginDateList.Count - 1 do
    TcxMaskEdit( FSubBeginDateList[aIndex] ).Text := ADate;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.MakeSubEndDateSameValue(ADate: String);
var
  aIndex: Integer;
begin
  if ( ADate = EmptyStr ) then
    ADate := '2020/12/31';
  for aIndex := 0 to FSubEndDateList.Count - 1 do
    TcxMaskEdit( FSubEndDateList[aIndex] ).Text := ADate;
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.VdDbUserIdAndPassword: Boolean;
begin
  Result := True;
  if ( txtDbUserId.Text = EmptyStr )  then
  begin
    Result := False;
    WarningMsg( '請輸入 User Name。' );
    if ( txtDbUserId.CanFocusEx ) then txtDbUserId.SetFocus;
    Exit;
  end;
  {}
  if ( txtDbPassword.Text = EmptyStr )  then
  begin
    Result := False;
    WarningMsg( '請輸入 Password。' );
    if ( txtDbPassword.CanFocusEx ) then txtDbPassword.SetFocus;
    Exit;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.VdSmartCardAndStb(const ACmdText: String): Boolean;
var
  aMsg1, aMsg2: String;
begin
  Result := False;
  if ( Length( txtNewStepBox.Text ) < 10 ) then
    aMsg1 := '請輸入機上盒序號(機上盒序號長度必須為10碼以上)';
  {}
  if ( Length( txtNewSmartCard.Text ) < 10 ) then
    aMsg2 := '請輸入智慧卡號碼(智慧卡號碼長度必須為10碼以上)';
  {}
  if ( aMsg1 + aMsg2 ) <> EmptyStr then
  begin
    WarningMsg( aMsg1 + #13#10 + aMsg2 );
  end; 
  {}
  if ( aMsg1 <> EmptyStr ) then
  begin
    if txtNewStepBox.CanFocusEx then txtNewStepBox.SetFocus;
    Exit;
  end else
  if ( aMsg2 <> EmptyStr ) then
  begin
    if txtNewSmartCard.CanFocusEx then txtNewSmartCard.SetFocus;
    Exit;
  end;
  {}
  if ( ACmdText = 'A6' ) or ( ACmdText = 'A8' ) or ( ACmdText = 'A9' ) then
  begin
    if ( Length( txtOldStepBox.Text ) < 10 ) then
      aMsg1 := '請輸入舊機上盒序號(機上盒序號長度必須為10碼以上)';
    {}
    if ( Length( txtOldSmartCard.Text ) < 10 ) then
      aMsg2 := '請輸入舊智慧卡號碼(智慧卡號碼長度必須為10碼以上)';
    {}
    if ( aMsg1 + aMsg2 ) <> EmptyStr then
    begin
      WarningMsg( aMsg1 + #13#10 + aMsg2 );
    end;

    {}
    if ( aMsg1 <> EmptyStr ) then
    begin
      if txtOldStepBox.CanFocusEx then txtOldSmartCard.SetFocus;
      Exit;
    end else
    if ( aMsg2 <> EmptyStr ) then
    begin
      if txtOldSmartCard.CanFocusEx then txtOldSmartCard.SetFocus;
      Exit;
    end;
  end;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.VdParameter(const ACmdText: String): Boolean;
begin
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.ConnectToOracle: Boolean;
begin
  Screen.Cursor := crSQLWait;
  try
    OraSession.Close;
    OraSession.ConnectString := FSoInfo.ConnectionString;
    try
      OraSession.Open;
    except
      on E: Exception do
      begin
        ErrorMsg( Format( '資料庫連結失敗, 訊息:%s', [E.Message] ) );
      end;
    end;
  finally
    Screen.Cursor := crDefault;
  end;
  Result := OraSession.Connected;
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.IsNotNeedParameterCommand(const ACmdText: String): Boolean;
begin
  Result :=
    ( ACmdText = 'A2' ) or
    ( ACmdText = 'A3' ) or
    ( ACmdText = 'A4' ) or
    ( ACmdText = 'B3' );
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.GetCompCode: String;
begin
  OraReader.Close;
  OraReader.SQL.Text := Format(
    ' SELECT CODENO FROM CD039 WHERE UPPER( TABLEOWNER ) = ''%s'' ',
    [UpperCase( FSoInfo.DbAccount )] );
  OraReader.Open;
  Result := OraReader.Fields[0].AsString;
  OraReader.Close;
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.GetSeqNo: String;
begin
  OraReader.Close;
  OraReader.SQL.Text := ' SELECT S_SENDNAGRA_SEQNO.NEXTVAL FROM DUAL ';
  OraReader.Open;
  Result := OraReader.Fields[0].AsString;
  FCurrentActiveSeqNo := Result;
  OraReader.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.WriteCommand(const ACmdText: String);
begin
  if ( ACmdText = 'A1' ) then
    WriteA1
  else if ( ACmdText = 'A2' ) then
    WriteA2
  else if ( ACmdText = 'A3' ) then
    WriteA3
  else if ( ACmdText = 'A4' ) then
    WriteA4
  else if ( ACmdText = 'A6' ) then
    WriteA6
  else if ( ACmdText = 'A8' ) then
    WriteA8
  else if ( ACmdText = 'A9' ) then
    WriteA9
  else if ( ACmdText = 'B1' ) then
    WriteB1
  else if ( ACmdText = 'B2' ) then
    WriteB2
  else if ( ACmdText = 'B3' ) then
    WriteB3
  else if ( ACmdText = 'B4' ) then
    WriteB4
  else if ( ACmdText = 'B5' ) then
    WriteB5
  else if ( ACmdText = 'B6' ) then
    WriteB6
  else if ( ACmdText = 'B7' ) then
    WriteB7
  else if ( ACmdText = 'B9' ) then
    WriteB9
  else if ( ACmdText = 'C4' ) then
    WriteC4
  else if ( ACmdText = 'E1' ) then
    WriteE1
  else if ( ACmdText = 'E2' ) then
    WriteE2
  else if ( ACmdText = 'E3' ) then
    WriteE3
  else if ( ACmdText = 'E9' ) then
    WriteE9;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.WriteA1;
var
  aIndex: Integer;
  aSeqNo: String;
begin
  aSeqNo := GetSeqNo;
  OraWriter.SQL.Text :=
    ' insert into send_nagra (  ' +
    '   high_level_cmd_id,      ' + //1
    '   icc_no,                 ' + //2
    '   stb_no,                 ' + //3
    '   zip_code,               ' + //4
    '   mis_ird_cmd_data,       ' + //5
    '   pincode,                ' + //6
    '   cmd_status,             ' + //7
    '   operator,               ' + //8
    '   updtime,                ' +
    '   seqno,                  ' + //9
    '   compcode,               ' + //10
    '   resenttimes,            ' + //11
    '   stb_flag )              ' + //12
    ' values (                  ' +
    '   :1, :2, :3, :4,         ' +
    '   :5, :6, :7, :8,         ' +
    '   sysdate, :9, :10, :11, :12 ) ';
  {}  
  for aIndex := 0 to OraWriter.Params.Count - 1 do
  begin
    OraWriter.Params.Items[aIndex].DataType := ftString;
    OraWriter.Params.Items[aIndex].ParamType := ptInput;
  end;
  {}
  OraWriter.ParamByName( '1' ).AsString := 'A1';
  OraWriter.ParamByName( '2' ).AsString := txtNewSmartCard.Text;
  OraWriter.ParamByName( '3' ).AsString := txtNewStepBox.Text;
  OraWriter.ParamByName( '4' ).AsString := txtA1ZipCode.Text;
  OraWriter.ParamByName( '5' ).AsString := txtA1BatId.Text;
  OraWriter.ParamByName( '6' ).AsString := '0000';
  OraWriter.ParamByName( '7' ).AsString := 'W';
  OraWriter.ParamByName( '8' ).AsString := 'Nick';
  OraWriter.ParamByName( '9' ).AsString := aSeqNo;
  OraWriter.ParamByName( '10' ).AsString := FSoInfo.CompCode;
  OraWriter.ParamByName( '11' ).AsString := '0';
  OraWriter.ParamByName( '12' ).AsString := '1';
  {}
  OraWriter.Execute;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.WriteA2;
var
  aIndex: Integer;
  aSeqNo: String;
begin
  aSeqNo := GetSeqNo;
  OraWriter.SQL.Text :=
    ' insert into send_nagra (  ' +
    '   high_level_cmd_id,      ' + //1
    '   icc_no,                 ' + //2
    '   stb_no,                 ' + //3
    '   cmd_status,             ' + //4
    '   operator,               ' + //5
    '   updtime,                ' +
    '   seqno,                  ' + //6
    '   compcode,               ' + //7
    '   resenttimes,            ' + //8
    '   stb_flag )              ' + //9
    ' values (                  ' +
    '   :1, :2, :3, :4,         ' +
    '   :5, sysdate, :6,        ' +
    '   :7, :8, :9  )           ';
  {}  
  for aIndex := 0 to OraWriter.Params.Count - 1 do
  begin
    OraWriter.Params.Items[aIndex].DataType := ftString;
    OraWriter.Params.Items[aIndex].ParamType := ptInput;
  end;
  {}
  OraWriter.ParamByName( '1' ).AsString := 'A2';
  OraWriter.ParamByName( '2' ).AsString := txtNewSmartCard.Text;
  OraWriter.ParamByName( '3' ).AsString := EmptyStr;
  OraWriter.ParamByName( '4' ).AsString := 'W';
  OraWriter.ParamByName( '5' ).AsString := 'Nick';
  OraWriter.ParamByName( '6' ).AsString := aSeqNo;
  OraWriter.ParamByName( '7' ).AsString := FSoInfo.CompCode;
  OraWriter.ParamByName( '8' ).AsString := '0';
  OraWriter.ParamByName( '9' ).AsString := '1';
  {}
  OraWriter.Execute;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.WriteA3;
var
  aIndex: Integer;
  aSeqNo: String;
begin
  aSeqNo := GetSeqNo;
  OraWriter.SQL.Text :=
    ' insert into send_nagra (  ' +
    '   high_level_cmd_id,      ' + //1
    '   icc_no,                 ' + //2
    '   stb_no,                 ' + //3
    '   cmd_status,             ' + //4
    '   operator,               ' + //5
    '   updtime,                ' +
    '   seqno,                  ' + //6
    '   compcode,               ' + //7
    '   resenttimes,            ' + //8
    '   stb_flag )              ' + //9
    ' values (                  ' +
    '   :1, :2, :3, :4,         ' +
    '   :5, sysdate, :6,        ' +
    '   :7, :8, :9  )           ';
  {}  
  for aIndex := 0 to OraWriter.Params.Count - 1 do
  begin
    OraWriter.Params.Items[aIndex].DataType := ftString;
    OraWriter.Params.Items[aIndex].ParamType := ptInput;
  end;
  {}
  OraWriter.ParamByName( '1' ).AsString := 'A3';
  OraWriter.ParamByName( '2' ).AsString := txtNewSmartCard.Text;
  OraWriter.ParamByName( '3' ).AsString := EmptyStr;
  OraWriter.ParamByName( '4' ).AsString := 'W';
  OraWriter.ParamByName( '5' ).AsString := 'Nick';
  OraWriter.ParamByName( '6' ).AsString := aSeqNo;
  OraWriter.ParamByName( '7' ).AsString := FSoInfo.CompCode;
  OraWriter.ParamByName( '8' ).AsString := '0';
  OraWriter.ParamByName( '9' ).AsString := '1';
  {}
  OraWriter.Execute;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.WriteA4;
var
  aIndex: Integer;
  aSeqNo: String;
begin
  aSeqNo := GetSeqNo;
  OraWriter.SQL.Text :=
    ' insert into send_nagra (  ' +
    '   high_level_cmd_id,      ' + //1
    '   icc_no,                 ' + //2
    '   stb_no,                 ' + //3
    '   cmd_status,             ' + //4
    '   operator,               ' + //5
    '   updtime,                ' +
    '   seqno,                  ' + //6
    '   compcode,               ' + //7
    '   resenttimes,            ' + //8
    '   stb_flag )              ' + //9
    ' values (                  ' +
    '   :1, :2, :3, :4,         ' +
    '   :5, sysdate, :6,        ' +
    '   :7, :8, :9  )           ';
  {}  
  for aIndex := 0 to OraWriter.Params.Count - 1 do
  begin
    OraWriter.Params.Items[aIndex].DataType := ftString;
    OraWriter.Params.Items[aIndex].ParamType := ptInput;
  end;
  {}
  OraWriter.ParamByName( '1' ).AsString := 'A4';
  OraWriter.ParamByName( '2' ).AsString := txtNewSmartCard.Text;
  OraWriter.ParamByName( '3' ).AsString := EmptyStr;
  OraWriter.ParamByName( '4' ).AsString := 'W';
  OraWriter.ParamByName( '5' ).AsString := 'Nick';
  OraWriter.ParamByName( '6' ).AsString := aSeqNo;
  OraWriter.ParamByName( '7' ).AsString := FSoInfo.CompCode;
  OraWriter.ParamByName( '8' ).AsString := '0';
  OraWriter.ParamByName( '9' ).AsString := '1';
  {}
  OraWriter.Execute;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.WriteA6;

  { --------------------------------- }

  function BuildNotes: String;
  var
    aNote1, aNote2, aProducts, aPId: String;
  begin
    aNote1 := Format( '%s/%s@%s/%s',
      [txtOldStepBox.Text, txtOldSmartCard.Text,
       txtNewStepBox.Text, txtNewSmartCard.Text] );
    {}
    aProducts := txtA6ProductId.Text;
    repeat
      aPId := Trim( ExtractValue( aProducts ) );
      if ( aPId <> EmptyStr ) then
      begin
        aNote2 :=( aNote2 + Format( 'B~%s~%s~%s',
          [aPId,
           TrimChar( PadDateText( txtA6StartDate.Text ), ['/'] ),
           TrimChar( PadDateText( txtA6EndDate.Text ), ['/'] ) ] ) + ',' );
      end;
    until ( aProducts = EmptyStr );
    if IsDelimiter( ',', aNote2, Length( aNote2 ) ) then
      Delete( aNote2, Length( aNote2 ), 1 );
    Result := ( aNote1 + ';' + aNote2 );
  end;

  { --------------------------------- }

var
  aIndex: Integer;
  aSeqNo: String;
begin
  aSeqNo := GetSeqNo;
  OraWriter.SQL.Text :=
    ' insert into send_nagra (  ' +
    '   high_level_cmd_id,      ' + //1
    '   icc_no,                 ' + //2
    '   stb_no,                 ' + //3
    '   zip_code,               ' + //4
    '   mis_ird_cmd_data,       ' + //5
    '   pincode,                ' + //6
    '   cmd_status,             ' + //7
    '   operator,               ' + //8
    '   updtime,                ' +
    '   seqno,                  ' + //9
    '   compcode,               ' + //10
    '   resenttimes,            ' + //11
    '   stb_flag,               ' + //12
    '   notes   )               ' + //13
    ' values (                  ' +
    '   :1, :2, :3, :4,         ' +
    '   :5, :6, :7, :8,         ' +
    '   sysdate, :9, :10, :11,  ' +
    '   :12, :13 )              ';
  {}
  for aIndex := 0 to OraWriter.Params.Count - 1 do
  begin
    OraWriter.Params.Items[aIndex].DataType := ftString;
    OraWriter.Params.Items[aIndex].ParamType := ptInput;
  end;
  {}
  OraWriter.ParamByName( '1' ).AsString := 'A6';
  OraWriter.ParamByName( '2' ).AsString := EmptyStr;
  OraWriter.ParamByName( '3' ).AsString := EmptyStr;
  OraWriter.ParamByName( '4' ).AsString := txtA6ZipCode.Text;
  OraWriter.ParamByName( '5' ).AsString := txtA6BatId.Text;
  OraWriter.ParamByName( '6' ).AsString := '0000';
  OraWriter.ParamByName( '7' ).AsString := 'W';
  OraWriter.ParamByName( '8' ).AsString := 'Nick';
  OraWriter.ParamByName( '9' ).AsString := aSeqNo;
  OraWriter.ParamByName( '10' ).AsString := FSoInfo.CompCode;
  OraWriter.ParamByName( '11' ).AsString := '0';
  OraWriter.ParamByName( '12' ).AsString := '1';
  OraWriter.ParamByName( '13' ).AsString := BuildNotes;
  {}
  OraWriter.Execute;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.WriteA8;

  { --------------------------------- }

  function BuildNotes: String;
  begin
    Result := Format( '%s/%s@%s/%s',
      [txtOldStepBox.Text, txtOldSmartCard.Text,
       txtNewStepBox.Text, txtNewSmartCard.Text] );
  end;

  { --------------------------------- }

var
  aIndex: Integer;
  aSeqNo: String;
begin
  aSeqNo := GetSeqNo;
  OraWriter.SQL.Text :=
    ' insert into send_nagra (  ' +
    '   high_level_cmd_id,      ' + //1
    '   icc_no,                 ' + //2
    '   stb_no,                 ' + //3
    '   mis_ird_cmd_data,       ' + //4
    '   pincode,                ' + //5
    '   cmd_status,             ' + //6
    '   operator,               ' + //7
    '   updtime,                ' +
    '   seqno,                  ' + //8
    '   compcode,               ' + //9
    '   resenttimes,            ' + //10
    '   stb_flag,               ' + //11
    '   notes   )               ' + //12
    ' values (                  ' +
    '   :1, :2, :3, :4,         ' +
    '   :5, :6, :7,             ' +
    '   sysdate, :8, :9, :10,   ' +
    '   :11, :12  )             ';
  {}
  for aIndex := 0 to OraWriter.Params.Count - 1 do
  begin
    OraWriter.Params.Items[aIndex].DataType := ftString;
    OraWriter.Params.Items[aIndex].ParamType := ptInput;
  end;
  {}
  OraWriter.ParamByName( '1' ).AsString := 'A8';
  OraWriter.ParamByName( '2' ).AsString := EmptyStr;
  OraWriter.ParamByName( '3' ).AsString := EmptyStr;
  OraWriter.ParamByName( '4' ).AsString := txtA8BatId.Text;
  OraWriter.ParamByName( '5' ).AsString := '0000';
  OraWriter.ParamByName( '6' ).AsString := 'W';
  OraWriter.ParamByName( '7' ).AsString := 'Nick';
  OraWriter.ParamByName( '8' ).AsString := aSeqNo;
  OraWriter.ParamByName( '9' ).AsString := FSoInfo.CompCode;
  OraWriter.ParamByName( '10' ).AsString := '0';
  OraWriter.ParamByName( '11' ).AsString := '1';
  OraWriter.ParamByName( '12' ).AsString := BuildNotes;
  {}
  OraWriter.Execute;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.WriteA9;

  { --------------------------------- }

  function BuildNotes: String;
  var
    aNote1, aNote2, aProducts, aPId: String;
  begin
    aNote1 := Format( '%s/%s@%s/%s',
      [txtOldStepBox.Text, txtOldSmartCard.Text,
       txtNewStepBox.Text, txtNewSmartCard.Text] );
    {}
    aProducts := txtA9ProductId.Text;
    repeat
      aPId := Trim( ExtractValue( aProducts ) );
      if ( aPId <> EmptyStr ) then
      begin
        aNote2 :=( aNote2 + Format( 'B~%s~%s~%s',
          [aPId,
           TrimChar( PadDateText( txtA9StartDate.Text ), ['/'] ),
           TrimChar( PadDateText( txtA9EndDate.Text ), ['/'] )] ) + ',' );
      end;
    until ( aProducts = EmptyStr );
    if IsDelimiter( ',', aNote2, Length( aNote2 ) ) then
      Delete( aNote2, Length( aNote2 ), 1 );
    Result := ( aNote1 + ';' + aNote2 );
  end;

  { --------------------------------- }

var
  aIndex: Integer;
  aSeqNo: String;
begin
  aSeqNo := GetSeqNo;
  OraWriter.SQL.Text :=
    ' insert into send_nagra (  ' +
    '   high_level_cmd_id,      ' + //1
    '   icc_no,                 ' + //2
    '   stb_no,                 ' + //3
    '   zip_code,               ' + //4
    '   cmd_status,             ' + //5
    '   operator,               ' + //6
    '   updtime,                ' +
    '   seqno,                  ' + //7
    '   compcode,               ' + //8
    '   resenttimes,            ' + //9
    '   stb_flag,               ' + //10
    '   notes   )               ' + //11
    ' values (                  ' +
    '   :1, :2, :3, :4,         ' +
    '   :5, :6, sysdate,        ' +
    '   :7, :8, :9, :10, :11 )  '; 
  {}
  for aIndex := 0 to OraWriter.Params.Count - 1 do
  begin
    OraWriter.Params.Items[aIndex].DataType := ftString;
    OraWriter.Params.Items[aIndex].ParamType := ptInput;
  end;
  {}
  OraWriter.ParamByName( '1' ).AsString := 'A9';
  OraWriter.ParamByName( '2' ).AsString := EmptyStr;
  OraWriter.ParamByName( '3' ).AsString := EmptyStr;
  OraWriter.ParamByName( '4' ).AsString := txtA9ZipCode.Text;
  OraWriter.ParamByName( '5' ).AsString := 'W';
  OraWriter.ParamByName( '6' ).AsString := 'Nick';
  OraWriter.ParamByName( '7' ).AsString := aSeqNo;
  OraWriter.ParamByName( '8' ).AsString := FSoInfo.CompCode;
  OraWriter.ParamByName( '9' ).AsString := '0';
  OraWriter.ParamByName( '10' ).AsString := '1';
  OraWriter.ParamByName( '11' ).AsString := BuildNotes;
  {}
  OraWriter.Execute;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.WriteB1;

  { --------------------------------- }

  function BuildNotes: String;
  var
    aProducts, aPId: String;
  begin
    Result := EmptyStr;
    aProducts := txtB1ProductId.Text;
    repeat
      aPId := Trim( ExtractValue( aProducts ) );
      if ( aPId <> EmptyStr ) then
      begin
        Result :=( Result + Format( 'B~%s~%s~%s',
          [aPId,
           TrimChar( PadDateText( txtB1StartDate.Text ), ['/'] ),
           TrimChar( PadDateText( txtB1EndDate.Text ), ['/'] )] ) + ',' );
      end;
    until ( aProducts = EmptyStr );
    if IsDelimiter( ',', Result, Length( Result ) ) then
      Delete( Result, Length( Result ), 1 );
  end;

  { --------------------------------- }

var
  aIndex: Integer;
  aSeqNo: String;
begin
  aSeqNo := GetSeqNo;
  OraWriter.SQL.Text :=
    ' insert into send_nagra (  ' +
    '   high_level_cmd_id,      ' + //1
    '   icc_no,                 ' + //2
    '   stb_no,                 ' + //3
    '   cmd_status,             ' + //5
    '   operator,               ' + //6
    '   updtime,                ' +
    '   seqno,                  ' + //7
    '   compcode,               ' + //8
    '   resenttimes,            ' + //9
    '   stb_flag,               ' + //10
    '   notes   )               ' + //11
    ' values (                  ' +
    '   :1, :2, :3,             ' +
    '   :5, :6, sysdate,        ' +
    '   :7, :8, :9, :10, :11 )  '; 
  {}
  for aIndex := 0 to OraWriter.Params.Count - 1 do
  begin
    OraWriter.Params.Items[aIndex].DataType := ftString;
    OraWriter.Params.Items[aIndex].ParamType := ptInput;
  end;
  {}
  OraWriter.ParamByName( '1' ).AsString := 'B1';
  OraWriter.ParamByName( '2' ).AsString := txtNewSmartCard.Text;
  OraWriter.ParamByName( '3' ).AsString := EmptyStr;
  OraWriter.ParamByName( '5' ).AsString := 'W';
  OraWriter.ParamByName( '6' ).AsString := 'Nick';
  OraWriter.ParamByName( '7' ).AsString := aSeqNo;
  OraWriter.ParamByName( '8' ).AsString := FSoInfo.CompCode;
  OraWriter.ParamByName( '9' ).AsString := '0';
  OraWriter.ParamByName( '10' ).AsString := '1';
  OraWriter.ParamByName( '11' ).AsString := BuildNotes;
  {}
  OraWriter.Execute;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.WriteB2;

  { --------------------------------- }

  function BuildNotes: String;
  var
    aProducts, aPId: String;
  begin
    Result := EmptyStr;
    aProducts := txtB2ProductId.Text;
    repeat
      aPId := Trim( ExtractValue( aProducts ) );
      if ( aPId <> EmptyStr ) then
      begin
        Result := ( Result + aPId + ',' );
      end;
    until ( aProducts = EmptyStr );
    if IsDelimiter( ',', Result, Length( Result ) ) then
      Delete( Result, Length( Result ), 1 );
  end;

  { --------------------------------- }

var
  aIndex: Integer;
  aSeqNo: String;
begin
  aSeqNo := GetSeqNo;
  OraWriter.SQL.Text :=
    ' insert into send_nagra (  ' +
    '   high_level_cmd_id,      ' + //1
    '   icc_no,                 ' + //2
    '   stb_no,                 ' + //3
    '   cmd_status,             ' + //5
    '   operator,               ' + //6
    '   updtime,                ' +
    '   seqno,                  ' + //7
    '   compcode,               ' + //8
    '   resenttimes,            ' + //9
    '   stb_flag,               ' + //10
    '   notes   )               ' + //11
    ' values (                  ' +
    '   :1, :2, :3,             ' +
    '   :5, :6, sysdate,        ' +
    '   :7, :8, :9, :10, :11 )  ';
  {}
  for aIndex := 0 to OraWriter.Params.Count - 1 do
  begin
    OraWriter.Params.Items[aIndex].DataType := ftString;
    OraWriter.Params.Items[aIndex].ParamType := ptInput;
  end;
  {}
  OraWriter.ParamByName( '1' ).AsString := 'B2';
  OraWriter.ParamByName( '2' ).AsString := txtNewSmartCard.Text;
  OraWriter.ParamByName( '3' ).AsString := EmptyStr;
  OraWriter.ParamByName( '5' ).AsString := 'W';
  OraWriter.ParamByName( '6' ).AsString := 'Nick';
  OraWriter.ParamByName( '7' ).AsString := aSeqNo;
  OraWriter.ParamByName( '8' ).AsString := FSoInfo.CompCode;
  OraWriter.ParamByName( '9' ).AsString := '0';
  OraWriter.ParamByName( '10' ).AsString := '1';
  OraWriter.ParamByName( '11' ).AsString := BuildNotes;
  {}
  OraWriter.Execute;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.WriteB3;
var
  aIndex: Integer;
  aSeqNo: String;
begin
  aSeqNo := GetSeqNo;
  OraWriter.SQL.Text :=
    ' insert into send_nagra (  ' +
    '   high_level_cmd_id,      ' + //1
    '   icc_no,                 ' + //2
    '   stb_no,                 ' + //3
    '   cmd_status,             ' + //5
    '   operator,               ' + //6
    '   updtime,                ' +
    '   seqno,                  ' + //7
    '   compcode,               ' + //8
    '   resenttimes,            ' + //9
    '   stb_flag )              ' + //10
    ' values (                  ' +
    '   :1, :2, :3,             ' +
    '   :5, :6, sysdate,        ' +
    '   :7, :8, :9, :10 )       ';
  {}
  for aIndex := 0 to OraWriter.Params.Count - 1 do
  begin
    OraWriter.Params.Items[aIndex].DataType := ftString;
    OraWriter.Params.Items[aIndex].ParamType := ptInput;
  end;
  {}
  OraWriter.ParamByName( '1' ).AsString := 'B3';
  OraWriter.ParamByName( '2' ).AsString := txtNewSmartCard.Text;
  OraWriter.ParamByName( '3' ).AsString := EmptyStr;
  OraWriter.ParamByName( '5' ).AsString := 'W';
  OraWriter.ParamByName( '6' ).AsString := 'Nick';
  OraWriter.ParamByName( '7' ).AsString := aSeqNo;
  OraWriter.ParamByName( '8' ).AsString := FSoInfo.CompCode;
  OraWriter.ParamByName( '9' ).AsString := '0';
  OraWriter.ParamByName( '10' ).AsString := '1';
  {}
  OraWriter.Execute;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.WriteB4;

  { --------------------------------- }

  function BuildNotes: String;
  var
    aProducts, aPId: String;
  begin
    Result := EmptyStr;
    aProducts := txtB4ProductId.Text;
    repeat
      aPId := Trim( ExtractValue( aProducts ) );
      if ( aPId <> EmptyStr ) then
      begin
        Result := ( Result + aPId + ',' );
      end;
    until ( aProducts = EmptyStr );
    if IsDelimiter( ',', Result, Length( Result ) ) then
      Delete( Result, Length( Result ), 1 );
  end;

  { --------------------------------- }

var
  aIndex: Integer;
  aSeqNo: String;
begin
  aSeqNo := GetSeqNo;
  OraWriter.SQL.Text :=
    ' insert into send_nagra (  ' +
    '   high_level_cmd_id,      ' + //1
    '   icc_no,                 ' + //2
    '   stb_no,                 ' + //3
    '   cmd_status,             ' + //5
    '   operator,               ' + //6
    '   updtime,                ' +
    '   seqno,                  ' + //7
    '   compcode,               ' + //8
    '   resenttimes,            ' + //9
    '   stb_flag,               ' + //10
    '   notes   )               ' + //11
    ' values (                  ' +
    '   :1, :2, :3,             ' +
    '   :5, :6, sysdate,        ' +
    '   :7, :8, :9, :10, :11 )  ';
  {}
  for aIndex := 0 to OraWriter.Params.Count - 1 do
  begin
    OraWriter.Params.Items[aIndex].DataType := ftString;
    OraWriter.Params.Items[aIndex].ParamType := ptInput;
  end;
  {}
  OraWriter.ParamByName( '1' ).AsString := 'B4';
  OraWriter.ParamByName( '2' ).AsString := txtNewSmartCard.Text;
  OraWriter.ParamByName( '3' ).AsString := EmptyStr;
  OraWriter.ParamByName( '5' ).AsString := 'W';
  OraWriter.ParamByName( '6' ).AsString := 'Nick';
  OraWriter.ParamByName( '7' ).AsString := aSeqNo;
  OraWriter.ParamByName( '8' ).AsString := FSoInfo.CompCode;
  OraWriter.ParamByName( '9' ).AsString := '0';
  OraWriter.ParamByName( '10' ).AsString := '1';
  OraWriter.ParamByName( '11' ).AsString := BuildNotes;
  {}
  OraWriter.Execute;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.WriteB5;

  { --------------------------------- }

  function BuildNotes: String;
  var
    aProducts, aPId: String;
  begin
    Result := EmptyStr;
    aProducts := txtB5ProductId.Text;
    repeat
      aPId := Trim( ExtractValue( aProducts ) );
      if ( aPId <> EmptyStr ) then
      begin
        Result :=( Result + Format( 'B~%s~%s~%s',
          [aPId,
           TrimChar( PadDateText( txtB5StartDate.Text ), ['/'] ),
           TrimChar( PadDateText( txtB5EndDate.Text ), ['/'] )] ) + ',' );
      end;
    until ( aProducts = EmptyStr );
    if IsDelimiter( ',', Result, Length( Result ) ) then
      Delete( Result, Length( Result ), 1 );
  end;

  { --------------------------------- }

var
  aIndex: Integer;
  aSeqNo: String;
begin
  aSeqNo := GetSeqNo;
  OraWriter.SQL.Text :=
    ' insert into send_nagra (  ' +
    '   high_level_cmd_id,      ' + //1
    '   icc_no,                 ' + //2
    '   stb_no,                 ' + //3
    '   cmd_status,             ' + //5
    '   operator,               ' + //6
    '   updtime,                ' +
    '   seqno,                  ' + //7
    '   compcode,               ' + //8
    '   resenttimes,            ' + //9
    '   stb_flag,               ' + //10
    '   notes   )               ' + //11
    ' values (                  ' +
    '   :1, :2, :3,             ' +
    '   :5, :6, sysdate,        ' +
    '   :7, :8, :9, :10, :11 )  ';
  {}
  for aIndex := 0 to OraWriter.Params.Count - 1 do
  begin
    OraWriter.Params.Items[aIndex].DataType := ftString;
    OraWriter.Params.Items[aIndex].ParamType := ptInput;
  end;
  {}
  OraWriter.ParamByName( '1' ).AsString := 'B5';
  OraWriter.ParamByName( '2' ).AsString := txtNewSmartCard.Text;
  OraWriter.ParamByName( '3' ).AsString := EmptyStr;
  OraWriter.ParamByName( '5' ).AsString := 'W';
  OraWriter.ParamByName( '6' ).AsString := 'Nick';
  OraWriter.ParamByName( '7' ).AsString := aSeqNo;
  OraWriter.ParamByName( '8' ).AsString := FSoInfo.CompCode;
  OraWriter.ParamByName( '9' ).AsString := '0';
  OraWriter.ParamByName( '10' ).AsString := '1';
  OraWriter.ParamByName( '11' ).AsString := BuildNotes;
  {}
  OraWriter.Execute;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.WriteB6;

  { --------------------------------- }

  function BuildNotes: String;
  var
    aProducts, aPId: String;
  begin
    Result := EmptyStr;
    aProducts := txtB6ProductId.Text;
    repeat
      aPId := Trim( ExtractValue( aProducts ) );
      if ( aPId <> EmptyStr ) then
      begin
        Result :=( Result + Format( 'B~%s~%s~%s',
          [aPId,
           FormatDateTime( 'yyyymmdd', Date ),
           TrimChar( PadDateText( txtB6EndDate.Text ), ['/'] )] ) + ',' );
      end;
    until ( aProducts = EmptyStr );
    if IsDelimiter( ',', Result, Length( Result ) ) then
      Delete( Result, Length( Result ), 1 );
  end;

  { --------------------------------- }

var
  aIndex: Integer;
  aSeqNo: String;
begin
  aSeqNo := GetSeqNo;
  OraWriter.SQL.Text :=
    ' insert into send_nagra (  ' +
    '   high_level_cmd_id,      ' + //1
    '   icc_no,                 ' + //2
    '   stb_no,                 ' + //3
    '   cmd_status,             ' + //5
    '   operator,               ' + //6
    '   updtime,                ' +
    '   seqno,                  ' + //7
    '   compcode,               ' + //8
    '   resenttimes,            ' + //9
    '   stb_flag,               ' + //10
    '   notes   )               ' + //11
    ' values (                  ' +
    '   :1, :2, :3,             ' +
    '   :5, :6, sysdate,        ' +
    '   :7, :8, :9, :10, :11 )  '; 
  {}
  for aIndex := 0 to OraWriter.Params.Count - 1 do
  begin
    OraWriter.Params.Items[aIndex].DataType := ftString;
    OraWriter.Params.Items[aIndex].ParamType := ptInput;
  end;
  {}
  OraWriter.ParamByName( '1' ).AsString := 'B6';
  OraWriter.ParamByName( '2' ).AsString := txtNewSmartCard.Text;
  OraWriter.ParamByName( '3' ).AsString := EmptyStr;
  OraWriter.ParamByName( '5' ).AsString := 'W';
  OraWriter.ParamByName( '6' ).AsString := 'Nick';
  OraWriter.ParamByName( '7' ).AsString := aSeqNo;
  OraWriter.ParamByName( '8' ).AsString := FSoInfo.CompCode;
  OraWriter.ParamByName( '9' ).AsString := '0';
  OraWriter.ParamByName( '10' ).AsString := '1';
  OraWriter.ParamByName( '11' ).AsString := BuildNotes;
  {}
  OraWriter.Execute;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.WriteB7;

  { --------------------------------- }

  function BuildNotes: String;
  var
    aNote1, aNote2: String;
    aIccs, aIId: String;
    aProducts, aPId: String;
  begin
    aNote1 := EmptyStr;
    aIccs := txtB7IccNo.Text;
    repeat
      aIId := Trim( ExtractValue( aIccs ) );
      if ( aIId <> EmptyStr ) then
      begin
        aNote1 :=( aNote1 + aIId + ';' );
      end;
    until ( aIccs = EmptyStr );
    if IsDelimiter( ';', aNote1, Length( aNote1 ) ) then
      Delete( aNote1, Length( aNote1 ), 1 );
    {}
    aNote2 := EmptyStr;
    aProducts := txtB7ProductId.Text;
    repeat
      aPId := Trim( ExtractValue( aProducts ) );
      if ( aPId <> EmptyStr ) then
      begin
        aNote2 :=( aNote2 + Format( '%s^%s^%s',
          [aPId,
           TrimChar( PadDateText( txtB7StartDate.Text ), ['/'] ),
           TrimChar( PadDateText( txtB7EndDate.Text ), ['/'] )] ) + ',' );
      end;
    until ( aProducts = EmptyStr );
    if IsDelimiter( ',', aNote2, Length( aNote2 ) ) then
      Delete( aNote2, Length( aNote2 ), 1 );
    Result := ( aNote2 + '~' + aNote1 );
  end;

  { --------------------------------- }

var
  aIndex: Integer;
  aSeqNo: String;
begin
  aSeqNo := GetSeqNo;
  OraWriter.SQL.Text :=
    ' insert into send_nagra (  ' +
    '   high_level_cmd_id,      ' + //1
    '   icc_no,                 ' + //2
    '   stb_no,                 ' + //3
    '   cmd_status,             ' + //5
    '   operator,               ' + //6
    '   updtime,                ' +
    '   seqno,                  ' + //7
    '   compcode,               ' + //8
    '   resenttimes,            ' + //9
    '   stb_flag,               ' + //10
    '   notes   )               ' + //11
    ' values (                  ' +
    '   :1, :2, :3,             ' +
    '   :5, :6, sysdate,        ' +
    '   :7, :8, :9, :10, :11 )  '; 
  {}
  for aIndex := 0 to OraWriter.Params.Count - 1 do
  begin
    OraWriter.Params.Items[aIndex].DataType := ftString;
    OraWriter.Params.Items[aIndex].ParamType := ptInput;
  end;
  {}
  OraWriter.ParamByName( '1' ).AsString := 'B7';
  OraWriter.ParamByName( '2' ).AsString := EmptyStr;
  OraWriter.ParamByName( '3' ).AsString := EmptyStr;
  OraWriter.ParamByName( '5' ).AsString := 'W';
  OraWriter.ParamByName( '6' ).AsString := 'Nick';
  OraWriter.ParamByName( '7' ).AsString := aSeqNo;
  OraWriter.ParamByName( '8' ).AsString := FSoInfo.CompCode;
  OraWriter.ParamByName( '9' ).AsString := '0';
  OraWriter.ParamByName( '10' ).AsString := '1';
  OraWriter.ParamByName( '11' ).AsString := BuildNotes;
  {}
  OraWriter.Execute;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.WriteB9;

  { --------------------------------- }

  function BuildNotes: String;
  var
    aProducts, aPId: String;
  begin
    Result := EmptyStr;
    aProducts := txtB9ProductId.Text;
    repeat
      aPId := Trim( ExtractValue( aProducts ) );
      if ( aPId <> EmptyStr ) then
      begin
        Result :=( Result + Format( 'B~%s~%s~%s',
          [aPId,
           TrimChar( PadDateText( txtB9StartDate.Text ), ['/'] ),
           TrimChar( PadDateText( txtB9EndDate.Text ), ['/'] )] ) + ',' );
      end;
    until ( aProducts = EmptyStr );
    if IsDelimiter( ',', Result, Length( Result ) ) then
      Delete( Result, Length( Result ), 1 );
  end;

  { --------------------------------- }

var
  aIndex: Integer;
  aSeqNo: String;
begin
  aSeqNo := GetSeqNo;
  OraWriter.SQL.Text :=
    ' insert into send_nagra (  ' +
    '   high_level_cmd_id,      ' + //1
    '   icc_no,                 ' + //2
    '   stb_no,                 ' + //3
    '   zip_code,               ' + //4
    '   mis_ird_cmd_data,       ' + //5
    '   pincode,                ' + //6
    '   cmd_status,             ' + //7
    '   operator,               ' + //8
    '   updtime,                ' +
    '   seqno,                  ' + //9
    '   compcode,               ' + //10
    '   resenttimes,            ' + //11
    '   stb_flag,               ' + //12
    '   notes   )               ' + //13
    ' values (                  ' +
    '   :1, :2, :3, :4,         ' +
    '   :5, :6, :7, :8,         ' +
    '   sysdate, :9, :10, :11,  ' +
    '   :12, :13 )              ';
  {}
  for aIndex := 0 to OraWriter.Params.Count - 1 do
  begin
    OraWriter.Params.Items[aIndex].DataType := ftString;
    OraWriter.Params.Items[aIndex].ParamType := ptInput;
  end;
  {}
  OraWriter.ParamByName( '1' ).AsString := 'B9';
  OraWriter.ParamByName( '2' ).AsString := txtNewSmartCard.Text;
  OraWriter.ParamByName( '3' ).AsString := txtNewStepBox.Text;
  OraWriter.ParamByName( '4' ).AsString := txtB9ZipCode.Text;
  OraWriter.ParamByName( '5' ).AsString := txtB9BatId.Text;
  OraWriter.ParamByName( '6' ).AsString := '0000';
  OraWriter.ParamByName( '7' ).AsString := 'W';
  OraWriter.ParamByName( '8' ).AsString := 'Nick';
  OraWriter.ParamByName( '9' ).AsString := aSeqNo;
  OraWriter.ParamByName( '10' ).AsString := FSoInfo.CompCode;
  OraWriter.ParamByName( '11' ).AsString := '0';
  OraWriter.ParamByName( '12' ).AsString := '1';
  OraWriter.ParamByName( '13' ).AsString := BuildNotes;
  {}
  OraWriter.Execute;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.WriteC4;
var
  aIndex: Integer;
  aSeqNo: String;
begin
  aSeqNo := GetSeqNo;
  OraWriter.SQL.Text :=
    ' insert into send_nagra (  ' +
    '   high_level_cmd_id,      ' + //1
    '   icc_no,                 ' + //2
    '   stb_no,                 ' + //3
    '   zip_code,               ' + //4
    '   cmd_status,             ' + //5
    '   operator,               ' + //6
    '   updtime,                ' +
    '   seqno,                  ' + //7
    '   compcode,               ' + //8
    '   resenttimes,            ' + //9
    '   stb_flag )              ' + //10
    ' values (                  ' +
    '   :1, :2, :3, :4,         ' +
    '   :5, :6, sysdate,        ' +
    '   :7, :8, :9, :10  )      ';
  {}  
  for aIndex := 0 to OraWriter.Params.Count - 1 do
  begin
    OraWriter.Params.Items[aIndex].DataType := ftString;
    OraWriter.Params.Items[aIndex].ParamType := ptInput;
  end;
  {}
  OraWriter.ParamByName( '1' ).AsString := 'C4';
  OraWriter.ParamByName( '2' ).AsString := txtNewSmartCard.Text;
  OraWriter.ParamByName( '3' ).AsString := EmptyStr;
  OraWriter.ParamByName( '4' ).AsString := txtC4ZipCode.Text;
  OraWriter.ParamByName( '5' ).AsString := 'W';
  OraWriter.ParamByName( '6' ).AsString := 'Nick';
  OraWriter.ParamByName( '7' ).AsString := aSeqNo;
  OraWriter.ParamByName( '8' ).AsString := FSoInfo.CompCode;
  OraWriter.ParamByName( '9' ).AsString := '0';
  OraWriter.ParamByName( '10' ).AsString := '1';
  {}
  OraWriter.Execute;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.WriteE1;
var
  aIndex: Integer;
  aSeqNo: String;
begin
  aSeqNo := GetSeqNo;
  OraWriter.SQL.Text :=
    ' insert into send_nagra (  ' +
    '   high_level_cmd_id,      ' + //1
    '   icc_no,                 ' + //2
    '   stb_no,                 ' + //3
    '   pincode,                ' + //4
    '   cmd_status,             ' + //5
    '   operator,               ' + //6
    '   updtime,                ' +
    '   seqno,                  ' + //7
    '   compcode,               ' + //8
    '   resenttimes,            ' + //9
    '   stb_flag )              ' + //10
    ' values (                  ' +
    '   :1, :2, :3, :4,         ' +
    '   :5, :6, sysdate,        ' +
    '   :7, :8, :9, :10  )      ';
  {}  
  for aIndex := 0 to OraWriter.Params.Count - 1 do
  begin
    OraWriter.Params.Items[aIndex].DataType := ftString;
    OraWriter.Params.Items[aIndex].ParamType := ptInput;
  end;
  {}
  OraWriter.ParamByName( '1' ).AsString := 'E1';
  OraWriter.ParamByName( '2' ).AsString := txtNewSmartCard.Text;
  OraWriter.ParamByName( '3' ).AsString := txtNewStepBox.Text;
  OraWriter.ParamByName( '4' ).AsString := txtE1PinCode.Text;
  OraWriter.ParamByName( '5' ).AsString := 'W';
  OraWriter.ParamByName( '6' ).AsString := 'Nick';
  OraWriter.ParamByName( '7' ).AsString := aSeqNo;
  OraWriter.ParamByName( '8' ).AsString := FSoInfo.CompCode;
  OraWriter.ParamByName( '9' ).AsString := '0';
  OraWriter.ParamByName( '10' ).AsString := '1';
  {}
  OraWriter.Execute;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.WriteE2;
var
  aIndex: Integer;
  aSeqNo: String;
begin
  aSeqNo := GetSeqNo;
  OraWriter.SQL.Text :=
    ' insert into send_nagra (  ' +
    '   high_level_cmd_id,      ' + //1
    '   icc_no,                 ' + //2
    '   stb_no,                 ' + //3
    '   mis_ird_cmd_data,       ' + //4
    '   cmd_status,             ' + //5
    '   operator,               ' + //6
    '   updtime,                ' +
    '   seqno,                  ' + //7
    '   compcode,               ' + //8
    '   resenttimes,            ' + //9
    '   stb_flag,               ' + //10
    '   notes  )                ' + //11
    ' values (                  ' +
    '   :1, :2, :3, :4,         ' +
    '   :5, :6, sysdate,        ' +
    '   :7, :8, :9, :10, :11  ) ';
  {}  
  for aIndex := 0 to OraWriter.Params.Count - 1 do
  begin
    OraWriter.Params.Items[aIndex].DataType := ftString;
    OraWriter.Params.Items[aIndex].ParamType := ptInput;
  end;
  {}
  OraWriter.ParamByName( '1' ).AsString := 'E2';
  OraWriter.ParamByName( '2' ).AsString := txtNewSmartCard.Text;
  OraWriter.ParamByName( '3' ).AsString := txtNewStepBox.Text;
  OraWriter.ParamByName( '4' ).AsString := IntToStr( txtE2Priority.ItemIndex );
  OraWriter.ParamByName( '5' ).AsString := 'W';
  OraWriter.ParamByName( '6' ).AsString := 'Nick';
  OraWriter.ParamByName( '7' ).AsString := aSeqNo;
  OraWriter.ParamByName( '8' ).AsString := FSoInfo.CompCode;
  OraWriter.ParamByName( '9' ).AsString := '0';
  OraWriter.ParamByName( '10' ).AsString := '1';
  OraWriter.ParamByName( '11' ).AsString := txtE2Message.Text;
  {}
  OraWriter.Execute;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.WriteE3;

  { --------------------------------- }

  function BuildMisIrdData: String;
  begin
    Result := Format( '%s~%s~%s',
      [txtE3NetworkId.Text,
       txtE3TransportId.Text,
       txtE3ServiceId.Text] );
  end;

  { --------------------------------- }
    
var
  aIndex: Integer;
  aSeqNo: String;
begin
  aSeqNo := GetSeqNo;
  OraWriter.SQL.Text :=
    ' insert into send_nagra (  ' +
    '   high_level_cmd_id,      ' + //1
    '   icc_no,                 ' + //2
    '   stb_no,                 ' + //3
    '   mis_ird_cmd_data,       ' + //4
    '   cmd_status,             ' + //5
    '   operator,               ' + //6
    '   updtime,                ' +
    '   seqno,                  ' + //7
    '   compcode,               ' + //8
    '   resenttimes,            ' + //9
    '   stb_flag,               ' + //10
    '   notes  )                ' + //11
    ' values (                  ' +
    '   :1, :2, :3, :4,         ' +
    '   :5, :6, sysdate,        ' +
    '   :7, :8, :9, :10, :11  ) ';
  {}  
  for aIndex := 0 to OraWriter.Params.Count - 1 do
  begin
    OraWriter.Params.Items[aIndex].DataType := ftString;
    OraWriter.Params.Items[aIndex].ParamType := ptInput;
  end;
  {}
  OraWriter.ParamByName( '1' ).AsString := 'E3';
  OraWriter.ParamByName( '2' ).AsString := txtNewSmartCard.Text;
  OraWriter.ParamByName( '3' ).AsString := txtNewStepBox.Text;
  OraWriter.ParamByName( '4' ).AsString := BuildMisIrdData;
  OraWriter.ParamByName( '5' ).AsString := 'W';
  OraWriter.ParamByName( '6' ).AsString := 'Nick';
  OraWriter.ParamByName( '7' ).AsString := aSeqNo;
  OraWriter.ParamByName( '8' ).AsString := FSoInfo.CompCode;
  OraWriter.ParamByName( '9' ).AsString := '0';
  OraWriter.ParamByName( '10' ).AsString := '1';
  OraWriter.ParamByName( '11' ).AsString := EmptyStr;
  {}
  OraWriter.Execute;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.WriteE9;
var
  aIndex: Integer;
  aSeqNo: String;
begin
  aSeqNo := GetSeqNo;
  OraWriter.SQL.Text :=
    ' insert into send_nagra (  ' +
    '   high_level_cmd_id,      ' + //1
    '   icc_no,                 ' + //2
    '   stb_no,                 ' + //3
    '   mis_ird_cmd_data,       ' + //4
    '   cmd_status,             ' + //5
    '   operator,               ' + //6
    '   updtime,                ' +
    '   seqno,                  ' + //7
    '   compcode,               ' + //8
    '   resenttimes,            ' + //9
    '   stb_flag,               ' + //10
    '   notes  )                ' + //11
    ' values (                  ' +
    '   :1, :2, :3, :4,         ' +
    '   :5, :6, sysdate,        ' +
    '   :7, :8, :9, :10, :11  ) ';
  {}  
  for aIndex := 0 to OraWriter.Params.Count - 1 do
  begin
    OraWriter.Params.Items[aIndex].DataType := ftString;
    OraWriter.Params.Items[aIndex].ParamType := ptInput;
  end;
  {}
  OraWriter.ParamByName( '1' ).AsString := 'E9';
  OraWriter.ParamByName( '2' ).AsString := txtNewSmartCard.Text;
  OraWriter.ParamByName( '3' ).AsString := txtNewStepBox.Text;
  OraWriter.ParamByName( '4' ).AsString := txtE9BatId.Text;
  OraWriter.ParamByName( '5' ).AsString := 'W';
  OraWriter.ParamByName( '6' ).AsString := 'Nick';
  OraWriter.ParamByName( '7' ).AsString := aSeqNo;
  OraWriter.ParamByName( '8' ).AsString := FSoInfo.CompCode;
  OraWriter.ParamByName( '9' ).AsString := '0';
  OraWriter.ParamByName( '10' ).AsString := '1';
  OraWriter.ParamByName( '11' ).AsString := EmptyStr;
  {}
  OraWriter.Execute;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.btnConnectClick(Sender: TObject);
begin
  if OraSession.Connected then
  begin
    DisconnectFromOracle;
    btnWriteSendNagra.Enabled := False;
    cmbTNS.Enabled := True;
    txtDbUserId.Enabled := True;
    txtDbPassword.Enabled := True;
    btnConnect.Caption := '連線';
    lblConnectStatus.Caption := '已離線';
  end else
  begin
    if not VdDbUserIdAndPassword then Exit;
    FSoInfo.DbAliase := cmbTNS.Text;
    FSoInfo.DbAccount := txtDbUserId.Text;
    FSoInfo.DbPassword := txtDbPassword.Text;
    FSoInfo.ConnectionString := Format( '%s/%s@%s', [
      FSoInfo.DbAccount, FSoInfo.DbPassword, FSoInfo.DbAliase] );
    if ( ConnectToOracle ) then
    begin
      FSoInfo.CompCode := GetCompCode;
      btnWriteSendNagra.Enabled := True;
      cmbTNS.Enabled := False;
      txtDbUserId.Enabled := False;
      txtDbPassword.Enabled := False;
      btnConnect.Caption := '離線';
      lblConnectStatus.Caption := '已連線成功';
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.CommandPageChange(Sender: TObject);
var
  aCmdText: String;
begin
  aCmdText := CommandPage.ActivePage.Name;
  if IsNotNeedParameterCommand( aCmdText ) then
    lbNoParamNeeded.Parent := CommandPage.ActivePage;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.btnWriteSendNagraClick(Sender: TObject);

  { -------------------------------------- }

  function WaitExecute(var aStatus, aCode, aMsg: String): Boolean;
  begin
    aStatus := EmptyStr;
    aCode := EmptyStr;
    aMsg := EmptyStr;
    {}
    OraReader.Close;
    OraReader.SQL.Text := Format(
      ' SELECT CMD_STATUS, ERR_CODE, ERR_MSG ' +
      '   FROM SEND_NAGRA                    ' +
      '  WHERE COMPCODE = ''%s''             ' +
      '    AND SEQNO = ''%s''                ',
      [FSoInfo.CompCode, FCurrentActiveSeqNo] );
    OraReader.Open;
    aStatus := OraReader.FieldByName( 'CMD_STATUS' ).AsString;
    Result := ( aStatus = 'C' ) or ( aStatus = 'E' );
    if Result  then
    begin
      aCode := OraReader.FieldByName( 'ERR_CODE' ).AsString;
      aMsg :=  OraReader.FieldByName( 'ERR_MSG' ).AsString;
    end;
    OraReader.Open;
  end;

  { -------------------------------------- }

  procedure WriteTimeOut;
  begin
    OraReader.Close;
    OraReader.SQL.Text := Format(
      ' UPDATE SEND_NAGRA               ' +
      '    SET CMD_STATUS = ''E'',      ' +
      '        ERR_CODE = ''J01'',      ' +
      '        ERR_MSG = ''作業逾時''   ' +  
      '  WHERE COMPCODE = ''%s''        ' +
      '    AND SEQNO = ''%s''           ',
      [FSoInfo.CompCode, FCurrentActiveSeqNo] );
    OraReader.Execute;
  end;

  { -------------------------------------- }
  
var
  aCmdStatus, aErrCode, aErrMsg: String;
  aExecuteTime: TDateTime;
begin
  if not VdSmartCardAndStb( CommandPage.ActivePage.Name ) then
    Exit;
  if not VdParameter( CommandPage.ActivePage.Name ) then
    Exit;
  {}
  Screen.Cursor := crSQLWait;
  try
    btnWriteSendNagra.Enabled := False;
    try
      aExecuteTime := Now;
      {}
      WriteCommand( CommandPage.ActivePage.Name  );
      {}
      lbSendResult.Caption := '等候Gateway回應....';
      lbSendResult.Font.Color := clWindowText;
      lbErrCode.Caption := 'N/A';
      lbErrMsg.Caption := 'N/A';
      {}
      while True do
      begin
        Delay( 1000 );
        if WaitExecute( aCmdStatus, aErrCode, aErrMsg ) then
        begin
          if ( aCmdStatus = 'C' ) then
          begin
            lbSendResult.Caption := '指令傳送結果完成';
            lbSendResult.Font.Color := clBlue;
            lbErrCode.Caption := 'N/A';
            lbErrMsg.Caption := 'N/A';
          end else
          begin
            lbSendResult.Caption := '指令傳送結果發生錯誤';
            lbSendResult.Font.Color := clRed;
            lbErrCode.Caption := aErrCode;
            lbErrMsg.Caption := aErrMsg;
          end;
          Break;
        end else
        begin
          if ( SecondsBetween( Now, aExecuteTime ) > 60 ) then
          begin
            WriteTimeOut;
          end;
        end;
        Application.ProcessMessages;
      end;
      {}
    finally
      btnWriteSendNagra.Enabled := True;
    end;  
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

end.
