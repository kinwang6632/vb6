unit cbMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, cxLookAndFeels, cxControls, cxContainer, cxEdit,
  cxTextEdit, ExtCtrls, Menus, cxLookAndFeelPainters, ImgList, cxButtons,
  cxGroupBox, IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient,
  cxPC, ComCtrls, cxListView, IniFiles, cxMemo, cxGraphics, cxMaskEdit,
  cxDropDownEdit, dxSkinsCore, dxSkinBlack, dxSkinBlue, dxSkinCaramel,
  dxSkinCoffee, dxSkinGlassOceans, dxSkiniMaginary, dxSkinLilian,
  dxSkinLiquidSky, dxSkinLondonLiquidSky, dxSkinMcSkin, dxSkinMoneyTwins,
  dxSkinOffice2007Black, dxSkinOffice2007Blue, dxSkinOffice2007Green,
  dxSkinOffice2007Pink, dxSkinOffice2007Silver, dxSkinSilver, dxSkinStardust,
  dxSkinsDefaultPainters, dxSkinValentine, dxSkinXmas2008Blue,
  dxSkinscxPCPainter;

type
  PDvnBuffer = ^TDvnBuffer;
  TDvnBuffer = record
    Data: array [0..1023] of Char;
  end;

  TfmMain = class(TForm)
    SocketImage: TImage;
    txtDvnIp: TcxTextEdit;
    cxLookAndFeelController1: TcxLookAndFeelController;
    Label1: TLabel;
    Label2: TLabel;
    txtDvnPort: TcxTextEdit;
    Bevel1: TBevel;
    Panel1: TPanel;
    Panel2: TPanel;
    Label8: TLabel;
    lblConnectStatus: TLabel;
    btnConnect: TcxButton;
    SocketImageList: TImageList;
    gbCmdHeader: TcxGroupBox;
    Label3: TLabel;
    txtSC: TcxTextEdit;
    Label4: TLabel;
    txtSTB: TcxTextEdit;
    Label5: TLabel;
    txtAreaCode: TcxTextEdit;
    Label6: TLabel;
    txtStartTime: TcxTextEdit;
    Label7: TLabel;
    txtExpiryTime: TcxTextEdit;
    IdSocket: TIdTCPClient;
    cmdPage: TcxPageControl;
    cxTabSheet1: TcxTabSheet;
    cxTabSheet2: TcxTabSheet;
    Label9: TLabel;
    Bevel2: TBevel;
    btnSendCmd: TcxButton;
    cxTabSheet3: TcxTabSheet;
    cxTabSheet4: TcxTabSheet;
    cxTabSheet5: TcxTabSheet;
    cxTabSheet6: TcxTabSheet;
    cxTabSheet7: TcxTabSheet;
    cxTabSheet8: TcxTabSheet;
    cxTabSheet9: TcxTabSheet;
    cxTabSheet10: TcxTabSheet;
    Label10: TLabel;
    txtFrameNumber: TcxTextEdit;
    AckList: TcxMemo;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    txtClientName: TcxTextEdit;
    txtAccountNo: TcxTextEdit;
    txtTelNo: TcxTextEdit;
    txtEmail: TcxTextEdit;
    Label15: TLabel;
    txtAddress: TcxTextEdit;
    Label16: TLabel;
    Label17: TLabel;
    Label18: TLabel;
    txtBankType: TcxTextEdit;
    Label19: TLabel;
    cmbProductType: TcxComboBox;
    Label20: TLabel;
    txtNumberProduct: TcxTextEdit;
    Label21: TLabel;
    txtProductName: TcxTextEdit;
    Label22: TLabel;
    txtGatewayName: TcxTextEdit;
    Label23: TLabel;
    txtBankType2: TcxTextEdit;
    Label24: TLabel;
    cmbProductType2: TcxComboBox;
    Label25: TLabel;
    txtNumberProduct2: TcxTextEdit;
    Label26: TLabel;
    txtProductName2: TcxTextEdit;
    Label27: TLabel;
    txtGatewayName2: TcxTextEdit;
    txtMessage: TcxMemo;
    Label28: TLabel;
    Label29: TLabel;
    txtAddAmount: TcxTextEdit;
    Label30: TLabel;
    txtDecAmount: TcxTextEdit;
    Label31: TLabel;
    txtUserDefineValue: TcxTextEdit;
    Label32: TLabel;
    cmbConfigCode: TcxComboBox;
    Label33: TLabel;
    txtNumberOfValue: TcxTextEdit;
    Label34: TLabel;
    txtConfigValue: TcxTextEdit;
    Label35: TLabel;
    cxTextEdit13: TcxTextEdit;
    Label36: TLabel;
    txtGlobalAccount: TcxTextEdit;
    Label37: TLabel;
    cmbTokenAction: TcxComboBox;
    Label38: TLabel;
    cmbTokenAction2: TcxComboBox;
    Label39: TLabel;
    txtSMSName: TcxTextEdit;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure btnConnectClick(Sender: TObject);
    procedure btnSendCmdClick(Sender: TObject);
  private
    { Private declarations }
    FrameNumber: Integer;
    FDataBuffer: PDvnBuffer;
    FTokenSequence: Integer;
    function NextFrameNumber: String;
    function CurrentFrameNumber: String;
    function NextTokenConfirmCode: String;
    function CurrentTokenConfirmCode: String;
    procedure SocketStatusChange(const aConnected: Boolean);
    procedure LoadConfg;
    procedure SaveConfig;
    procedure SendToDVN;
    procedure RecvFromDVN;
    function BuildCommand(const aIndex: Integer): String;
  public
    { Public declarations }
  end;

var
  fmMain: TfmMain;

implementation

{$R *.dfm}

uses cbUtilis;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.FormCreate(Sender: TObject);
begin
  FrameNumber := 0;
  FTokenSequence := 0;
  New( FDataBuffer );
  SocketStatusChange( IdSocket.Connected );
  AckList.Lines.Clear;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.FormShow(Sender: TObject);
begin
  LoadConfg;
  txtDvnIp.Text := '210.0.189.13';
  txtDvnPort.Text := '2100';
  txtAreaCode.Text := '0-0-0-0';
  txtFrameNumber.Text := Format( '%s-%s', ['0', CurrentFrameNumber] )
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  CanClose := True;
  SaveConfig;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.FormDestroy(Sender: TObject);
begin
  Dispose( FDataBuffer );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.SocketStatusChange(const aConnected: Boolean);
begin
  SocketImageList.GetIcon( Ord( aConnected ), SocketImage.Picture.Icon );
  lblConnectStatus.Caption := 'Disconnect';
  lblConnectStatus.Font.Color := clRed;
  btnSendCmd.Enabled := False;
  btnConnect.Caption := 'Connect';
  if aConnected then
  begin
    lblConnectStatus.Caption := 'Connected';
      lblConnectStatus.Font.Color := clBlue;
    btnSendCmd.Enabled := True;
    btnConnect.Caption := 'Disconnect';
  end;
end;

{ ---------------------------------------------------------------------------- }


function TfmMain.NextFrameNumber: String;
begin
  Inc( FrameNumber );
  Result := CurrentFrameNumber;
  txtFrameNumber.Text := Result; 
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.CurrentFrameNumber: String;
begin
  Result := IntToStr( FrameNumber );
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.CurrentTokenConfirmCode: String;
begin
  Result := '01' + FormatDateTime( 'yyyymmdd', Date ) +
    Lpad( IntToStr( FTokenSequence ), 6, '0' );
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.NextTokenConfirmCode: String;
begin
  Inc( FTokenSequence );
  if ( FTokenSequence > 9999 ) then FTokenSequence := 1;
  Result := CurrentTokenConfirmCode;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.SaveConfig;
var
  aIniFile: TIniFile;
  aFileName: String;
begin
  aFileName := IncludeTrailingPathDelimiter( ExtractFilePath(
    Application.ExeName ) ) + 'Config.ini';
  aIniFile := TIniFile.Create( aFileName );
  try
    aIniFile.WriteInteger( 'Common', 'FrameNumber', FrameNumber );
    aIniFile.WriteInteger( 'Common', 'TokenSequence', FTokenSequence );
  finally
    aIniFile.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.LoadConfg;
var
  aIniFile: TIniFile;
  aFileName: String;
begin
  aFileName := IncludeTrailingPathDelimiter( ExtractFilePath(
    Application.ExeName ) ) + 'Config.ini';
  aIniFile := TIniFile.Create( aFileName );
  try
    FrameNumber := aIniFile.ReadInteger( 'Common', 'FrameNumber', 0 );
    FTokenSequence := aIniFile.ReadInteger( 'Common', 'TokenSequence', 0 );
  finally
    aIniFile.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.btnConnectClick(Sender: TObject);
var
  aBytes: Integer;
begin
  if ( IdSocket.Connected ) then
  begin
    IdSocket.Disconnect;
  end else
  begin
    IdSocket.Host := txtDvnIp.Text;
    IdSocket.Port := StrToIntDef( txtDvnPort.Text, 2100 );
    IdSocket.Connect( 3000 );
  end;
  SocketStatusChange( IdSocket.Connected );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.SendToDVN;
var
  aCommand: String;
begin
  aCommand := BuildCommand( cmdPage.ActivePageIndex );
  ZeroMemory( FDataBuffer, SizeOf( FDataBuffer^ ) );
  Move( PChar( aCommand )^, FDataBuffer^, Length( aCommand ) );
  IdSocket.WriteBuffer( FDataBuffer^, Length( aCommand ), True );
  AckList.Lines.Add( Format( 'SEND:%s', [aCommand] ) );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.RecvFromDVN;
var
  aBytes: Integer;
  aAckText: String;
begin
  ZeroMemory( FDataBuffer, SizeOf( FDataBuffer^ ) );
  aBytes := IdSocket.ReadFromStack( False, -1, False );
  Move( IdSocket.InputBuffer.Memory^, FDataBuffer^, aBytes );
  IdSocket.InputBuffer.Remove( aBytes );
  SetString( aAckText, PChar( FDataBuffer ), aBytes );
  AckList.Lines.Add( Format( 'ACK:%s',[aAckText] ) );
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.BuildCommand(const aIndex: Integer): String;
var
  aSc, aStb, aArea, aStartTime, aExpiryTime: String;
  aFrameNo, aProductType, aBankType, aProductType2, aBankType2: String;
  aMailText, aTokenAction, aTokenAction2: String;
  aConfigCode, aSMSName: String;
begin
  aSc := txtSC.Text;
  aStb := ( '0-' + txtAreaCode.Text + '-' + txtSTB.Text );
  aStartTime := FormatDateTime( 'yyyy/mm/dd hh:nn:ss', Now );
  aExpiryTime := txtExpiryTime.Text;
  aFrameNo := Format( '0-%s', [NextFrameNumber] );
  {}
  aProductType := Nvl( cmbProductType.Text, 'GP' );
  aBankType :=  Nvl( txtBankType.Text, '0' );
  {}
  aProductType2 := Nvl( cmbProductType2.Text, 'GP' );
  aBankType2 :=  Nvl( txtBankType2.Text, '0' );
  {}
  aMailText := 'A' + TrimChar( txtMessage.Lines.Text, [#13,#10] );
  {}
  aTokenAction := Copy( cmbTokenAction.Text, 1, 1 );
  aTokenAction2 := Copy( cmbTokenAction2.Text, 1, 1 );
  {}
  aConfigCode := Copy( cmbConfigCode.Text, 1, 1 );
  {}
  case aIndex of
    { PAIRING_STB ( FrameNo; Version;Password; SmartCardNo; StartingTime;
        ExpiryTime; STB_ID ; ClientName; AccountNo; TelephoneNo; EmailID; Address ) }
    0: Result := Format(
      'PAIRING_STB(%s;%s;%s;%s;%s;%s;%s;%s;%s;%s;%s;%s)',
      [aFrameNo, '2', aStb, aSc, aStartTime, EmptyStr, aStb, txtClientName.Text,
       txtGlobalAccount.Text, txtTelNo.Text, txtEmail.Text, txtAddress.Text] );
    { STB_ON ( FrameNo; Version; STB_ID; SmartCardNo; StartingTime; ExpiryTime ) }
    1: Result := Format(
      'STB_ON(%s;%s;%s;%s;%s;%s)',
      [aFrameNo, '2', aStb, aSc, aExpiryTime, EmptyStr] );
    { STB_OFF ( FrameNo; Version; STB_ID; SmartCardNo; StartingTime; ExpiryTime ) }
    2: Result := Format(
      'STB_OFF(%s;%s;%s;%s;%s;%s)',
      [aFrameNo, '2', aStb, aSc, aExpiryTime, EmptyStr] );
    { ADD_PRODUCT ( FrameNo; Version; STB_ID; SmartCardNo; StartingTime;
        ExpiryTime; BankType; Product Type Identifier; Numbers of Products, Product Name;
        Name of Target SMS Gateway) }
    3: begin
         aSMSName := txtSMSName.Text;
         if ( aProductType = 'GP' ) then aSMSName := EmptyStr;
         aStartTime := txtStartTime.Text;
         aExpiryTime := txtExpiryTime.Text;
         Result := Format(
           'ADD_PRODUCT(%s;%s;%s;%s;%s;%s;%s;%s;%s,%s;%s)',
           [aFrameNo, '2', aStb, aSc, aStartTime, aExpiryTime, aBankType, aProductType,
           txtNumberProduct.Text, txtProductName.Text, aSMSName] );
       end;    
    { REMOVE_PRODUCT (FrameNo; Version; STB_ID; SmartCardNo; StartingTime;
        ExpiryTime; BankType; Product Type Identifier; Numbers of Products, Product Name;
        Name of Target SMS Gateway }
    4: begin
         aSMSName := txtSMSName.Text;
         if ( aProductType2 = 'GP' ) then aSMSName := EmptyStr;
         aStartTime := txtStartTime.Text;
         aExpiryTime := txtExpiryTime.Text;
         Result := Format(
           'REMOVE_PRODUCT(%s;%s;%s;%s;%s;%s;%s;%s;%s,%s;%s)',
           [aFrameNo, '2', aStb, aSc, aStartTime, aExpiryTime, aBankType2, aProductType2,
            txtNumberProduct2.Text, txtProductName2.Text, aSMSName] );
       end;     
    { MAIL_MESSAGE (FrameNo; Version; STB_ID; SmartCardNo; StartingTime;
        ExpiryTime; Content ) }
    5: Result := Format(
      'MAIL_MESSAGE(%s;%s;%s;%s;%s;%s;%s)',
      [aFrameNo, '2', aStb, aSc, aStartTime, EmptyStr, aMailText] );
    { ADD_TOKEN (FrameNo; Version; STB_ID; SmartCardNo; StartingTime; ExpiryTime;
        Account Code; Action Flag; Confirm Code;Amount ) }
    6: Result := Format(
      'ADD_TOKEN(%s;%s;%s;%s;%s;%s;%s;%s;%s;%s)',
      [aFrameNo, '2', aStb, aSc, aStartTime, EmptyStr, txtGlobalAccount.Text,
       aTokenAction, NextTokenConfirmCode, txtAddAmount.Text] );
    { DEDUCT_TOKEN (FrameNo; Version; STB_ID; SmartCardNo; StartingTime; ExpiryTime;
        Account Code; Action Flag; Confirm Code;Amount ) }
    7: Result := Format(
      'DEDUCT_TOKEN(%s;%s;%s;%s;%s;%s;%s;%s;%s;%s)',
      [aFrameNo, '2', aStb, aSc, aStartTime, EmptyStr, txtGlobalAccount.Text,
       aTokenAction2, NextTokenConfirmCode, txtDecAmount.Text] );
    { USER_DEFINED(FrameNo; Version; STB_ID; SmartCardNo; StartingTime;
         ExpiryTime; Content ) }
    8: Result := Format(
      'USER_DEFINED(%s;%s;%s;%s;%s;%s;%s)',
      [aFrameNo, '2', aStb, aSc, EmptyStr, EmptyStr, txtUserDefineValue.Text] );
    { SYS_CONFIG(FrameNo; Version; STB_ID; SmartCardNo; Config. Code ID; Numbers of
       value, Code Value; Name of target SMS Gateway) }
    9: begin
         aSMSName := txtSMSName.Text;
         Result := Format(
           'SYS_CONFIG(%s;%s;%s;%s;%s;%s,%s;%s)',
           [aFrameNo, '2', aStb, aSc, aConfigCode, txtNumberOfValue.Text,
            txtConfigValue.Text, aSMSName] );
       end;     
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.btnSendCmdClick(Sender: TObject);
begin
  SendToDVN;
  Sleep( 1500 );
  RecvFromDVN;
end;

end.
