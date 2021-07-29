unit cbMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, cxControls, cxContainer, cxEdit, cxTextEdit, cxMemo, ExtCtrls,
  cxLookAndFeels, Menus, cxLookAndFeelPainters, StdCtrls, cxButtons,
  IdBaseComponent, IdComponent, IdTCPServer, cxGraphics, IdStack,
  cxCustomData, cxStyles, cxTL, cxInplaceContainer, ImgList, SyncObjs, cxClasses,
  dxSkinsCore, dxSkinsDefaultPainters;

type

  TStringListEx = class(TStringList)
  protected
    procedure SetTextStr(const Value: string); override;
  end;

  TfmMain = class(TForm)
    cxLookAndFeelController1: TcxLookAndFeelController;
    ConsoleStyleController: TcxEditStyleController;
    Panel1: TPanel;
    Panel2: TPanel;
    btnRun: TcxButton;
    IdTCPServer: TIdTCPServer;
    CmdList: TcxTreeList;
    CmdCol1: TcxTreeListColumn;
    CmdCol2: TcxTreeListColumn;
    CmdCol3: TcxTreeListColumn;
    btnStop: TcxButton;
    Label1: TLabel;
    CmdImageList: TImageList;
    cxStyleRepository: TcxStyleRepository;
    cxStyle1: TcxStyle;
    cxStyle2: TcxStyle;
    cxStyle3: TcxStyle;
    cxStyle4: TcxStyle;
    cxStyle5: TcxStyle;
    cxStyle6: TcxStyle;
    cxStyle7: TcxStyle;
    cxStyle8: TcxStyle;
    cxStyle9: TcxStyle;
    cxStyle10: TcxStyle;
    cxStyle11: TcxStyle;
    cxStyle12: TcxStyle;
    cxStyle13: TcxStyle;
    cxStyle14: TcxStyle;
    cxStyle15: TcxStyle;
    cxStyle16: TcxStyle;
    TreeListStyleSheetConsoleBlack: TcxTreeListStyleSheet;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnRunClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure IdTCPServerConnect(AThread: TIdPeerThread);
    procedure IdTCPServerDisconnect(AThread: TIdPeerThread);
    procedure IdTCPServerExecute(AThread: TIdPeerThread);
    procedure btnStopClick(Sender: TObject);
  private
    { Private declarations }
    FIsRun: Boolean;
    FExecuteLock: TCriticalSection;
    FConnectStateLock: TCriticalSection;
    FList: TStringListEx;
    FRecvCmd: String;
    FAck: String;
    FActiveNode: TcxTreeListNode;
    FConnectStateThread: TIdPeerThread;
    procedure OnCmdNotify;
    procedure OnConnectNotify;
    procedure OnDisconnectNotify;
    function AddRecvCmd(const aCmd: String): TcxTreeListNode;
    procedure AddRecvCmdAck(const aNode:  TcxTreeListNode; const aAck: String);
    function BuildAck(aCmd: String; const aIsValidate: Boolean): String;
    function CheckCmd(aCmd: String): Boolean;
    procedure SocketStateChange;
  public
    { Public declarations }
    function StartSocket: Boolean;
    function EndSocket: Boolean;
  end;

var
  fmMain: TfmMain;

implementation

{$R *.dfm}

uses cbUtilis;

{ ---------------------------------------------------------------------------- }

{ TStringListEx }

procedure TStringListEx.SetTextStr(const Value: string);
var
  P, Start: PChar;
  S: string;
begin
  BeginUpdate;
  try
    Clear;
    P := Pointer(Value);
    if P <> nil then
      while P^ <> #0 do
      begin
        Start := P;
        while not (P^ in [#0, #10, #13, #59]) do Inc(P);
        SetString(S, Start, P - Start);
        Add(S);
        if P^ = #13 then Inc(P);
        if P^ = #10 then Inc(P);
        if P^ = #59 then Inc(P);
      end;
  finally
    EndUpdate;
  end;
end;
{ ---------------------------------------------------------------------------- }

procedure TfmMain.FormCreate(Sender: TObject);
begin
  FIsRun := False;
  CmdList.Clear;
  FExecuteLock := TCriticalSection.Create;
  FConnectStateLock := TCriticalSection.Create;
  FList := TStringListEx.Create;
  FList.Delimiter := ';';
  btnRun.Enabled := not FIsRun;
  btnStop.Enabled := FIsRun;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.FormDestroy(Sender: TObject);
begin
  FExecuteLock.Free;
  FConnectStateLock.Free;
  FList.Free;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  EndSocket;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.btnRunClick(Sender: TObject);
begin
  if ( StartSocket ) then
  begin
    btnRun.Enabled := False;
    btnStop.Enabled := True;
    FIsRun := True;
    SocketStateChange;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.btnStopClick(Sender: TObject);
begin
  EndSocket;
  btnStop.Enabled := False;
  btnRun.Enabled := True;
  FIsRun := False;
  SocketStateChange;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.SocketStateChange;
var
  aNode: TcxTreeListNode;
  aText: String;
  aImgIdx: Integer;
begin
  if ( IdTCPServer.Active ) then
  begin
    aText := 'DVN CA Simulator Start.';
    aImgIdx := 2;
  end else
  begin
    aText := 'DVN CA Simulator Stop.';
    aImgIdx := 3;
  end;  
  aNode := CmdList.Add;
  aNode.ImageIndex := aImgIdx;
  aNode.SelectedIndex := aImgIdx;
  aNode.Values[0] := FormatDateTime( 'yyyy-mm-dd hh:nn:ss', Now );
  aNode.Values[1] := aText;
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.StartSocket: Boolean;
begin
  if not IdTCPServer.Active then
  begin
    try
      IdTCPServer.Active := True;
    except
      on E: Exception do ErrorMsg( E.Message );
    end;
  end;
  Result := IdTCPServer.Active;  
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.EndSocket: Boolean;
begin
  if IdTCPServer.Active then
  begin
    try
      IdTCPServer.Active := False;
    except
      on E: Exception do ErrorMsg( E.Message );
    end;
  end;
  Result := ( not IdTCPServer.Active );  
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.IdTCPServerConnect(AThread: TIdPeerThread);
begin
  FConnectStateLock.Enter;
  try
    FConnectStateThread := AThread;
    AThread.Synchronize( OnConnectNotify );
    FConnectStateThread := nil;
  finally
    FConnectStateLock.Leave;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.IdTCPServerDisconnect(AThread: TIdPeerThread);
begin
  if AThread.Connection.Server.Active then
  begin
    FConnectStateLock.Enter;
    try
      FConnectStateThread := AThread;
      AThread.Synchronize( OnDisconnectNotify );
      FConnectStateThread := nil;
    finally
      FConnectStateLock.Leave;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.IdTCPServerExecute(AThread: TIdPeerThread);
var
  aByte: Integer;
  aCmdValid: Boolean;
begin
  FExecuteLock.Enter;
  try
    FRecvCmd := EmptyStr;
    FAck := EmptyStr;
    if not AThread.Terminated and AThread.Connection.Connected then
    begin
      try
        aByte := AThread.Connection.ReadFromStack( False, 100, False );
      except
        aByte := 0;
      end;
      if ( aByte > 0 ) then
      begin
        SetString( FRecvCmd, PChar( AThread.Connection.InputBuffer.Memory ), aByte );
        AThread.Connection.InputBuffer.Remove( aByte );
        {}
        AThread.Synchronize( OnCmdNotify );
        aCmdValid := CheckCmd( FRecvCmd );
        FAck := BuildAck( FRecvCmd, aCmdValid );
        FRecvCmd := EmptyStr;
        {}
        AThread.Connection.Write( FAck );
        AThread.Synchronize( OnCmdNotify );
        FAck := EmptyStr;
      end;
    end;
  finally
    FExecuteLock.Leave;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.OnCmdNotify;
begin
  if ( FRecvCmd <> EmptyStr ) then
    FActiveNode := AddRecvCmd( FRecvCmd )
  else if ( FAck <> EmptyStr ) then
    AddRecvCmdAck( FActiveNode, FAck );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.OnConnectNotify;
var
  aNode: TcxTreeListNode;
begin
  aNode := CmdList.Add;
  aNode.ImageIndex := 0;
  aNode.SelectedIndex := 0;
  aNode.Values[0] := FormatDateTime( 'yyyy-mm-dd hh:nn:ss', Now );
//  aNode.Values[1] := Format( 'SMS Gateway %s (%s) Connected.',
//    [FConnectStateThread.Connection.LocalName, GStack.LocalAddress] );
  aNode.MakeVisible;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.OnDisconnectNotify;
var
  aNode: TcxTreeListNode;
begin
  aNode := CmdList.Add;
  aNode.ImageIndex := 0;
  aNode.SelectedIndex := 0;
  aNode.Values[0] := FormatDateTime( 'yyyy-mm-dd hh:nn:ss', Now );
//  aNode.Values[1] := Format( 'SMS Gateway %s (%s) Disconnected.',
//    [FConnectStateThread.Connection.LocalName, GStack.LocalAddress] );
  aNode.MakeVisible;
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.AddRecvCmd(const aCmd: String): TcxTreeListNode;
begin
  Result := CmdList.Add;
  Result.ImageIndex := 1;
  Result.SelectedIndex := Result.ImageIndex;
  Result.Values[0] := FormatDateTime( 'yyyy-mm-dd hh:nn:ss', Now );
  Result.Values[1] := aCmd;
  Result.MakeVisible;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.AddRecvCmdAck(const aNode: TcxTreeListNode;
  const aAck: String);
begin
  if Assigned( aNode ) then
    aNode.Values[2] := aAck;
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.BuildAck(aCmd: String;  const aIsValidate: Boolean): String;
const
  aStatus: array[Boolean] of String = ( 'Invalid_Cmd', 'OK' );
var
  aPos1, aPos2: Integer;
  aBody, aFrameNo, aStb, aApi, aProtocol: String;
begin
  aPos1 := AnsiPos( '(', aCmd );
  aPos2 := AnsiPos( ')', aCmd );
  aApi := Copy( aCmd, 1, aPos1 - 1 );
  aBody := Copy( aCmd, aPos1 + 1, aPos2 -1 );
  FList.DelimitedText := aBody;
  aFrameNo := FList[0];
  aProtocol := FList[1];
  aStb := FList[2];
  Result := Format( 'ACK(%s;%s;%s;%s;%s;)',
    [aFrameNo, aProtocol, aStb, aApi, aStatus[aIsValidate]] );
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.CheckCmd(aCmd: String): Boolean;
var
  aPos1, aPos2: Integer;
  aBody, aFrameNo, aStb, aApi, aProtocol, aIcc: String;
  aStartTime, aExpiryTime, aProductType, aBankType, aProdId, aMsg: String;
  aProdCount, aKeyIdx, aNum: Integer;
  aTime1, aTime2: TDateTime;
  aActionFlag, aConfirmCode, aCodeValue: String;
  aCredit: Double;
  aText: String;
begin
  Result := False;
  aPos1 := AnsiPos( '(', aCmd );
  aPos2 := AnsiPos( ')', aCmd );
  if ( aPos1 <= 0 ) or ( aPos2 <= 0 ) then Exit;
  aApi := Copy( aCmd, 1, aPos1 - 1 );
  {}
  if not
    ( ( aApi = 'PAIRING_STB' ) or ( aApi = 'STB_ON' ) or
      ( aApi = 'STB_OFF' ) or ( aApi = 'ADD_PRODUCT' ) or
      ( aApi = 'REMOVE_PRODUCT' ) or ( aApi = 'MAIL_MESSAGE' ) or
      ( aApi = 'ADD_TOKEN' ) or ( aApi = 'DEDUCT_TOKEN' ) or
      ( aApi = 'USER_DEFINED' ) or ( aApi = 'SYS_CONFIG' )
    ) then
     Exit;
  {}
  aBody := Copy( aCmd, aPos1 + 1, ( aPos2 - aPos1 - 1 ) );
  FList.Text := aBody;
  aFrameNo := FList[0];
  if ( aFrameNo = EmptyStr ) then Exit;
  aProtocol := FList[1];
  if ( aProtocol <> '2' ) then Exit;
  aStb := FList[2];
  if ( aStb = EmptyStr ) then Exit;
  aIcc := FList[3];
  if ( aIcc = EmptyStr ) then Exit;
  {}
  if ( aApi = 'ADD_PRODUCT' ) or ( aApi = 'REMOVE_PRODUCT' ) then
  begin
    aStartTime := FList[4];
    aExpiryTime := FList[5];
    if not TryStrToDateTime( aStartTime, aTime1 ) then Exit;
    if not TryStrToDateTime( aExpiryTime, aTime2 ) then Exit;
    if ( aTime1 >= aTime2 ) then Exit;
    aBankType := FList[6];
    if ( aBankType <> '1' ) then Exit;
    aProductType := FList[7];
    if ( aProductType <> 'PP' ) then Exit;
    aText := FList[8];
    aProdCount := StrToIntDef( ExtractValue( aText ), 0 );
    aProdId := aText;
    if ( aProdId = EmptyStr ) then Exit;
    repeat
      ExtractValue( aProdId );
      Dec( aProdCount );
    until ( aProdId = EmptyStr );
    if ( aProdCount <> 0 ) then Exit;
  end else
  if ( aApi = 'PAIRING_STB' ) then
  begin
    aStartTime := FList[4];
    if not TryStrToDateTime( aStartTime, aTime1 ) then Exit;
  end else
  if ( aApi = 'MAIL_MESSAGE' ) then
  begin
    aStartTime := FList[4];
    if not TryStrToDateTime( aStartTime, aTime1 ) then Exit;
    aMsg := FList[6];
    if ( aMsg = EmptyStr ) then Exit;
    if not ( ( aMsg[1] = 'A' ) or ( aMsg[1] = 'B' ) or ( aMsg[1] = '0' ) ) then
      Exit;
  end else
  if ( aApi = 'ADD_TOKEN' ) or ( aApi = 'DEDUCT_TOKEN' ) then
  begin
    aStartTime := FList[4];
    if not TryStrToDateTime( aStartTime, aTime1 ) then Exit;
    aActionFlag := FList[7];
    if not ( ( aActionFlag = '0' ) or ( aActionFlag = '1' ) or ( aActionFlag = '2' ) ) then Exit;
    aConfirmCode := FList[8];
    if ( aConfirmCode = EmptyStr ) then Exit;
    if not TryStrToFloat( FList[9], aCredit ) then Exit;
  end else
  if ( aApi = 'SYS_CONFIG' ) then
  begin
    if not TryStrToInt( FList[4], aKeyIdx ) then Exit;
    if not ( aKeyIdx in [1..7] ) then Exit;
    aText := FList[5];
    if not TryStrToInt( ExtractValue( aText ), aNum ) then Exit;
    aCodeValue := aText;
    if ( aCodeValue = EmptyStr ) then Exit;
    repeat
      ExtractValue( aCodeValue );
      Dec( aNum );
    until ( aCodeValue = EmptyStr );
    if ( aNum <> 0 ) then Exit;
  end;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

end.
