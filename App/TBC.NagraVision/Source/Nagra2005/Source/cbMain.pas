unit cbMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  ExtCtrls, StdCtrls, ComCtrls, DB, Math, IniFiles,
  { Indy 9.0.50 }
  IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient,
  { App }
  cbStyleModule, cbConfigModule, cbLiceMgr, cbClass, cbControl, cbGuiCleanup,
  { Developer Express Common }
  cxClasses, cxGraphics, cxContainer, cxControls, cxStyles, cxCustomData, 
  { Developer Express Bar & Docking Control }
  dxBar, dxBarExtItems, dxDockControl, dxDockPanel,
  { Developer Express PageControl & TreeList }
  cxTL, cxListView,
  { Developer Express Editor }
  cxInplaceContainer, cxImageComboBox, cxEdit, cxMaskEdit, cxDropDownEdit,
  cxTextEdit, cxCheckBox, cxProgressBar,
  { Developer Express Skin Painter }
  cxLookAndFeels, cxLookAndFeelPainters, dxSkinsCore, dxSkinBlack,
  dxSkinBlue, dxSkinCaramel, dxSkinCoffee, dxSkinsDefaultPainters,
  dxSkinsdxBarPainter, dxSkinsdxDockControlPainter, Dialogs;


type
  TfmMain = class(TForm)
    dxBarManager: TdxBarManager;
    dxMenuItem1: TdxBarSubItem;
    dxPlay: TdxBarButton;
    dxStop: TdxBarButton;
    dxNagraServer: TdxBarStatic;
    dxTabContainerDockSite2: TdxTabContainerDockSite;
    dxTabContainerDockSite3: TdxTabContainerDockSite;
    dxDockingManager: TdxDockingManager;
    dxDockHost: TdxDockSite;
    dxLayoutDockSite1: TdxLayoutDockSite;
    dxLayoutDockSite2: TdxLayoutDockSite;
    dxLayoutDockSite3: TdxLayoutDockSite;
    dxDockPanel1: TdxDockPanel;
    dxDockPanel4: TdxDockPanel;
    dxDockPanel2: TdxDockPanel;
    dxDockPanel3: TdxDockPanel;
    dxDockPanel5: TdxDockPanel;
    dxDockPanel6: TdxDockPanel;
    dxDockPanel7: TdxDockPanel;
    AutoExecTimer: TTimer;
    PlayCmdGroup: TdxBarGroup;
    StopCmdGroup: TdxBarGroup;
    LoadConfigTimer: TTimer;
    StartupMsgList: TcxListView;
    Panel1: TPanel;
    Panel2: TPanel;
    SoDatabaseMsgList: TcxListView;
    ControlSocket: TIdTCPClient;
    FeedbackSocket: TIdTCPClient;
    ControlMsgList: TcxListView;
    lblConfigName: TLabel;
    SoTree: TcxTreeList;
    SoTreeCol1: TcxTreeListColumn;
    SoTreeCol2: TcxTreeListColumn;
    SoTreeCol3: TcxTreeListColumn;
    dxAutoScroll: TdxBarButton;
    ConsoleTree: TcxTreeList;
    dxConsole: TdxBarButton;
    ControlConsoleTreeColumn1: TcxTreeListColumn;
    ControlConsoleTreeColumn2: TcxTreeListColumn;
    ControlConsoleTreeColumn3: TcxTreeListColumn;
    ControlConsoleTreeColumn4: TcxTreeListColumn;
    ControlConsoleTreeColumn5: TcxTreeListColumn;
    dxDockPanel8: TdxDockPanel;
    FeedbackMsgList: TcxListView;
    dxDbTree: TdxBarButton;
    dxControl: TdxBarButton;
    FeedbackTree: TcxTreeList;
    FeedbackTreecColumn1: TcxTreeListColumn;
    FeedbackTreecColumn2: TcxTreeListColumn;
    FeedbackTreecColumn3: TcxTreeListColumn;
    FeedbackTreecColumn5: TcxTreeListColumn;
    FeedbackTreecColumn6: TcxTreeListColumn;
    FeedbackTreecColumn4: TcxTreeListColumn;
    dxFeedback: TdxBarButton;
    ControlSendTree: TcxTreeList;
    ControlSendTreeColumn1: TcxTreeListColumn;
    ControlSendTreeColumn2: TcxTreeListColumn;
    ControlSendTreeColumn3: TcxTreeListColumn;
    ControlSendTreeColumn4: TcxTreeListColumn;
    ControlSendTreeColumn5: TcxTreeListColumn;
    ControlSendTreeColumn6: TcxTreeListColumn;
    ControlSendTreeColumn7: TcxTreeListColumn;
    ControlSendTreeColumn8: TcxTreeListColumn;
    ControlSendTreeColumn9: TcxTreeListColumn;
    ControlSendTreeColumn10: TcxTreeListColumn;
    dxCmdExpand: TdxBarButton;
    dxCmdCollapse: TdxBarButton;
    ControlSendGroup: TdxBarGroup;
    dxMenuItem11: TdxBarButton;
    dxEnableControlSend: TdxBarStatic;
    dxEnableFeedbackRecv: TdxBarStatic;
    cmbConfigList: TcxImageComboBox;
    dxMenuItem22: TdxBarButton;
    dxMenuItem2: TdxBarSubItem;
    dxMenuItem23: TdxBarButton;
    dxMenuItem24: TdxBarButton;
    dxMenuItem25: TdxBarButton;
    dxMenuItem26: TdxBarButton;
    dxMenuItem27: TdxBarButton;
    dxMenuItem28: TdxBarButton;
    dxMenuItem29: TdxBarButton;
    dxMenuItem21: TdxBarButton;
    dxMenuItem3: TdxBarButton;
    ControlSendLEDTimer: TTimer;
    FeedbackLEDTimer: TTimer;
    FeedbackTreecColumn7: TcxTreeListColumn;
    SoTreeCol4: TcxTreeListColumn;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormDestroy(Sender: TObject);
    procedure LoadConfigTimerTimer(Sender: TObject);
    procedure AutoExecTimerTimer(Sender: TObject);
    procedure dxPlayClick(Sender: TObject);
    procedure dxStopClick(Sender: TObject);
    procedure dxAutoScrollClick(Sender: TObject);
    procedure OnStartupMessage(aSubject: TSubject);
    procedure OnSoDatabaseMessage(aSubject: TSubject);
    procedure OnInfoMsgNotify(aSubject: TSubject);
    procedure OnControlMessage(aSubject: TSubject);
    procedure OnControlCmdUpdate(aSubject: TSubject);
    procedure OnConsoleUpdate(aSubject: TSubject);
    procedure OnFeedbackCmdUpdate(aSubject: TSubject);
    procedure OnFeedbackMessage(aSubject: TSubject);
    procedure OnControlLed(aSubject: TSubject);
    procedure OnFeedbackLed(aSubject: TSubject);
    procedure ControlSendTreeCustomDrawCell(Sender: TObject;
      ACanvas: TcxCanvas; AViewInfo: TcxTreeListEditCellViewInfo;
      var ADone: Boolean);
    procedure dxConsoleClick(Sender: TObject);
    procedure dxDbTreeClick(Sender: TObject);
    procedure dxControlClick(Sender: TObject);
    procedure dxDockPanelResize(Sender: TObject);
    procedure dxCmdExpandClick(Sender: TObject);
    procedure dxCmdCollapseClick(Sender: TObject);
    procedure dxMenuItem11Click(Sender: TObject);
    procedure cmbConfigListPropertiesChange(Sender: TObject);
    procedure dxMenuItem21Click(Sender: TObject);
    procedure dxMenuItem22Click(Sender: TObject);
    procedure dxDockPanel2Activate(Sender: TdxCustomDockControl;
      Active: Boolean);
    procedure dxMenuItem3Click(Sender: TObject);
    procedure ControlSendLEDTimerTimer(Sender: TObject);
    procedure FeedbackLEDTimerTimer(Sender: TObject);
    procedure ConsoleTreeKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private
    { Private declarations }
    FLicenceManager: TLicenceManager;
    FSoThreadManager: TThreadManager;
    FControlManager: TThreadManager;
    FFeedbackDbManager: TThreadManager;
    FFeedbackManager: TThreadManager;
    FSoDatabaseObSrv: TObServer;
    FControlMsgObSrv: TObServer;
    FMsgSubject: TMessageExSubject;
    FStartupObSrv: TObServer;
    //FInfoObServer: TObServer;
    FControlCmdObSrv: TObServer;
    FConsoleObSrv: TObServer;
    FFeedbackCmdObSrv: TObServer;
    FFeedbackMsgObSrv: TObServer;
    FControlLedOberver: TObServer;
    FFeedbackLedObServer: TObServer;
    FSoName: array of String;
    FAutoScroll: Boolean;
    FDbWarning: TFlickerControl;
    FControlSendWarning: TFlickerControl;
    FFeedbackRecvWarning: TFlickerControl;
    FDocklayoutStream: TMemoryStream;
    FGuiCleanupThread: TGuiCleanup;
    FIniFile: TIniFile;
    function CreateControlSocketThread: Boolean;
    function CreateFeedbackSocketThread: Boolean;
    function CreateAppIniFile: TIniFile;
    procedure InitiateUI;
    procedure InitiateUILanguage;
    procedure LoadWindowPositionFromFile;
    procedure SaveWindowPositionToFile;
    procedure LoadToolBarStateFromFile;
    procedure SaveToolBarStateToFile;
    procedure ReleaseControlSocketThread;
    procedure ReleaseFeedbackSocketThread;
    procedure ReleaseGuiCleanupThread;
    procedure CreateGuiCleanupThread;
    procedure CreateSoDbThread;
    procedure ReleaseSoDbThread;
    procedure CreateFeedbackDbThread;
    procedure ReleaseFeedbackDbThread;
    procedure AllocateControlResource;
    procedure CleanupTaskResource;
    procedure GuaranteeControlSendExecute;
    procedure GuaranteeFeedbackRecvExecute;
    procedure WMSocket(var Msg: TWMSocket); message WM_SOCKET;
    procedure WMDatabase(var Msg: TMessage); message WM_DATABASE;
  public
    { Public declarations }
    property StartupObServer: TObServer read FStartupObSrv;
    property SoDbObServer: TObserver read FSoDatabaseObSrv;
    procedure BuildSoDbTree;
    procedure ClearSoDbTree;
    procedure BuildFeedbackDbTree;
    procedure BuildConfigFileList;
    procedure ClearConfigFileList;
    procedure AppointConfigFile;
    procedure BeginExecute;
    procedure EndExecute;
    procedure UpdateSoTreeStatus(aSoInfo: Pointer);
    function CheckLicence: Boolean;
  end;


var
  fmMain: TfmMain;

  { 以下全域變數會在所有的 Thread 中共享 }

  { 準備傳送出去的指令 }
  { ControlSend Thread 會負責把它送出去 }
  ControlSendManager: TCommandListMagager;

  { 已經送出去的指令, 由 SoDbThread 讀出來後更新回 Send_Nagar }
  ControlSendDoneManager: TCommandListMagager;

  { 送出去的指令, 當 Nagra Ack 回來後, 會放到這裏,
    SoDbThread 會負責將它讀出來後, 更新回 Send_Nagar  }
  ControlRecvManager: TCommandListMagager;

  { 傳送指令出去的 Log 通通會在這裏, 包含 Send 出去 及 Ack 回來的 Log,
    SoDbThread 會負責將它讀出來後寫進 Log Table }
  ControlSendLogManager: TCommandListMagager;

  { CommunicationCheck List }
  ControlCommunicationCheck: TThreadStringList;
  FeedbackCommunicationCheck: TThreadStringList;

  { 傳送指令 Ack 回來的 List }
  ControlAck: TThreadStringList;

  { Feedback Ack 回來的 List }
  FeedbackAck: TThreadStringList;

  { 接收 Feedback 資料的 List }
  FeedbackRecvList: TThreadStringList;

  { 回應接收資料要 Ack 回去的 List }
  FeedbackSendList: TThreadStringList;

  { 指令處理的上限計數器,
    每傳送一筆低階指令, 計數器 - 1
    當有 ACK or NACK 回來, 計數器 + 1
    當計數器等於0, 則不可以傳送指令 }

  ControlSendMaxCounter: Integer;


  { 以上全域變數由 Main Form 負責建立及釋放 }


implementation

{$R *.dfm}

uses
  cbResStr, cbUtilis, cbSoDbThread, cbControlThread, cbFeedbackDbThread,
  cbFeedbackThread, cbAbout;

var
  aControlKeyValue: Integer = 0;
  aFeedbackKeyValue: Integer = 0;


{ ---------------------------------------------------------------------------- }

function GetMsgImageIndex(const aType: Longint): Integer;
begin
  case aType of
    MB_OK: Result := 1;
    MB_ICONERROR: Result := 2;
    MB_ICONWARNING: Result := 3;
  else
    Result := 0;  
  end;
end;

{ ---------------------------------------------------------------------------- }

function GetSoImageIndex(const aStatus: TDbConnectStatus): Integer;
begin
  case aStatus of
    dbOK, dbActive: Result := 3;
    dbError: Result := 4;
    dbNone: Result := 2;
  else
    Result := 1;
  end;
end;

{ ---------------------------------------------------------------------------- }

function GetSoStateIndex(const aStatus: TDbConnectStatus): Integer;
begin
  case aStatus of
    dbOK, dbError: Result := 6;
    dbActive: Result := 7;
  else
    Result := 5;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.OnStartupMessage(aSubject: TSubject);
var
  aText: String;
  aItem: TListItem;
begin
  if aSubject.RunState in [rsStartup, rsStop] then
  begin
    aText := Format( ' %s >   %s', [FormatDateTime( 'yyyy-mm-dd hh:mm', Now ),
      TMessageExSubject( aSubject ).NotifyMessage] );
    aItem := StartupMsgList.Items.Add;
    aItem.Caption := aText;
    aItem.ImageIndex := GetMsgImageIndex( TMessageSubject( aSubject ).NotifyType );
    aItem.MakeVisible( True );
    Application.ProcessMessages;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.OnSoDatabaseMessage(aSubject: TSubject);
var
  aText: String;
  aItem: TListItem;
begin
  if aSubject.RunState in [rsRunning] then
  begin
    aText := Format( ' %s >   %s', [FormatDateTime( 'yyyy-mm-dd hh:mm', Now ),
      TMessageExSubject( aSubject ).NotifyMessage] );
    aItem := SoDatabaseMsgList.Items.Add;
    aItem.Caption := aText;
    aItem.ImageIndex := GetMsgImageIndex( TMessageSubject( aSubject ).NotifyType );
    aItem.MakeVisible( True );
    Application.ProcessMessages;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.OnInfoMsgNotify(aSubject: TSubject);
begin
  { .... }
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.OnControlMessage(aSubject: TSubject);
var
  aText: String;
  aItem: TListItem;
begin
  if aSubject.RunState in [rsRunning] then
  begin
    aText := Format( ' %s >   %s', [FormatDateTime( 'yyyy-mm-dd hh:mm', Now ),
      TMessageExSubject( aSubject ).NotifyMessage] );
    aItem := ControlMsgList.Items.Add;
    aItem.Caption := aText;
    aItem.ImageIndex := GetMsgImageIndex( TMessageSubject( aSubject ).NotifyType );
    aItem.MakeVisible( True );
    Application.ProcessMessages;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.OnControlCmdUpdate(aSubject: TSubject);
const
  aHighImageIndex = 8;
  aLowImageIndex = 1;

  { ------------------------------------------------------ }

  function GetHighCmdDes(aCmdId: String): String;
  begin
    Result := EmptyStr;
    if ConfigModule.HighCmdEnv.Locate( 'HIGH_LEVEL_CMD', aCmdId, [] ) then
      Result := ConfigModule.HighCmdEnv.FieldByName( 'DESCRIPTION' ).AsString;
  end;

  { ------------------------------------------------------ }

  function GetLowCmdDesc(aCmdId: String): String;
  begin
    Result := EmptyStr;
    if ConfigModule.LowCmdEnv.Locate( 'LOW_LEVEL_CMD', aCmdId, [] ) then
      Result := ConfigModule.LowCmdEnv.FieldByName( 'DESCRIPTION' ).AsString;
  end;

  { ------------------------------------------------------ }

  function DoAddHighCmd(aObj: PSendNagra): TcxTreeListNode;
  begin
    Result := ControlSendTree.Add;
    Result.ImageIndex := aHighImageIndex;
    Result.SelectedIndex := aHighImageIndex;
    Result.AssignValues( [
      FSoName[StrToInt(aObj.CompCode)],            { COMPNAME }
      aObj.SeqNo,                                  { SEQNO }
      EmptyStr,                                    { ICCNO }
      Format( '%s - ( %s )', [aObj.HighCmdId,
        GetHighCmdDes( aObj.HighCmdId )] ),     { CMDID }
      aObj.SmsSendStatus,                             { SENDSTATUS }
      aObj.SmsSendStatus,                             { RECVSTATUS }
      aObj.Operator,                               { OPERATOR }
      EmptyStr,                                    { ERRORCODE }
      EmptyStr,                                    { ERRORMSG }
      FormatDateTime( 'yyyy-mm-dd hh:mm:ss', aObj.UpdTime )] ); { UPDTIME }
    aObj.Data := Result;
    aObj.ParentData := nil;
  end;

  { ------------------------------------------------------ }

  function DoAddLowCmd(aObj: PSendNagra): TcxTreeListNode;
  begin
    Result := ControlSendTree.AddChild( TcxTreeListNode( aObj.ParentData ) );
    Result.ImageIndex := aLowImageIndex;
    Result.SelectedIndex := aLowImageIndex;
    Result.AssignValues( [
      EmptyStr,                                     { COMPNAME }
      aObj.SmsTransactionNumber,                    { SEQNO }
      aObj.IccNo,                                   { ICCNO }
      Format( '    %s - ( %s )',
        [aObj.LowCmdId, GetLowCmdDesc( aObj.LowCmdId )] ),       { CMDID }
      //Format( '    %s - ( %s )', [EmptyStr,'低階'] ),       { CMDID }
      aObj.SmsSendStatus,                           { SENDSTATUS }
      aObj.SmsSendStatus,                           { RECVSTATUS }
      EmptyStr,                                     { OPERATOR }
      EmptyStr,                                     { ERRORCODE }
      EmptyStr,                                     { ERRORMSG }
      FormatDateTime( '    yyyy-mm-dd hh:mm:ss', aObj.UpdTime )] ); { UPDTIME }
    if ( aObj.SmsSendStatus = 'P' ) or
       ( aObj.SmsSendStatus = 'C' ) then
    begin
      if Result.Parent.Values[ControlSendTreeColumn5.ItemIndex] = 'W' then
        Result.Parent.Values[ControlSendTreeColumn5.ItemIndex] := aObj.SmsSendStatus;
    end;
    aObj.Data := Result;
  end;

  { ------------------------------------------------------ }

  procedure MakeSureNodeVisible(aNode: TcxTreeListNode);
  begin
    if Assigned( aNode.Parent ) then
      if not aNode.Parent.Expanded then
        aNode.Parent.Expand( True );
    aNode.MakeVisible;
  end;

  { ------------------------------------------------------ }

var
  aCmdSubject: TCommandSubject;
  aCmdObj: PSendNagra;
  aNode: TcxTreeListNode;
  aJustAdded, aIsLowCmd: Boolean;
begin
  aCmdSubject := TCommandSubject( aSubject );
  if aCmdSubject.UpdateCommandType <> ncCommand then Exit;
  aCmdObj := aCmdSubject.UpdateCommandObject;
  aJustAdded := ( not Assigned( aCmdObj.Data ) );
  aIsLowCmd := ( aCmdObj.SmsTransactionNumber <> EmptyStr );
  ControlSendTree.BeginUpdate;
  try
    if aJustAdded then
    begin
      { 低階指令還是高階 ? }
      if aIsLowCmd then
        aNode := DoAddLowCmd( aCmdObj )
      else
        aNode := DoAddHighCmd( aCmdObj );
    end else
    begin
      aNode := TcxTreeListNode( aCmdObj.Data );
      aNode.Values[ControlSendTreeColumn6.ItemIndex] := aCmdObj.SmsSendStatus;
      if aCmdObj.ErrCode <> EmptyStr then
        aNode.Values[ControlSendTreeColumn8.ItemIndex] := aCmdObj.ErrCode;
      if aCmdObj.ErrMsg <> EmptyStr then
        aNode.Values[ControlSendTreeColumn9.ItemIndex] := aCmdObj.ErrMsg;
    end;
  finally
    ControlSendTree.EndUpdate;
  end;
  if FAutoScroll and aJustAdded and aIsLowCmd then
    MakeSureNodeVisible( aNode );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.OnFeedbackCmdUpdate(aSubject: TSubject);
var
  aCmdSubject: TCommandSubject;
  aCmdObj: PRecvNagra;
  aNode: TcxTreeListNode;

  { ------------------------------------------------------ }

  function GetHighCmdDes(aCmdId: String): String;
  begin
    Result := EmptyStr;
    if ConfigModule.HighCmdEnv.Locate( 'HIGH_LEVEL_CMD', aCmdId, [] ) then
      Result := ConfigModule.HighCmdEnv.FieldByName( 'DESCRIPTION' ).AsString;
  end;

  { ------------------------------------------------------ }

begin
  aCmdSubject := TCommandSubject( aSubject );
  if aCmdSubject.UpdateCommandType <> ncCommand then Exit;
  aCmdObj := aCmdSubject.UpdateCommandObject;
  FeedbackTree.BeginUpdate;
  try
    if ( not Assigned( aCmdObj.Data ) ) then
    begin
      aNode := FeedbackTree.Add;
      aNode.ImageIndex := 2;
      aNode.SelectedIndex := aNode.ImageIndex;
      aNode.Values[FeedbackTreecColumn1.ItemIndex] := aCmdObj.ResponseTransactionNumber;
      aNode.Values[FeedbackTreecColumn2.ItemIndex] := aCmdObj.IccNo;
      aNode.Values[FeedbackTreecColumn3.ItemIndex] := Format( '%s - %s - ( %s )', [
        aCmdObj.HighLevelCmd, aCmdObj.LowLevelCmd, GetHighCmdDes( aCmdObj.HighLevelCmd)] );
      aNode.Values[FeedbackTreecColumn4.ItemIndex] := aCmdObj.CmdStatus;
      aNode.Values[FeedbackTreecColumn5.ItemIndex] := aCmdObj.SendStatus;
      aNode.Values[FeedbackTreecColumn6.ItemIndex] := aCmdObj.UpdTime;
      aNode.Values[FeedbackTreecColumn7.ItemIndex] := aCmdObj.StbNo;
      aCmdObj.Data := aNode;
      aCmdObj.ParentData := nil;
      if FAutoScroll then aNode.MakeVisible;
    end else
    begin
      aNode := TcxTreeListNode( aCmdObj.Data );
      aNode.Values[FeedbackTreecColumn4.ItemIndex] := aCmdObj.CmdStatus;
      aNode.Values[FeedbackTreecColumn5.ItemIndex] := aCmdObj.SendStatus;
    end;
  finally
    FeedbackTree.EndUpdate;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.OnFeedbackMessage(aSubject: TSubject);
var
  aText: String;
  aItem: TListItem;
begin
  if aSubject.RunState in [rsRunning] then
  begin
    aText := Format( ' %s >   %s', [FormatDateTime( 'yyyy-mm-dd hh:mm', Now ),
      TMessageExSubject( aSubject ).NotifyMessage] );
    aItem := FeedbackMsgList.Items.Add;
    aItem.Caption := aText;
    aItem.ImageIndex := GetMsgImageIndex( TMessageSubject( aSubject ).NotifyType );
    aItem.MakeVisible( True );
    Application.ProcessMessages;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.OnConsoleUpdate(aSubject: TSubject);
var
  aCmdSubject: TCommandSubject;
  aCmdObj: PSendNagra;
  aCmdObj2: PRecvNagra;
  aNode: TcxTreeListNode;
  aText: String;
  aNodeIndex: Integer;

  { ------------------------------------------------ }

  function AddNode: TcxTreeListNode;
  var
    aTransNum, aCmdText, aResponeNum: String;
  begin
    if ( aCmdSubject.ThreadType in [tdControlSend,tdControlRecv] ) then
    begin
      aTransNum := aCmdObj.SmsTransactionNumber;
      aCmdText := aCmdObj.SmsCommandText;
      aResponeNum := EmptyStr;
    end else
    if ( aCmdSubject.ThreadType in [tdFeedbackSend] ) then
    begin
      aTransNum := aCmdObj2.TransactionNumber;
      aCmdText := aCmdObj2.CommandText;
      aResponeNum := aCmdObj2.ResponseTransactionNumber;
    end else
    if ( aCmdSubject.ThreadType in [tdFeedbackRecv] ) then
    begin
      aTransNum := aCmdObj2.TransactionNumber;
      aCmdText := aCmdObj2.ResponseCommandText;
      aResponeNum := aCmdObj2.ResponseTransactionNumber;
    end;
    Result := ConsoleTree.Add;
    if ( aCmdSubject.UpdateCommandType = ncCommunication ) then
      Result.ImageIndex := 9
    else begin
      Result.ImageIndex := 1;
      if aCmdSubject.ThreadType in [tdFeedbackSend, tdFeedbackRecv] then
        Result.ImageIndex := 2;
    end;  
    Result.SelectedIndex := Result.ImageIndex;
    //aCmdText := Lpad( EmptyStr, Length( aCmdText ), '*' );
    Result.AssignValues( [aTransNum, aResponeNum, aCmdText, SUnknow] );
  end;

  { ------------------------------------------------ }

begin
  aCmdSubject := TCommandSubject( aSubject );
  { 傳送指令的 Command }
  if aCmdSubject.ThreadType in [tdControlSend,tdControlRecv] then
  begin
    aCmdObj := aCmdSubject.UpdateCommandObject;
    if aCmdObj.SmsTransactionNumber = EmptyStr then Exit;
    aNodeIndex := aCmdObj.Index;
    if ( aNodeIndex < 0 ) then
    begin
      aNode := AddNode;
      aCmdObj.Index := Integer( aNode );
      aCmdObj.ParentIndex := -1;
      if FAutoScroll then aNode.MakeVisible;
    end else
    begin
      aNode := TcxTreeListNode( aNodeIndex );
      if Assigned( aNode ) then
      begin
        aText := SUnknow;
        if aCmdObj.SmsSendStatus = 'C' then
          aText := SAck
        else if aCmdObj.SmsSendStatus = 'E' then
          aText := SNAck;
        aNode.Texts[ControlConsoleTreeColumn2.ItemIndex] :=
          aCmdObj.CaAckTransactionNumber;
        aNode.Texts[ControlConsoleTreeColumn4.ItemIndex] := aText;
      end;
    end;
  end else
  { 回傳機制的 Command }
  begin
    aCmdObj2 := aCmdSubject.UpdateCommandObject;
    aNodeIndex := aCmdObj2.Index;
    if ( aNodeIndex < 0 ) then
    begin
      aNode := AddNode;
      aCmdObj2.Index := Integer( aNode );
      aCmdObj2.ParentIndex := -1;
      if FAutoScroll then aNode.MakeVisible;
    end else
    begin
      aNode := TcxTreeListNode( aNodeIndex );
      if Assigned( aNode ) then
      begin
        aText := SUnknow;
        if aCmdObj2.SendStatus = 'C' then
          aText := SAck
        else if aCmdObj2.SendStatus = 'E' then
          aText := SNAck;
        aNode.Texts[ControlConsoleTreeColumn1.ItemIndex] :=
          aCmdObj2.TransactionNumber;
        aNode.Texts[ControlConsoleTreeColumn2.ItemIndex] :=
          aCmdObj2.ResponseTransactionNumber;
        aNode.Texts[ControlConsoleTreeColumn4.ItemIndex] := aText;
      end;
    end;
  end;
  { 把每一筆顯示的 Node, 加上時間 }
  if Assigned( aNode ) then
  begin
    aNode.Texts[ControlConsoleTreeColumn5.ItemIndex] :=
      FormatDateTime( 'yyyy/mm/dd hh:nn:ss', Now );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.ControlSendLEDTimerTimer(Sender: TObject);
begin
  ControlSendLEDTimer.Enabled := False;
  dxEnableControlSend.ImageIndex := 17;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.OnControlLed(aSubject: TSubject);
var
  aCmdSubject: TCommandSubject;
begin
  aCmdSubject := TCommandSubject( aSubject );
  { 傳送指令的 Command }
  if aCmdSubject.ThreadType in [tdControlSend] then
  begin
    { 利用 Timer 來做閃動的效果 }
    dxEnableControlSend.ImageIndex := 19;
    ControlSendLEDTimer.Enabled := True;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.FeedbackLEDTimerTimer(Sender: TObject);
begin
  FeedbackLEDTimer.Enabled := False;
  dxEnableFeedbackRecv.ImageIndex := 17;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.OnFeedbackLed(aSubject: TSubject);
var
  aCmdSubject: TCommandSubject;
begin
  aCmdSubject := TCommandSubject( aSubject );
  { 傳送指令的 Command }
  if aCmdSubject.ThreadType in [tdFeedbackSend] then
  begin
    dxEnableFeedbackRecv.ImageIndex := 19;
    { 利用 Timer 來做閃動的效果 }
    FeedbackLEDTimer.Enabled := True;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.FormCreate(Sender: TObject);
begin
  Application.Title := SAppTitle;
  StyleModule := TStyleModule.Create( nil );
  {}
  FSoThreadManager := TThreadManager.Create;
  FControlManager := TThreadManager.Create;
  FFeedbackDbManager := TThreadManager.Create;
  FFeedbackManager := TThreadManager.Create;
  {}
  ControlSendManager := TCommandListMagager.Create;
  ControlSendDoneManager := TCommandListMagager.Create;
  ControlRecvManager := TCommandListMagager.Create;
  ControlSendLogManager := TCommandListMagager.Create;
  { .. }
  ControlCommunicationCheck := TThreadStringList.Create;
  FeedbackCommunicationCheck := TThreadStringList.Create;
  { ... }
  ControlCommunicationCheck.Sorted := True;
  FeedbackCommunicationCheck.Sorted := True;
  { ... }
  ControlAck := TThreadStringList.Create;
  FeedbackAck := TThreadStringList.Create;
  FeedbackRecvList := TThreadStringList.Create;
  FeedbackSendList := TThreadStringList.Create;
  { 回饋訊息 }
  FStartupObSrv := TObServer.Create;
  FStartupObSrv.OnUpdate := OnStartupMessage;
  FSoDatabaseObSrv := TObServer.Create;
  FSoDatabaseObSrv.OnUpdate := OnSoDatabaseMessage;
  {}
  //FInfoObServer := TObServer.Create;
  //FInfoObServer.OnUpdate := OnInfoMsgNotify;
  {}
  { 傳送及接收回傳資料, 主畫面指令更新 }
  FControlCmdObSrv := TObServer.Create;
  FControlCmdObSrv.OnUpdate := OnControlCmdUpdate;
  FFeedbackCmdObSrv := TObServer.Create;
  FFeedbackCmdObSrv.OnUpdate := OnFeedbackCmdUpdate;
  { 命列模式 }
  FConsoleObSrv := TObServer.Create;
  FConsoleObSrv.OnUpdate := OnConsoleUpdate;
  { 訊息 }
  FControlMsgObSrv := TObServer.Create;
  FControlMsgObSrv.OnUpdate := OnControlMessage;
  FFeedbackMsgObSrv := TObServer.Create;
  FFeedbackMsgObSrv.OnUpdate := OnFeedbackMessage;
  { Splash Led }
  FControlLedOberver := TObServer.Create;
  FControlLedOberver.OnUpdate := OnControlLed;
  FFeedbackLedObServer := TObServer.Create;
  FFeedbackLedObServer.OnUpdate := OnFeedbackLed;
  {}
  { 主畫面用訊息顯示 }
  FMsgSubject := TMessageExSubject.Create;
  FMsgSubject.AddObServer( FStartupObSrv );
  FMsgSubject.AutoNotify := True;
  {}
  FDbWarning := TFlickerControl.Create( Self );
  FControlSendWarning := TFlickerControl.Create( Self );
  FFeedbackRecvWarning := TFlickerControl.Create( Self );
  {}
  FAutoScroll := False;
  { Indy Default ReadTimeout is Infinite }
  ControlSocket.ReadTimeout := 3000;
  FeedbackSocket.ReadTimeout := 3000;
  {}
  FIniFile := CreateAppIniFile;
  InitiateUILanguage;
  InitiateUI;
  {}
  ConfigModule := TConfigModule.Create;
  FGuiCleanupThread := nil;
  {}
  FLicenceManager := TLicenceManager.Create( nil );
  Application.ShowMainForm := False;
  if not CheckLicence then
  begin
    ErrorMsg( Format( SLicenceInvalid, [SVendor] ) );
    Application.Terminate;
    Exit;
  end;
  Application.ShowMainForm := True;
  { D2007 Bug }
  Panel1.Align := alBottom;
  Panel1.Align := alTop;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.FormShow(Sender: TObject);
 var UTC:_SYSTEMTIME;
begin
//  LocalDateTimeToUTCDateTime()
  LoadConfigTimer.Enabled := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  CanClose := ConfirmMsg( SDlgConfimQuit );
  if CanClose then
  begin
    ReleaseGuiCleanupThread;
    EndExecute;
    SaveToolBarStateToFile;
    SaveWindowPositionToFile;
    dxDockingManager.SaveLayoutToIniFile( DEFAULT_DOCKLAYOUTFILE );
    FLicenceManager.Update;
  end;  
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.FormDestroy(Sender: TObject);
begin
  FSoThreadManager.Free;
  FControlManager.Free;
  FFeedbackDbManager.Free;
  FFeedbackManager.Free;
  {}
  FDocklayoutStream.Free;
  {}
  ControlSendManager.Free;
  ControlSendDoneManager.Free;
  ControlRecvManager.Free;
  ControlSendLogManager.Free;
  ControlCommunicationCheck.Free;
  FeedbackCommunicationCheck.Free;
  ControlAck.Free;
  FeedbackAck.Free;
  FeedbackRecvList.Free;
  FeedbackSendList.Free;
  {}
  //FInfoObServer.Free;
  FSoDatabaseObSrv.Free;
  FStartupObSrv.Free;
  FControlMsgObSrv.Free;
  FControlCmdObSrv.Free;
  FConsoleObSrv.Free;
  FFeedbackMsgObSrv.Free;
  FFeedbackCmdObSrv.Free;
  FControlLedOberver.Free;
  FFeedbackLedObServer.Free;
  {}
  FDbWarning.Free;
  FControlSendWarning.Free;
  FFeedbackRecvWarning.Free;
  {}
  FMsgSubject.Free;
  StyleModule.Free;
  ConfigModule.Free;
  {}
  if Assigned( FIniFile ) then FIniFile.Free;
  {}
  FLicenceManager.Free;
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.CreateAppIniFile: TIniFile;
var
  aFileName: String;
begin
  if not Assigned( FIniFile ) then
  begin
    aFileName := IncludeTrailingPathDelimiter( ExtractFilePath(
      ParamStr( 0 ) ) ) + DEFAULT_APPFILE;
    FIniFile := TIniFile.Create( aFileName );  
  end;
  Result := FIniFile;
end;

{ ---------------------------------------------------------------------------- }


procedure TfmMain.LoadWindowPositionFromFile;
begin
  Self.WindowState := TWindowState( FIniFile.ReadInteger(
    'WindowPositon', 'State', Ord( wsNormal ) ) );
  if ( Self.WindowState = wsNormal ) then
  begin
    Self.Width := FIniFile.ReadInteger( 'WindowPositon', 'Width', Self.Width );
    Self.Height := FIniFile.ReadInteger( 'WindowPositon', 'Height', Self.Height );
    Self.Left := FIniFile.ReadInteger( 'WindowPositon', 'Left', Self.Left );
    Self.Top := FIniFile.ReadInteger( 'WindowPositon', 'Top', Self.Top );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.SaveWindowPositionToFile;
begin
  FIniFile.WriteInteger( 'WindowPositon', 'State', Ord( Self.WindowState ) );
  if ( Self.WindowState = wsNormal ) then
  begin
    FIniFile.WriteInteger( 'WindowPositon', 'Width', Self.Width );
    FIniFile.WriteInteger( 'WindowPositon', 'Height', Self.Height );
    FIniFile.WriteInteger( 'WindowPositon', 'Left', Self.Left );
    FIniFile.WriteInteger( 'WindowPositon', 'Top', Self.Top );
    FIniFile.UpdateFile;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.LoadToolBarStateFromFile;
begin
  dxDbTree.Down := FIniFile.ReadBool( 'ToolBarState', 'DbTree', False );
  dxControl.Down := FIniFile.ReadBool( 'ToolBarState', 'Control', False );
  dxConsole.Down := FIniFile.ReadBool( 'ToolBarState', 'Console', False );
  dxFeedback.Down := FIniFile.ReadBool( 'ToolBarState', 'Feedback', False );
  dxAutoScroll.Down := FIniFile.ReadBool( 'ToolBarState', 'AutoScroll', False );
  FAutoScroll := dxAutoScroll.Down;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.SaveToolBarStateToFile;
begin
  FIniFile.WriteBool( 'ToolBarState', 'DbTree', dxDbTree.Down );
  FIniFile.WriteBool( 'ToolBarState', 'Control', dxControl.Down );
  FIniFile.WriteBool( 'ToolBarState', 'Console', dxConsole.Down );
  FIniFile.WriteBool( 'ToolBarState', 'Feedback', dxFeedback.Down );
  FIniFile.WriteBool( 'ToolBarState', 'AutoScroll', dxAutoScroll.Down );
  FIniFile.UpdateFile;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.InitiateUI;
begin
  ClearSoDbTree;
  StartupMsgList.Clear;
  PlayCmdGroup.Enabled := False;
  StopCmdGroup.Enabled := False;
  ControlSendTree.Styles.StyleSheet := StyleModule.TreeListStyleSheetStormVGA;
  ConsoleTree.Styles.StyleSheet := StyleModule.TreeListStyleSheetConsoleBlack;
  ConsoleTree.OptionsView.Headers := False;
  FDbWarning.Parent := dxDockPanel1;
  FControlSendWarning.Parent := dxDockPanel2;
  FFeedbackRecvWarning.Parent := dxDockPanel4;
  cmbConfigList.Left := ( lblConfigName.Left + lblConfigName.Width + 5 );
  { 先將 Design time layout 好的 UI 位置儲存於 Stream 中 }
  FDocklayoutStream := TMemoryStream.Create;
  dxDockingManager.SaveLayoutToStream( FDocklayoutStream );
  { 然後讀取 user 自行拖放畫面儲存起來的設定檔 }
  dxDockingManager.LoadLayoutFromIniFile( DEFAULT_DOCKLAYOUTFILE );
  { 回復上次儲存的視窗位置及大小 }
  LoadWindowPositionFromFile;
  dxDockPanel2Activate( dxDockPanel2, False );
  { ToolBar Button State}
  LoadToolBarStateFromFile;
  {}
  dxDockPanel2.Visible := False;
  dxDockPanel4.Visible := False;
  dxDockPanel7.Visible := False;
  dxDockPanel8.Visible := False;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.InitiateUILanguage;
begin
  Caption := Format( '%s-%s', [SVendor, SAppTitle] );
  lblConfigName.Caption := SlblConfigFileName;
  dxNagraServer.Caption := Format( SlblConfigCASAddr,
    [EmptyStr, SControlSend, EmptyStr, SControlRecv, EmptyStr] );
  dxEnableControlSend.Caption := SlblEnableControlSend;
  dxEnableFeedbackRecv.Caption := SlblEnableFeedbackRecv;
  {}
  dxDockPanel1.Caption := SDockPanel1;
  dxDockPanel2.Caption := SDockPanel2;
  dxDockPanel3.Caption := SDockPanel3;
  dxDockPanel4.Caption := SDockPanel4;
  dxDockPanel5.Caption := SDockPanel5;
  dxDockPanel6.Caption := SDockPanel6;
  dxDockPanel7.Caption := SDockPanel7;
  dxDockPanel8.Caption := SDockPanel8;
  { 傳送指令 }
  ControlSendTreeColumn1.Caption.Text := SControlSendTreeColumn1;
  ControlSendTreeColumn2.Caption.Text := SControlSendTreeColumn2;
  ControlSendTreeColumn3.Caption.Text := SControlSendTreeColumn3;
  ControlSendTreeColumn4.Caption.Text := SControlSendTreeColumn4;
  ControlSendTreeColumn5.Caption.Text := SControlSendTreeColumn5;
  ControlSendTreeColumn6.Caption.Text := SControlSendTreeColumn6;
  ControlSendTreeColumn7.Caption.Text := SControlSendTreeColumn7;
  ControlSendTreeColumn8.Caption.Text := SControlSendTreeColumn8;
  ControlSendTreeColumn9.Caption.Text := SControlSendTreeColumn9;
  ControlSendTreeColumn10.Caption.Text := SControlSendTreeColumn10;
  { 命令模式}
  ControlConsoleTreeColumn1.Caption.Text := SControlConsoleTreeColumn1;
  ControlConsoleTreeColumn2.Caption.Text := SControlConsoleTreeColumn2;
  ControlConsoleTreeColumn3.Caption.Text := SControlConsoleTreeColumn3;
  ControlConsoleTreeColumn4.Caption.Text := SControlConsoleTreeColumn4;
  { 回傳資料 }
  FeedbackTreecColumn1.Caption.Text := SFeedbackTreecColumn1;
  FeedbackTreecColumn2.Caption.Text := SFeedbackTreecColumn2;
  FeedbackTreecColumn3.Caption.Text := SFeedbackTreecColumn3;
  FeedbackTreecColumn4.Caption.Text := SFeedbackTreecColumn4;
  FeedbackTreecColumn5.Caption.Text := SFeedbackTreecColumn5;
  FeedbackTreecColumn6.Caption.Text := SFeedbackTreecColumn6;
  FeedbackTreecColumn7.Caption.Text := SFeedbackTreecColumn7;
  { 系統}
  dxMenuItem1.Caption := SMenuItem1;
  dxMenuItem11.Caption := SMenuItem11;            
  { 檢視 }
  dxMenuItem2.Caption := SMenuItem2;
  dxMenuItem21.Caption := SMenuItem21;
  dxMenuItem22.Caption := SMenuItem22;
  dxMenuItem23.Caption := SMenuItem23;
  dxMenuItem24.Caption := SMenuItem24;
  dxMenuItem25.Caption := SMenuItem25;
  dxMenuItem26.Caption := SMenuItem26;
  dxMenuItem27.Caption := SMenuItem27;
  dxMenuItem28.Caption := SMenuItem28;
  dxMenuItem29.Caption := SMenuItem29;
  { 關於 }
  dxMenuItem3.Caption := SMenuItem3;
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.CreateControlSocketThread: Boolean;
var
  aThread: TSMSSocketThread;
  aBool1, aBool2: Boolean;
begin
  FMsgSubject.Info( Format( SThreadStart, [SControlSend] ) );
  Delay( COMMON_DELAY );
  aBool1 := False;
  aBool2 := False;
  try
  { 傳送指令回應的 Thread }
    aThread := TControlSendThread.Create( ControlSocket );
    aThread.MainFormHandle := Self.Handle;
    aThread.CommEnv := ConfigModule.CommEnv;
    aThread.HighCmdEnv := ConfigModule.HighCmdEnv;
    aThread.CmdTransEnv := ConfigModule.CmdTransEnv;
    aThread.CASEnv := ConfigModule.CASEnv;
    aThread.CmdError := ConfigModule.CmdError;
    aThread.MessageSubject.AddObServer( FStartupObSrv );
    aThread.MessageSubject.AddObServer( FControlMsgObSrv );
    aThread.CommandSubject.AddObServer( FControlLedOberver );
    if dxControl.Down then
      aThread.CommandSubject.AddObServer( FControlCmdObSrv );
    if dxConsole.Down then
      aThread.CommandSubject.AddObServer( FConsoleObSrv );
    FControlManager.Add( aThread );
    aBool1 := True;
    FMsgSubject.Info( Format( SThreadReady, [SControlSend] ) );
    Delay( COMMON_DELAY );
  except
    on E: Exception do
    begin
      FMsgSubject.Error( Format( SThreadStartError, [SControlSend, E.Message] ) );
      Delay( COMMON_DELAY );
    end;
  end;
  { 傳送指令執行緒建立成功, 才可建立接收回應的執行緒 }
  if aBool1 then
  begin
    { 接收傳送指令回應的 Thread }
    FMsgSubject.Info( Format( SThreadStart, [SControlRecv] ) );
    Delay( COMMON_DELAY );
    try
      aThread := TControlRecvThread.Create( ControlSocket );
      aThread.MainFormHandle := Self.Handle;
      aThread.CommEnv := ConfigModule.CommEnv;
      aThread.HighCmdEnv := ConfigModule.HighCmdEnv;
      aThread.CmdTransEnv := ConfigModule.CmdTransEnv;
      aThread.CASEnv := ConfigModule.CASEnv;
      aThread.CmdError := ConfigModule.CmdError;
      aThread.MessageSubject.AddObServer( FStartupObSrv );
      aThread.MessageSubject.AddObServer( FControlMsgObSrv );
      if dxControl.Down then
        aThread.CommandSubject.AddObServer( FControlCmdObSrv );
      if dxConsole.Down then
        aThread.CommandSubject.AddObServer( FConsoleObSrv );
      FControlManager.Add( aThread );
      aBool2 := True;
      FMsgSubject.Info( Format( SThreadReady, [SControlRecv] ) );
      Delay( COMMON_DELAY );
    except
      on E: Exception do
      begin
        FMsgSubject.Error( Format( SThreadStartError, [E.Message] ) );
        Delay( COMMON_DELAY );
      end;
    end;
  end;
  Result := ( aBool1 and aBool2 );
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.CreateFeedbackSocketThread: Boolean;
var
  aThread: TSMSSocketThread;
  aBool1, aBool2: Boolean;
  aIndex: Integer;
begin
  FMsgSubject.Info( Format( SThreadStart, [SFeedbackSend] ) );
  Delay( COMMON_DELAY );
  aBool1 := False;
  aBool2 := False;
  try
    { 傳送指令回應的 Thread }
    aThread := TFeedbackSendThread.Create( FeedbackSocket );
    aThread.MainFormHandle := Self.Handle;
    aThread.CommEnv := ConfigModule.CommEnv;
    aThread.HighCmdEnv := ConfigModule.HighCmdEnv;
    aThread.CmdTransEnv := ConfigModule.CmdTransEnv;
    aThread.CASEnv := ConfigModule.CASEnv;
    aThread.CmdError := ConfigModule.CmdError;
    aThread.MessageSubject.AddObServer( FStartupObSrv );
    aThread.MessageSubject.AddObServer( FFeedbackMsgObSrv );
    aThread.CommandSubject.AddObServer( FFeedbackLedObServer );
    if dxFeedback.Down then
      aThread.CommandSubject.AddObServer( FFeedbackCmdObSrv );
    if dxConsole.Down then
      aThread.CommandSubject.AddObServer( FConsoleObSrv );
    FFeedbackManager.Add( aThread );
    aBool1 := True;
    FMsgSubject.Info( Format( SThreadReady, [SFeedbackSend] ) );
    Delay( COMMON_DELAY );
  except
    on E: Exception do
    begin
      FMsgSubject.Error( Format( SThreadStartError, [SFeedbackSend, E.Message] ) );
      Delay( COMMON_DELAY );
    end;
  end;
  { 傳送指令執行緒建立成功, 才可建立接收回應的執行緒 }
  if aBool1 then
  begin
    FMsgSubject.Info( Format( SThreadStart, [SFeedbackRecv] ) );
    Delay( COMMON_DELAY );
    { 傳送指令回應的 Thread }
    try
      aThread := TFeedbackRecvThread.Create( FeedbackSocket );
      aThread.MainFormHandle := Self.Handle;
      aThread.CommEnv := ConfigModule.CommEnv;
      aThread.HighCmdEnv := ConfigModule.HighCmdEnv;
      aThread.CmdTransEnv := ConfigModule.CmdTransEnv;
      aThread.CASEnv := ConfigModule.CASEnv;
      aThread.CmdError := ConfigModule.CmdError;
      aThread.MessageSubject.AddObServer( FStartupObSrv );
      aThread.MessageSubject.AddObServer( FFeedbackMsgObSrv );
      if dxFeedback.Down then
        aThread.CommandSubject.AddObServer( FFeedbackCmdObSrv );
      if dxConsole.Down then
        aThread.CommandSubject.AddObServer( FConsoleObSrv );
      FFeedbackManager.Add( aThread );
      aBool2 := True;
      FMsgSubject.Info( Format( SThreadReady, [SFeedbackRecv] ) );
      Delay( COMMON_DELAY );
    except
      on E: Exception do
      begin
        FMsgSubject.Error( Format( SThreadStartError , [SFeedbackRecv, E.Message] ) );
        Delay( COMMON_DELAY );
      end;
    end;
  end;
  Result := ( aBool1 and aBool2 );
  { if both thread created is ok, than statr the thread }
  if Result then
  begin
    for aIndex := 0 to FFeedbackManager.Count - 1 do
      TSMSSocketThread( FFeedbackManager.Items[aIndex] ).Resume;
    Delay( COMMON_DELAY );
    for aIndex := 0 to FFeedbackManager.Count - 1 do
      PostMessage( TSMSSocketThread( FFeedbackManager.Items[aIndex] ).WndHandle,
      WM_PLAY, 0, 0 );
    Delay( COMMON_DELAY );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.ReleaseControlSocketThread;
var
  aIndex: Integer;
begin
  if FControlManager.Count <= 0 then Exit;
  { 先打出停止的訊息 }
  for aIndex := 0 to FControlManager.Count - 1 do
    { 2010.01.20 By Lawrence Add Try...except...end}
    begin
      try
        PostMessage( TSMSSocketThread(
          FControlManager.Items[aIndex] ).WndHandle, WM_STOP, 0, 0 );
      except
        on E: Exception do
        begin
          FMsgSubject.Error( Format( SSocketThreadError, [SReleaseControl, E.Message] ) );
          Delay( COMMON_DELAY );
        end;
      end;
    end;
  Delay( 1000 );
  FMsgSubject.Info( Format( SThreadStop, [SControlSend] ) );
  Delay( COMMON_DELAY );
  { 停止傳送指令執行緒 }
  try
    FControlManager.StopThread( 0 );
    FMsgSubject.Info( Format( SThreadEnd, [SControlSend] ) );
    Delay( COMMON_DELAY );
  except
    on E: Exception do
    begin
      FMsgSubject.Error( Format( SThreadStopError, [SControlSend, E.Message] ) );
      Delay( COMMON_DELAY );
    end;
  end;
  { 有可能在建立接收回應的執行緒時失敗, 所以 ThreadManager 只有傳送的 }
  FMsgSubject.Info( Format( SThreadStop, [SControlRecv] ) );
  Delay( COMMON_DELAY );
  if FControlManager.Count = 0 then Exit;
  try
    FControlManager.StopThread( 1 );
    FMsgSubject.Info( Format( SThreadEnd, [SControlRecv] ) );
    Delay( COMMON_DELAY );
  except
    on E: Exception do
    begin
      FMsgSubject.Error( Format( SThreadStopError, [SControlRecv, E.Message] ) );
      Delay( COMMON_DELAY );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.ReleaseFeedbackSocketThread;
var
  aIndex: Integer;
begin
  if FFeedbackManager.Count <= 0 then Exit;
  { 先打出停止的訊息 }
  for aIndex := 0 to FFeedbackManager.Count - 1 do
    PostMessage( TSMSSocketThread(
      FFeedbackManager.Items[aIndex] ).WndHandle, WM_STOP, 0, 0 );
  Delay( 1000 );
  FMsgSubject.Info( Format( SThreadStop, [SFeedbackSend] ) );
  Delay( COMMON_DELAY );
  { 停止傳送指令執行緒 }
  try
    FFeedbackManager.StopThread( 0 );
    FMsgSubject.Info( Format( SThreadEnd, [SFeedbackSend] ) );
    Delay( COMMON_DELAY );
  except
    on E: Exception do
    begin
      FMsgSubject.Error( Format( SThreadStopError, [SFeedbackSend, E.Message] ) );
      Delay( COMMON_DELAY );
    end;
  end;
  { 有可能在建立接收回應的執行緒時失敗, 所以 ThreadManager 只有傳送的 }
  FMsgSubject.Info( Format( SThreadStop, [SFeedbackRecv] ) );
  Delay( COMMON_DELAY );
  if FFeedbackManager.Count = 0 then Exit;
  try
    FFeedbackManager.StopThread( 1 );
    FMsgSubject.Info( Format( SThreadEnd, [SFeedbackRecv] ) );
    Delay( COMMON_DELAY );
  except
    on E: Exception do
    begin
      FMsgSubject.Error( Format( SThreadStopError, [SFeedbackRecv, E.Message] ) );
      Delay( COMMON_DELAY );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.CreateSoDbThread;
var
  aSoDbThread: TSoDbThread;
  aIndex: Integer;
  aList: TList;
begin
  FMsgSubject.Info( Format( SThreadStart, [SSoDb] ) );
  Delay( COMMON_DELAY );
  { 各系統台自獨立的執行緒 }
  if ConfigModule.CommEnv.DbMultiThread then
  begin
    aList := ConfigModule.SoList.LockList;
    try
      for aIndex := 0 to aList.Count - 1 do
      begin
        if PSoInfo( aList.Items[aIndex] ).Selected then
        begin
          aSoDbThread := TSoDbThread.Create( PSoInfo( aList.Items[aIndex] ) );
          aSoDbThread.MainFormHandle := Self.Handle;
          aSoDbThread.UpdateGUI := dxDbTree.Down;
          aSoDbThread.CommEnv := ConfigModule.CommEnv;
          aSoDbThread.HighCmdEnv := ConfigModule.HighCmdEnv;
          aSoDbThread.CASEnv := ConfigModule.CASEnv;
          aSoDbThread.CmdError := ConfigModule.CmdError;
          aSoDbThread.CommandSubject.ThreadType := tdControlDb;
          if dxControl.Down then
            aSoDbThread.CommandSubject.AddObServer( FControlCmdObSrv );
          aSoDbThread.MessageSubject.AddObServer( FStartupObSrv );
          aSoDbThread.MessageSubject.AddObServer( FSoDatabaseObSrv );
          FSoThreadManager.Add( aSoDbThread );
        end;
      end;
    finally
      ConfigModule.SoList.UnlockList;
    end;
  end else
  begin
    aSoDbThread := TSoDbThread.Create( ConfigModule.SoList );
    aSoDbThread.MainFormHandle := Self.Handle;
    aSoDbThread.UpdateGUI := dxDbTree.Down;
    aSoDbThread.CommEnv := ConfigModule.CommEnv;
    aSoDbThread.HighCmdEnv := ConfigModule.HighCmdEnv;
    aSoDbThread.CASEnv := ConfigModule.CASEnv;
    aSoDbThread.CmdError := ConfigModule.CmdError;
    aSoDbThread.CommandSubject.ThreadType := tdControlDb;
    if dxControl.Down then
      aSoDbThread.CommandSubject.AddObServer( FControlCmdObSrv );
    aSoDbThread.MessageSubject.AddObServer( FStartupObSrv );
    aSoDbThread.MessageSubject.AddObServer( FSoDatabaseObSrv );
    FSoThreadManager.Add( aSoDbThread );
  end;
  FMsgSubject.Info( SDbConnectStart );
  Delay( COMMON_DELAY );
  { 所有系統台執行緒連接資料庫 }
  for aIndex := 0 to FSoThreadManager.Count - 1 do
  begin
    if TSoDbThread( FSoThreadManager.Items[aIndex] ).Suspended then
      TSoDbThread( FSoThreadManager.Items[aIndex] ).Resume;
  end;
  Delay( 5000 );
  FMsgSubject.Info( SDbConnectEnd );
  Delay( COMMON_DELAY );
  for aIndex := 0 to FSoThreadManager.Count - 1 do
  begin
    PostMessage( TSoDbThread( FSoThreadManager.Items[aIndex] ).WndHandle,
      WM_PLAY, 0, 0 );
  end;
  Delay( COMMON_DELAY );
  FMsgSubject.Info( Format( SThreadReady, [SSoDb] ) );
  Delay( COMMON_DELAY );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.ReleaseSoDbThread;
var
  aIndex: Integer;
begin
  if FSoThreadManager.Count <= 0 then Exit;
  FMsgSubject.Info( Format( SThreadStop, [SSoDb] ) );
  Delay( COMMON_DELAY );
  FMsgSubject.Info( SDbDisConnectStart );
  { 先打出停止的訊息 }
  for aIndex := 0 to FSoThreadManager.Count - 1 do
    PostMessage( TSoDbThread( FSoThreadManager.Items[aIndex] ).WndHandle, WM_STOP, 0, 0 );
  Delay( 3000 );
  { 正式停止 }
  for aIndex := 0 to FSoThreadManager.Count - 1 do
  begin
    try
      FSoThreadManager.StopThread( aIndex );
      Delay( COMMON_DELAY );
    except
      on E: Exception do
      begin
        FMsgSubject.Error( Format( SThreadStopError, [SSoDb,E.Message] ) );
        Delay( COMMON_DELAY );
      end;
    end;
  end;
  FMsgSubject.Info( Format( SThreadEnd, [SSoDb] ) );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.CreateFeedbackDbThread;
var
  aThread: TFeedbackDbThread;
begin
  FMsgSubject.Info( Format( SThreadStart, [SFeedbackDb] ) );
  Delay( COMMON_DELAY );
  aThread := TFeedbackDbThread.Create( ConfigModule.FeedbackSo );
  aThread.MainFormHandle := Self.Handle;
  aThread.UpdateGUI := dxDbTree.Down;
  aThread.CommEnv := ConfigModule.CommEnv;
  aThread.HighCmdEnv := ConfigModule.HighCmdEnv;
  aThread.CASEnv := ConfigModule.CASEnv;
  aThread.CmdError := ConfigModule.CmdError;
  aThread.CommandSubject.ThreadType := tdFeedbackDb;
  aThread.MessageSubject.AddObServer( FStartupObSrv );
  aThread.MessageSubject.AddObServer( FSoDatabaseObSrv );
  FFeedbackDbManager.Add( aThread );
  FMsgSubject.Info( Format( SThreadReady, [SFeedbackDb] ) );
  Delay( COMMON_DELAY );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.ReleaseFeedbackDbThread;
var
  aIndex: Integer;
begin
  if FFeedbackDbManager.Count <= 0 then Exit;
  FMsgSubject.Info( Format( SThreadStop, [SFeedbackDb] ) );
  Delay( COMMON_DELAY );
  FMsgSubject.Info( SFeedbackDbDisConnectStart );
  { 先打出停止的訊息 }
  for aIndex := 0 to FFeedbackDbManager.Count - 1 do
    PostMessage( TSoDbThread( FFeedbackDbManager.Items[aIndex] ).WndHandle, WM_STOP, 0, 0 );
  Delay( 3000 );
  { 正式停止 }
  for aIndex := 0 to FFeedbackDbManager.Count - 1 do
  begin
    try
      FFeedbackDbManager.StopThread( aIndex );
      Delay( COMMON_DELAY );
    except
      on E: Exception do
      begin
        FMsgSubject.Error( Format( SThreadStopError, [SFeedbackDb, E.Message] ) );
        Delay( COMMON_DELAY );
      end;
    end;
  end;  
  FMsgSubject.Info( Format( SThreadEnd, [SFeedbackDb] ) );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.LoadConfigTimerTimer(Sender: TObject);
begin
  LoadConfigTimer.Enabled := False;
  AutoExecTimer.Enabled := False;
  PlayCmdGroup.Enabled := False;
  cmbConfigList.Enabled := False;
  Screen.Cursor := crHourGlass;
  try
    FMsgSubject.Info( SConfigSearchFile );
    Delay( 500 );
    ConfigModule.ConfigFileManager.Refresh;
    ClearConfigFileList;
    BuildConfigFileList;
    AppointConfigFile;
    Delay( 500 );
    FMsgSubject.Info( Format( SConfigSearchCompelete,
      [ConfigModule.ConfigFileManager.FileCount] ) );
    Delay( 500 );
    ConfigModule.ConfigLoadFromFile;
    Delay( 500 );
    ClearSoDbTree;
    Delay( 300 );
    if ( ConfigModule.SoCount > 0 ) then
    begin
      FMsgSubject.Info( SConfigBuildSoTreeStart );
      Delay( 500 );
      BuildSoDbTree;
      FMsgSubject.Info( SConfigBuildSoTreeEnd );
      Delay( COMMON_DELAY );
      BuildFeedbackDbTree;
      Delay( COMMON_DELAY );
      if ( ConfigModule.CmdTransEnv.EnableControlSend or
           ConfigModule.CmdTransEnv.EnableFeedbackRecv ) = False then
      begin
        FMsgSubject.Error( Format( SUnknowExec, [SControlSend, SFeedbackRecv] ) );
      end else
      begin
        if ConfigModule.CmdTransEnv.EnableControlSend then
          FMsgSubject.Info( Format( SExecMode, [SControlSend] ) );
        Delay( COMMON_DELAY );
        if ConfigModule.CmdTransEnv.EnableFeedbackRecv then
          FMsgSubject.Info( Format( SExecMode, [SFeedbackRecv] ) );
        Delay( COMMON_DELAY );
        AutoExecTimer.Enabled := ConfigModule.AppStartupInfo.AutoExecute;
        if AutoExecTimer.Enabled then
          FMsgSubject.Info( SAutoExec )
        else begin
          FMsgSubject.Info( SNoneAutoExec );
          PlayCmdGroup.Enabled := True;
        end;
      end;
    end;
    cmbConfigList.Enabled := True;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.AutoExecTimerTimer(Sender: TObject);
begin
  { 將 Start 這個 Group 的 Button Enable 起來 }
  dxPlay.Click;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.BuildSoDbTree;

  { -------------------------------------------------- }

  function GetPosNode(const aPosName: String; const aCanAdd: Boolean = True):
    TcxTreeListNode;
  begin
    Result := nil;
    if ( aPosName = EmptyStr ) then Exit;
    Result := SoTree.FindNodeByText( aPosName, SoTreeCol2 );
    if not Assigned( Result ) and aCanAdd then
    begin
      Result := SoTree.Add( nil );
      Result.AssignValues( [EmptyStr, aPosName, EmptyStr] );
      Result.ImageIndex := 0;
      Result.SelectedIndex := 0;
    end;
  end;

  { -------------------------------------------------- }

var
  aIndex, aMaxCompCode: Integer;
  aList: TList;
  aSo: PSoInfo;
  aNodePos, aNodeComp: TcxTreeListNode;
begin
  TcxProgressBarProperties( SoTreeCol4.Properties ).Max := ConfigModule.CommEnv.DbWaterMark;
  aList := ConfigModule.SoList.LockList;
  try
    aMaxCompCode := 0;
    for aIndex := 0 to aList.Count - 1 do
    begin
      aSo := PSoInfo( aList.Items[aIndex] );
      aMaxCompCode := Max( StrToInt( aSo.CompCode ), aMaxCompCode );
      { 取平台節點, 此系統台要掛在那一個平台下 }
      aNodePos := GetPosNode( aSo.PosName );
      aNodeComp := SoTree.AddChild( aNodePos, aSO );
      aSo.Data := aNodeComp;
      aNodeComp.AssignValues( [aSo.CompCode, aSo.CompName, EmptyStr, aSo.BufferCount] );
      aNodeComp.ImageIndex := GetSoImageIndex( aSo.DbConnectStatus );
      aNodeComp.StateIndex := GetSoStateIndex( aSo.DbConnectStatus );
      aNodeComp.SelectedIndex := aNodeComp.ImageIndex;
      aNodePos.Expanded := True;
      FMsgSubject.Info( Format( SConfigBuildSoTree, [aSo.CompName] ) );
      Delay( COMMON_DELAY );
    end;
    { 給顯示指令用的系統台名稱 }
    SetLength( FSoName, 0 );
    SetLength( FSoName, aMaxCompCode + 1 );
    { 將系統台名稱依照 CompCode 位置寫到對應位置的 array 裏 }
    { FSoName 這個 array , 是給 Show 傳送指令用 }
    for aIndex := 0 to SoTree.Nodes.Count - 1 do
    begin
      if SoTree.Nodes.Items[aIndex].Level > 0 then
      begin
        aSo := PSoInfo( SoTree.Nodes.Items[aIndex].Data );
        FSoName[StrToInt( aSo.CompCode )] := aSo.CompName;
        SoTree.Nodes.Items[aIndex].Data := nil;
      end;  
    end;
  finally
    ConfigModule.SoList.UnlockList;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.ClearSoDbTree;
begin
  SoTree.BeginUpdate;
  try
    SoTree.Clear;
  finally
    SoTree.EndUpdate;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.BuildFeedbackDbTree;

  { -------------------------------------------------- }

  function GetPosNode(const aPosName: String; const aCanAdd: Boolean = True):
    TcxTreeListNode;
  begin
    Result := nil;
    if ( aPosName = EmptyStr ) then Exit;
    Result := SoTree.FindNodeByText( aPosName, SoTreeCol2 );
    if not Assigned( Result ) and aCanAdd then
    begin
      Result := SoTree.Add( nil );
      Result.AssignValues( [EmptyStr, aPosName, EmptyStr] );
      Result.ImageIndex := 0;
      Result.SelectedIndex := 0;
    end;
  end;

  { -------------------------------------------------- }

var
  aNodePos, aNodeComp: TcxTreeListNode;
begin
  if ( ConfigModule.FeedbackSo.Pos < 0 ) then Exit;
  aNodePos := GetPosNode( ConfigModule.FeedbackSo.PosName );
  aNodeComp := SoTree.AddChild( aNodePos, ConfigModule.FeedbackSo );
  //ConfigModule.FeedbackSo.ItemIndex := aNodeComp.AbsoluteIndex;
  ConfigModule.FeedbackSo.Data := aNodeComp;  
  aNodeComp.AssignValues( [EmptyStr, ConfigModule.FeedbackSo.CompName, EmptyStr] );
  aNodeComp.ImageIndex := GetSoImageIndex( ConfigModule.FeedbackSo.DbConnectStatus );
  aNodeComp.StateIndex := GetSoStateIndex( ConfigModule.FeedbackSo.DbConnectStatus );
  aNodeComp.SelectedIndex := aNodeComp.ImageIndex;
  aNodePos.Expanded := True;
  FMsgSubject.Info( Format( SConfigBuildSoTree,
    [ConfigModule.FeedbackSo.CompName] ) );
  Delay( COMMON_DELAY );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.BuildConfigFileList;
var
  aIndex, aPos: Integer;
  aItem: TcxImageComboBoxItem;
begin
  cmbConfigList.Properties.OnChange := nil;
  try
    for aIndex := 0 to ConfigModule.ConfigFileManager.FileCount - 1 do
    begin
      aPos := LastDelimiter( '.', ConfigModule.ConfigFileManager.Files[aIndex] );
      aItem := cmbConfigList.Properties.Items.Add;
      aItem.ImageIndex := 7;
      aItem.Value := ConfigModule.ConfigFileManager.Files[aIndex];
      aItem.Description := Copy( ConfigModule.ConfigFileManager.Files[aIndex],
        1, aPos - 1 );
    end;
  finally
    cmbConfigList.Properties.OnChange := cmbConfigListPropertiesChange;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.ClearConfigFileList;
begin
  cmbConfigList.Properties.OnChange := nil;
  try
    cmbConfigList.Properties.Items.Clear;
  finally
    cmbConfigList.Properties.OnChange := cmbConfigListPropertiesChange;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.AppointConfigFile;
var
  aIndex: Integer;
  aCompareText: String;
begin
  cmbConfigList.Properties.OnChange := nil;
  try
    cmbConfigList.ItemIndex := -1;
    for aIndex := 0 to cmbConfigList.Properties.Items.Count - 1 do
    begin
      aCompareText := cmbConfigList.Properties.Items[aIndex].Description +
        DEFAULT_CONFIGFILE_EXT;
      if AnsiCompareText( aCompareText,
         ConfigModule.ConfigFileManager.ActiveFile ) = 0 then
      begin
        cmbConfigList.ItemIndex := aIndex;
        Break;
      end;
    end;
  finally
    cmbConfigList.Properties.OnChange := cmbConfigListPropertiesChange;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.BeginExecute;
begin
  FDbWarning.Flickered := False;
  FControlSendWarning.Flickered := False;
  FFeedbackRecvWarning.Flickered := False;
  CreateGuiCleanupThread;
  { 啟用傳送指令 }
  if ConfigModule.CmdTransEnv.EnableControlSend then
  begin
    { CAS 指令計數器 } 
    ControlSendMaxCounter := ConfigModule.CASEnv.CmdMaxCounter;
    AllocateControlResource;
    CreateControlSocketThread;
    CreateSoDbThread;
    GuaranteeControlSendExecute;
  end;
  { 啟用接收回傳資料 }
  if ConfigModule.CmdTransEnv.EnableFeedbackRecv then
  begin
    CreateFeedbackDbThread;
    GuaranteeFeedbackRecvExecute;
  end;
  if ( ConfigModule.CmdTransEnv.EnableControlSend or
       ConfigModule.CmdTransEnv.EnableFeedbackRecv ) then
  begin
    StopCmdGroup.Enabled := True;
    PlayCmdGroup.Enabled := False;
    cmbConfigList.Enabled := False;
  end;  
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.EndExecute;
begin
  ReleaseGuiCleanupThread;
  ReleaseControlSocketThread;
  ReleaseSoDbThread;
  ReleaseFeedbackSocketThread;
  ReleaseFeedbackDbThread;
  CleanupTaskResource;
  if ConfigModule.CmdTransEnv.EnableControlSend then
    dxEnableControlSend.ImageIndex := 17;
  if ConfigModule.CmdTransEnv.EnableFeedbackRecv then
    dxEnableFeedbackRecv.ImageIndex := 17;
  if ( ConfigModule.CmdTransEnv.EnableControlSend or
       ConfigModule.CmdTransEnv.EnableFeedbackRecv ) then
  begin
    StopCmdGroup.Enabled := False;
    PlayCmdGroup.Enabled := True;
    cmbConfigList.Enabled := True;
  end;
  FDbWarning.Flickered := False;
  FControlSendWarning.Flickered := False;
  FFeedbackRecvWarning.Flickered := False;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.GuaranteeControlSendExecute;
var
  aIndex: Integer;
begin
  { 1.先確定傳送指令的 Thread 可以連接上 CAS 主機,
    2.再啟動連接資料庫抓取指令的 Thread }
  { 恢復傳送及接收回應的執行緒, 連線狀態由執行緒內自行管理 }
  for aIndex := 0 to FControlManager.Count - 1 do
    TSMSSocketThread( FControlManager.Items[aIndex] ).Resume;
  Delay( COMMON_DELAY );
  { 傳送指令 --> 接收, 打出執行訊號 }
  PostMessage( TSMSSocketThread( FControlManager.Items[0] ).WndHandle, WM_PLAY, 0, 0 );
  { 傳送指令 --> 傳送, 打出執行訊號 }
  PostMessage( TSMSSocketThread( FControlManager.Items[1] ).WndHandle, WM_PLAY, 0, 0 );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.GuaranteeFeedbackRecvExecute;
begin
  { 1.先恢復執行負責寫入回傳資料的 Db Thread,  }
  FMsgSubject.Info( SFeedbackDbConnectStart );
  FFeedbackDbManager.Items[0].Resume;
  PostMessage( TFeedbackDbThread( FFeedbackDbManager.Items[0] ).WndHandle,
    WM_PLAY, 0, 0 );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.UpdateSoTreeStatus(aSoInfo: Pointer);
var
  aNode: TcxTreeListNode;
begin
  if not Assigned( aSoInfo ) then Exit;
  if not Assigned( PSoInfo( aSoInfo ).Data ) then Exit;
  aNode := TcxTreeListNode( PSoInfo( aSoInfo ).Data );
  SoTree.BeginUpdate;
  try
    aNode.ImageIndex := GetSoImageIndex( PSoInfo( aSoInfo ).DbConnectStatus );
    aNode.SelectedIndex := aNode.ImageIndex;
    aNode.StateIndex := GetSoStateIndex( PSoInfo( aSoInfo ).DbConnectStatus );
    if PSoInfo( aSoInfo ).DbConnectStatus = dbNone then
    begin
      aNode.Values[SoTreeCol3.ItemIndex] := EmptyStr;
      aNode.Values[SoTreeCol4.ItemIndex] := 0;
    end else
    begin
      aNode.Values[SoTreeCol3.ItemIndex] := PSoInfo( aSoInfo ).RecordCount;
      aNode.Values[SoTreeCol4.ItemIndex] := PSoInfo( aSoInfo ).BufferCount;
    end;
  finally
    SoTree.EndUpdate;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.AllocateControlResource;
var
  aList: TList;
  aIndex: Integer;
  aCmdList: TThreadStringList;
begin
  aList := ConfigModule.SoList.LockList;
  try
    for aIndex := 0 to aList.Count - 1 do
    begin
      if PSoInfo( aList.Items[aIndex] ).Selected then
      begin
        { 1. 建立 ControlSend 指令傳送的 List, 各系統台獨立 }
        aCmdList := TThreadStringList.Create;
        { ControlSendManager 的識別碼是 SEQNO, 而且有指令傳送先後順序的問題,
          所以 Sort 的屬性不可設 }
        aCmdList.Sorted := False;
        ControlSendManager.Add( aCmdList );
        { 2. 建立 ControlSendDoneManager 已經傳送指令出去傳送的 List, 各系統台獨立 }
        aCmdList := TThreadStringList.Create;
        { ControlSendDoneManager, 是存放已經送出去的指令,
          識別碼是 SEQNO, 可以使用 Sort 屬性加快搜尋速度 }
        aCmdList.Sorted := True;
        ControlSendDoneManager.Add( aCmdList );
        { 3. 建立 Ack 回來的 List, 各系統台獨立 }
        aCmdList := TThreadStringList.Create;
        { 可以使用 Sort, 雖然識別碼是 SEQNO, 可是已經沒有傳送指令順序問題,
          使用 Sort 屬性, 可以加快搜尋速度  }
        aCmdList.Sorted := True;
        ControlRecvManager.Add( aCmdList );
        { 4. 建立寫 Log List, 各系統台獨立 }
        aCmdList := TThreadStringList.Create;
        aCmdList.Sorted := True;
        ControlSendLogManager.Add( aCmdList );
      end;
    end;
  finally
    ConfigModule.SoList.UnlockList;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.CleanupTaskResource;
var
  aManagerIndex, aIndex, aIndex2: Integer;
  aCommandManager: TCommandListMagager;
  aManagerName: String;
begin
  { 在這裏最後清除尚未處理的 List 資料 }
  for aManagerIndex := 1 to 4 do
  begin
    try
      case aManagerIndex of
        1: begin
             aCommandManager := ControlSendManager; { 1.清除傳送指令的 List }
             aManagerName := SControlSendManager;
           end;
        2: begin
             aCommandManager := ControlSendDoneManager; { 2.指令已送出去的 List }
             aManagerName := SControlSendDoneManager;
           end;
        3: begin
             aCommandManager := ControlRecvManager;  { 3.指令Ack回來的 List }
             aManagerName := SControlRecvManager;
           end;
        4: begin
             aCommandManager := ControlSendLogManager; { 4.實際傳送出去的低階指令(寫Log用) }
             aManagerName := SControlSendLogManager;
           end;
      else
        aCommandManager := nil;
      end;
      for aIndex := 0 to aCommandManager.ItemCount - 1 do
      begin
        for aIndex2 := 0 to aCommandManager.Items[aIndex].Count - 1 do
        begin
          Dispose( PSendNagra( aCommandManager.Items[aIndex].Objects[aIndex2] ) );
          aCommandManager.Items[aIndex].Objects[aIndex2] := nil;
        end;
      end;
      aCommandManager.Clear;
    except
      on E: Exception do
      begin
        FMsgSubject.Error( Format( SCleanupCommandManagerError,
          [aManagerName, E.Message] ) );
        Delay( 100 );
      end;
    end;
  end;
  { 傳送指令的 CommunicationCheck List }
  for aIndex := 0 to ControlCommunicationCheck.Count - 1 do
    Dispose( PSendNagra( ControlCommunicationCheck.Objects[aIndex] ) );
  ControlCommunicationCheck.Clear;
  { 接收 Callback 資料的 CommunicationCheck List }
  for aIndex := 0 to FeedbackCommunicationCheck.Count - 1 do
    Dispose( PRecvNagra( FeedbackCommunicationCheck.Objects[aIndex] ) );
  FeedbackCommunicationCheck.Clear;
  { 傳送指令 Ack List }
  for aIndex := 0 to ControlAck.Count - 1 do
    Dispose( PSendNagra( ControlAck.Objects[aIndex] ) );
  ControlAck.Clear;
  { 回傳資料 Ack List }
  for aIndex := 0 to FeedbackAck.Count - 1 do
    Dispose( PRecvNagra( FeedbackAck.Objects[aIndex] ) );
  FeedbackAck.Clear;
  { 回應接收資料的 List }
  for aIndex := 0 to FeedbackSendList.Count - 1 do
    Dispose( PRecvNagra( FeedbackSendList.Objects[aIndex] ) );
  FeedbackSendList.Clear;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.ControlSendTreeCustomDrawCell(Sender: TObject;
  ACanvas: TcxCanvas; AViewInfo: TcxTreeListEditCellViewInfo;
  var ADone: Boolean);
var
  aIconLeft, aIconTop, aFlag, aOffsetX: Integer;
  aTextRect: TRect;
begin
  if AViewInfo.Node.Level = 0 then
  begin
    ACanvas.Brush.Color := $00E1E6E8;
    ACanvas.FillRect( AViewInfo.VisibleRect );
    if AViewInfo.Column = ControlSendTreeColumn7 then
    begin
      aIconLeft := AViewInfo.VisibleRect.Left + 3;
      aIconTop := AViewInfo.VisibleRect.Top;
      ACanvas.DrawImage( StyleModule.CmdTreeImageList, aIconLeft,
        aIconTop, 7 );
      aFlag := ( cxSingleLine	or cxShowEndEllipsis or cxAlignLeft	or cxAlignVCenter );
      aTextRect := AViewInfo.VisibleRect;
      aOffsetX := StyleModule.CmdTreeImageList.Width + 8;
      InflateRect( aTextRect, ( 0 - aOffsetX ), 0 );
      ACanvas.DrawText( AViewInfo.DisplayValue, aTextRect, aFlag );
      ADone := True;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.dxPlayClick(Sender: TObject);
begin
  BeginExecute;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.dxStopClick(Sender: TObject);
begin
  EndExecute;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.dxAutoScrollClick(Sender: TObject);
begin
  FAutoScroll := dxAutoScroll.Down;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.dxControlClick(Sender: TObject);
var
  aIndex: Integer;
begin
  for aIndex := FControlManager.Count - 1 downto 0 do
    PostMessage( TCommonThread( FControlManager.Items[aIndex] ).WndHandle,
      WM_UPDATEGUI, Integer( FControlCmdObSrv ), Ord( dxControl.Down ) );
  for aIndex := FSoThreadManager.Count - 1 downto 0 do
    PostMessage( TCommonThread( FSoThreadManager.Items[aIndex] ).WndHandle,
      WM_UPDATEGUI, Integer( FControlCmdObSrv ), Ord( dxControl.Down ) );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.dxConsoleClick(Sender: TObject);
var
  aIndex: Integer;
begin
  for aIndex := 0 to FControlManager.Count - 1 do
  begin
    PostMessage( TCommonThread( FControlManager.Items[aIndex] ).WndHandle,
      WM_UPDATEGUI, Integer( FConsoleObSrv ), Ord( dxConsole.Down ) );
  end;
  for aIndex := 0 to FFeedbackManager.Count - 1 do
  begin
    PostMessage( TCommonThread( FFeedbackManager.Items[aIndex] ).WndHandle,
      WM_UPDATEGUI, Integer( FConsoleObSrv ), Ord( dxConsole.Down ) );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.dxDbTreeClick(Sender: TObject);
var
  aIndex: Integer;
begin
  for aIndex := 0 to FSoThreadManager.Count - 1 do
    TSoDbThread( FSoThreadManager.Items[aIndex] ).UpdateGUI := dxDbTree.Down;
  if FFeedbackDbManager.Count > 0 then
    TFeedbackDbThread( FFeedbackDbManager.Items[0] ).UpdateGUI := dxDbTree.Down;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.WMDatabase(var Msg: TMessage);
begin
  { 顯示警告畫面 }
  FDbWarning.Flickered := not Boolean( Msg.LParam );
  if ( Msg.WParam = Ord( tdFeedbackDb ) ) then
  begin
    { 通知接收回傳資料的執行緒, 停止接收回傳資料 }
    if not Boolean( Msg.LParam ) then
      ReleaseFeedbackSocketThread
    else begin
      if FFeedbackManager.Count <= 0 then
        CreateFeedbackSocketThread;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.WMSocket(var Msg: TWMSocket);
var
  aThreadType: TThreadType;
  aCommandType: TNotifyCommandType;
  aIndex: Integer;
begin
  aThreadType := TThreadType( Msg.ThreadType );
  aCommandType := TNotifyCommandType( Msg.NotifyCommandType );
  if ( aThreadType in [tdControlSend, tdControlRecv] ) then
  begin
    FControlSendWarning.Flickered := ( aCommandType in [ncNack, ncSocketError] );
    { 對各系統台的 DB Thread 發送訊息 }
    for aIndex := 0 to FSoThreadManager.Count - 1 do
    begin
      PostMessage( TSoDbThread( FSoThreadManager.Items[aIndex] ).WndHandle,
        Msg.Msg, TMessage( Msg ).WParam, TMessage( Msg ).LParam );
    end;
  end else
  if ( aThreadType in [tdFeedbackSend, tdFeedbackRecv] ) then
  begin
    FFeedbackRecvWarning.Flickered := ( aCommandType in [ncNack, ncSocketError] );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.dxDockPanelResize(Sender: TObject);
var
  aControl: TFlickerControl;
begin
  aControl := nil;
  if Sender = dxDockPanel1 then
    aControl := FDbWarning
  else
  if Sender = dxDockPanel2 then
    aControl := FControlSendWarning
  else
  if Sender = dxDockPanel4 then
    aControl := FFeedbackRecvWarning;
  if Assigned( aControl ) then
  begin
    aControl.Top := ( aControl.Parent.Height - aControl.Height ) div 2;
    aControl.Left := ( aControl.Parent.Width - aControl.Width ) div 2;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.dxCmdExpandClick(Sender: TObject);
begin
  ControlSendTree.FullExpand;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.dxCmdCollapseClick(Sender: TObject);
begin
 ControlSendTree.FullCollapse;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.dxMenuItem11Click(Sender: TObject);
begin
  { 正在執行中不可按選項 }
  if ConfigModule.ShowConfigForm = mrOk then
  begin
    LoadConfigTimer.Interval := 500;
    LoadConfigTimer.Enabled := True;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.cmbConfigListPropertiesChange(Sender: TObject);
var
  aFileName: String;
begin
  { 完整檔案路徑 + 名稱, Ex: D:\App\北平台.cfg }
  aFileName := IncludeTrailingPathDelimiter( ExtractFilePath( ParamStr( 0 ) ) ) +
    cmbConfigList.Properties.Items[cmbConfigList.ItemIndex].Description +
      DEFAULT_CONFIGFILE_EXT;
  if not FileExists( aFileName ) then
  begin
    FMsgSubject.Info( Format( SConfigFileNotExists, [aFileName] ) );
  end else
  begin
    FMsgSubject.Info( Format( SConfigFileChange, [aFileName] ) );
    ConfigModule.ConfigFileManager.ActiveFile :=
      cmbConfigList.Properties.Items[cmbConfigList.ItemIndex].Description +
        DEFAULT_CONFIGFILE_EXT;
    LoadConfigTimer.Interval := 500;
    LoadConfigTimer.Enabled := True;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.dxMenuItem21Click(Sender: TObject);
begin
  FDocklayoutStream.Position := 0;
  dxDockingManager.LoadLayoutFromStream( FDocklayoutStream );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.dxMenuItem22Click(Sender: TObject);
var
  aDockPanel: TdxDockPanel;
begin
  aDockPanel := nil;
  if Sender = dxMenuItem22 then
  begin
    aDockPanel := dxDockPanel1;
  end else
  if Sender = dxMenuItem23 then
  begin
    aDockPanel := dxDockPanel2;
  end else
  if Sender = dxMenuItem24 then
  begin
    aDockPanel := dxDockPanel3;
  end else
  if Sender = dxMenuItem25 then
  begin
    aDockPanel := dxDockPanel4;
  end else
  if Sender = dxMenuItem26 then
  begin
    aDockPanel := dxDockPanel5;
  end else
  if Sender = dxMenuItem27 then
  begin
    aDockPanel := dxDockPanel6;
  end else
  if Sender = dxMenuItem28 then
  begin
    aDockPanel := dxDockPanel7;
  end else
  if Sender = dxMenuItem29 then
  begin
    aDockPanel := dxDockPanel8;
  end;
  if Assigned( aDockPanel ) then
  begin
    if not aDockPanel.Visible then aDockPanel.Visible := True;
    aDockPanel.Controller.ActiveDockControl := aDockPanel;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.CreateGuiCleanupThread;
begin
  FGuiCleanupThread := TGuiCleanup.Create;
  FGuiCleanupThread.Resume;
  PostMessage( FGuiCleanupThread.WndHandle, WM_PLAY, 0, 0 );
  FMsgSubject.Info( Format( SThreadReady, [SGuiCleanup] ) );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.ReleaseGuiCleanupThread;
begin
  if not Assigned( FGuiCleanupThread ) then Exit;
  PostMessage( FGuiCleanupThread.WndHandle, WM_STOP, 0, 0 );
  Delay( 300 );
  try
    FGuiCleanupThread.Terminate;
    FGuiCleanupThread.WaitFor;
    FGuiCleanupThread.Free;
    FGuiCleanupThread := nil;
    FMsgSubject.Info( Format( SThreadEnd, [SGuiCleanup] ) );
  except
    on E: Exception do
      FMsgSubject.Error( Format( SThreadStopError, [SGuiCleanup, E.Message] ) );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.dxDockPanel2Activate(Sender: TdxCustomDockControl;
  Active: Boolean);
begin
  ControlSendGroup.Enabled := Active;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.dxMenuItem3Click(Sender: TObject);
begin
  fmAbout :=  TfmAbout.Create( nil );
  try
    fmAbout.ShowModal;
  finally
    fmAbout.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.ConsoleTreeKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if ( ssCtrl in Shift  ) then
  begin
    if ( Chr( Key ) = 'C' ) then
      ConsoleTree.CopySelectedToClipboard
    else if ( Chr( Key ) = 'A' ) then
      ConsoleTree.SelectAll;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.CheckLicence: Boolean;
begin
  Result := FLicenceManager.LoadLicenceFile;
  if not Result then
    Result := FLicenceManager.ActiveUserInputForm;
  if Result then
    Result := FLicenceManager.ValidateLicence;  
end;

{ ---------------------------------------------------------------------------- }

end.
