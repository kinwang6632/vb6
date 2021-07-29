unit frmMainU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, ComCtrls, Menus, ImgList, StdCtrls, Buttons, DB,
  DBClient, Mask, DBCtrls, ActnList, ActnMan, ToolWin, ActnCtrls, ConstU,
  Grids, DBGrids;
type
  TfrmMain = class(TForm)
    StatusBar1: TStatusBar;
    pnlMaster: TPanel;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    TabSheet3: TTabSheet;
    ImageList1: TImageList;
    btnExit: TBitBtn;
    RadioGroup1: TRadioGroup;
    cdsSysParam: TClientDataSet;
    dsrSysParam: TDataSource;
    StaticText1: TStaticText;
    DBEdit1: TDBEdit;
    ActionManager1: TActionManager;
    Action1: TAction;
    ActionToolBar2: TActionToolBar;
    cdsSysParamnPort: TIntegerField;
    cdsSysParamnVersion: TIntegerField;
    cdsSysParamsSecurityType: TStringField;
    DBEdit2: TDBEdit;
    StaticText2: TStaticText;
    DBRadioGroup1: TDBRadioGroup;
    Button1: TButton;
    pnlMonitor: TPanel;
    dbgMonotor: TDBGrid;
    BitBtn1: TBitBtn;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Action1Execute(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure pnlMonitorClick(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
  private
    { Private declarations }
    sG_ExePath : String;
    procedure saveCDS(I_CDS : TClientDataSet; sI_FileName : String);
    function transNum2Hex(nI_Num, nI_Digits : Integer):String;
    function getVersion:String;
    function getSecurityType:String;
    function getLength(sI_Cmd:String) : String;
    function getSignature:String;
    function getFromID:String;
    function getToID:String;
    function getConnectionID:String;
    function getDateString:String;
    function getSequenceID:String;
    function getActionType:String;
    function getActionPriority:String;
    function getPriorityReassignment:String;
    function getMainKey(nI_SubscriberID : Integer):String;
    function getActionCode():String;
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;

implementation

uses Ustru, UdateTimeu, dtmMainU;

{$R *.dfm}

procedure TfrmMain.FormClose(Sender: TObject; var Action: TCloseAction);
begin

    Action := caFree;
    Action1Execute(Sender);
end;

procedure TfrmMain.Action1Execute(Sender: TObject);
begin
    saveCDS(cdsSysParam,sG_ExePath + SYS_PARAM_FILE_NAME);
    Application.Terminate;
end;

procedure TfrmMain.FormCreate(Sender: TObject);
begin
    sG_ExePath := ExtractFilePath(Application.ExeName);
    if FileExists(sG_ExePath + SYS_PARAM_FILE_NAME) then
      cdsSysParam.LoadFromFile(sG_ExePath + SYS_PARAM_FILE_NAME);
end;

procedure TfrmMain.saveCDS(I_CDS: TClientDataSet; sI_FileName : String);
begin
    if (I_CDS.State in [dsEdit, dsInsert]) then
      I_CDS.Post;
    if (I_CDS.ChangeCount>0) then
      I_CDS.SaveToFile(sI_FileName);
end;

function TfrmMain.transNum2Hex(nI_Num, nI_Digits: Integer): String;
var
    sL_Result : String;
begin
    //將十進位數字轉成16進位的字串ex : 十進位的數字32,會轉成'20'
    sL_Result := IntToHex(nI_Num, nI_Digits);
    sL_Result := TUstr.AddString(sL_Result,'0',false,nI_Digits*2);
    result := sL_Result;
end;

procedure TfrmMain.Button1Click(Sender: TObject);
var
    sL_Cmd : String;
begin
    sL_Cmd := getVersion + getSecurityType + getLength('testCmdBody') ;
    ShowMessage(sL_Cmd);
    sL_Cmd := getFromID + getConnectionID + getToID + getDateString ;
    ShowMessage(sL_Cmd);
    sL_Cmd := getSequenceID + getActionType + getActionPriority + getPriorityReassignment;
    ShowMessage(sL_Cmd);    
    sL_Cmd := getMainKey(1234567) + getActionCode()+ getSignature;
    ShowMessage(sL_Cmd);
end;

function TfrmMain.getVersion: String;
var
    nL_Version : Integer;
    sL_Version : String;
begin
    nL_Version := cdsSysParam.FieldByName('nVersion').AsInteger;
    sL_Version := transNum2Hex(nL_Version,2);
    result := sL_Version;
end;

function TfrmMain.getSecurityType: String;
var
    sL_SecurityType : String;
begin
    sL_SecurityType := cdsSysParam.FieldByName('sSecurityType').AsString;
    result := sL_SecurityType;
end;

function TfrmMain.getSignature: String;
var
    nL_Signature : Integer;
    sL_Signature : String;
begin

    nL_Signature := 20;
    sL_Signature := transNum2Hex(nL_Signature,8);
    result := sL_Signature;
end;

function TfrmMain.getLength(sI_Cmd: String): String;
var
    nL_Length : Integer;
    sL_Length : String;
begin
    nL_Length := length(sI_Cmd);
    sL_Length := transNum2Hex(nL_Length,2);
    result := sL_Length;

end;

function TfrmMain.getFromID: String;
var
    nL_FromID : Integer;
    sL_FromID : String;
begin
    nL_FromID := REGULAR_SMS_COMMAND_ID;
    sL_FromID := transNum2Hex(nL_FromID,2);
    result := sL_FromID;

end;

function TfrmMain.getToID: String;
var
    nL_ToID : Integer;
    sL_ToID : String;
begin
    nL_ToID := EMMG_ID;
    sL_ToID := transNum2Hex(nL_ToID,2);
    result := sL_ToID;

end;

function TfrmMain.getConnectionID: String;
var
    nL_ConnectionID : Integer;
    sL_ConnectionID : String;
begin
    nL_ConnectionID := 19;
    sL_ConnectionID := transNum2Hex(nL_ConnectionID,1);
    result := sL_ConnectionID;

end;

function TfrmMain.getDateString: String;
begin
    result := TUdateTime.GetPureDateTimeStr(now);
end;

function TfrmMain.getSequenceID: String;
var
    nL_SequenceID : Integer;
    sL_SequenceID : String;
begin
    nL_SequenceID := 1;
    sL_SequenceID := transNum2Hex(nL_SequenceID,2);
    result := sL_SequenceID;

end;

function TfrmMain.getActionType: String;
begin
    result := ACTION_TYPE_1;
end;

function TfrmMain.getActionPriority: String;
begin
    result := ACTION_PRIORITY_1;
end;

function TfrmMain.getPriorityReassignment: String;
begin
    result := ACTION_PRIORITY_2;
end;

function TfrmMain.getMainKey(nI_SubscriberID: Integer): String;
var
    sL_SubscriberID : String;
begin
    sL_SubscriberID := transNum2Hex(nI_SubscriberID,4);
    result := sL_SubscriberID;

end;

function TfrmMain.getActionCode: String;
begin
end;

procedure TfrmMain.pnlMonitorClick(Sender: TObject);
begin
    if (not dbgMonotor.Visible) then
      dbgMonotor.Visible := true;
end;

procedure TfrmMain.BitBtn1Click(Sender: TObject);
begin
    showmessage(dtmMain.getSequence(now));
end;

end.
