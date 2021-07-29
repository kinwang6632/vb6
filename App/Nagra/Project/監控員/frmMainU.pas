unit frmMainU;

interface

uses
  {
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Psock, NMEcho, Buttons, ComCtrls;
  }
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, TLHelp32, ComCtrls, ExtCtrls, Buttons, ImgList, Psock, NMEcho, inifiles,
  IdEcho, IdFinger, IdBaseComponent, IdComponent, IdTCPConnection,
  IdTCPClient, ScktComp, DB, DBClient;

const
    NULL_PROC_ID = -999;
    MONITOR_LOG_FILE_PATH = 'CaLog';
    MONITOR_LOG_FILE_NAME = 'Monitor.log';
    PORT=8123;      
type
  TfrmMain = class(TForm)
    Panel1: TPanel;
    memInfo: TMemo;
    Panel2: TPanel;
    SaveDialog1: TSaveDialog;
    timAutoInvoke: TTimer;
    Panel3: TPanel;
    CheckBox1: TCheckBox;
    CheckBox2: TCheckBox;
    chbAcceptClient: TCheckBox;
    ServerSocket: TServerSocket;
    Panel4: TPanel;
    BitBtn3: TBitBtn;
    BitBtn2: TBitBtn;
    BitBtn1: TBitBtn;
    edtIniFileName: TEdit;
    StaticText1: TStaticText;
    SpeedButton1: TSpeedButton;
    BitBtn4: TBitBtn;
    OpenDialog1: TOpenDialog;
    stb: TStatusBar;
    cdsParam: TClientDataSet;
    cdsParamsServerIP: TStringField;
    cdsParamnSPortNo: TIntegerField;
    cdsParamnRPortNo: TIntegerField;
    cdsParamsSysName: TStringField;
    cdsParamsSysVersion: TStringField;
    cdsParamsMisCommandPath: TStringField;
    cdsParamnTimeOut: TIntegerField;
    cdsParamnMaxCommandCount: TIntegerField;
    cdsParambCommandLog: TBooleanField;
    cdsParambErrorLog: TBooleanField;
    cdsParamnResponseLog: TBooleanField;
    cdsParamdUptTime: TDateTimeField;
    cdsParamsUptName: TStringField;
    cdsParambSetZipCode: TBooleanField;
    cdsParamnCreditLimit: TIntegerField;
    cdsParamnSourceID: TIntegerField;
    cdsParamnDestID: TIntegerField;
    cdsParamnMopPPID: TIntegerField;
    cdsParamsBroadcastSDate: TStringField;
    cdsParamsBroadcastEDate: TStringField;
    cdsParamnCmdRefreshRate: TIntegerField;
    cdsParamnCmdRefreshRate2: TIntegerField;
    cdsParamnCmdResentTimes: TIntegerField;
    cdsParambShowUI: TBooleanField;
    cdsParambAssignBroadcastDate: TBooleanField;
    cdsParamsSHotTime: TStringField;
    cdsParamsEHotTime: TStringField;
    procedure FormShow(Sender: TObject);

    procedure BitBtn1Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure BitBtn3Click(Sender: TObject);
    procedure timAutoInvokeTimer(Sender: TObject);
    procedure CheckBox2Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure ServerSocketAccept(Sender: TObject;
      Socket: TCustomWinSocket);
    procedure ServerSocketClientRead(Sender: TObject;
      Socket: TCustomWinSocket);
    procedure ServerSocketClientDisconnect(Sender: TObject;
      Socket: TCustomWinSocket);
    procedure BitBtn4Click(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);

  private
    { Private declarations }
    G_ExePathInfoStrList : TStringList;
    sG_RemoteAddress, sG_RemoteHost : String;
    nG_CmdRefreshRate1, nG_CmdRefreshRate2 : Integer;

    sG_TimeDefine : String;
    nG_CaAppId : Integer;
    nG_OperFlag : Integer;
    nG_InvokeCaTimer : Integer;
//    nG_PingRetryTimer : Integer;
//    nG_PingTimer :Integer;
//    sG_EchoIP : String;
//    nG_EchoPort : Integer;
//    nG_UrgentRetryTimes : Integer;
    nG_MemoMaxLineCount : Integer;
    G_MonitorLogStrList : TStringList;
    function reDefineFetchDataInterval(dI_Time:TDate):Integer;    
    procedure GetSysParam;    
    function isCaActive:boolean;
    function getCaTimeStamp:String;
    procedure setDefaultSetting;        
    function getRendomFileName:String;
    procedure writeMonitorLog(sI_Info:String);
    procedure setIniParam;
    procedure runCaCmdGateway(sI_ExePath : String);
    function getSysIniSettiing:WideString;
    function getProcessId(sI_ProcessName:String):Integer;

    procedure setProcessMonitor(bI_Status:boolean);
    procedure setEchoInfo;
    procedure killCaGateway(sI_ExePath : String);

  public
    { Public declarations }
    G_TimeStampStrList : TStringList;
    procedure RunCA(sI_IniFileName:String);    
    procedure InvokeMonitor(sI_Method, sI_User: OleVariant;  var sI_ActionStatusCode, sI_ActionStatusMsg: OleVariant);

  end;

var
  frmMain: TfrmMain;

implementation

uses ConstObjU, Uotheru, Ustru, UdateTimeu;

{$R *.dfm}

{ TfrmMain }

procedure TfrmMain.killCaGateway(sI_ExePath : String);
var
    sL_CaExeName, sL_MemInfo : String;
    L_Handle : THANDLE;
    L_ExitCode : Cardinal;
begin
    sL_CaExeName := ExtractFileName(sI_ExePath);
    sL_MemInfo := '開始停止' + sI_ExePath + ' -- '+ DateTimeToStr(now);

    nG_CaAppId := getProcessId(sL_CaExeName);
    if nG_CaAppId<>NULL_PROC_ID then
    begin
      memInfo.Lines.Insert(0,sL_MemInfo);
      L_Handle := OpenProcess(PROCESS_TERMINATE,true, nG_CaAppId);
      GetExitCodeProcess(L_Handle, L_ExitCode);
      TerminateProcess(L_Handle, L_ExitCode);
      CloseHandle(L_Handle);
    end;
end;

procedure TfrmMain.FormShow(Sender: TObject);
var
    sL_LogFullPath : String;
    sL_ExeFileName, sL_ExePath, sL_FullPathLogFileName : String;
begin
    GetSysParam;
    setIniParam;
    setDefaultSetting;
    //down, 開啟 Log File
    G_MonitorLogStrList := TStringList.Create;
    sL_ExeFileName := Application.ExeName;
    sL_ExePath := ExtractFileDir(sL_ExeFileName);


    sL_LogFullPath := sL_ExePath + '\' + MONITOR_LOG_FILE_PATH;
    if not DirectoryExists(sL_LogFullPath) then
      CreateDir(sL_LogFullPath);

    sL_FullPathLogFileName := sL_LogFullPath + '\' + MONITOR_LOG_FILE_NAME;
    if FileExists(sL_FullPathLogFileName) then
      G_MonitorLogStrList.LoadFromFile(sL_FullPathLogFileName);
    //up, 開啟 Log File

    ServerSocket.Port := PORT;
    ServerSocket.Active := True;
    G_TimeStampStrList := TStringList.Create;
end;

procedure TfrmMain.setEchoInfo;
var
    bL_PriorStatus : boolean;
begin




    bL_PriorStatus := timAutoInvoke.Enabled;
    timAutoInvoke.Enabled := false;
    timAutoInvoke.Interval := nG_InvokeCaTimer;
    timAutoInvoke.Enabled := bL_PriorStatus;


end;


procedure TfrmMain.BitBtn1Click(Sender: TObject);
begin
    ServerSocket.Close;
    Application.Terminate;
end;

procedure TfrmMain.BitBtn2Click(Sender: TObject);
begin
    if SaveDialog1.Execute then
      memInfo.Lines.SaveToFile(SaveDialog1.FileName);
end;

{
show 出系統正在 run 的所有 process..

procedure TfrmMain.getAllProcess;
var
   i:integer;
   ContinueLoop:BOOL;
   NewItem : TListItem;
   FSnapshotHandle:THandle;
   FProcessEntry32:TProcessEntry32;
begin
   lsvProcess.Items.Clear;
   FSnapshotHandle:=CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS,0);
   FProcessEntry32.dwSize:=Sizeof(FProcessEntry32);
   ContinueLoop:=Process32First(FSnapshotHandle,FProcessEntry32);
   while integer(ContinueLoop)<>0 do
   begin
        NewItem:=lsvProcess.Items.add;
        NewItem.Caption:=ExtractFileName(FProcessEntry32.szExeFile);
        NewItem.subItems.Add(IntToHex(FProcessEntry32.th32ProcessID,4));
        NewItem.subItems.Add(FProcessEntry32.szExeFile);
        ContinueLoop:=Process32Next(FSnapshotHandle,FProcessEntry32);
   end;
   CloseHandle(FSnapshotHandle);
end;
}
function TfrmMain.getProcessId(sI_ProcessName: String):Integer;
var
   sL_ProcessName : String;
   bL_ContinueLoop:BOOL;
   FSnapshotHandle:THandle;
   FProcessEntry32:TProcessEntry32;
   nL_ProcessID : Integer;
begin

    nL_ProcessID := NULL_PROC_ID;
    FSnapshotHandle:=CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS,0);
    FProcessEntry32.dwSize:=Sizeof(FProcessEntry32);
    bL_ContinueLoop:=Process32First(FSnapshotHandle,FProcessEntry32);
    while integer(bL_ContinueLoop)<>0 do
    begin
      sL_ProcessName := ExtractFileName(FProcessEntry32.szExeFile);
//      if AnsiPos(sI_ProcessName, sL_ProcessName)>0then
      if AnsiPos(sL_ProcessName, sI_ProcessName)>0then
      begin
        nL_ProcessID := FProcessEntry32.th32ProcessID;
//      NewItem.subItems.Add(IntToHex(FProcessEntry32.th32ProcessID,4));
        break;
      end;
      bL_ContinueLoop:=Process32Next(FSnapshotHandle,FProcessEntry32);
    end;
    CloseHandle(FSnapshotHandle);


    result := nL_ProcessID;
end;
function TfrmMain.getSysIniSettiing: WideString;
var
    L_IniFile : TIniFile;
    sL_CaExePath, sL_ExeFileName, sL_ExePath, sL_IniFileName : STring;

begin
    sL_ExeFileName := Application.ExeName;
    sL_ExePath := ExtractFileDir(sL_ExeFileName);
    if edtIniFileName.Text='' then
      sL_IniFileName := sL_ExePath + '\' + CABLE_NAGRA_INI_FILE_NAME
    else
      sL_IniFileName := sL_ExePath + '\' + edtIniFileName.Text;


    if not FileExists(sL_IniFileName) then
    begin
      result := '-1: 讀取連線參數檔<'+sL_IniFileName +'>失敗';
      exit;
    end;
    L_IniFile := TIniFile.Create(sL_IniFileName);

//    nG_DefaultApFlag := L_IniFile.ReadInteger('MONITOR','DEFAULT_AP_FLAG',1);
    sL_CaExePath := L_IniFile.ReadString('MONITOR','CA_EXE_PATH','');
    G_ExePathInfoStrList := TUstr.ParseStrings(sL_CaExePath,'~');

//    sG_CaProcessorPath := L_IniFile.ReadString('MONITOR','CA_PROCESSOR_PATH','');

//    nG_PingRetryTimer := L_IniFile.ReadInteger('MONITOR','PING_RETRY_TIMER',10)*1000;
//    nG_PingTimer := L_IniFile.ReadInteger('MONITOR','PING_TIMER',300)*1000;
//    sG_EchoIP := L_IniFile.ReadString('MONITOR','PING_IP','');
//    nG_EchoPort := L_IniFile.ReadInteger('MONITOR','ECHO_PORT',0);
//    nG_UrgentRetryTimes := L_IniFile.ReadInteger('MONITOR','URGENT_RETRY_TIMES',10);
    nG_MemoMaxLineCount := L_IniFile.ReadInteger('MONITOR','MEMO_MAX_LINE_COUNT',100);
    nG_InvokeCaTimer := L_IniFile.ReadInteger('MONITOR','INVOKE_CA_TIMER',120)*1000;




    L_IniFile.Free;
    result := '';
end;

procedure TfrmMain.runCaCmdGateway(sI_ExePath : String);
var
    sL_MemInfo1, sL_MemInfo2, sL_MemInfo3 : String;
    bL_Status : boolean;
    sI_ErrorMsg, sL_ApFilePath : String;
begin
    sL_ApFilePath := sI_ExePath;


    bL_Status := TUother.RunShell(sL_ApFilePath,AUTO_INVOKE + ' ' + edtIniFileName.Text ,true, false, sI_ErrorMsg);
    if bL_Status then

      memInfo.Lines.Insert(0,sL_MemInfo1)
    else
    begin
      memInfo.Lines.Insert(0,'        錯誤訊息:' + sI_ErrorMsg);
      memInfo.Lines.Insert(0,sL_MemInfo2)

    end;
end;

procedure TfrmMain.setIniParam;
var
    sL_Result : String;
begin
    //down, 讀取ini 檔之設定
    sL_Result := getSysIniSettiing;
    if sL_Result<>'' then
    begin
      MessageDlg(sL_Result,mtError,[mbOK],0);
      Application.Terminate;
      exit;
    end;
    //up, 讀取ini 檔之設定
    setEchoInfo;
end;

procedure TfrmMain.BitBtn3Click(Sender: TObject);
begin
    setIniParam;
end;



procedure TfrmMain.timAutoInvokeTimer(Sender: TObject);
var
    bG_RestartAllExe : boolean;
    sL_FullpathCaExeName, sL_CaExeName : String;
    bL_IsCaActive : boolean;
    nL_CaAppId, ii : Integer;
    sL_MemInfo1, sL_MemInfo2, sL_MemInfo3, sL_MemInfo4 : String;
begin
    timAutoInvoke.Enabled := false;
    bG_RestartAllExe := false;
    
    for ii:=0 to G_ExePathInfoStrList.Count -1 do
    begin
      sL_FullpathCaExeName := G_ExePathInfoStrList.Strings[ii];
      sL_CaExeName := ExtractFileName(sL_FullpathCaExeName);


      nL_CaAppId := getProcessId(sL_CaExeName);


      if nL_CaAppId=NULL_PROC_ID then
      begin
        bG_RestartAllExe := true;
        writeMonitorLog('偵測到 ' + sL_FullpathCaExeName + ' 尚未執行.');
        writeMonitorLog('即將停止所有 CA 之運作,然後重新啟動.');
        writeMonitorLog('======================================================================');

        break;
      end
    end;
    if bG_RestartAllExe then
    begin
      for ii:=0 to G_ExePathInfoStrList.Count -1 do
      begin
        sL_FullpathCaExeName := G_ExePathInfoStrList.Strings[ii];
        killCaGateway(sL_FullpathCaExeName);
        runCaCmdGateway(sL_FullpathCaExeName);
      end;
      //down, 讓程式畫面產生 refresh 的效果, 若不做,CA有時會無法開始傳送指令

      self.WindowState := wsMinimized;
      self.WindowState := wsNormal;
      self.Repaint;
      //up, 讓程式畫面產生 refresh 的效果, 若不做,CA有時會無法開始傳送指令
    end;

    timAutoInvoke.Enabled := true;
end;
{
procedure TfrmMain.timAutoInvokeTimer(Sender: TObject);
var
    sL_FullpathCaExeName, sL_CaExeName : String;
    bL_IsCaActive : boolean;
    nL_CaAppId, ii : Integer;
    sL_MemInfo1, sL_MemInfo2, sL_MemInfo3, sL_MemInfo4 : String;
begin
    timAutoInvoke.Enabled := false;

    for ii:=0 to G_ExePathInfoStrList.Count -1 do
    begin
      sL_FullpathCaExeName := G_ExePathInfoStrList.Strings[ii];
      sL_CaExeName := ExtractFileName(sL_FullpathCaExeName);

      sL_MemInfo1 := '偵測到 ' + sL_FullpathCaExeName + ' 尚未執行';
      sL_MemInfo2 := '開始呼叫 ' + sL_FullpathCaExeName;
      sL_MemInfo3 := '呼叫 ' + sL_FullpathCaExeName + ' 成功';
      sL_MemInfo4 := '偵測到 ' + sL_FullpathCaExeName + ' 停滯過久';





      nL_CaAppId := getProcessId(sL_CaExeName);


      if nL_CaAppId=NULL_PROC_ID then
      begin
        writeMonitorLog(sL_MemInfo1);
        writeMonitorLog(sL_MemInfo2);
        runCaCmdGateway(sL_FullpathCaExeName);
        writeMonitorLog(sL_MemInfo3);
        writeMonitorLog('======================================================================');
      end
      else
      begin
        //down, 若CA 已經在執行中 => 檢查CA是否已經很久都沒有傳送指令 => 若是,則重新啟動CA
        bL_IsCaActive := IsCaActive;
        if not bL_IsCaActive then
        begin
          writeMonitorLog(sL_MemInfo4);
          killCaGateway(sL_FullpathCaExeName);
          writeMonitorLog(sL_MemInfo2);
          runCaCmdGateway(sL_FullpathCaExeName);
          writeMonitorLog(sL_MemInfo3);
          writeMonitorLog('======================================================================');
        end;
        //up, 若CA 已經在執行中 => 檢查CA是否已經很久都沒有傳送指令 => 若是,則重新啟動CA
      end;


    end;
    //down, 讓程式畫面產生 refresh 的效果, 若不做,CA有時會無法開始傳送指令

    self.WindowState := wsMinimized;
    self.WindowState := wsNormal;
    self.Repaint;
    //up, 讓程式畫面產生 refresh 的效果, 若不做,CA有時會無法開始傳送指令
    timAutoInvoke.Enabled := true;
end;

}
procedure TfrmMain.setProcessMonitor(bI_Status: boolean);
begin
    timAutoInvoke.Enabled := false;
    timAutoInvoke.Interval := nG_InvokeCaTimer;
    timAutoInvoke.Enabled := bI_Status;

end;

procedure TfrmMain.CheckBox2Click(Sender: TObject);
begin
    setProcessMonitor(CheckBox2.Checked);
end;

procedure TfrmMain.FormClose(Sender: TObject; var Action: TCloseAction);
begin

    
    G_MonitorLogStrList.Free;
    G_TimeStampStrList.Free;
end;

procedure TfrmMain.writeMonitorLog(sI_Info: String);
var
    sL_RandomFileName : String;
    sL_ExeFileName, sL_ExePath : String;
    sL_RealInfo, sL_FullPathLogFileName : String;
begin
    sL_ExeFileName := Application.ExeName;
    sL_ExePath := ExtractFileDir(sL_ExeFileName);
    sL_FullPathLogFileName := sL_ExePath + '\' + MONITOR_LOG_FILE_PATH + '\' + MONITOR_LOG_FILE_NAME;

    if Copy(sI_Info,1,5)='=====' then
      sL_RealInfo := sI_Info
    else
      sL_RealInfo := datetimetostr(now) + '=>' + sI_Info;

    G_MonitorLogStrList.Add(sL_RealInfo);


    G_MonitorLogStrList.SaveToFile(sL_FullPathLogFileName);
    G_MonitorLogStrList.LoadFromFile(sL_FullPathLogFileName);
    if G_MonitorLogStrList.Count>nG_MemoMaxLineCount then
    begin
      sL_RandomFileName := getRendomFileName;

      RenameFile(sL_FullPathLogFileName, sL_RandomFileName);

      G_MonitorLogStrList.Clear;
    end;
end;

function TfrmMain.getRendomFileName: String;
var
    sL_ExeFileName, sL_ExePath : String;
    sL_TmpXmlFileName : String;
    nL_Seed : Integer;
begin
    sL_ExeFileName := Application.ExeName;
    sL_ExePath := ExtractFileDir(sL_ExeFileName);

    Randomize;
    nL_Seed := 100000;
    sL_TmpXmlFileName := sL_ExePath + '\' + TUstr.GetPureStr(DateToStr(Date),'/') + '-' + IntToStr(Random(nL_Seed)) + '.log';
//    sL_TmpXmlFileName := sL_ExePath + '\' + IntToStr(Random(nL_Seed)) + '.log';

    while FileExists(sL_TmpXmlFileName) do
      getRendomFileName;

    result := sL_TmpXmlFileName;

end;

procedure TfrmMain.InvokeMonitor(sI_Method, sI_User: OleVariant;
  var sI_ActionStatusCode, sI_ActionStatusMsg: OleVariant);
var
    sL_CaExeName, sL_FullpathCaExeName : String;
    nL_CaAppId, ii : Integer;
    sL_MemInfo1, sL_MemInfo2, sL_MemInfo3 : String;
    sL_MemInfo4, sL_MemInfo5, sL_MemInfo6 : String;
    sL_MemInfo7, sL_MemInfo8 : String;
begin
    for ii:=0 to G_ExePathInfoStrList.Count -1 do
    begin
      sL_FullpathCaExeName := G_ExePathInfoStrList.Strings[ii];
      sL_CaExeName := ExtractFileName(sL_FullpathCaExeName);
      sL_MemInfo1 := sL_CaExeName + ' 尚未執行.';
      sL_MemInfo2 := sL_CaExeName + ' 正常執行中.' + ' [ ' + edtIniFileName.Text  +' ]';
      sL_MemInfo3 := sL_CaExeName + ' 尚未執行.所以無法停止.';
      sL_MemInfo4 := sL_CaExeName + ' 即將於一分鐘內完全停止執行.';
      sL_MemInfo5 := sL_CaExeName + ' 即將於一分鐘內開始執行.';
      sL_MemInfo6 := sL_CaExeName + ' 沒有辦法停止執行.';
      sL_MemInfo7 := sL_CaExeName + ' 無法執行.';
      sL_MemInfo8 := sL_CaExeName + ' 執行中.所以無法重複執行.';



      case StrToInt(sI_Method) of
        1:// Ping CA Command Gateway
         begin
          nL_CaAppId := getProcessId(sL_CaExeName);


          if nL_CaAppId=NULL_PROC_ID then   // 表示 CA Command Gateway 沒有執行
          begin
            sI_ActionStatusCode := '-1';
            sI_ActionStatusMsg := sL_MemInfo1;
          end
          else
          begin
            sI_ActionStatusCode := INVOKE_MONITOR_OK_CODE;
            sI_ActionStatusMsg := sL_MemInfo2;
          end
         end;
        2:// Terminate CA Command Gateway
         begin
          nL_CaAppId := getProcessId(sL_CaExeName);


          if nL_CaAppId=NULL_PROC_ID then   // 表示 CA Command Gateway 沒有執行
          begin
            sI_ActionStatusCode := '-1';
            sI_ActionStatusMsg := sL_MemInfo3;
            Exit;
          end
          else
          begin
            try
              killCaGateway(sL_FullpathCaExeName);
              sI_ActionStatusCode := INVOKE_MONITOR_OK_CODE;
              sI_ActionStatusMsg := sL_MemInfo4;
            except
              sI_ActionStatusCode := '-1';
              sI_ActionStatusMsg := sL_MemInfo6;
            end;
          end;
         end;
        3:// Run CA Command Gateway
         begin
          nL_CaAppId := getProcessId(sL_CaExeName);


          if nL_CaAppId=NULL_PROC_ID then   // 表示 CA Command Gateway 沒有執行
          begin
            try
              runCaCmdGateway(sL_FullpathCaExeName);
              sI_ActionStatusCode := INVOKE_MONITOR_OK_CODE;
              sI_ActionStatusMsg := sL_MemInfo5;
            except
              sI_ActionStatusCode := '-1';
              sI_ActionStatusMsg := sL_MemInfo7;
            end;
          end
          else
          begin
            sI_ActionStatusCode := '-1';
            sI_ActionStatusMsg := sL_MemInfo8;
          end;
         end;
      end;
    end;
end;

procedure TfrmMain.ServerSocketAccept(Sender: TObject;
  Socket: TCustomWinSocket);
var
    sL_RemoteHost, sL_RemoteAddress : String;
begin
    sL_RemoteHost := Socket.RemoteHost;
    sL_RemoteAddress := Socket.RemoteAddress;
    memInfo.Lines.Insert(0, 'Host : '+sL_RemoteHost + '(' + sL_RemoteAddress + ') 於 ' + DateTimeToStr(now)  + ' - 開始連線');
end;

procedure TfrmMain.ServerSocketClientRead(Sender: TObject;
  Socket: TCustomWinSocket);
var
    nL_Method : Integer;
    sL_Msg, sL_Method, sL_User, sL_Info,sL_ValidMethod : String;
    L_StrList : TStringList;
    sL_ActionStatusCode, sL_ActionStatusMsg: OleVariant;
    sL_RemoteHost, sL_RemoteAddress, sL_TimeStamp : String;


begin
    sL_Info := Socket.ReceiveText;

    sL_ValidMethod := '1,2,3,4';

    L_StrList := TUStr.ParseStrings(sL_Info,',');


    if (L_StrList=nil) or(L_StrList.Count <> 2) or (AnsiPos(L_StrList.Strings[0],sL_ValidMethod)=0 ) then
    begin
      Socket.SendText('-99,傳入之參數錯誤!');
    end
    else
    begin
      sL_Method := L_StrList.Strings[0];
      nL_Method := StrToInt(sL_Method);
      if (nL_Method<>1) and (not chbAcceptClient.Checked) then
      begin
        sL_ActionStatusCode := '-99';
        sL_ActionStatusMsg := 'CA Command Gateway 不接受前端控制';
        Socket.SendText(sL_ActionStatusCode + ',' + sL_ActionStatusMsg);
        Exit;
      end;


      sL_User := L_StrList.Strings[1];
      InvokeMonitor(sL_Method, sL_User, sL_ActionStatusCode, sL_ActionStatusMsg);
      Socket.SendText(sL_ActionStatusCode + ',' + sL_ActionStatusMsg);

      if sL_ActionStatusCode=INVOKE_MONITOR_OK_CODE then
      begin
        sL_RemoteHost := Socket.RemoteHost;
        sL_RemoteAddress := Socket.RemoteAddress;

        case nL_Method of
          2://停止CA
           begin
            sL_Msg := 'Host : '+ sL_RemoteHost + '(' + sL_RemoteAddress + ')  (' + sL_User + ') 於 ' + DateTimeToStr(now)  + ' - 停止CA';
            memInfo.Lines.Insert(0, sL_Msg);
           end;
          3://啟動CA
           begin
            sL_Msg := 'Host : '+ sL_RemoteHost + '(' + sL_RemoteAddress + ')  (' + sL_User + ') 於 ' + DateTimeToStr(now)  + ' - 啟動CA';
            memInfo.Lines.Insert(0, sL_Msg);
           end;
        end;
      end;

    end;

end;

procedure TfrmMain.ServerSocketClientDisconnect(Sender: TObject;
  Socket: TCustomWinSocket);
var
    sL_RemoteHost, sL_RemoteAddress : String;
begin
    sL_RemoteHost := Socket.RemoteHost;
    sL_RemoteAddress := Socket.RemoteAddress;
    memInfo.Lines.Insert(0, 'Host : '+sL_RemoteHost + '(' + sL_RemoteAddress + ') 於 ' + DateTimeToStr(now)  + ' - 結束連線');

end;

procedure TfrmMain.BitBtn4Click(Sender: TObject);
var
    ii : Integer;
begin
   if MessageDlg('確定要清除所有監控資料?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
   begin
     for ii:=0 to stb.Panels.Count -1 do
       stb.Panels[ii].Text := '';
     memInfo.Clear;
   end;
end;

procedure TfrmMain.SpeedButton1Click(Sender: TObject);
begin
    if OpenDialog1.Execute then
      edtIniFileName.Text := ExtractFileName(OpenDialog1.FileName);
end;

procedure TfrmMain.setDefaultSetting;
begin
    timAutoInvoke.Enabled := false;
end;



function TfrmMain.isCaActive: boolean;
var
    nL_FetchDataInterval, nL_Diff : Integer;
    bL_Result : boolean;
    sL_CaLastActiveTime : String;
    dL_Diff, dL_CaLastActiveTime : TDate;
begin
    bL_Result := true;
    sL_CaLastActiveTime := getCaTimeStamp;
    if sL_CaLastActiveTime <> '' then
    begin
      try
        dL_CaLastActiveTime := StrToDateTime(sL_CaLastActiveTime);
        dL_Diff := now - dL_CaLastActiveTime;
        nL_Diff := Round(dL_Diff*24*60*60);

        nL_FetchDataInterval := reDefineFetchDataInterval(now);

        if nL_Diff > (nL_FetchDataInterval)*10 then //若上次CA "活動"的時間距離現在已經超過指令每次讀取時間的十倍時, 就判定CA死掉了.
        begin
          bL_Result := false;
        end;
      except
        bL_Result := true;
      end;

    end;

    result := bL_Result;
end;

function TfrmMain.getCaTimeStamp: String;
var
    sL_ExeFileName, sL_ExePath, sL_TimeStampFileName : String;
    sL_CaTimeStamp, sL_TmpFileName : String;
begin
    sL_CaTimeStamp := '';
    sL_ExeFileName := Application.ExeName;
    sL_ExePath := ExtractFileDir(sL_ExeFileName);
    sL_TimeStampFileName := sL_ExePath + '\' + TIME_STAMP_FILE_NAME;
    sL_TmpFileName :=  sL_ExePath + '\' + TMP_TIME_STAMP_FILE_NAME;
    
    if FileExists(sL_TimeStampFileName) then
    begin
      if FileExists(sL_TmpFileName) then
        DeleteFile(sL_TmpFileName);

      RenameFile(sL_TimeStampFileName, sL_TmpFileName);
      if FileExists(sL_TmpFileName) then
      begin
        G_TimeStampStrList.LoadFromFile(sL_TmpFileName);
        sL_CaTimeStamp := G_TimeStampStrList.Strings[0];
      end;
    end;
    result := sL_CaTimeStamp;
end;

procedure TfrmMain.GetSysParam;
var
    sL_Path : String;
begin

    //down,給預設值
    nG_CmdRefreshRate1 := 3;
    nG_CmdRefreshRate2 := 30;
    sG_TimeDefine := '0700-2300';
    //up,給預設值

    with cdsParam do
    begin
      if (State = dsInactive) then
        CreateDataSet;
      sL_Path := ExtractFilePath(Application.ExeName);
      if (FileExists(sL_Path + SYS_PARAM_FILENAME)) then
        LoadFromFile(sL_Path + SYS_PARAM_FILENAME);

      nG_CmdRefreshRate1 := FieldByName('nCmdRefreshRate1').AsInteger; //尖峰時段?秒讀取一次指令資料庫
      nG_CmdRefreshRate2 := FieldByName('nCmdRefreshRate2').AsInteger; //離峰時段?秒讀取一次指令資料庫
      sG_TimeDefine := FieldByName('sSHotTime').AsString + '-' + FieldByName('sEHotTime').AsString; // 尖峰時段定義
      Close;        
    end;        
end;

function TfrmMain.reDefineFetchDataInterval(dI_Time: TDate): Integer;
var
    nL_Result, nL_Time1, nL_Time2, nL_Time3 : Integer;
    sL_Time1, sL_Time2 : String;
    wL_Hour, wL_Min, wL_Sec, wL_MSec : word;
begin
    sL_Time1 := Copy(sG_TimeDefine,1,4);
    sL_Time2 := Copy(sG_TimeDefine,6,4);
    DecodeTime(dI_Time,wL_Hour,wL_Min, wL_Sec, wL_MSec);
    nL_Time1 := StrToInt(sL_Time1);
    nL_Time2 := StrToInt(sL_Time2);
    nL_Time3 := wL_Hour*100+wL_Min;

    if (nL_Time3>nL_Time1) and (nL_Time3<nL_Time2) then
      nL_Result := nG_CmdRefreshRate1 //傳入之時間為尖峰時段
    else
      nL_Result := nG_CmdRefreshRate2; //傳入之時間為離峰時段

    result := nL_Result;
end;

procedure TfrmMain.RunCA(sI_IniFileName:String);
var
    ii : Integer;

    sL_FullpathExeName, sL_RealIniFileName : String;
begin

{
    nG_OperFlag := StrToIntDef(Copy(sI_IniFileName,1,1),0);
    sL_RealIniFileName := Copy(sI_IniFileName,2,length(sI_IniFileName)-1);
}
    sL_RealIniFileName := sI_IniFileName;
    edtIniFileName.Text := sL_RealIniFileName;
    CheckBox2.Checked := true;
    chbAcceptClient.Checked := true;
    sleep(3000);
    BitBtn3Click(nil);
    
    for ii:=0 to G_ExePathInfoStrList.Count -1 do
    begin
      sL_FullpathExeName := G_ExePathInfoStrList.Strings[ii];
      killCaGateway(sL_FullpathExeName);
      //mark with nick, 2005/05/23 runCaCmdGateway(sL_FullpathExeName);

    end;
end;

end.
