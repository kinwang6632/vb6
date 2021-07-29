unit frmMainU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, Buttons, DB, ADODB, ScktComp;

const
    PARAM_COUNT= 4;
    IP_FILE_NAME = 'ip.dat';
    PORT=8123;
type
  TfrmMain = class(TForm)
    Panel1: TPanel;
    StaticText1: TStaticText;
    edtIP: TEdit;
    btnSo8D10: TBitBtn;
    btnSo8D20A: TBitBtn;
    btnSo8D20B: TBitBtn;
    btnExit: TBitBtn;
    ClientSocket: TClientSocket;
    btnConn: TBitBtn;
    procedure btnExitClick(Sender: TObject);
    procedure btnSo8D10Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btnSo8D20AClick(Sender: TObject);
    procedure btnSo8D20BClick(Sender: TObject);
    procedure btnConnClick(Sender: TObject);
    procedure ClientSocketError(Sender: TObject; Socket: TCustomWinSocket;
      ErrorEvent: TErrorEvent; var ErrorCode: Integer);
    procedure ClientSocketRead(Sender: TObject; Socket: TCustomWinSocket);
    procedure ClientSocketConnect(Sender: TObject;
      Socket: TCustomWinSocket);
  private
    { Private declarations }
    sG_CompName, sG_Priv1, sG_Priv2, sG_User : String;
    procedure invokeServer(sI_Method, sI_User:String);
    function getCaption(sI_FormID, sI_FormName : String):String;
    procedure saveIP;

  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;

implementation

uses prjMonitor_TLB, Ustru;

{$R *.dfm}

procedure TfrmMain.btnExitClick(Sender: TObject);
begin
//    ClientSocket.Close;
    saveIP;
    Application.Terminate;
end;

procedure TfrmMain.btnSo8D10Click(Sender: TObject);
var
    Intf : IMonitor;
    sL_IP, sL_Msg : String;
    sL_ActionStatusCode: OleVariant;
    sL_ActionStatusMsg: OleVariant;
begin

    invokeServer('1', sG_User);

    {
    if edtIP.Text ='' then
    begin
      MessageDlg('請輸入IP.', mtWarning, [mbOK],0);
      edtIP.SetFocus;
      Exit;
    end
    else
      sL_IP := Trim(edtIP.Text);

    Intf := CoMonitor.CreateRemote(sL_IP);
    Intf.InvokeMonitor('1',sG_User,sL_ActionStatusCode, sL_ActionStatusMsg);
    if sL_ActionStatusCode='-99' then
    begin
      Intf := nil;
      MessageDlg(sL_ActionStatusMsg, mtError, [mbOK],0);
      Exit;
    end;

    if sL_ActionStatusCode='0' then
      sL_Msg := 'CA 正常執行中'
    else
      sL_Msg := sL_ActionStatusMsg;
    Intf := nil;
    MessageDlg(sL_Msg, mtInformation, [mbOK],0);
    }
end;

procedure TfrmMain.FormShow(Sender: TObject);
var
    L_StrList : TStringList;
    sL_ExePath, sL_TmpFileName : String;
begin

    sL_ExePath := ExtractFilePath(Application.ExeName);
    sL_TmpFileName := sL_ExePath + IP_FILE_NAME;

    edtIP.SetFocus;
    if FileExists(sL_TmpFileName) then
    begin
      L_StrList := TStringList.Create;
      L_StrList.LoadFromFile(sL_TmpFileName);
      edtIP.Text := L_StrList.Strings[0];
      L_StrList.Free;
    end;

    //down, 設定權限...
    if sG_Priv1='Y' then
      btnSo8D10.Enabled :=  true
    else
      btnSo8D10.Enabled :=  false;

    if sG_Priv2='Y' then
    begin
      btnSo8D20A.Enabled :=  true;
      btnSo8D20B.Enabled :=  true;
    end
    else
    begin
      btnSo8D20A.Enabled :=  false;
      btnSo8D20B.Enabled := false;
    end;
    //up, 設定權限...

    self.Caption := getCaption('SO8D00','CA Client 模組');

end;

procedure TfrmMain.FormCreate(Sender: TObject);
begin
    if ParamCount <> PARAM_COUNT then
    begin
      MessageDlg('傳入之參數個數錯誤!', mtError, [mbOK],0);
      Application.Terminate;
    end
    else
    begin
      //陽明山 Howard y y => 有權限與否,由參數傳進來
      sG_CompName := Uppercase(ParamStr(1)); //公司名稱
      sG_User := Uppercase(ParamStr(2)); //傳入User Name
      sG_Priv1 := Uppercase(ParamStr(3)); // Y=> 表示user 有偵測CA的權限
      sG_Priv2 := Uppercase(ParamStr(4)); // Y=> 表示user 有停止與啟動CA的權限
    end;

end;

procedure TfrmMain.btnSo8D20AClick(Sender: TObject);
var
    Intf : IMonitor;
    sL_IP, sL_Msg : String;
    sL_ActionStatusCode: OleVariant;
    sL_ActionStatusMsg: OleVariant;

begin
    if MessageDlg('是否確定要停止CA之運作?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      invokeServer('2', sG_User);
    end;  
    {
    if edtIP.Text ='' then
    begin
      MessageDlg('請輸入IP.', mtWarning, [mbOK],0);
      edtIP.SetFocus;
      Exit;
    end
    else
      sL_IP := Trim(edtIP.Text);
    if MessageDlg('是否確定要停止CA之運作?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      Intf := CoMonitor.CreateRemote(sL_IP);
      Intf.InvokeMonitor('2',sG_User,sL_ActionStatusCode, sL_ActionStatusMsg);
      if sL_ActionStatusCode='-99' then
      begin
        Intf := nil;
        MessageDlg(sL_ActionStatusMsg, mtError, [mbOK],0);
        Exit;
      end;

      if sL_ActionStatusCode='0' then
        sL_Msg := 'CA 停止執行 ok.'
      else
        sL_Msg := sL_ActionStatusMsg;
      Intf := nil;
      MessageDlg(sL_Msg, mtInformation, [mbOK],0);
    end;
    }
end;

procedure TfrmMain.saveIP;
var
    L_StrList : TStringList;
    sL_ExePath, sL_TmpFileName : String;
begin
    if edtIP.Text <>'' then
    begin
      sL_ExePath := ExtractFilePath(Application.ExeName);
      sL_TmpFileName := sL_ExePath + IP_FILE_NAME;
      DeleteFile(sL_TmpFileName);
      L_StrList := TStringList.Create;
      L_StrList.Add(trim(edtIP.Text));
      L_StrList.SaveToFile(sL_TmpFileName);
      L_StrList.Free;
    end;
end;

procedure TfrmMain.btnSo8D20BClick(Sender: TObject);
var
    Intf : IMonitor;
    sL_IP, sL_Msg : String;
    sL_ActionStatusCode: OleVariant;
    sL_ActionStatusMsg: OleVariant;

begin
    if MessageDlg('是否確定要啟動CA?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      invokeServer('3', sG_User);
    end;
    {
    if edtIP.Text ='' then
    begin
      MessageDlg('請輸入IP.', mtWarning, [mbOK],0);
      edtIP.SetFocus;
      Exit;
    end
    else
      sL_IP := Trim(edtIP.Text);
    if MessageDlg('是否確定要啟動CA?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      Intf := CoMonitor.CreateRemote(sL_IP);
      Intf.InvokeMonitor('3',sG_User,sL_ActionStatusCode, sL_ActionStatusMsg);
      if sL_ActionStatusCode='-99' then
      begin
        Intf := nil;
        MessageDlg(sL_ActionStatusMsg, mtError, [mbOK],0);
        Exit;
      end;

      if sL_ActionStatusCode='0' then
        sL_Msg := 'CA 將在一分鐘內執行.一分鐘後請偵測CA是否正常運作'
      else
        sL_Msg := sL_ActionStatusMsg;
      Intf := nil;
      MessageDlg(sL_Msg, mtInformation, [mbOK],0);
    end;
    }
end;

function TfrmMain.getCaption(sI_FormID, sI_FormName: String): String;
begin
    result := sG_CompName + '-' + sG_User + '--' + sI_FormName + '-[' +  sI_FormID + ']';
end;

procedure TfrmMain.invokeServer(sI_Method, sI_User: String);
var
    sL_Info : String;
begin
    if not ClientSocket.Active then
    begin
      MessageDlg('請先執行連接!', mtWarning, [mbOK],0);
      btnConn.SetFocus;
      Exit;
    end;

    sL_Info :=  sI_Method + ',' +  sI_User;
    ClientSocket.Socket.SendText(sL_Info);


end;

procedure TfrmMain.btnConnClick(Sender: TObject);
var
    sL_IP : String;
begin
    if edtIP.Text ='' then
    begin
      MessageDlg('請輸入IP.', mtWarning, [mbOK],0);
      edtIP.SetFocus;
      Exit;
    end
    else
      sL_IP := Trim(edtIP.Text);

    if not ClientSocket.Active then
    begin
      ClientSocket.Host := sL_IP;
      ClientSocket.Port := PORT;
      try
        ClientSocket.Active := true;

      except

      end;
      
    end;


end;

procedure TfrmMain.ClientSocketError(Sender: TObject;
  Socket: TCustomWinSocket; ErrorEvent: TErrorEvent;
  var ErrorCode: Integer);
begin
    MessageDlg('連接失敗', mtError, [mbOK],0);
end;

procedure TfrmMain.ClientSocketRead(Sender: TObject;
  Socket: TCustomWinSocket);
var
    nL_ActionCode : Integer;
    sL_Info, sL_ActionMsg : String;
    L_StrList : TStringList;
begin
    sL_Info := Socket.ReceiveText;
//    showmessage(sL_Info);
    L_StrList := TUStr.ParseStrings(sL_Info,',');
    if (L_StrList<>nil) and (L_StrList.Count = 2) then
    begin

      nL_ActionCode := StrToInt(L_StrList.Strings[0]);
      sL_ActionMsg := L_StrList.Strings[1];
      if nL_ActionCode<0 then
        MessageDlg(sL_ActionMsg, mtError, [mbOK],0)
      else
        MessageDlg(sL_ActionMsg, mtInformation, [mbOK],0);        
    end;


end;

procedure TfrmMain.ClientSocketConnect(Sender: TObject;
  Socket: TCustomWinSocket);
begin
      MessageDlg('連接完成', mtInformation, [mbOK],0);
end;

end.
