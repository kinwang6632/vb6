unit frmMainMenuU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Menus, DB, ADODB, ExtCtrls ,IniFiles;

const
    //Delphi7 的 PramCount = 12
    PARAM_COUNT= 12;   //傳入參數個數, 正式的值

type
  TSnapSHotIniData = class(TObject)
    nDataAreaNo : Integer;
    sDataArea : String;
    sAlias : String;
    sUserID : String;
    sPassword : String;
    sCompCode : String;
    sCompName : String;
  end;

  TfrmMainMenu = class(TForm)
    MainMenu1: TMainMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    btnSO8B30: TMenuItem;
    btnSO8B10: TMenuItem;
    btnSO8B50: TMenuItem;
    btnSO8B40: TMenuItem;
    btnSO8B20: TMenuItem;
    N10: TMenuItem;
    ADOConnection1: TADOConnection;
    qryPriv: TADOQuery;
    qryCommon: TADOQuery;
    btnSO8A10: TMenuItem;
    btnSO8A20: TMenuItem;
    btnSO8A30: TMenuItem;
    btnSO8A40: TMenuItem;
    btnSO8A50: TMenuItem;
    btnSO8A60: TMenuItem;
    btnSO8F10: TMenuItem;
    btnSO8F20: TMenuItem;
    Panel1: TPanel;
    btnSo8C10: TMenuItem;
    btnSo8C20: TMenuItem;
    btnSo8C30: TMenuItem;
    btnSo8C40: TMenuItem;
    Panel2: TPanel;
    Panel3: TPanel;
    Panel4: TPanel;
    Panel5: TPanel;
    procedure btnSO8B30Click(Sender: TObject);
    procedure btnSO8B10Click(Sender: TObject);
    procedure btnSO8B20Click(Sender: TObject);
    procedure btnSO8B40Click(Sender: TObject);
    procedure btnSO8B50Click(Sender: TObject);
    procedure N10Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btnSO8A10Click(Sender: TObject);
    procedure btnSO8A20Click(Sender: TObject);
    procedure btnSO8A30Click(Sender: TObject);
    procedure btnSO8A40Click(Sender: TObject);
    procedure btnSO8A50Click(Sender: TObject);
    procedure btnSO8A60Click(Sender: TObject);
    procedure btnSO8F10Click(Sender: TObject);
    procedure btnSO8F20Click(Sender: TObject);
    procedure btnSo8C10Click(Sender: TObject);
    procedure btnSo8C20Click(Sender: TObject);
    procedure btnSo8C30Click(Sender: TObject);
    procedure btnSo8C40Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    sG_ServiceType,sG_IsSupervisor,sG_UserID,sG_User,sG_CompCode,sG_CompName : String;
    sG_DbUserName,sG_DbPassword,sG_DbAlias,sG_Server_IP,sG_NeedLogin : String;


    sG_LoginID,sG_LoginName : String;
    function getDbConnInfo:WideString;
    procedure setDateTimeFormat;
    procedure createDtmMain1;
    procedure createDtmMain2;
    procedure createDtmMain3;
    procedure createDtmMain4;
    function connectToDB(sI_Comcpde : String):WideString;
    function connectToDB2(sI_Comcpde : String):WideString;
    procedure activePrivDataset(sI_CompCode,sI_UserID:String);
    function checkPriv(sI_Mid:String):boolean;
    function setCaption(sI_FunctionID, sI_Desc : String) : String;
    function TransTmpIniFile(sI_IniFileName : String) : String;
  end;

var
  frmMainMenu: TfrmMainMenu;

implementation

uses frmSO8B30U, frmSO8B10U, frmSO8B20U, frmSO8B40U, frmSO8B50U, dtmMain1U,
  frmSO8A10U, frmSO8A20U, frmSO8A30U, frmSO8A40U, frmSO8A50U, frmSO8A60U,
  dtmMain2U, Emc_Separate_TLB, frmSO8F10_1U, frmSO8F20U, dtmMain4U,
  frmLoginU, frmSO8C10U, frmSO8C20U, frmSO8C30U, frmSO8C40_1U, dtmMain3U,
  UCommonU, DevPower_TLB;

{$R *.dfm}

procedure TfrmMainMenu.btnSO8B30Click(Sender: TObject);
var
    sL_Result : String;
    L_Frm : TfrmSO8B30;
begin
    createDtmMain1;
{
    L_Frm := TfrmSO8B30.Create(Application);
    L_Frm.ShowModal;
    L_Frm.Free;
}

    If not Assigned(frmSO8B30) then
      Application.CreateForm(TfrmSO8B30,frmSO8B30);
    frmSO8B30.ShowModal;

end;

procedure TfrmMainMenu.btnSO8B10Click(Sender: TObject);
begin
    createDtmMain1;

	frmSO8B10 := TfrmSO8B10.Create(Application);
 	frmSO8B10.ShowModal;
end;

procedure TfrmMainMenu.btnSO8B20Click(Sender: TObject);
begin
    createDtmMain1;
    
    Application.CreateForm(TfrmSO8B20,frmSO8B20);
    frmSO8B20.ShowModal;
end;

procedure TfrmMainMenu.btnSO8B40Click(Sender: TObject);
begin
    createDtmMain1;

    Application.CreateForm(TfrmSO8B40,frmSO8B40);
    frmSO8B40.ShowModal;
end;

procedure TfrmMainMenu.btnSO8B50Click(Sender: TObject);
Var
 L_Frm : TfrmSO8B50;
begin
    createDtmMain1;

    L_Frm := TfrmSO8B50.Create(Application);
    L_Frm.ShowModal;
    L_Frm.Free;
end;

procedure TfrmMainMenu.N10Click(Sender: TObject);
begin
    if MessageDlg('是否確定要結束程式?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
      Close;
end;

procedure TfrmMainMenu.FormCreate(Sender: TObject);
var
    sL_Result : String;
    bL_LoginOk,bL_IsEnable : Boolean;
begin
    //down, 從 ini 檔讀取連線至 DB 的 information
    sL_Result := getDbConnInfo;
    if sL_Result<>'' then
    begin
      showmessage(sL_Result);
      Application.Terminate;
    end;
    //up, 從 ini 檔讀取連線至 DB 的 information


    sL_Result := connectToDB(frmMainMenu.sG_CompCode);
    if sL_Result <> '' then
    begin
        showmessage(sL_Result);
        exit;
    end;

    //設定系統日期格式
    setDateTimeFormat;

    //down, 設定權限...
    if frmMainMenu.sG_IsSupervisor='Y' then
    begin
      //佣獎金
      btnSO8B10.Enabled := true;
      btnSO8B20.Enabled := true;
      btnSO8B30.Enabled := true;
      btnSO8B40.Enabled := true;
      btnSO8B50.Enabled := true;

      //拆帳
      btnSO8A10.Enabled := true;
      btnSO8A20.Enabled := true;
      btnSO8A30.Enabled := true;
      btnSO8A40.Enabled := true;
      btnSO8A50.Enabled := true;
      btnSO8A60.Enabled := true;

      //手開單
      btnSo8C10.Enabled :=  true;
      btnSo8C20.Enabled :=  true;
      btnSo8C30.Enabled :=  true;
      btnSo8C40.Enabled :=  true;
    end
    else
    begin
      activePrivDataset(frmMainMenu.sG_CompCode,frmMainMenu.sG_UserID);
      //佣獎金
      btnSO8B10.Enabled :=  checkPriv('SO8B10');
      btnSO8B20.Enabled :=  checkPriv('SO8B20');
      btnSO8B30.Enabled :=  checkPriv('SO8B30');
      btnSO8B40.Enabled :=  checkPriv('SO8B40');
      btnSO8B50.Enabled := checkPriv('SO8B50');

      //拆帳
      btnSO8A10.Enabled :=  checkPriv('SO8A10');
      btnSO8A20.Enabled :=  checkPriv('SO8A20');
      btnSO8A30.Enabled :=  checkPriv('SO8A30');
      btnSO8A40.Enabled :=  checkPriv('SO8A40');
      btnSO8A50.Enabled :=  checkPriv('SO8A50');
      btnSO8A60.Enabled :=  checkPriv('SO8A60');

      //手開單
      btnSo8C10.Enabled :=  checkPriv('SO8C10');
      btnSo8C20.Enabled :=  checkPriv('SO8C20');
      btnSo8C30.Enabled :=  checkPriv('SO8C30');
      btnSo8C40.Enabled :=  checkPriv('SO8C40');
    end;


    //二階代碼
    if (frmMainMenu.sG_IsSupervisor='Y') or (frmMainMenu.sG_NeedLogin='Y') then
    begin
      btnSO8F20.Enabled := true;
    end
    else
    begin
      frmMainMenu.activePrivDataset(frmMainMenu.sG_CompCode,frmMainMenu.sG_UserID);
      bL_IsEnable := frmMainMenu.checkPriv('SO8F20');
      btnSO8F20.Enabled := bL_IsEnable;
    end;
    //up, 設定權限...



    if sG_NeedLogin='N' then   //各系統台使用
    begin


    end
    else    //營管使用
    begin
      frmLogin := TfrmLogin.Create(Application);
      with frmLogin do
      begin

        if ShowModal=mrCancel then
        begin
          bL_LoginOk := false;
        end
        else
        begin
          bL_LoginOk := true;
          sG_LoginID := sG_USerID;
          sG_LoginName := sG_UserName;
          Close;
          free;
        end;
      end;

      if not bL_LoginOk then
        Application.Terminate;

    end;
end;

function TfrmMainMenu.getDbConnInfo: WideString;
begin
    //若為各 SO 公司則抓取 CC&B (VB) 傳來之參數
    if ParamCount = PARAM_COUNT then
    begin
      //Y MAINTAIN Admin 16 北桃園 EMCNTY EMCNTY MIS CIP 127.0.0.1 N 0
      //是否為SuperUser LoginUserID LoginUserName CompCode CompName DBUser DbPassword DbAlias Server_IP
      sG_IsSupervisor := Uppercase(ParamStr(1));
      sG_UserID := ParamStr(2);
      sG_User := ParamStr(3);
      sG_CompCode := ParamStr(4);               
      sG_CompName := ParamStr(5);
      sG_DbUserName := ParamStr(6);
      sG_DbPassword := ParamStr(7);
      sG_DbAlias := ParamStr(8);
      sG_ServiceType := ParamStr(9);
      sG_Server_IP := ParamStr(10);
      sG_NeedLogin := Uppercase(ParamStr(11));

{
SHOWMESSAGE('ParamCount   ' + IntToStr(ParamCount));
SHOWMESSAGE('sG_IsSupervisor(1)   ' + sG_IsSupervisor);
SHOWMESSAGE('sG_UserID(2)   ' + sG_UserID);
SHOWMESSAGE('sG_User(3)   ' + sG_User);
SHOWMESSAGE('sG_CompCode(4)   ' + sG_CompCode);
SHOWMESSAGE('sG_CompName(5)   ' + sG_CompName);
SHOWMESSAGE('sG_DbUserName(6)   ' + sG_DbUserName);
SHOWMESSAGE('sG_DbPassword(7)   ' + sG_DbPassword);
SHOWMESSAGE('sG_DbAlias(8)   ' + sG_DbAlias);
}
    end
    else
    begin
      result := '傳入之參數個數錯誤!';
      EXIT;
    end;
end;

procedure TfrmMainMenu.setDateTimeFormat;
begin
    DateSeparator := '/';
    TimeSeparator := ':';

    LongDateFormat := 'yyyy/mm/dd';
    ShortDateFormat := 'yyyy/mm/dd';
    ShortTimeFormat :='hh:nn:ss';
    LongTimeFormat :='hh:nn:ss';
    TimeAMString := '';
    TimePMString := '';
end;


procedure TfrmMainMenu.createDtmMain1;
begin
    //執行等待狀態
    TUCommonFun.setWaitingCursor;

    if not Assigned(dtmMain1) then
      dtmMain1 := TdtmMain1.Create(nil);

    dtmMain1.dataSetActive;

    //回復原狀態
    TUCommonFun.setDefaultCursor;
end;

function TfrmMainMenu.connectToDB(sI_Comcpde: String): WideString;
var
    ii,nL_DataAreaNo : Integer;
    sL_DbConnStr : String;
    sL_DbPassword, sL_DbAlias, sL_DbUserName,sL_CompCode : String;
begin
    try

      sL_DbPassword := frmMainMenu.sG_DbPassword;
      sL_DbAlias := frmMainMenu.sG_DbAlias;
      sL_DbUserName := frmMainMenu.sG_DbUserName;


      if not ADOConnection1.Connected then
      begin
        sL_DbConnStr := 'Provider=MSDAORA.1;Password=' + sL_DbPassword + ';User ID='+sL_DbUserName+';Data Source='+sL_DbAlias+';Persist Security Info=True'; //適用於其他 PC
        ADOConnection1.ConnectionString := sL_DbConnStr;
        ADOConnection1.Connected := true;
      end;

    except
      result := '-1: 與CC&B資料庫連線失敗, 資料庫別名=<'+sL_DbAlias+'>, DB User=<'+sL_DbUserName +'>';
      exit;
    end;
    result := '';

end;

procedure TfrmMainMenu.activePrivDataset(sI_CompCode, sI_UserID: String);
var
    sL_GroupID, sL_SQL : String;
begin
    with qryCommon do
    begin
      Close;
      SQL.Clear;
      sL_SQL :='select GROUPID from So026 where USERID=' + STR_SEP + sI_UserID + STR_SEP + ' and COMPCODE=' + sI_CompCode;
      SQL.Add(sL_SQL);
      Open;
      if RecordCount=1 then
        sL_GroupID := FieldByName('GROUPID').AsString
      else
        sL_GroupID := '';
      Close;
    end;

    if sL_GroupID <> ''then
    begin
      sL_SQL := 'SELECT MID FROM SO029 WHERE GROUP' + sL_GroupID + '=1 And COMPCODE=' + sI_CompCode;
      with qryPriv do
      begin
        Close;
        SQL.Clear;
        SQL.Add(sL_SQL);
        Open;
      end;
    end;
end;

function TfrmMainMenu.checkPriv(sI_Mid: String): boolean;
begin
    if qryPriv.State = dsInactive then
    begin
      result := false;
      exit;
    end;

    with qryPriv do

    begin
      if Locate('MID', sI_Mid, [loPartialKey]) then
        result := true
      else
        result := false;
    end;
end;

procedure TfrmMainMenu.btnSO8A10Click(Sender: TObject);
begin
    createDtmMain2;

    Application.CreateForm(TfrmSO8A10,frmSO8A10);
    frmSO8A10.ShowModal;
end;

procedure TfrmMainMenu.btnSO8A20Click(Sender: TObject);
begin
    createDtmMain2;

    Application.CreateForm(TfrmSO8A20,frmSO8A20);
    frmSO8A20.ShowModal;

end;

procedure TfrmMainMenu.btnSO8A30Click(Sender: TObject);
begin
    createDtmMain2;

    Application.CreateForm(TfrmSO8A30,frmSO8A30);
    frmSO8A30.ShowModal;
end;

procedure TfrmMainMenu.btnSO8A40Click(Sender: TObject);
begin
    createDtmMain2;

    Application.CreateForm(TfrmSO8A40,frmSO8A40);
    frmSO8A40.ShowModal;
end;

procedure TfrmMainMenu.btnSO8A50Click(Sender: TObject);
begin
    createDtmMain2;

    Application.CreateForm(TfrmSO8A50,frmSO8A50);
    frmSO8A50.ShowModal;
end;

procedure TfrmMainMenu.btnSO8A60Click(Sender: TObject);
begin
    createDtmMain2;

    Application.CreateForm(TfrmSO8A60,frmSO8A60);
    frmSO8A60.ShowModal;
end;

procedure TfrmMainMenu.createDtmMain2;
var
    sL_Result : String;
begin
    //執行等待狀態
    TUCommonFun.setWaitingCursor;

    
    if not Assigned(dtmMain2) then
      dtmMain2 := TdtmMain2.Create(nil);

    sL_Result := connectToDB2(frmMainMenu.sG_CompCode);
    if sL_Result <> '' then
    begin
        showmessage(sL_Result);
        exit;
    end;

    //回復原狀態
    TUCommonFun.setDefaultCursor;
end;

function TfrmMainMenu.connectToDB2(sI_Comcpde: String): WideString;
var
    L_Intf : IEmc_Separate1;
    sL_DBResult : OleVariant;
begin
    dtmMain2.connToServer;
    //if not dtmMain2.DCOMConnection1.Connected then
    //  dtmMain2.DCOMConnection1.Connected := true;

    L_Intf := IUnknown(dtmMain2.DCOMConnection1.AppServer) as IEmc_Separate1;

    L_Intf.connectToDB(sG_User,sG_CompCode,sG_DbUserName,sG_DbPassword,sG_DbAlias,sL_DBResult);


    if sL_DBResult <> '' then
    begin

        MessageDlg(sL_DBResult,mtError, [mbOK],0);
        Application.Terminate;
        Exit;
    end;
end;

procedure TfrmMainMenu.btnSO8F10Click(Sender: TObject);
begin
    createDtmMain4;

    frmSO8F10_1 := TfrmSO8F10_1.Create(Application);
    frmSO8F10_1.ShowModal;
    frmSO8F10_1.Free;
end;

procedure TfrmMainMenu.btnSO8F20Click(Sender: TObject);
begin
    createDtmMain4;
    
    frmSO8F20 := TfrmSO8F20.Create(Application);
    frmSO8F20.ShowModal;
    frmSO8F20.Free;
end;

procedure TfrmMainMenu.createDtmMain4;
begin
    //執行等待狀態
    TUCommonFun.setWaitingCursor;

    if not Assigned(dtmMain4) then
      dtmMain4 := TdtmMain4.Create(nil);

    //回復原狀態
    TUCommonFun.setDefaultCursor;
end;

procedure TfrmMainMenu.btnSo8C10Click(Sender: TObject);
var
    L_Frm : TfrmSO8C10;
begin
    createDtmMain3;

    L_Frm := TfrmSO8C10.Create(Application);
    L_Frm.ShowModal;
    L_Frm.Free;
end;

procedure TfrmMainMenu.btnSo8C20Click(Sender: TObject);
var
    L_Frm : TfrmSO8C20;
begin
    createDtmMain3;

    L_Frm := TfrmSO8C20.Create(Application);
    L_Frm.ShowModal;
    L_Frm.Free;
end;

procedure TfrmMainMenu.btnSo8C30Click(Sender: TObject);
var
    L_Frm : TfrmSO8C30;
begin
    createDtmMain3;

    L_Frm := TfrmSO8C30.Create(Application);
    L_Frm.ShowModal;
    L_Frm.Free;
end;

procedure TfrmMainMenu.btnSo8C40Click(Sender: TObject);
var
    L_Frm : TfrmSO8C40_1;
begin
    createDtmMain3;

    L_Frm := TfrmSO8C40_1.Create(Application);
    L_Frm.ShowModal;
    L_Frm.Free;
end;

procedure TfrmMainMenu.createDtmMain3;
begin
    //執行等待狀態
    TUCommonFun.setWaitingCursor;

    if not Assigned(dtmMain3) then
      dtmMain3 := TdtmMain3.Create(nil);

    //回復原狀態
    TUCommonFun.setDefaultCursor;
end;

function TfrmMainMenu.setCaption(sI_FunctionID, sI_Desc: String): String;
begin
    Result := sG_CompName + '-' + sG_User + '--' + sI_Desc + '-[' + sI_FunctionID + ']';
end;

procedure TfrmMainMenu.FormShow(Sender: TObject);
begin
    self.Caption := sG_CompName + '-' + sG_User + '--[輔助功能]';
end;

function TfrmMainMenu.TransTmpIniFile(sI_IniFileName: String) : String;
var
    L_StrList, L_TmpStrList : TStringList;
    sL_ExeFileName, sL_ExePath ,sL_IniFileName,sL_TempIniFileName : String;
    ii : Integer;
    L_Intf : _Encrypt;
    sL_EncKey : WideString;
    //sL_DataAresSection : STring;
    sL_Temp : WideString;
begin
    //sL_DataAresSection := 'DATAAREA';
    sL_EncKey := 'CS';
    L_Intf := CoEncrypt.Create;
    sL_ExeFileName := Application.ExeName;
    sL_ExePath := ExtractFileDir(sL_ExeFileName);

    sL_ExeFileName := Application.ExeName;

    sL_ExePath := ExtractFileDir(sL_ExeFileName);

    sL_IniFileName := sL_ExePath + '\' + sI_IniFileName;

    if not FileExists(sL_IniFileName) then
    begin
      result := '讀取資料庫連線參數檔<'+sL_IniFileName +'>失敗';
      exit;
    end;


    L_StrList := TStringList.Create;
    L_TmpStrList := TStringList.Create;

    L_StrList.LoadFromFile(sL_IniFileName);
    for ii:=0 to L_StrList.Count-1 do
    begin
       if (Copy(L_StrList.Strings[ii],1,1)=';') or (L_StrList.Strings[ii]='')
          or (Copy(L_StrList.Strings[ii],1,8)='COMPCODE') or (Copy(L_StrList.Strings[ii],1,8)='COMPNAME') then
       begin
         L_TmpStrList.Add(L_StrList.Strings[ii]);
       end
       else
       begin
         //L_TmpStrList.Add(L_Intf.Decrypt(L_StrList.Strings[ii], sL_EncKey))
         sL_Temp := L_StrList.Strings[ii];
         L_TmpStrList.Add(L_Intf.Decrypt(sL_Temp));
       end;

    end;
    L_TmpStrList.SaveToFile(sL_ExePath + '\' + TMP_SYS_INI_FILE_NAME);
    L_TmpStrList.Free;

    L_Intf := nil;
{
var
    L_StrList, L_TmpStrList : TStringList;
    sL_ExeFileName, sL_ExePath : String;
    ii : Integer;
    //L_Intf : _Encrypt;
    sL_EncKey : WideString;
    sL_Temp : WideString;
begin
    sL_EncKey := 'CS';
    //L_Intf := CoEncrypt.Create;
    sL_ExeFileName := Application.ExeName;

    sL_ExePath := ExtractFileDir(sL_ExeFileName);

    L_StrList := TStringList.Create;
    L_TmpStrList := TStringList.Create;

    L_StrList.LoadFromFile(sL_ExePath + '\' + sI_IniFileName);
    for ii:=0 to L_StrList.Count-1 do
    begin
       if (Copy(L_StrList.Strings[ii],1,1)=';') or (L_StrList.Strings[ii]='')
          or (Copy(L_StrList.Strings[ii],1,8)='COMPCODE') or (Copy(L_StrList.Strings[ii],1,8)='COMPNAME') then
       begin
         L_TmpStrList.Add(L_StrList.Strings[ii]);
       end
       else
       begin
         //L_TmpStrList.Add(L_Intf.Decrypt(L_StrList.Strings[ii], sL_EncKey))
         sL_Temp := L_StrList.Strings[ii];
         //L_TmpStrList.Add(L_Intf.Decrypt(sL_Temp));
       end;

    end;
    L_TmpStrList.SaveToFile(TMP_SYS_INI_FILE_NAME);
    L_TmpStrList.Free;

    //L_Intf := nil;

}
end;

end.
