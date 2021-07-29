unit dtmMainU;

interface

uses
  SysUtils, Classes, DB, DBClient ,Forms ,Variants;

type
  TdtmMain = class(TDataModule)
    cdsUser: TClientDataSet;
    cdsUsersID: TStringField;
    cdsUsersName: TStringField;
    cdsUsersPassword: TStringField;
    cdsParam: TClientDataSet;
    ClientDataSet1: TClientDataSet;
    cdsParamsServerIP: TStringField;
    cdsParamnSPortNo: TIntegerField;
    cdsParamnRPortNo: TIntegerField;
    cdsParamsSysName: TStringField;
    cdsParamsSysVersion: TStringField;
    cdsParamsMisCommandPath: TStringField;
    cdsParamnTimeOut: TIntegerField;
    cdsParamnMaxCommandCount: TIntegerField;
    cdsParambCommandLog: TBooleanField;
    cdsParamnResponseLog: TBooleanField;
    cdsParamdUptTime: TDateTimeField;
    cdsParamsUptName: TStringField;
    cdsParamnCmdRefreshRate: TIntegerField;
    cdsParamnCmdRefreshRate2: TIntegerField;
    cdsParamnCmdResentTimes: TIntegerField;
    cdsParambShowUI: TBooleanField;
    cdsParamsSHotTime: TStringField;
    cdsParamsEHotTime: TStringField;
    ClientDataSet2: TClientDataSet;
    StringField1: TStringField;
    IntegerField1: TIntegerField;
    IntegerField2: TIntegerField;
    StringField2: TStringField;
    StringField3: TStringField;
    StringField4: TStringField;
    IntegerField3: TIntegerField;
    IntegerField4: TIntegerField;
    BooleanField1: TBooleanField;
    IntegerField5: TIntegerField;
    DateTimeField1: TDateTimeField;
    StringField5: TStringField;
    IntegerField6: TIntegerField;
    IntegerField7: TIntegerField;
    IntegerField8: TIntegerField;
    BooleanField2: TBooleanField;
    StringField6: TStringField;
    StringField7: TStringField;
    IntegerField9: TIntegerField;
    StringField8: TStringField;
    IntegerField10: TIntegerField;
    IntegerField11: TIntegerField;
    cdsParamsServerIP2: TStringField;
    cdsParamnSPortNo2: TIntegerField;
    cdsParamnRPortNo2: TIntegerField;
    cdsParamsSysName2: TStringField;
    cdsParamsSysVersion2: TStringField;
    cdsParamsLogPath: TStringField;
    cdsParamnTimeOut2: TIntegerField;
    cdsParamnMaxCommandCount2: TIntegerField;
    cdsParambCommandLog2: TBooleanField;
    cdsParamdUptTime2: TDateTimeField;
    cdsParamsUptName2: TStringField;
    cdsParamnCmdRefreshRate1: TIntegerField;
    cdsParamnCmdRefreshRate22: TIntegerField;
    cdsParamnCmdResentTimes2: TIntegerField;
    cdsParambShowUI2: TBooleanField;
    cdsParamsSHotTime2: TStringField;
    cdsParamsEHotTime2: TStringField;
    cdsParamnVersion: TIntegerField;
    cdsParamsSecuritytype: TStringField;
    cdsParamnFormID: TIntegerField;
    cdsParamnToID: TIntegerField;
    cdsParambResponseLog: TBooleanField;
    cdsParamnConnectionID: TIntegerField;
    cdsParamsKey: TStringField;
    cdsParamnForwardTolerance: TIntegerField;
    cdsParamnBackwardTolerance: TIntegerField;
    cdsParamnCurrency: TIntegerField;
    procedure DataModuleCreate(Sender: TObject);
    procedure DataModuleDestroy(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    G_UserIDStrList : TStringList;
    function getExePath : String;
    procedure activeCDS;
    function checkUserIdIsOnly(sI_UserID : String) : Boolean;
    function getAllUserID : TStringList;
  end;

var
  dtmMain: TdtmMain;

implementation

uses ConstObjU, frmLoginU, frmSysParamU;

{$R *.dfm}

{ TdtmMain }

procedure TdtmMain.activeCDS;
var
    sL_FileName, sL_ExePath : String;

begin
    
    sL_ExePath := dtmMain.getExePath;

    cdsUser.FileName := sL_ExePath + '\' + USER_INFO_FILENAME;
    if not cdsUser.Active then
      //cdsUser.CreateDataSet;
    cdsUser.Active := TRUE;

    sL_FileName := sL_ExePath + '\' + SYS_PARAM_FILENAME;
    cdsParam.LoadFromFile(sL_FileName);
    if not cdsParam.Active then
      cdsParam.CreateDataSet;
      //cdsParam.Active := TRUE;
end;





function TdtmMain.checkUserIdIsOnly(sI_UserID: String): Boolean;
var
    nL_Ndx : Integer;
begin
    //檢查此UserID是否為唯一值
    nL_Ndx := G_UserIDStrList.IndexOf(sI_UserID);

    if nL_Ndx <> -1 then
      Result := false
    else
      Result := true;
end;



function TdtmMain.getAllUserID: TStringList;
begin
    //將 User 資料存於 StringList
    G_UserIDStrList.Clear;

    cdsUser.First;
    while not cdsUser.Eof do
    begin
      G_UserIDStrList.Add(cdsUser.FieldByName('SID').AsString);

      cdsUser.Next;
    end;
end;

function TdtmMain.getExePath: String;
var
    sL_ExeFileName : String;
    sL_ExePath : String;
begin
    //取得執行檔路徑
    sL_ExeFileName := Application.ExeName;

    Result := ExtractFileDir(sL_ExeFileName);
end;

procedure TdtmMain.DataModuleCreate(Sender: TObject);
begin
    G_UserIDStrList := TStringList.Create;
    G_UserIDStrList.Clear;
end;

procedure TdtmMain.DataModuleDestroy(Sender: TObject);
begin
    G_UserIDStrList.Free;
end;



end.
