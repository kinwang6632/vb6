//Provider=MSDAORA.1;Password=v30;User ID=v30;Data Source=mis;Persist Security Info=True
unit dtmMainU;

interface

uses
  SysUtils, Classes, DB, ADODB, DBClient, Dialogs,StdCtrls ,Forms;

const
    STR_SEP = '''';
    NO_FILTER_FLAG = '0';
    PARAM_COUNT= 9;   //傳入參數個數

type
  TdtmMain = class(TDataModule)
    adoConnection: TADOConnection;
    adoCD067A: TADOQuery;
    adoCD067ATABLENAME: TStringField;
    adoCD067ATABLEDESCRIPTION: TStringField;
    adoCD067AOPERATOR: TStringField;
    adoCD067AUPDTIME: TStringField;
    adoCD067B: TADOQuery;
    adoCD067C: TADOQuery;
    adoCD067BCODENO: TIntegerField;
    adoCD067BTABLENAME: TStringField;
    adoCD067BDESCRIPTION: TStringField;
    adoCD067BREFNO: TIntegerField;
    adoCD067BSERVICETYPE: TStringField;
    adoCD067BSTOPFLAG: TIntegerField;
    adoCD067BOPERATOR: TStringField;
    adoCD067BUPDTIME: TStringField;
    adoCD067CCOMPCODE: TIntegerField;
    adoCD067CTABLENAME: TStringField;
    adoCD067CMASTERCODENO: TIntegerField;
    adoCD067CDETAILCODENO: TIntegerField;
    adoCD067CSTOPFLAG: TIntegerField;
    adoCD067COPERATOR: TStringField;
    adoCD067CUPDTIME: TStringField;
    adoCD067BServiceTypeName: TStringField;
    adoCD067BStopFlagName: TStringField;
    adoCD067CStopFlagName: TStringField;
    adoCD067CrefCompName: TStringField;
    cdsCompInfo: TClientDataSet;
    cdsCompInfoCompCode: TIntegerField;
    cdsCompInfoCompName: TStringField;
    adoCodeCD067A: TADOQuery;
    adoRptData: TADOQuery;
    qryCommon: TADOQuery;
    qryPriv: TADOQuery;
    cdsCompInfoIndex: TIntegerField;
    procedure adoCD067BCalcFields(DataSet: TDataSet);
    procedure adoCD067BAfterScroll(DataSet: TDataSet);
    procedure adoCD067CCalcFields(DataSet: TDataSet);
    procedure DataModuleCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    sG_IsSupervisor,sG_UserID,sG_User,sG_CompCode,sG_CompName : String;
    sG_DbUserName,sG_DbPassword,sG_DbAlias,sG_NeedLogin : String;
    procedure filterCd067C(sI_CompCode:String);
    procedure activeCd067C(sI_TableName, sI_CodeNo, sI_CompCode:String);
    procedure activeCd067B(sI_TableName:String);
    procedure activeDataSet(I_DataSet : TDataSet);    
    function connToDb(sI_DbAlias, sI_DbUserID, sI_DbPasswd : String):boolean;
    function getCurCobRecID(I_Cob:TComboBox): String;
    function getRptData(sI_TableName,sI_CompCode : String) : Boolean;
    function getDbConnInfo : WideString;
    function setCaption(sI_FunctionID, sI_Desc : String) : String;
    procedure activePrivDataset(sI_CompCode,sI_UserID:String);
    function checkPriv(sI_Mid:String):boolean;
    function checkCodeDataCounts(sI_SourceTableName,sI_QueryTableName : String) : Integer;
    function checkDetailDataCounts(sI_QueryTableName,sI_QueryMasterCode : String) : Integer;
  end;

var
  dtmMain: TdtmMain;

implementation

uses frmCateCode2U, UCommonU, UObjectu;

{$R *.dfm}

{ TdtmMain }

procedure TdtmMain.activeCd067B(sI_TableName: String);
var
    sL_TableName, sL_CodeNo : String;
    nL_Cd067BCount : Integer;
begin
    //down, open CD067B...
    with adoCD067B do
    begin
      Close;
      SQL.Clear;
      SQL.Add('select * from Cd067B where TABLENAME=' + STR_SEP +  sI_TableName + STR_SEP );
      SQL.Add(' order by CODENO');
      Open;
      nL_Cd067BCount := RecordCount;
    end;
    //up, open Cd067B...

    if nL_Cd067BCount>0 then
    begin
{
      sL_CodeNo := adoCD067B.fieldByName('CODENO').AsString;
      sL_TableName := adoCD067B.fieldByName('TABLENAME').AsString;
      activeCd067C(sL_TableName,sL_CodeNo);
}
    end;
end;

procedure TdtmMain.activeDataSet(I_DataSet: TDataSet);
begin
    if not I_DataSet.Active then
      I_DataSet.Active := true;

end;

function TdtmMain.connToDb(sI_DbAlias, sI_DbUserID,
  sI_DbPasswd: String): boolean;
var
    bL_Result : boolean;
begin
    bL_Result := true;
    if adoConnection.Connected then
    begin
      result := true;
      Exit;
    end;

//Provider=MSDAORA.1;Password=may;User ID=gicmis;Data Source=mis;Persist Security Info=True
    try
      adoConnection.ConnectionString := 'Provider=MSDAORA.1;Password='+sI_DbPasswd+';User ID='+sI_DbUserID+';Data Source='+sI_DbAlias+';Persist Security Info=True';
      adoConnection.Connected := true;
    except
      bL_Result := false;
    end;
    result := bL_Result;
end;

procedure TdtmMain.adoCD067BCalcFields(DataSet: TDataSet);
begin
    if DataSet.FieldByName('SERVICETYPE').AsString ='C' then
      DataSet.FieldByName('ServiceTypeName').AsString := '有線電視'
    else
      DataSet.FieldByName('ServiceTypeName').AsString := 'CM';

    if DataSet.FieldByName('STOPFLAG').AsInteger = 1 then
      DataSet.FieldByName('StopFlagName').AsString := '是'
    else
      DataSet.FieldByName('StopFlagName').AsString := '否';
end;

procedure TdtmMain.activeCd067C(sI_TableName, sI_CodeNo,sI_CompCode: String);
begin
    //down, open CD067C...
    with adoCD067C do
    begin
      Close;
      if (sI_TableName<>'') and (sI_CodeNo<>'') then
      begin
        SQL.Clear;
        SQL.Add('select * from Cd067C where TABLENAME=' + STR_SEP +  sI_TableName + STR_SEP + ' and MASTERCODENO=' + sI_CodeNo);
        if sI_CompCode<>'' then
          SQL.Add(' and CompCode=' + sI_CompCode);
        Open;
      end;
    end;

    //up, open Cd067C...
end;

procedure TdtmMain.adoCD067BAfterScroll(DataSet: TDataSet);
var
    sL_CodeNo, sL_TableName ,sL_CompCode : String;
begin
    sL_CodeNo := DataSet.fieldByName('CODENO').AsString;
    sL_TableName := DataSet.fieldByName('TABLENAME').AsString;

    //如果營管使用則顯示所有公司資料
    if sG_NeedLogin='Y' then
      sL_CompCode := ''
    else
      sL_CompCode := sG_CompCode;

    activeCd067C(sL_TableName,sL_CodeNo, sL_CompCode);
    frmCateCode2.ChangeDetailBtnState;
end;

procedure TdtmMain.adoCD067CCalcFields(DataSet: TDataSet);
var
    nL_CompCode : Integer;
begin
    if DataSet.FieldByName('STOPFLAG').AsInteger = 1 then
      DataSet.FieldByName('StopFlagName').AsString := '是'
    else
      DataSet.FieldByName('StopFlagName').AsString := '否';

    nL_CompCode := DataSet.FieldByName('COMPCODE').AsInteger;

    if cdsCompInfo.Locate('COMPCODE',nL_CompCode,[loCaseInsensitive]) then
      DataSet.FieldByName('refCompName').AsString := cdsCompInfo.FieldByName('CompName').AsString;
end;

procedure TdtmMain.filterCd067C(sI_CompCode: String);
var
    sL_FilterStr : String;
    sL_CodeNo, sL_TableName, sL_CompCode : String;
begin
    sL_CodeNo := adoCD067B.fieldByName('CODENO').AsString;
    sL_TableName := adoCD067B.fieldByName('TABLENAME').AsString;

    if sI_CompCode <> NO_FILTER_FLAG then
      sL_CompCode := sI_CompCode
    else
      sL_CompCode := '';
    activeCd067C(sL_TableName,sL_CodeNo, sL_CompCode);
    frmCateCode2.ChangeDetailBtnState;


end;

function TdtmMain.getCurCobRecID(I_Cob: TComboBox): String;
begin
  Result := (I_Cob.Items.Objects[I_Cob.ItemIndex] as TNormalObj).s_Code;
end;

function TdtmMain.getRptData(sI_TableName,sI_CompCode : String): Boolean;
var
    sL_SQL,sL_Where : String;
begin
    sL_SQL := 'SELECT A.TableName,A.CodeNo,A.Description,A.ServiceType,B.CompCode,B.DetailCodeNo ' +
              'FROM Cd067B A,Cd067C B ';
    if sI_TableName <> '0' then    //某一個代碼Table
    begin
      if sI_CompCode = '0' then
      begin
        sL_Where := 'WHERE A.TableName=''' + sI_TableName + ''' AND B.MasterCodeNo=A.CodeNo ' +
                    'AND A.StopFlag=0 AND B.StopFlag=0 ORDER BY B.CompCode,A.TableName,A.CodeNo,B.DetailCodeNo';
      end
      else
      begin
        sL_Where := 'WHERE A.TableName=''' + sI_TableName + ''' AND B.MasterCodeNo=A.CodeNo ' +
                    'AND A.StopFlag=0 AND B.StopFlag=0 AND COMPCODE=' + sI_CompCode +
                    ' ORDER BY B.CompCode,A.TableName,A.CodeNo,B.DetailCodeNo';
      end;
    end
    else    //全部代碼Table
    begin
      if sI_CompCode = '0' then
      begin
        sL_Where := 'WHERE B.MasterCodeNo=A.CodeNo ' +
                    'AND A.StopFlag=0 AND B.StopFlag=0 ORDER BY B.CompCode,A.TableName,A.CodeNo,B.DetailCodeNo';
      end
      else
      begin
        sL_Where := 'WHERE B.MasterCodeNo=A.CodeNo ' +
                    'AND A.StopFlag=0 AND B.StopFlag=0 AND COMPCODE=' + sI_CompCode +
                    ' ORDER BY B.CompCode,A.TableName,A.CodeNo,B.DetailCodeNo';
      end;
    end;

    sL_SQL := sL_SQL + sL_Where;

    with adoRptData do
    begin
      SQL.Clear;
      SQL.Add(sL_SQL);
      Open;
    end;
end;


function TdtmMain.getDbConnInfo: WideString;
begin
    //若為各 SO 公司則抓取 CC&B (VB) 傳來之參數
    if ParamCount = PARAM_COUNT then
    begin
      //N 1 Howard 1 觀昇有限 gicmis may mis 127.0.0.1
      //是否為SuperUser LoginUserID LoginUserName CompCode CompName DBUser DbPassword DbAlias Server_IP
      sG_IsSupervisor := Uppercase(ParamStr(1));
      sG_UserID := ParamStr(2);
      sG_User := ParamStr(3);
      sG_CompCode := ParamStr(4);
      sG_CompName := ParamStr(5);
      sG_DbUserName := ParamStr(6);
      sG_DbPassword := ParamStr(7);
      sG_DbAlias := ParamStr(8);
      sG_NeedLogin := Uppercase(ParamStr(9));
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

procedure TdtmMain.DataModuleCreate(Sender: TObject);
var
    sL_Result : String;
begin
    //down, 讀取連線至 DB 的 information
    sL_Result := getDbConnInfo;
    if sL_Result<>'' then
    begin
      showmessage(sL_Result);
      Application.Terminate;
    end;
    //up, 讀取連線至 DB 的 information
end;

function TdtmMain.setCaption(sI_FunctionID, sI_Desc: String): String;
begin
    Result := sG_CompName + '-' + sG_User + '--' + sI_Desc + '-[' + sI_FunctionID + ']';
end;

procedure TdtmMain.activePrivDataset(sI_CompCode, sI_UserID: String);
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

function TdtmMain.checkPriv(sI_Mid: String): boolean;
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

function TdtmMain.checkCodeDataCounts(sI_SourceTableName,sI_QueryTableName : String): Integer;
var
    sL_SQL : String;
begin
    sL_SQL := 'SELECT COUNT(*) COUNTS FROM ' + sI_SourceTableName +
              ' WHERE TableName=''' + sI_QueryTableName + '''';

    with qryCommon do
    begin
      Close;
      SQL.Clear;
      SQL.Add(sL_SQL);
      Open;

      Result := FieldByName('COUNTS').AsInteger;
    end;
end;

function TdtmMain.checkDetailDataCounts(sI_QueryTableName,
  sI_QueryMasterCode: String): Integer;
var
    sL_SQL : String;
begin
    sL_SQL := 'SELECT COUNT(*) COUNTS FROM Cd067C ' +
              ' WHERE TableName=''' + sI_QueryTableName +
              ''' AND MasterCodeNo=' + sI_QueryMasterCode;

    with qryCommon do
    begin
      Close;
      SQL.Clear;
      SQL.Add(sL_SQL);
      Open;

      Result := FieldByName('COUNTS').AsInteger;
    end;

end;

end.


