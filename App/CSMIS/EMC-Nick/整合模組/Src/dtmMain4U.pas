//Provider=MSDAORA.1;Password=v30;User ID=v30;Data Source=mis;Persist Security Info=True
unit dtmMain4U;

interface

uses
  SysUtils, Classes, DB, ADODB, DBClient, Dialogs,StdCtrls ,Forms;

const
    STR_SEP = '''';
    NO_FILTER_FLAG = '999';
    ALL_COMPCODE = '999';
    PARAM_COUNT= 9;   //傳入參數個數

type
  TdtmMain4 = class(TDataModule)
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
    cdsIniInfo: TClientDataSet;
    cdsIniInfoDataAreaNo: TStringField;
    cdsIniInfoDataArea: TStringField;
    cdsIniInfoAlias: TStringField;
    cdsIniInfoUserID: TStringField;
    cdsIniInfoPassword: TStringField;
    cdsIniInfoCompCode: TStringField;
    cdsIniInfoCompName: TStringField;
    adoQryCD067A: TADOQuery;
    adoQryCD067B: TADOQuery;
    adoQryCD067C: TADOQuery;
    cdsDBError: TClientDataSet;
    cdsDBErrorErrorMsg: TStringField;
    cdsDBErrorCompCode: TStringField;
    ADOConnection2: TADOConnection;
    procedure adoCD067BCalcFields(DataSet: TDataSet);
    procedure adoCD067BAfterScroll(DataSet: TDataSet);
    procedure adoCD067CCalcFields(DataSet: TDataSet);
    procedure DataModuleCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    //sG_IsSupervisor,sG_UserID,sG_User,sG_CompCode,sG_CompName : String;
    //sG_DbUserName,sG_DbPassword,sG_DbAlias,sG_NeedLogin : String;
    procedure filterCd067C(sI_CompCode:String);
    procedure activeCd067C(sI_TableName, sI_CodeNo, sI_CompCode:String);
    procedure activeCd067B(sI_TableName:String);
    procedure activeDataSet(I_DataSet : TDataSet);    
    function getCurCobRecID(I_Cob:TComboBox): String;
    function getRptData(sI_TableName,sI_CompCode : String) : Boolean;
    function checkCodeDataCounts(sI_SourceTableName,sI_QueryTableName : String) : Integer;
    function checkDetailDataCounts(sI_QueryTableName,sI_QueryMasterCode : String) : Integer;
    function connectToCommDB(bI_IsSuperUser : Boolean) : WideString;
  end;

var
  dtmMain4: TdtmMain4;

implementation

uses frmSO8F10_2U, UCommonU, UObjectu, frmMainMenuU, frmSO8F10_1U;

{$R *.dfm}

{ TdtmMain }

procedure TdtmMain4.activeCd067B(sI_TableName: String);
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

procedure TdtmMain4.activeDataSet(I_DataSet: TDataSet);
begin
    if not I_DataSet.Active then
      I_DataSet.Active := true;

end;


procedure TdtmMain4.adoCD067BCalcFields(DataSet: TDataSet);
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

procedure TdtmMain4.activeCd067C(sI_TableName, sI_CodeNo,sI_CompCode: String);
var
    sL_SQL : String;
begin
    //down, open CD067C...
    with adoCD067C do
    begin
      Close;
      if (sI_TableName<>'') and (sI_CodeNo<>'') then
      begin
        SQL.Clear;
        sL_SQL := 'select * from Cd067C where TABLENAME=' + STR_SEP +  sI_TableName + STR_SEP + ' and MASTERCODENO=' + sI_CodeNo;

        if sI_CompCode<>'' then
          sL_SQL := sL_SQL + ' and CompCode=' + sI_CompCode;

        SQL.Add(sL_SQL);
        Open;
      end;
    end;

    //up, open Cd067C...
end;

procedure TdtmMain4.adoCD067BAfterScroll(DataSet: TDataSet);
var
    sL_CodeNo, sL_TableName ,sL_CompCode : String;
begin
    sL_CodeNo := DataSet.fieldByName('CODENO').AsString;
    sL_TableName := DataSet.fieldByName('TABLENAME').AsString;

    //如果營管使用則顯示所有公司資料
    {
    if (frmMainMenu.sG_NeedLogin='Y') OR (frmSO8F10_2.sG_CobCompCode='999')then
      sL_CompCode := ''
    else
      sL_CompCode := frmMainMenu.sG_CompCode;
    }

    if frmSO8F10_2.sG_CobCompCode=ALL_COMPCODE then
      sL_CompCode := ''
    else
      sL_CompCode := frmSO8F10_2.sG_CobCompCode;

    activeCd067C(sL_TableName,sL_CodeNo, sL_CompCode);
    frmSO8F10_2.ChangeDetailBtnState;
end;

procedure TdtmMain4.adoCD067CCalcFields(DataSet: TDataSet);
var
    nL_CompCode : Integer;
begin
    if DataSet.FieldByName('STOPFLAG').AsInteger = 1 then
      DataSet.FieldByName('StopFlagName').AsString := '是'
    else
      DataSet.FieldByName('StopFlagName').AsString := '否';

    nL_CompCode := DataSet.FieldByName('COMPCODE').AsInteger;
       

    if nL_CompCode = StrToInt(ALL_COMPCODE) then
      DataSet.FieldByName('refCompName').AsString := ALL_COMPCODE + '-所有'
     else if cdsCompInfo.Locate('COMPCODE',nL_CompCode,[loCaseInsensitive]) then
      DataSet.FieldByName('refCompName').AsString := cdsCompInfo.FieldByName('CompName').AsString;
end;

procedure TdtmMain4.filterCd067C(sI_CompCode: String);
var
    sL_FilterStr : String;
    sL_CodeNo, sL_TableName, sL_CompCode : String;
begin
    sL_CodeNo := adoCD067B.fieldByName('CODENO').AsString;
    sL_TableName := adoCD067B.fieldByName('TABLENAME').AsString;

    if sI_CompCode <> ALL_COMPCODE then
      sL_CompCode := sI_CompCode
    else
      sL_CompCode := '';
    activeCd067C(sL_TableName,sL_CodeNo, sL_CompCode);
    frmSO8F10_2.ChangeDetailBtnState;


end;

function TdtmMain4.getCurCobRecID(I_Cob: TComboBox): String;
begin
  Result := (I_Cob.Items.Objects[I_Cob.ItemIndex] as TNormalObj).s_Code;
end;

function TdtmMain4.getRptData(sI_TableName,sI_CompCode : String): Boolean;
var
    sL_SQL,sL_Where : String;
begin
    sL_SQL := 'SELECT A.TableName,A.CodeNo,A.Description,A.ServiceType,B.CompCode,B.DetailCodeNo ' +
              'FROM Cd067B A,Cd067C B ';
    if sI_TableName <> NO_FILTER_FLAG then    //某一個代碼Table
    begin
      if sI_CompCode = ALL_COMPCODE then
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
      if sI_CompCode = ALL_COMPCODE then
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




procedure TdtmMain4.DataModuleCreate(Sender: TObject);
var
    sL_Result : String;
begin

end;




function TdtmMain4.checkCodeDataCounts(sI_SourceTableName,sI_QueryTableName : String): Integer;
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

function TdtmMain4.checkDetailDataCounts(sI_QueryTableName,
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

function TdtmMain4.connectToCommDB(bI_IsSuperUser : Boolean): WideString;
var
    ii,nL_DataAreaNo : Integer;
    sL_DbConnStr : String;
    sL_DbPassword, sL_DbAlias, sL_DbUserName,sL_CompCode : String;
begin
    try
      if bI_IsSuperUser then
      begin
        frmSO8F10_1.getCommDBIniInfo(sL_DbAlias,sL_DbUserName,sL_DbPassword)
      end
      else
      begin
        sL_DbPassword := frmMainMenu.sG_DbPassword;
        sL_DbAlias := frmMainMenu.sG_DbAlias;
        sL_DbUserName := frmMainMenu.sG_DbUserName;
      end;

      if not ADOConnection2.Connected then
      begin
        sL_DbConnStr := 'Provider=MSDAORA.1;Password=' + sL_DbPassword + ';User ID='+sL_DbUserName+';Data Source='+sL_DbAlias+';Persist Security Info=True'; //適用於其他 PC
        ADOConnection2.ConnectionString := sL_DbConnStr;
        ADOConnection2.Connected := true;
      end
      else
      begin
        ADOConnection2.Connected := false;
        sL_DbConnStr := 'Provider=MSDAORA.1;Password=' + sL_DbPassword + ';User ID='+sL_DbUserName+';Data Source='+sL_DbAlias+';Persist Security Info=True'; //適用於其他 PC
        ADOConnection2.ConnectionString := sL_DbConnStr;
        ADOConnection2.Connected := true;
      end;

    except
      result := '-1: 與CC&B資料庫連線失敗, 資料庫別名=<'+sL_DbAlias+'>, DB User=<'+sL_DbUserName +'>';
      exit;
    end;
    result := '';


end;

end.


