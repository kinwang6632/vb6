unit dtmMainU;

interface

uses
  SysUtils, Classes, DB, ADODB, inifiles, Forms, Encryption_TLB, Variants;
const
    TMP_ADDFUN_INI_FILE_NAME = 'TmpAddFun.ini';
    ADDFUN_INI_FILE_NAME = 'AddFun.ini';
    STR_SEP = '''';
    SO_SEP = ',';
    
type
  TdtmMain = class(TDataModule)
    adoConn: TADOConnection;
    qryComm: TADOQuery;
    qryGroupData: TADOQuery;
    qryGroupDataCODENO: TIntegerField;
    qryGroupDataSONAME: TStringField;
    qryGroupDataSOCODE: TStringField;
    qryGroupDataSTOPFLAG: TIntegerField;
    qryGroupDataUPDTIME: TDateTimeField;
    qryGroupDataUPDEN: TStringField;
    qryGroupDataSTOPFLAGDESC: TStringField;
    qryUserData: TADOQuery;
    qryUserDataUSERID: TStringField;
    qryUserDataUSERNAME: TStringField;
    qryUserDataUSERPASSWD: TStringField;
    qryUserDataUSERGROUP: TIntegerField;
    qryUserDataISSUPERVISOR: TIntegerField;
    qryUserDataISSYSOP: TIntegerField;
    qryUserDataSTOPFLAG: TIntegerField;
    qryUserDataUPDTIME: TDateTimeField;
    qryUserDataUPDEN: TStringField;
    qryUserDataCOMPCODE: TIntegerField;
    qryUserDataDEPTCODE: TIntegerField;
    qryUserDataTEL: TStringField;
    qryUserDataTELEXT: TIntegerField;
    qryUserDataEMAIL: TStringField;
    qryUserDataCREATER: TStringField;
    qryUserDataCREATEDATE: TDateTimeField;
    qryCompCode: TADOQuery;
    qryDeptCode: TADOQuery;
    qryGroupCode: TADOQuery;
    qryUserDataSTOPFLAGDESC: TStringField;
    qryUserDataISSUPERVISORDESC: TStringField;
    qryUserDataISSYSOPDESC: TStringField;
    qryUserDataCOMPCODEDESC: TStringField;
    qryUserDataDEPTCODEDESC: TStringField;
    qryUserDataGROUPCODEDESC: TStringField;
    qryAllGroupCode: TADOQuery;
    qryAllCompCode: TADOQuery;
    qryAllDeptCode: TADOQuery;
    qryGroupDataDESCRIPTION: TStringField;
    procedure DataModuleCreate(Sender: TObject);
    procedure qryGroupDataAfterScroll(DataSet: TDataSet);
    procedure qryGroupDataAfterInsert(DataSet: TDataSet);
    procedure qryGroupDataCalcFields(DataSet: TDataSet);
    procedure qryUserDataCalcFields(DataSet: TDataSet);
    procedure qryUserDataAfterInsert(DataSet: TDataSet);
  private
    { Private declarations }
    sG_UserDataSQL, sG_OrderByOption : String;
    sG_UserName,sG_UserGroup,sG_IsSupervisor,sG_IsSysOp : String;
    function connectToDB : String;
    procedure TransTmpIniFile;

  public
    { Public declarations }
    function getUserName:String;
    function isValidUser(sI_UserID, sI_Passwd : String):boolean;
    procedure getAllGroupData;
    function getUserData(nI_QueryKey:Integer; sI_QueryValue:String):Integer;overload;
    procedure getUserData(sI_OrderByStr:String);overload;        
  end;

var
  dtmMain: TdtmMain;

implementation

uses frmGroupDataU;

{$R *.dfm}

{ TdtmMain }

function TdtmMain.connectToDB: String;
var
    L_IniFile : TIniFile;
    sL_ExeFileName, sL_ExePath, sL_IniFileName : STring;
    nL_DbCount, ii : Integer;
    sL_Result, sL_DbConnStr, sL_CompCode : String;
    sL_DbAlias, sL_DbUserName, sL_DbPassword  : String;
begin
    TransTmpIniFile;
    sL_ExeFileName := Application.ExeName;
    sL_ExePath := ExtractFileDir(sL_ExeFileName);
    sL_IniFileName := sL_ExePath + '\' + TMP_ADDFUN_INI_FILE_NAME;


    sL_Result := '';
    if not FileExists(sL_IniFileName) then
    begin
      sL_DbConnStr := '';
      result := '讀取資料庫連線參數檔<'+sL_IniFileName + '>失敗';
      exit;
    end;

    try
      L_IniFile := TIniFile.Create(sL_IniFileName);





      //down,  connect 到 EBT 寫入 data 的 db
      sL_DbAlias := L_IniFile.ReadString('EBT_DB','ALIAS','');
      sL_DbUserName := L_IniFile.ReadString('EBT_DB','USERID','');
      sL_DbPassword := L_IniFile.ReadString('EBT_DB','PASSWORD','');

      sL_DbConnStr := 'Provider=MSDAORA.1;Password=' + sL_DbPassword + ';User ID='+sL_DbUserName+';Data Source='+sL_DbAlias+';Persist Security Info=True'; //適用於其他 PC

      adoConn.Close;
      adoConn.LoginPrompt := false;
      adoConn.ConnectionString := sL_DbConnStr;
      adoConn.Connected := true;



      //up,  connect 到 EBT 寫入 data 的 db



      L_IniFile.Free;
      DeleteFile(sL_ExePath + '\' + TMP_ADDFUN_INI_FILE_NAME);
    except
      result := '資料庫連線失敗!請檢查ini檔之設定!(並請確定已經建立所有的database object)' + ' Alias:'+ sL_DbAlias + '  UserID:' +sL_DbUserName;
    end;
end;

procedure TdtmMain.TransTmpIniFile;
var
    L_StrList, L_TmpStrList : TStringList;
    sL_ExeFileName, sL_ExePath : String;
    ii : Integer;
    L_Intf : _Password;
    sL_EncKey : WideString;
    sL_DataAresSection : STring;
begin
    sL_DataAresSection := 'DATAAREA';
    sL_EncKey := 'CS';
    L_Intf := CoPassword.Create;
    sL_ExeFileName := Application.ExeName;
    sL_ExePath := ExtractFileDir(sL_ExeFileName);

    L_StrList := TStringList.Create;
    L_TmpStrList := TStringList.Create;

    L_StrList.LoadFromFile(sL_ExePath + '\' + ADDFUN_INI_FILE_NAME);
    for ii:=0 to L_StrList.Count-1 do
    begin
       if (Copy(L_StrList.Strings[ii],1,2)<>'//') then
         L_TmpStrList.Add(L_Intf.Decrypt(L_StrList.Strings[ii], sL_EncKey))
       else
         L_TmpStrList.Add(L_StrList.Strings[ii]);
    end;
    L_TmpStrList.SaveToFile(sL_ExePath + '\' + TMP_ADDFUN_INI_FILE_NAME);
    L_TmpStrList.Free;

    L_Intf := nil;


end;

procedure TdtmMain.DataModuleCreate(Sender: TObject);
begin
    connectToDB;
end;

function TdtmMain.isValidUser(sI_UserID, sI_Passwd: String): boolean;
var
    bL_Result : boolean;
    sL_SQL : String;
begin
    bL_Result := true;
    sL_SQL := 'select UserName,UserGroup,IsSupervisor,IsSysOp from EMC_EBT002 where STOPFLAG=0 and UserID=' + STR_SEP + sI_UserID + STR_SEP +' and UserPasswd=' + STR_SEP + sI_Passwd + STR_SEP ;
    with qryComm do
    begin
     SQL.Clear;
     SQL.Add(sL_SQL);
     Open;

     if RecordCount=0 then
       bL_Result := false
     else
     begin
       sG_UserName := FieldByName('UserName').asString;
       sG_UserGroup := FieldByName('UserGroup').asString;
       sG_IsSupervisor := FieldByName('IsSupervisor').asString;
       sG_IsSysOp := FieldByName('IsSysOp').asString;
     end;
    end;

    result := bL_Result;
end;

function TdtmMain.getUserName: String;
begin
    result := sG_UserName;
end;

procedure TdtmMain.getAllGroupData;
var

    sL_SQL : String;
begin

    sL_SQL := 'select * from EMC_EBT003 order by CODENO';
    with qryGroupData do
    begin
     SQL.Clear;
     SQL.Add(sL_SQL);
     Open;

    end;


end;

procedure TdtmMain.qryGroupDataAfterScroll(DataSet: TDataSet);
begin
    if DataSet.RecordCount>0 then
      frmGroupData.setSoUi;
end;

procedure TdtmMain.qryGroupDataAfterInsert(DataSet: TDataSet);
begin
    DataSet.FieldByName('STOPFLAG').AsInteger := 0;
end;

procedure TdtmMain.qryGroupDataCalcFields(DataSet: TDataSet);
begin
    if DataSet.FieldByName('STOPFLAG').AsInteger=0 then
      DataSet.FieldByName('STOPFLAGDESC').AsString := '否'
    else
      DataSet.FieldByName('STOPFLAGDESC').AsString := '是';    
end;

procedure TdtmMain.qryUserDataCalcFields(DataSet: TDataSet);
var
    nL_CompCode, nL_DeptCode, nL_GroupCode : Integer;
begin
    if DataSet.FieldByName('STOPFLAG').AsInteger =0 then
      DataSet.FieldByName('STOPFLAGDESC').AsString := '否'
    else
      DataSet.FieldByName('STOPFLAGDESC').AsString := '是';

    if DataSet.FieldByName('IsSupervisor').AsInteger =0 then
      DataSet.FieldByName('IsSupervisorDesc').AsString := '否'
    else
      DataSet.FieldByName('IsSupervisorDesc').AsString := '是';

    if DataSet.FieldByName('IsSysOp').AsInteger =0 then
      DataSet.FieldByName('IsSysOpDesc').AsString := '否'
    else
      DataSet.FieldByName('IsSysOpDesc').AsString := '是';

    nL_CompCode := DataSet.FieldByName('COMPCODE').AsInteger;
    nL_DeptCode := DataSet.FieldByName('DEPTCODE').AsInteger;
    nL_GroupCode := DataSet.FieldByName('USERGROUP').AsInteger;

    with qryCompCode do
    begin
      qryCompCode.First;
      while not qryCompCode.Eof do
      begin
        if qryCompCode.FieldByName('CODENO').AsInteger = nL_CompCode then
        begin
          DataSet.FieldByName('COMPCODEDESC').AsString := qryCompCode.fieldByName('DESCRIPTION').AsString;
          break;
        end;
        qryCompCode.Next;
      end;
    end;

    with qryDeptCode do
    begin
      qryDeptCode.First;
      while not qryDeptCode.Eof do
      begin
        if qryDeptCode.FieldByName('CODENO').AsInteger = nL_DeptCode then
        begin
          DataSet.FieldByName('DEPTCODEDESC').AsString := qryDeptCode.fieldByName('DESCRIPTION').AsString;
          break;
        end;
        qryDeptCode.Next;
      end;
    end;

    with qryGroupCode do
    begin
      qryGroupCode.First;
      while not qryGroupCode.Eof do
      begin
        if qryGroupCode.FieldByName('CODENO').AsInteger = nL_GroupCode then
        begin
          DataSet.FieldByName('GROUPCODEDESC').AsString := qryGroupCode.fieldByName('DESCRIPTION').AsString;
          break;
        end;
        qryGroupCode.Next;
      end;
    end;

    {
    if qryCompCode.Locate('CODENO', VarArrayOf([nL_CompCode]), [loPartialKey]) then
      DataSet.FieldByName('COMPCODEDESC').AsString := qryCompCode.fieldByName('DESCRIPTION').AsString;

    if qryDeptCode.Locate('CODENO', VarArrayOf([nL_DeptCode]), [loPartialKey]) then
      DataSet.FieldByName('DEPTCODEDESC').AsString := qryDeptCode.fieldByName('DESCRIPTION').AsString;
    }
end;

function TdtmMain.getUserData(nI_QueryKey: Integer;
  sI_QueryValue: String): Integer;
var
    sL_SQL, sL_WhereCond : String;
    nL_Result : Integer;
begin

    case nI_QueryKey of
     8://是否停用
        sL_WhereCond := ' where STOPFLAG=' + sI_QueryValue;     
     7://是否群組管理者
        sL_WhereCond := ' where ISSUPERVISOR=' + sI_QueryValue;
     6://是否系統管理者
        sL_WhereCond := ' where ISSYSOP=' + sI_QueryValue;
     5://全部資料
      sL_WhereCond := '';
     0://公司別
      begin
        sL_WhereCond := ' where COMPCODE=' + sI_QueryValue;
      end;

     1://部門
      begin
        sL_WhereCond := ' where DEPTCODE=' + sI_QueryValue;
      end;

     2://群組
      begin
        sL_WhereCond := ' where UserGroup=' + sI_QueryValue;      
      end;

     3://帳號
      begin
        sL_WhereCond := ' where UserID=' + STR_SEP + sI_QueryValue + STR_SEP;
      end;

     4://姓名
      begin
        sL_WhereCond := ' where UserName=' + STR_SEP + sI_QueryValue + STR_SEP;      
      end;

    end;

    sL_SQL := 'select * from EMC_EBT002 ' + sL_WhereCond;

    sG_UserDataSQL := sL_SQL;

    with qryUserData do
    begin
      Close;
      SQL.Clear;
      SQL.Add(sL_SQL);
      Open;
      nL_Result := RecordCount;
    end;

    result := nL_Result;
end;

procedure TdtmMain.qryUserDataAfterInsert(DataSet: TDataSet);
begin
    DataSet.FieldByName('IsSupervisor').AsInteger := 0;
    DataSet.FieldByName('IsSysOp').AsInteger := 0;
    DataSet.FieldByName('STOPFLAG').AsInteger := 0;
end;

procedure TdtmMain.getUserData(sI_OrderByStr: String);
var
    sL_SQL : String;

begin
    if sG_UserDataSQL='' then
      Exit
    else
    begin
      if TRIM(sG_OrderByOption)='DESC' then
        sG_OrderByOption := ' ASC'
      else
        sG_OrderByOption := ' DESC';
              
      sL_SQL := sG_UserDataSQL + sI_OrderByStr + sG_OrderByOption;
      with qryUserData do
      begin
        Close;
        SQL.Clear;
        SQL.Add(sL_SQL);
        Open;
      end;
    end;
end;

end.
