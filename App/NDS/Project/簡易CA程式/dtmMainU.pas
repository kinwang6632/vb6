unit dtmMainU;

interface

uses
  SysUtils, Classes, DB, ADODB, DBClient, inifiles, Forms, Provider;

const
    CHANNEL_DATA_FILENAME = 'cd024.dat';
    TMP_SYS_INFO_INI_FILE = 'Tmp.ini';
    SYS_INFO_INI_FILE = 'Connection.ini';

type
  TDataAreaObj = class(TObject)
  public
    s_DataAreaID : String;

  end;


  TdtmMain = class(TDataModule)
    ADOConnection1: TADOConnection;
    ADOQuery1: TADOQuery;
    cdsCD024: TClientDataSet;
    DataSetProvider1: TDataSetProvider;
    adoLogConnection: TADOConnection;
    qryReport1: TADOQuery;
    procedure DataModuleCreate(Sender: TObject);
  private
    { Private declarations }

  public
    { Public declarations }
    sG_ControlCompCode,sG_ControlCompName : String;
    procedure transTmpIniFile;
    procedure connectToDB(sI_DataArea:String; var sI_DbSID, sI_UserID, sI_Password, sI_CompCode:String);
    function transData(sI_SQL, sI_TargetDataFileName:String; I_TargetCDS : TClientDataSet):boolean;
    function insertDataToLog(sI_SubscriberID,sI_IccNo,sI_HighLevelCmdID,sI_Action,sI_ExpirationDate,sI_UpdTime,sI_Operator : String) : String;
    procedure getIniData(var I_StrList : TStringList;var sI_ControlCompCode,sI_ControlCompName : String);
  end;

var
  dtmMain: TdtmMain;

implementation

uses Encryption_TLB, frmMainU;

{$R *.dfm}

{ TdtmMain }

procedure TdtmMain.connectToDB(sI_DataArea: String; var sI_DbSID,
  sI_UserID, sI_Password, sI_CompCode: String);
var

    L_IniFile : TIniFile;
    sL_ExeFileName, sL_ExePath,sL_DataSourcePath : String;
begin
    TransTmpIniFile;
    sL_ExeFileName := Application.ExeName;
    sL_ExePath := ExtractFileDir(sL_ExeFileName);

    L_IniFile := TIniFile.Create(sL_ExePath + '\' + TMP_SYS_INFO_INI_FILE);

    sI_DbSID:= L_IniFile.ReadString(sI_DataArea,'SID','');
    sI_UserID:= L_IniFile.ReadString(sI_DataArea,'DB_USER','');
    sI_Password:= L_IniFile.ReadString(sI_DataArea,'DB_USER_PASSWORD','');
    sI_CompCode:= L_IniFile.ReadString(sI_DataArea,'COMP_CODE','');

    L_IniFile.Free;
    DeleteFile(sL_ExePath + '\' + TMP_SYS_INFO_INI_FILE);


    ADOConnection1.Connected := false;
    ADOConnection1.ConnectionString := 'Provider=MSDAORA.1;Password=' + sI_Password + ';User ID='+sI_UserID+';Data Source='+sI_DbSID+';Persist Security Info=True';
//    ADOConnection1.ConnectionString := 'Provider=OraOLEDB.Oracle.1;Password=' + sI_Password + ';Persist Security Info=True;User ID=' + sI_UserID + ';Data Source='+sI_DbSID;
    ADOConnection1.Connected := true;


end;

function TdtmMain.transData(sI_SQL, sI_TargetDataFileName: String; I_TargetCDS: TClientDataSet):boolean;
var
    bL_Result : boolean;
begin
    bL_Result := true;
    try
      if I_TargetCDS.Active then
        I_TargetCDS.EmptyDataSet;
      I_TargetCDS.Close;

      DataSetProvider1.DataSet := ADOQuery1;
      I_TargetCDS.ProviderName := 'DataSetProvider1';
      I_TargetCDS.CommandText := sI_SQL;
      I_TargetCDS.Open;
      I_TargetCDS.ProviderName := '';
      I_TargetCDS.SaveToFile(sI_TargetDataFileName);
    except
      bL_Result := false;
    end;

    result := bL_Result;
end;

procedure TdtmMain.transTmpIniFile;
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

    L_StrList.LoadFromFile(sL_ExePath + '\' + SYS_INFO_INI_FILE);
    for ii:=0 to L_StrList.Count-1 do
    begin
       if (Copy(L_StrList.Strings[ii],1,2)<>'//') then
         L_TmpStrList.Add(L_Intf.Decrypt(L_StrList.Strings[ii], sL_EncKey))
       else
         L_TmpStrList.Add(L_StrList.Strings[ii]);
    end;
    L_TmpStrList.SaveToFile(sL_ExePath + '\' + TMP_SYS_INFO_INI_FILE);
    L_TmpStrList.Free;

    L_Intf := nil;


end;


procedure TdtmMain.DataModuleCreate(Sender: TObject);
var
    sL_ExeFileName,sL_ExePath,sL_DataSourcePath : String;
    L_StrList : TStringList;
begin
    sL_ExeFileName := Application.ExeName;
    sL_ExePath := ExtractFileDir(sL_ExeFileName);

    //Log資料庫
    sL_DataSourcePath := sL_ExePath + '\SimpleCA.mdb';
    adoLogConnection.Connected := false;
    adoLogConnection.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=' + sL_DataSourcePath + ';Persist Security Info=False';
    adoLogConnection.Connected := true;

    //取得控制CA簡易程式公司Info
    L_StrList := TStringList.Create;
    getIniData(L_StrList,sG_ControlCompCode,sG_ControlCompName);
end;

function TdtmMain.insertDataToLog(sI_SubscriberID, sI_IccNo,
  sI_HighLevelCmdID, sI_Action, sI_ExpirationDate,
  sI_UpdTime,sI_Operator: String): String;
var
    sL_SQL : String;
begin
    if UpperCase(sI_ExpirationDate) = 'NULL' then
      sI_ExpirationDate := '';

    sL_SQL := 'INSERT INTO SimpleCaLog VALUES(''' + sI_SubscriberID +
              ''',''' + sI_IccNo + ''',''' + sI_HighLevelCmdID +
              ''',''' + sI_Action + ''',''' + sI_ExpirationDate +
              ''',''' + sI_UpdTime + ''',''' + sI_Operator + ''')';

    with qryReport1 do
    begin
      Close;
      SQL.Clear;
      SQL.Add(sL_SQL);
      ExecSQL;
    end;
end;

procedure TdtmMain.getIniData(var I_StrList: TStringList;
  var sI_ControlCompCode, sI_ControlCompName: String);
var
    L_IniFile : TIniFile;
    sL_ExeFileName, sL_ExePath : String;
    L_StrList : TStringList;
    ii : Integer;
    sL_DataAreaID, sL_DataArea,sL_DataAresSection : String;
    L_Obj : TDataAreaObj;
begin
    sL_DataAresSection := 'DATAAREA';
    L_StrList := TStringList.Create;
    sL_ExeFileName := Application.ExeName;
    sL_ExePath := ExtractFileDir(sL_ExeFileName);

    dtmMain.transTmpIniFile;
    L_IniFile := TIniFile.Create(sL_ExePath + '\' + TMP_SYS_INFO_INI_FILE);
    L_IniFile.ReadSection(sL_DataAresSection, L_StrList);
    for ii:=0 to L_StrList.Count-1 do
    begin
      sL_DataAreaID := L_StrList.Strings[ii];
      L_Obj := TDataAreaObj.Create;
      L_Obj.s_DataAreaID := sL_DataAreaID;
      sL_DataArea := L_IniFile.ReadString(sL_DataAresSection,sL_DataAreaID,'');
      I_StrList.AddObject(sL_DataArea, L_Obj);

    end;

    sI_ControlCompCode := L_IniFile.ReadString('SO_INFO','COMPCODE','');
    sI_ControlCompName := L_IniFile.ReadString('SO_INFO','COMPNAME','');

    L_IniFile.Free;
    L_StrList.Free;
    DeleteFile(sL_ExePath + '\' + TMP_SYS_INFO_INI_FILE);

end;

end.
