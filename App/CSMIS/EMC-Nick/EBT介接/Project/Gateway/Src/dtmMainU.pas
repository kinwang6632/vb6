unit dtmMainU;

interface

uses
  Windows, SysUtils, Classes, DB, DBClient, Forms, ADODB, inifiles, Controls,
  Provider;

type
  TdtmMain = class(TDataModule)
    cdsUser: TClientDataSet;
    cdsUsersID: TStringField;
    cdsUsersName: TStringField;
    cdsUsersPassword: TStringField;
    cdsParam: TClientDataSet;
    cdsParamSysName: TStringField;
    cdsParamnProcessPeriod: TSmallintField;
    cdsParambLogData: TBooleanField;
    cdsParambShowUI: TBooleanField;
    connCSIS: TADOConnection;
    ADOQuery1: TADOQuery;
    cdsParambAcceptNullData: TBooleanField;
    EBT007_UIDataSet: TClientDataSet;
    procedure DataModuleCreate(Sender: TObject);
    procedure DataModuleDestroy(Sender: TObject);
    procedure processSo151(sI_CompCode:String);
    procedure processSo153(sI_CompCode:String);
  private
    { Private declarations }
    G_CompCodeList : TStringList;
    G_CompMappingData : TStringList;
    G_AdoQueryForEbtData, G_AdoQueryForEbtDataUpdate : TADOQuery;
    G_AdoConnAry : array of TADOConnection;
    G_AdoConnAryForEbtData : array of TADOConnection;
    G_SoAdoQuery, G_SoAdoQueryForUpdate  : array of TADOQuery;
    procedure TransTmpIniFile;
    procedure ActiveCDS(I_CDS:TClientDataSet; sI_RefFileName:String; sI_DefaultStatus:String);
    function getEmcCompCode(sI_EbtCompCode:String):String;
    function IsEmcCompCodeExitis(const aEmcCompCode:String):Boolean;
    function RunSoSQL(nI_Mode:Integer; sI_CompCode, sI_SQL:String):TADOQuery;
    function RunSoSQLForUpdate(nI_Mode:Integer; sI_CompCode, sI_SQL:String):TADOQuery;
    function RunEbtSQL(nI_Mode:Integer; sI_SQL:String):TADOQuery;
    function RunEbtSQLForUpdate(nI_Mode:Integer; sI_SQL:String):TADOQuery;
    function hasSo152Data(sI_CompCode, sI_EbtContractNo:String):boolean;
  public
    { Public declarations }
    procedure SaveCDS(aCDS: TClientDataSet; aRefFileName:String);
    function getSysName:String;
    function getProcessPeriod:Integer;
    function getIfShowUI:boolean;
    function getIfLogData: boolean;
    function getIfAcceptNullData: boolean;
    function connectToDB : String;
    procedure CloseDbConn ;
    function getEmcEbt006Data:TDataSet;
    function getEmcEbt007Data:TDataSet;
    procedure updateEmcEbt006(sI_EbtContractNo, sI_ErrCode, sI_ErrMsg,
      sI_ProcessFlag: String);
    procedure UpdateEmcEbt007(aContractNo, aEbtModifySerialNo,
      aErrCode, aErrMsg, aProcessFlag: string);
    procedure dispatch006DataToSo(I_DataSet:TDataSet);
    procedure dispatch007DataToSo(I_DataSet:TDataSet);
    function getSoAdoQueryForUpdate(sI_CompCode:String):TADOQuery;
    function getSoAdoQuery(sI_CompCode:String):TADOQuery;
    procedure getCompMappingData;
    procedure processSoData;
    procedure setTirggerStatus;
    procedure GetCompanyCodeToStringList;
    procedure ReCreateUIDataSet;
    function GetMappingUISeq(const aDataSet: TClientDataSet;
      const aRowId: string ): string;
    function TrimInvalidateChar(const aSource: String ): String;
    procedure ProcessSo004(aCompCode, aCustId, aEbtConstractNo, aEbtModifySerialNo, aTag: String);
    function CheckCanProcs(aCompCode: String): Boolean;
  end;

var
  dtmMain: TdtmMain;

implementation

uses ConstObjU, Ustru, frmMainU, Encryption_TLB, Variants;

{$R *.dfm}

procedure TdtmMain.ActiveCDS(I_CDS: TClientDataSet; sI_RefFileName,
  sI_DefaultStatus: String);
var
    sL_Path : String;
begin
//    SetCurrentDir(ExtractFilePath(Application.ExeName));

    sL_Path := ExtractFilePath(Application.ExeName);

    if (I_CDS.State = dsInactive) then
      I_CDS.CreateDataSet;
    if (FileExists(sL_Path + sI_RefFileName)) then
      I_CDS.LoadFromFile(sL_Path + sI_RefFileName);

    if (I_CDS.RecordCount=0) then
      I_CDS.append
    else
      I_CDS.edit;

    if (sI_DefaultStatus='E') then //edit mode
    begin
    end
    else if (sI_DefaultStatus='B') then//browse mode
    begin
      I_CDS.Post;
    end;

end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMain.CloseDbConn;
var
  AIndex: Integer;
begin
  for AIndex := Low( G_SoAdoQuery ) to High( G_SoAdoQuery ) do
  begin
    G_SoAdoQuery[AIndex].Active := False;
    G_SoAdoQuery[AIndex].Free;
  end;
  for AIndex := Low( G_SoAdoQueryForUpdate ) to High( G_SoAdoQueryForUpdate ) do
  begin
    G_SoAdoQueryForUpdate[AIndex].Active := False;
    G_SoAdoQueryForUpdate[AIndex].Free;
  end;
  for AIndex := Low( G_AdoConnAry ) to High(G_AdoConnAry) do
  begin
    G_AdoConnAry[AIndex].Close;
    G_AdoConnAry[AIndex].Free;
  end;
  for AIndex := Low( G_AdoConnAryForEbtData ) to High( G_AdoConnAryForEbtData ) do
  begin
    G_AdoConnAryForEbtData[AIndex].Close;
    G_AdoConnAryForEbtData[AIndex].Free;
  end;
  SetLength( G_SoAdoQuery, 0 );
  SetLength( G_SoAdoQueryForUpdate, 0 );
  SetLength( G_AdoConnAry, 0 );
  SetLength( G_AdoConnAryForEbtData, 0 );
  FreeAndNil( G_AdoQueryForEbtData );
  FreeAndNil( G_AdoQueryForEbtDataUpdate );
end;

{ ---------------------------------------------------------------------------- }

function TdtmMain.connectToDB:String;
var
    L_IniFile : TIniFile;
    sL_ExeFileName, sL_ExePath, sL_IniFileName : STring;
    nL_DbCount, ii : Integer;
    sL_Result, sL_DbConnStr, sL_CompCode : String;
    sL_DbAlias, sL_DbUserName, sL_DbPassword  : String;
    aErr: String;
begin
    TransTmpIniFile;
    sL_ExeFileName := Application.ExeName;
    sL_ExePath := ExtractFileDir(sL_ExeFileName);
    sL_IniFileName := sL_ExePath + '\' + TMP_CSIS_INI_FILE_NAME;


    sL_Result := '';
    if not FileExists(sL_IniFileName) then
    begin
      sL_DbConnStr := '';
      result := '讀取資料庫連線參數檔<'+sL_IniFileName + '>失敗';
      exit;
    end;

    Screen.Cursor := crSQLWait;
    try

      try
        L_IniFile := TIniFile.Create(sL_IniFileName);
        nL_DbCount := L_IniFile.ReadInteger( 'DBINFO', 'DB_COUNT', 0 );

        setLength( G_AdoConnAry,nL_DbCount );
        setLength( G_SoAdoQueryForUpdate, nL_DbCount );
        setLength( G_SoAdoQuery,nL_DbCount );

        setLength(G_AdoConnAryForEbtData,1);

        //down,  connect 到 各系統台...

        for ii:=1 to nL_DbCount do
        begin
          sL_DbAlias := L_IniFile.ReadString('DBINFO','ALIAS_' + IntToStr(ii),'');
          sL_DbUserName := L_IniFile.ReadString('DBINFO','USERID_' +  IntToStr(ii),'');
          sL_DbPassword := L_IniFile.ReadString('DBINFO','PASSWORD_' + IntToStr(ii),'');
          sL_CompCode := L_IniFile.ReadString('DBINFO','COMPCODE_' + IntToStr(ii),'');

          { nick fix }
          //G_CompCodeList.Add(sL_CompCode);

          //適用於其他 PC
          sL_DbConnStr := 'Provider=MSDAORA.1;Password=' + sL_DbPassword +
            ';User ID='+sL_DbUserName+';Data Source='+sL_DbAlias+
            ';Persist Security Info=True';
          G_AdoConnAry[ii-1] := TAdoConnection.Create(nil);
          G_AdoConnAry[ii-1].LoginPrompt := false;
          G_AdoConnAry[ii-1].ConnectionString := sL_DbConnStr;
          G_AdoConnAry[ii-1].Connected := true;


          G_SoAdoQueryForUpdate[ii-1] := TAdoQuery.Create(nil);
          G_SoAdoQueryForUpdate[ii-1].Connection := G_AdoConnAry[ii-1];


          G_SoAdoQuery[ii-1] := TAdoQuery.Create(nil);
          G_SoAdoQuery[ii-1].Connection := G_AdoConnAry[ii-1];

        end;

       //down,  connect 到 EBT 寫入 data 的 db
       sL_DbAlias := L_IniFile.ReadString('EBT_DB','ALIAS','');
       sL_DbUserName := L_IniFile.ReadString('EBT_DB','USERID','');
       sL_DbPassword := L_IniFile.ReadString('EBT_DB','PASSWORD','');

       sL_DbConnStr := 'Provider=MSDAORA.1;Password=' + sL_DbPassword + ';User ID='+sL_DbUserName+';Data Source='+sL_DbAlias+';Persist Security Info=True'; //適用於其他 PC
       G_AdoConnAryForEbtData[0] := TAdoConnection.Create(Application);
       G_AdoConnAryForEbtData[0].LoginPrompt := false;
       G_AdoConnAryForEbtData[0].ConnectionString := sL_DbConnStr;
       G_AdoConnAryForEbtData[0].Connected := true;

       G_AdoQueryForEbtData := TADOQuery.Create(Application);
       G_AdoQueryForEbtData.Connection := G_AdoConnAryForEbtData[0];
       G_AdoQueryForEbtData.CursorLocation := clUseClient;
       G_AdoQueryForEbtData.CursorType := ctStatic;
       G_AdoQueryForEbtData.LockType := ltBatchOptimistic;
       G_AdoQueryForEbtData.Prepared := True;
       G_AdoQueryForEbtData.CacheSize := 1000;


       G_AdoQueryForEbtDataUpdate := TADOQuery.Create(Application);
       G_AdoQueryForEbtDataUpdate.Connection := G_AdoConnAryForEbtData[0];

       { Mapping UI 顯示的 Data 對應到實際 DataSet }
       
       ReCreateUIDataSet;
       
       L_IniFile.Free;
       DeleteFile( sL_ExePath + '\' + TMP_CSIS_INI_FILE_NAME );
     except
       on E: Exception do
       begin
         aErr := Format(
           '資料庫連線失敗!請檢查ini檔之設定!(並請確定已經建立所有的database object)'#13#10 +
           'Alias:%s, UserID:%s'#13#10 +
           '原因:%s', [sL_DbAlias, sL_DbUserName, E.Message] );
         Result := aErr;  
         //Result := '資料庫連線失敗!請檢查ini檔之設定!(並請確定已經建立所有的database object)' + ' Alias:'+ sL_DbAlias + '  UserID:' +sL_DbUserName;
       end;
     end;
   finally
      Screen.Cursor := crDefault;
   end;  
end;

procedure TdtmMain.DataModuleCreate(Sender: TObject);
begin
    ActiveCDS(dtmMain.cdsUser, USER_INFO_FILENAME, 'B');  //B:browse mode; E:edit mode
    ActiveCDS(dtmMain.cdsParam, PARAM_INFO_FILENAME, 'B');  //B:browse mode; E:edit mode
    G_CompCodeList := TStringList.Create;
    G_CompMappingData := TStringList.Create;


end;

procedure TdtmMain.DataModuleDestroy(Sender: TObject);
begin
    SaveCDS(dtmMain.cdsUser, USER_INFO_FILENAME);
    SaveCDS(dtmMain.cdsParam, PARAM_INFO_FILENAME);
    G_CompCodeList.Free;
    G_CompMappingData.Free;
end;

procedure TdtmMain.dispatch006DataToSo(I_DataSet: TDataSet);
var
    sL_RealEmcCompCode, sL_RealEmcCustID : String;
    sL_EmcCustID, sL_CatvValid, sL_EbtContractNo : String;
    sL_SQL : String;
    L_AdoQuery : TAdoQuery;

    aCanProcs: Boolean;

    dL_EbtContractBDate, dL_EbtContractEDate : TDate;
    sL_EbtCatvID, sL_EbtCustID, sL_EbtContractBDate, sL_EbtContractEDate, sL_EbtCustCName, sL_EbtCustContactPhone : String;
    sL_EbtCustContactMobile, sL_EbtCompOwnerName, sL_EbtContactPhone, sL_EbtItContactName, sL_EbtItContactPhone : String;
    sL_EbtItContactMobile, sL_EbtItEMail, sL_EbtInstAddr, sL_EbtCustAddr, sL_EbtBillAddr : String;
    sL_EbtContractStatusCode, sL_EbtContractStatusDesc, sL_EbtFeePeriodCode, sL_EbtFeePeriodDesc : String;
    sL_EbtServiceType, sL_EbtAgentName, sL_EbtAgentPhone,sL_EbtAgentID, sL_EbtAgentAddress : String;
    sL_EbtIdCardId, sL_EbtCompanyOwnerId : String;
    sL_EbtNonProfitCompanyId, sL_EbtCompanyId, sL_EbtNotes, sL_UpdEn, sL_UpdTime : String;
    aTag: String;

begin
    sL_EmcCustID := I_DataSet.FieldByName('EmcCustID').AsString;
    sL_EbtCatvID := I_DataSet.FieldByName('EbtCatvID').AsString;

    sL_CatvValid := UpperCase(I_DataSet.FieldByName('CatvValid').AsString);

    if sL_CatvValid='Y' then //表示此資料已經是 catv 用戶...
    begin
      sL_RealEmcCompCode := IntToStr(StrToIntDef(Copy(sL_EmcCustID,1,3),VALIDNESS_COMPCODE));
      sL_RealEmcCustID := IntToStr(StrToIntDef(Copy(sL_EmcCustID,4,8),0));
    end
    else
    begin
      sL_RealEmcCompCode := getEmcCompCode(sL_EbtCatvID);
    end;



    if sL_RealEmcCustID='' then
      sL_RealEmcCustID := 'null';

    if (sL_RealEmcCompCode<>IntToStr(VALIDNESS_COMPCODE))then
    begin


        sL_EbtContractNo := I_DataSet.FieldByName('EbtContractNo').AsString;

        sL_EbtCustID := I_DataSet.FieldByName('EbtCustID').AsString;


        sL_EbtContractBDate := I_DataSet.FieldByName('EbtContractBDate').AsString;
        if sL_EbtContractBDate<>'' then
          dL_EbtContractBDate := I_DataSet.FieldByName('EbtContractBDate').AsDateTime
        else
          dL_EbtContractBDate := 0;

        sL_EbtContractEDate := I_DataSet.FieldByName('EbtContractEDate').AsString;
        if sL_EbtContractEDate<>'' then
          dL_EbtContractEDate := I_DataSet.FieldByName('EbtContractEDate').AsDateTime
        else
          dL_EbtContractEDate := 0;


        sL_EbtCustCName := I_DataSet.FieldByName('EbtCustCName').AsString;
        sL_EbtCustContactPhone := I_DataSet.FieldByName('EbtCustContactPhone').AsString;
        sL_EbtCustContactMobile := I_DataSet.FieldByName('EbtCustContactMobile').AsString;
        sL_EbtCompOwnerName := I_DataSet.FieldByName('EbtCompOwnerName').AsString;
        sL_EbtContactPhone := I_DataSet.FieldByName('EbtContactPhone').AsString;
        sL_EbtItContactName := I_DataSet.FieldByName('EbtItContactName').AsString;
        sL_EbtItContactPhone := I_DataSet.FieldByName('EbtItContactPhone').AsString;
        sL_EbtItContactMobile := I_DataSet.FieldByName('EbtItContactMobile').AsString;
        sL_EbtItEMail := I_DataSet.FieldByName('EbtItEMail').AsString;
        sL_EbtInstAddr := I_DataSet.FieldByName('EbtInstAddr').AsString;
        sL_EbtCustAddr := I_DataSet.FieldByName('EbtCustAddr').AsString;

        sL_EbtBillAddr := I_DataSet.FieldByName('EbtBillAddr').AsString;
        sL_EbtContractStatusCode := I_DataSet.FieldByName('EbtContractStatusCode').AsString;
        sL_EbtContractStatusDesc := I_DataSet.FieldByName('EbtContractStatusDesc').AsString;
        sL_EbtFeePeriodCode := I_DataSet.FieldByName('EbtFeePeriodCode').AsString;
        sL_EbtFeePeriodDesc := I_DataSet.FieldByName('EbtFeePeriodDesc').AsString;
        sL_EbtServiceType := I_DataSet.FieldByName('EbtServiceType').AsString;
        sL_EbtAgentName := I_DataSet.FieldByName('EbtAgentName').AsString;
        sL_EbtAgentPhone := I_DataSet.FieldByName('EbtAgentPhone').AsString;
        sL_EbtAgentID := I_DataSet.FieldByName('EbtAgentID').AsString;
        sL_EbtAgentAddress := I_DataSet.FieldByName('EbtAgentAddress').AsString;

        sL_EbtIdCardId := I_DataSet.FieldByName('EbtIdCardId').AsString;
        sL_EbtCompanyOwnerId := I_DataSet.FieldByName('EbtCompanyOwnerId').AsString;
        sL_EbtNonProfitCompanyId := I_DataSet.FieldByName('EbtNonProfitCompanyId').AsString;
        sL_EbtCompanyId := I_DataSet.FieldByName('EbtCompanyId').AsString;
        sL_EbtNotes := I_DataSet.FieldByName('EbtNotes').AsString;
        sL_UpdEn := I_DataSet.FieldByName('UpdEn').AsString;
        sL_UpdTime := I_DataSet.FieldByName('UpdTime').AsString;

        aTag := I_DataSet.FieldByName( 'TAG' ).AsString;


        aCanProcs := CheckCanProcs( sL_RealEmcCompCode );
        if not aCanProcs  then
        begin
          updateEmcEbt006(sL_EbtContractNo, '-1001', '此系統台不須處理', PROCESS_FLAG_4 );
        end;

          

        if sL_CatvValid='Y' then //表示此資料已經是 catv 用戶 => 將資料存到各 SO之 SO152 中
        begin
          if aCanProcs then
          begin
          
            sL_SQL := ' INSERT INTO SO152 ( EMCCOMPCODE,EMCCUSTID,EBTCONTRACTNO,EBTCUSTID,EBTCONTRACTBDATE,EBTCONTRACTEDATE,EBTCUSTCNAME,EBTCUSTCONTACTPHONE, ';
            sL_SQL := sL_SQL +' EBTCUSTCONTACTMOBILE,EBTCOMPOWNERNAME,EBTCONTACTPHONE,EBTITCONTACTNAME,EBTITCONTACTPHONE,EBTITCONTACTMOBILE, ';
            sL_SQL := sL_SQL +' EBTITEMAIL,EBTINSTADDR,EBTCUSTADDR,EBTBILLADDR,EBTCONTRACTSTATUSCODE,EBTCONTRACTSTATUSDESC, ';
            sL_SQL := sL_SQL +' EBTFEEPERIODCODE,EBTFEEPERIODDESC,EBTSERVICETYPE,EBTAGENTNAME,EBTAGENTPHONE,EBTAGENTID,EBTAGENTADDRESS, ';
            sL_SQL := sL_SQL +' EBTIDCARDID,EBTCOMPANYOWNERID,EBTNONPROFITCOMPANYID,EBTCOMPANYID,EBTNOTES,UPDEN, UPDTIME, EBTREQID, EBTREQTYPE, TAG ) VALUES ( ';
            sL_SQL := sL_SQL + sL_RealEmcCompCode + ',' +  sL_RealEmcCustID + ',' + ORACLE_STR_SEP + sL_EbtContractNo + ORACLE_STR_SEP ;
            sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtCustID + ORACLE_STR_SEP +',' + TUstr.getOracleSQLDateStr(dL_EbtContractBDate);
            sL_SQL := sL_SQL + ',' + TUstr.getOracleSQLDateStr(dL_EbtContractEDate) + ',' +  ORACLE_STR_SEP + sL_EbtCustCName + ORACLE_STR_SEP;
            sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtCustContactPhone + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtCustContactMobile + ORACLE_STR_SEP;
            sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtCompOwnerName + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtContactPhone + ORACLE_STR_SEP;
            sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtItContactName + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtItContactPhone + ORACLE_STR_SEP;
            sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtItContactMobile + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtItEMail + ORACLE_STR_SEP;
            sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtInstAddr + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtCustAddr + ORACLE_STR_SEP;

            sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtBillAddr + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtContractStatusCode + ORACLE_STR_SEP;
            sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtContractStatusDesc + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtFeePeriodCode + ORACLE_STR_SEP;
            sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtFeePeriodDesc + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtServiceType + ORACLE_STR_SEP;
            sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtAgentName + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtAgentPhone + ORACLE_STR_SEP;
            sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtAgentID + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtAgentAddress + ORACLE_STR_SEP;
            sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtIdCardId + ORACLE_STR_SEP;
            sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtCompanyOwnerId + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtNonProfitCompanyId + ORACLE_STR_SEP;

            sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtCompanyId + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtNotes + ORACLE_STR_SEP;
            sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_UpdEn + ORACLE_STR_SEP + ',' + TUstr.getOracleSQLDateTimeStr(now) + ',' + 'NULL' + ',' + 'NULL' + ',' + ORACLE_STR_SEP + aTag + ORACLE_STR_SEP + ')';

            L_AdoQuery := runSoSQLForUpdate(IUD_MODE,sL_RealEmcCompCode,sL_SQL);
            updateEmcEbt006(sL_EbtContractNo,'','',PROCESS_FLAG_5);

            { 處理吊牌號碼 }

            if aTag <> EmptyStr then
               ProcessSo004( sL_RealEmcCompCode, sL_RealEmcCustID, sL_EbtContractNo, EmptyStr, aTag );
          end;
        end
        else
        begin//表示此資料不是 catv 用戶 => 將資料存到各 SO之 SO153 中

          if aCanProcs then
          begin

            sL_SQL := ' INSERT INTO SO153( EMCCOMPCODE,EMCCUSTID,EBTCONTRACTNO,EBTCUSTID,EBTCONTRACTBDATE,EBTCONTRACTEDATE,EBTCUSTCNAME,EBTCUSTCONTACTPHONE,';
            sL_SQL := sL_SQL + 'EBTCUSTCONTACTMOBILE,EBTCOMPOWNERNAME,EBTCONTACTPHONE,EBTITCONTACTNAME,EBTITCONTACTPHONE,EBTITCONTACTMOBILE,';
            sL_SQL := sL_SQL + 'EBTITEMAIL,EBTINSTADDR,EBTCUSTADDR,EBTBILLADDR,EBTCONTRACTSTATUSCODE,EBTCONTRACTSTATUSDESC,';
            sL_SQL := sL_SQL + 'EBTFEEPERIODCODE,EBTFEEPERIODDESC,EBTSERVICETYPE,EBTAGENTNAME,EBTAGENTPHONE,EBTAGENTID,EBTAGENTADDRESS,';
            sL_SQL := sL_SQL + 'EBTIDCARDID,EBTCOMPANYOWNERID,EBTNONPROFITCOMPANYID,EBTCOMPANYID,EBTNOTES,UPDEN, UPDTIME, EbtCatvID,ProcessFlag, SEQNO, IFSYNC, TAG ) ';
            sL_SQL := sL_SQL + ' VALUES ( ';
            sL_SQL := sL_SQL + sL_RealEmcCompCode + ',' +  sL_RealEmcCustID + ',' + ORACLE_STR_SEP + sL_EbtContractNo + ORACLE_STR_SEP ;
            sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtCustID + ORACLE_STR_SEP +',' + TUstr.getOracleSQLDateStr(dL_EbtContractBDate);
            sL_SQL := sL_SQL + ',' + TUstr.getOracleSQLDateStr(dL_EbtContractEDate) + ',' +  ORACLE_STR_SEP + sL_EbtCustCName + ORACLE_STR_SEP;
            sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtCustContactPhone + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtCustContactMobile + ORACLE_STR_SEP;
            sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtCompOwnerName + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtContactPhone + ORACLE_STR_SEP;
            sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtItContactName + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtItContactPhone + ORACLE_STR_SEP;
            sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtItContactMobile + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtItEMail + ORACLE_STR_SEP;
            sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtInstAddr + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtCustAddr + ORACLE_STR_SEP;

            sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtBillAddr + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtContractStatusCode + ORACLE_STR_SEP;
            sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtContractStatusDesc + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtFeePeriodCode + ORACLE_STR_SEP;
            sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtFeePeriodDesc + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtServiceType + ORACLE_STR_SEP;
            sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtAgentName + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtAgentPhone + ORACLE_STR_SEP;
            sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtAgentID + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtAgentAddress + ORACLE_STR_SEP;
            sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtIdCardId + ORACLE_STR_SEP;
            sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtCompanyOwnerId + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtNonProfitCompanyId + ORACLE_STR_SEP;

            sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtCompanyId + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtNotes + ORACLE_STR_SEP;
            sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_UpdEn + ORACLE_STR_SEP + ',' + TUstr.getOracleSQLDateTimeStr(now) ;
            sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtCatvID + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP +  PROCESS_FLAG_2 + ORACLE_STR_SEP ;
            sL_SQL := sL_SQL + ',S_So153_SeqNo.NextVal,' + ORACLE_STR_SEP + 'N' + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + aTag + ORACLE_STR_SEP + ')';

            L_AdoQuery := runSoSQLForUpdate(IUD_MODE,sL_RealEmcCompCode,sL_SQL);
            updateEmcEbt006(sL_EbtContractNo,'','',PROCESS_FLAG_2);
            { 這裏還沒辦法處理吊牌號碼, 必須等 SO153 付予客編完成後, 才可以處理 }

          end;
        end;

    end;

end;

procedure TdtmMain.dispatch007DataToSo(I_DataSet: TDataSet);
var

    aActionType : Integer;
    sL_RealEmcCompCode, sL_RealEmcCustID : String;
    sL_EmcCustID, sL_CatvValid, sL_EbtContractNo : String;
    sL_ProcessFlag,  sL_SQL, sL_EbtModifySerialNo : String;
    L_AdoQuery : TAdoQuery;

    aErrCode, aErrMsg: String;
    aCanProcs: Boolean;

    dL_EbtContractBDate, dL_EbtContractEDate : TDate;
    sL_EbtCatvID, sL_EbtCustID, sL_EbtContractBDate, sL_EbtContractEDate, sL_EbtCustCName, sL_EbtCustContactPhone : String;
    sL_EbtCustContactMobile, sL_EbtCompOwnerName, sL_EbtContactPhone, sL_EbtItContactName, sL_EbtItContactPhone : String;
    sL_EbtItContactMobile, sL_EbtItEMail, sL_EbtInstAddr, sL_EbtCustAddr, sL_EbtBillAddr : String;
    sL_EbtContractStatusCode, sL_EbtContractStatusDesc, sL_EbtFeePeriodCode, sL_EbtFeePeriodDesc : String;
    sL_EbtServiceType, sL_EbtAgentName, sL_EbtAgentPhone,sL_EbtAgentID, sL_EbtAgentAddress : String;
    sL_EbtIdCardId, sL_EbtCompanyOwnerId, sL_IfMoveToOtherSo : String;
    sL_EbtNonProfitCompanyId, sL_EbtCompanyId, sL_EbtNotes, sL_UpdEn, sL_UpdTime : String;
    sL_OldEmcCustID, sL_RealOldEmcCompCode, sL_RealOldEmcCustID : String;
    sL_EbtReqId, sL_EbtReqType, sL_Tag: String;


    procedure insertSO152(sI_CompCode:String);
    var
        sL_SQL : String;
        L_AdoQuery : TAdoQuery;
    begin
          sL_SQL := 'insert into SO152(EMCCOMPCODE,EMCCUSTID,EBTCONTRACTNO,EBTCUSTID,EBTCONTRACTBDATE,EBTCONTRACTEDATE,EBTCUSTCNAME,EBTCUSTCONTACTPHONE,';
          sL_SQL := sL_SQL +'EBTCUSTCONTACTMOBILE,EBTCOMPOWNERNAME,EBTCONTACTPHONE,EBTITCONTACTNAME,EBTITCONTACTPHONE,EBTITCONTACTMOBILE,';
          sL_SQL := sL_SQL +'EBTITEMAIL,EBTINSTADDR,EBTCUSTADDR,EBTBILLADDR,EBTCONTRACTSTATUSCODE,EBTCONTRACTSTATUSDESC,';
          sL_SQL := sL_SQL +'EBTFEEPERIODCODE,EBTFEEPERIODDESC,EBTSERVICETYPE,EBTAGENTNAME,EBTAGENTPHONE,EBTAGENTID,EBTAGENTADDRESS,';
          sL_SQL := sL_SQL +'EBTIDCARDID,EBTCOMPANYOWNERID,EBTNONPROFITCOMPANYID,EBTCOMPANYID,EBTNOTES,UPDEN, UPDTIME, EBTREQID, EBTREQTYPE, TAG ) VALUES ( ';
          sL_SQL := sL_SQL + sL_RealEmcCompCode + ',' +  sL_RealEmcCustID + ',' + ORACLE_STR_SEP + sL_EbtContractNo + ORACLE_STR_SEP ;
          sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtCustID + ORACLE_STR_SEP +',' + TUstr.getOracleSQLDateStr(dL_EbtContractBDate);
          sL_SQL := sL_SQL + ',' + TUstr.getOracleSQLDateStr(dL_EbtContractEDate) + ',' +  ORACLE_STR_SEP + sL_EbtCustCName + ORACLE_STR_SEP;
          sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtCustContactPhone + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtCustContactMobile + ORACLE_STR_SEP;
          sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtCompOwnerName + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtContactPhone + ORACLE_STR_SEP;
          sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtItContactName + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtItContactPhone + ORACLE_STR_SEP;
          sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtItContactMobile + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtItEMail + ORACLE_STR_SEP;
          sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtInstAddr + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtCustAddr + ORACLE_STR_SEP;

          sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtBillAddr + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtContractStatusCode + ORACLE_STR_SEP;
          sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtContractStatusDesc + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtFeePeriodCode + ORACLE_STR_SEP;
          sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtFeePeriodDesc + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtServiceType + ORACLE_STR_SEP;
          sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtAgentName + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtAgentPhone + ORACLE_STR_SEP;
          sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtAgentID + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtAgentAddress + ORACLE_STR_SEP;
          sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtIdCardId + ORACLE_STR_SEP;
          sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtCompanyOwnerId + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtNonProfitCompanyId + ORACLE_STR_SEP;

          sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtCompanyId + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtNotes + ORACLE_STR_SEP;
          sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_UpdEn + ORACLE_STR_SEP + ',' + TUstr.getOracleSQLDateTimeStr(now);
          sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtReqId + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtReqType + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_Tag + ORACLE_STR_SEP + ')';
          L_AdoQuery := runSoSQLForUpdate(IUD_MODE,sI_CompCode,sL_SQL);
    end;

    procedure updateSO152(sI_CompCode:String);
    var
        L_AdoQuery : TAdoQuery;
    begin
        if frmMain.getIfAcceptNullData then
        begin
          sL_SQL := 'UPDATE SO152 SET EMCCOMPCODE =' + sI_CompCode + ',EMCCUSTID=' + sL_RealEmcCustID ;
          sL_SQL := sL_SQL + ',EBTCUSTID=' + ORACLE_STR_SEP + sL_EbtCustID + ORACLE_STR_SEP + ',EBTCONTRACTBDATE='+ TUstr.getOracleSQLDateStr(dL_EbtContractBDate);
          sL_SQL := sL_SQL + ',EBTCONTRACTEDATE='+ TUstr.getOracleSQLDateStr(dL_EbtContractEDate) +',EBTCUSTCNAME=' + ORACLE_STR_SEP + sL_EbtCustCName + ORACLE_STR_SEP;
          sL_SQL := sL_SQL + ',EBTCUSTCONTACTPHONE=' + ORACLE_STR_SEP + sL_EbtCustContactPhone + ORACLE_STR_SEP;
          sL_SQL := sL_SQL + ',EBTCUSTCONTACTMOBILE=' + ORACLE_STR_SEP + sL_EbtCustContactMobile + ORACLE_STR_SEP + ',EBTCOMPOWNERNAME=' + ORACLE_STR_SEP + sL_EbtCompOwnerName + ORACLE_STR_SEP ;
          sL_SQL := sL_SQL + ',EBTCONTACTPHONE='+ ORACLE_STR_SEP + sL_EbtContactPhone + ORACLE_STR_SEP + ',EBTITCONTACTNAME='+ORACLE_STR_SEP + sL_EbtItContactName + ORACLE_STR_SEP;
          sL_SQL := sL_SQL + ',EBTITCONTACTPHONE='+ ORACLE_STR_SEP + sL_EbtItContactPhone + ORACLE_STR_SEP + ',EBTITCONTACTMOBILE='+ ORACLE_STR_SEP + sL_EbtItContactMobile + ORACLE_STR_SEP;
          sL_SQL := sL_SQL + ',EBTITEMAIL='+ ORACLE_STR_SEP + sL_EbtItEMail + ORACLE_STR_SEP + ',EBTINSTADDR='+ ORACLE_STR_SEP + sL_EbtInstAddr + ORACLE_STR_SEP;
          sL_SQL := sL_SQL + ',EBTCUSTADDR=' + ORACLE_STR_SEP + sL_EbtCustAddr + ORACLE_STR_SEP + ',EBTBILLADDR='+ ORACLE_STR_SEP + sL_EbtBillAddr + ORACLE_STR_SEP;
          sL_SQL := sL_SQL + ',EBTCONTRACTSTATUSCODE='+ ORACLE_STR_SEP + sL_EbtContractStatusCode + ORACLE_STR_SEP;
          sL_SQL := sL_SQL + ',EBTCONTRACTSTATUSDESC=' + ORACLE_STR_SEP + sL_EbtContractStatusDesc + ORACLE_STR_SEP;
          sL_SQL := sL_SQL + ',EBTFEEPERIODCODE=' +ORACLE_STR_SEP + sL_EbtFeePeriodCode + ORACLE_STR_SEP + ',EBTFEEPERIODDESC=' +ORACLE_STR_SEP + sL_EbtFeePeriodDesc + ORACLE_STR_SEP ;
          sL_SQL := sL_SQL + ',EBTSERVICETYPE='+ ORACLE_STR_SEP + sL_EbtServiceType + ORACLE_STR_SEP + ',EBTAGENTNAME=' + ORACLE_STR_SEP + sL_EbtAgentName + ORACLE_STR_SEP;
          sL_SQL := sL_SQL + ',EBTAGENTPHONE='+ ORACLE_STR_SEP + sL_EbtAgentPhone + ORACLE_STR_SEP;
          sL_SQL := sL_SQL + ',EBTAGENTID='+ ORACLE_STR_SEP + sL_EbtAgentID + ORACLE_STR_SEP;
          sL_SQL := sL_SQL + ',EBTAGENTADDRESS=' + ORACLE_STR_SEP + sL_EbtAgentAddress + ORACLE_STR_SEP;
          sL_SQL := sL_SQL + ',EBTIDCARDID='+ ORACLE_STR_SEP + sL_EbtIdCardId + ORACLE_STR_SEP;
          sL_SQL := sL_SQL + ',EBTCOMPANYOWNERID='+ ORACLE_STR_SEP + sL_EbtCompanyOwnerId + ORACLE_STR_SEP;
          sL_SQL := sL_SQL + ',EBTNONPROFITCOMPANYID='+ ORACLE_STR_SEP + sL_EbtNonProfitCompanyId + ORACLE_STR_SEP;
          sL_SQL := sL_SQL + ',EBTCOMPANYID=' + ORACLE_STR_SEP + sL_EbtCompanyId + ORACLE_STR_SEP;
          sL_SQL := sL_SQL + ',EBTNOTES='+ORACLE_STR_SEP + sL_EbtNotes + ORACLE_STR_SEP;
          sL_SQL := sL_SQL + ',UPDEN='+ ORACLE_STR_SEP + sL_UpdEn + ORACLE_STR_SEP;
          sL_SQL := sL_SQL + ',UPDTIME=' + TUstr.getOracleSQLDateTimeStr(now);
          sL_SQL := sL_SQL + ',EBTREQID=' + ORACLE_STR_SEP + sL_EbtReqId + ORACLE_STR_SEP;
          sL_SQL := sL_SQL + ',EBTREQTYPE=' + ORACLE_STR_SEP + sL_EbtReqType + ORACLE_STR_SEP;
          sL_SQL := sL_SQL + ',TAG=' + ORACLE_STR_SEP + sL_Tag + ORACLE_STR_SEP;
          sL_SQL := sL_SQL + ' WHERE EBTCONTRACTNO=' + ORACLE_STR_SEP + sL_EbtContractNo + ORACLE_STR_SEP;
        end
        else
        begin

          sL_SQL := ' UPDATE SO152 SET EMCCOMPCODE =' + sI_CompCode + ',EMCCUSTID=' + sL_RealEmcCustID ;

          if sL_EbtCustID<>'' then
            sL_SQL := sL_SQL + ',EBTCUSTID=' + ORACLE_STR_SEP + sL_EbtCustID + ORACLE_STR_SEP;

          if dL_EbtContractBDate<>0 then
            sL_SQL := sL_SQL + ',EBTCONTRACTBDATE='+ TUstr.getOracleSQLDateStr(dL_EbtContractBDate);

          if dL_EbtContractEDate<>0 then
            sL_SQL := sL_SQL + ',EBTCONTRACTEDATE='+ TUstr.getOracleSQLDateStr(dL_EbtContractEDate);

          if sL_EbtCustCName<>'' then
            sL_SQL := sL_SQL + ',EBTCUSTCNAME=' + ORACLE_STR_SEP + sL_EbtCustCName + ORACLE_STR_SEP;

          if sL_EbtCustContactPhone<>'' then
            sL_SQL := sL_SQL + ',EBTCUSTCONTACTPHONE=' + ORACLE_STR_SEP + sL_EbtCustContactPhone + ORACLE_STR_SEP;

          if sL_EbtCustContactMobile<>'' then
            sL_SQL := sL_SQL +',EBTCUSTCONTACTMOBILE=' + ORACLE_STR_SEP + sL_EbtCustContactMobile + ORACLE_STR_SEP ;

          if sL_EbtCompOwnerName<>'' then
            sL_SQL := sL_SQL + ',EBTCOMPOWNERNAME=' + ORACLE_STR_SEP + sL_EbtCompOwnerName + ORACLE_STR_SEP ;

          if sL_EbtContactPhone<>'' then
            sL_SQL := sL_SQL +',EBTCONTACTPHONE='+ ORACLE_STR_SEP + sL_EbtContactPhone + ORACLE_STR_SEP ;

          if sL_EbtItContactName<>'' then
            sL_SQL := sL_SQL +',EBTITCONTACTNAME='+ORACLE_STR_SEP + sL_EbtItContactName + ORACLE_STR_SEP;

          if sL_EbtItContactPhone<>'' then
            sL_SQL := sL_SQL +',EBTITCONTACTPHONE='+ ORACLE_STR_SEP + sL_EbtItContactPhone + ORACLE_STR_SEP ;

          if sL_EbtItContactMobile<>'' then
            sL_SQL := sL_SQL + ',EBTITCONTACTMOBILE='+ ORACLE_STR_SEP + sL_EbtItContactMobile + ORACLE_STR_SEP;

          if sL_EbtItEMail<>'' then
            sL_SQL := sL_SQL +',EBTITEMAIL='+ ORACLE_STR_SEP + sL_EbtItEMail + ORACLE_STR_SEP ;

          if sL_EbtInstAddr<>'' then
            sL_SQL := sL_SQL + ',EBTINSTADDR='+ ORACLE_STR_SEP + sL_EbtInstAddr + ORACLE_STR_SEP;

          if sL_EbtCustAddr<>'' then
            sL_SQL := sL_SQL +',EBTCUSTADDR=' + ORACLE_STR_SEP + sL_EbtCustAddr + ORACLE_STR_SEP;
            
          if sL_EbtBillAddr<>'' then
            sL_SQL := sL_SQL + ',EBTBILLADDR='+ ORACLE_STR_SEP + sL_EbtBillAddr + ORACLE_STR_SEP;

          if sL_EbtContractStatusCode<>'' then
            sL_SQL := sL_SQL +',EBTCONTRACTSTATUSCODE='+ ORACLE_STR_SEP + sL_EbtContractStatusCode + ORACLE_STR_SEP;

          if sL_EbtContractStatusDesc<>'' then
            sL_SQL := sL_SQL +',EBTCONTRACTSTATUSDESC=' + ORACLE_STR_SEP + sL_EbtContractStatusDesc + ORACLE_STR_SEP;

          if sL_EbtFeePeriodCode<>'' then
            sL_SQL := sL_SQL +',EBTFEEPERIODCODE=' +ORACLE_STR_SEP + sL_EbtFeePeriodCode + ORACLE_STR_SEP;

          if sL_EbtFeePeriodDesc<>'' then
            sL_SQL := sL_SQL +',EBTFEEPERIODDESC=' +ORACLE_STR_SEP + sL_EbtFeePeriodDesc + ORACLE_STR_SEP ;

          if sL_EbtServiceType<>'' then
            sL_SQL := sL_SQL +',EBTSERVICETYPE='+ ORACLE_STR_SEP + sL_EbtServiceType + ORACLE_STR_SEP ;

          if sL_EbtAgentName<>'' then
            sL_SQL := sL_SQL + ',EBTAGENTNAME=' + ORACLE_STR_SEP + sL_EbtAgentName + ORACLE_STR_SEP;

          if sL_EbtAgentPhone<>'' then
            sL_SQL := sL_SQL +',EBTAGENTPHONE='+ ORACLE_STR_SEP + sL_EbtAgentPhone + ORACLE_STR_SEP;

          if sL_EbtAgentID<>'' then
            sL_SQL := sL_SQL +',EBTAGENTID='+ ORACLE_STR_SEP + sL_EbtAgentID + ORACLE_STR_SEP;

          if sL_EbtAgentAddress<>'' then
            sL_SQL := sL_SQL +',EBTAGENTADDRESS=' + ORACLE_STR_SEP + sL_EbtAgentAddress + ORACLE_STR_SEP;

          if sL_EbtIdCardId<>'' then
            sL_SQL := sL_SQL +',EBTIDCARDID='+ ORACLE_STR_SEP + sL_EbtIdCardId + ORACLE_STR_SEP;

          if sL_EbtCompanyOwnerId<>'' then
            sL_SQL := sL_SQL +',EBTCOMPANYOWNERID='+ ORACLE_STR_SEP + sL_EbtCompanyOwnerId + ORACLE_STR_SEP;

          if sL_EbtNonProfitCompanyId<>'' then
            sL_SQL := sL_SQL +',EBTNONPROFITCOMPANYID='+ ORACLE_STR_SEP + sL_EbtNonProfitCompanyId + ORACLE_STR_SEP;

          if sL_EbtCompanyId<>'' then
            sL_SQL := sL_SQL +',EBTCOMPANYID=' + ORACLE_STR_SEP + sL_EbtCompanyId + ORACLE_STR_SEP;

          if sL_EbtNotes <> '' then
            sL_SQL := sL_SQL +',EBTNOTES='+ORACLE_STR_SEP + sL_EbtNotes + ORACLE_STR_SEP;

          if sL_UpdEn <> '' then
            sL_SQL := sL_SQL +',UPDEN='+ ORACLE_STR_SEP + sL_UpdEn + ORACLE_STR_SEP;

          if sL_EbtReqId <> '' then
            sL_SQL := sL_SQL +',EBTREQID='+ ORACLE_STR_SEP + sL_EbtReqId + ORACLE_STR_SEP;

          if sL_EbtReqType <> '' then
            sL_SQL := sL_SQL +',EBTREQTYPE='+ ORACLE_STR_SEP + sL_EbtReqType + ORACLE_STR_SEP;

          if sL_Tag <> '' then
            sL_SQL := sL_SQL +',TAG='+ ORACLE_STR_SEP + sL_Tag + ORACLE_STR_SEP;

          sL_SQL := sL_SQL +', UPDTIME=' + TUstr.getOracleSQLDateTimeStr(now);
          
          sL_SQL := sL_SQL + ' WHERE EBTCONTRACTNO=' + ORACLE_STR_SEP + sL_EbtContractNo + ORACLE_STR_SEP;

        end;

        L_AdoQuery := runSoSQLForUpdate(IUD_MODE,sI_CompCode,sL_SQL);

    end;

    
    procedure updateSO152MoveToOtherSo(sI_CompCode:String);
    begin
      sL_SQL := ' UPDATE SO152 SET IfMoveToOtherSo =' + ORACLE_STR_SEP + 'Y' + ORACLE_STR_SEP;
      sL_SQL := sL_SQL + ' WHERE EBTCONTRACTNO=' + ORACLE_STR_SEP + sL_EbtContractNo + ORACLE_STR_SEP;
      runSoSQLForUpdate( IUD_MODE, sI_CompCode, sL_SQL );
    end;

    
    procedure insertSO153(sI_CompCode:String);
    var
        L_AdoQuery : TAdoQuery;
    begin
        sL_SQL := 'insert into SO153(EMCCOMPCODE,EMCCUSTID,EBTCONTRACTNO,EBTCUSTID,EBTCONTRACTBDATE,EBTCONTRACTEDATE,EBTCUSTCNAME,EBTCUSTCONTACTPHONE,';
        sL_SQL := sL_SQL +'EBTCUSTCONTACTMOBILE,EBTCOMPOWNERNAME,EBTCONTACTPHONE,EBTITCONTACTNAME,EBTITCONTACTPHONE,EBTITCONTACTMOBILE,';
        sL_SQL := sL_SQL +'EBTITEMAIL,EBTINSTADDR,EBTCUSTADDR,EBTBILLADDR,EBTCONTRACTSTATUSCODE,EBTCONTRACTSTATUSDESC,';
        sL_SQL := sL_SQL +'EBTFEEPERIODCODE,EBTFEEPERIODDESC,EBTSERVICETYPE,EBTAGENTNAME,EBTAGENTPHONE,EBTAGENTID,EBTAGENTADDRESS,';
        sL_SQL := sL_SQL +'EBTIDCARDID,EBTCOMPANYOWNERID,EBTNONPROFITCOMPANYID,EBTCOMPANYID,EBTNOTES,UPDEN, UPDTIME, EbtCatvID,ProcessFlag,SeqNo, ifSync, EbtModifySerialNo, TAG ) values(';
        sL_SQL := sL_SQL + sI_CompCode + ',' +  sL_RealEmcCustID + ',' + ORACLE_STR_SEP + sL_EbtContractNo + ORACLE_STR_SEP ;
        sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtCustID + ORACLE_STR_SEP +',' + TUstr.getOracleSQLDateStr(dL_EbtContractBDate);
        sL_SQL := sL_SQL + ',' + TUstr.getOracleSQLDateStr(dL_EbtContractEDate) + ',' +  ORACLE_STR_SEP + sL_EbtCustCName + ORACLE_STR_SEP;
        sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtCustContactPhone + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtCustContactMobile + ORACLE_STR_SEP;
        sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtCompOwnerName + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtContactPhone + ORACLE_STR_SEP;
        sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtItContactName + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtItContactPhone + ORACLE_STR_SEP;
        sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtItContactMobile + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtItEMail + ORACLE_STR_SEP;
        sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtInstAddr + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtCustAddr + ORACLE_STR_SEP;

        sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtBillAddr + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtContractStatusCode + ORACLE_STR_SEP;
        sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtContractStatusDesc + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtFeePeriodCode + ORACLE_STR_SEP;
        sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtFeePeriodDesc + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtServiceType + ORACLE_STR_SEP;
        sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtAgentName + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtAgentPhone + ORACLE_STR_SEP;
        sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtAgentID + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtAgentAddress + ORACLE_STR_SEP;
        sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtIdCardId + ORACLE_STR_SEP;
        sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtCompanyOwnerId + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtNonProfitCompanyId + ORACLE_STR_SEP;

        sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtCompanyId + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtNotes + ORACLE_STR_SEP;
        sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_UpdEn + ORACLE_STR_SEP + ',' + TUstr.getOracleSQLDateTimeStr(now) ;
        sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtCatvID + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP +  PROCESS_FLAG_2 + ORACLE_STR_SEP ;
        sL_SQL := sL_SQL + ',S_So153_SeqNo.NextVal,' + ORACLE_STR_SEP + 'N' + ORACLE_STR_SEP ;
        sL_SQL := sL_SQL + ',' + ORACLE_STR_SEP + sL_EbtModifySerialNo + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_Tag + ORACLE_STR_SEP + ' ) ';
        L_AdoQuery := runSoSQLForUpdate(IUD_MODE,sI_CompCode,sL_SQL);
    end;

begin
    sL_EmcCustID := I_DataSet.FieldByName('EmcCustID').AsString;
    sL_EbtCatvID := I_DataSet.FieldByName('EbtCatvID').AsString;

    sL_CatvValid := UpperCase(I_DataSet.FieldByName('CatvValid').AsString);
    sL_OldEmcCustID := UpperCase(I_DataSet.FieldByName('OldEmcCustID').AsString);
    sL_IfMoveToOtherSo := UpperCase(I_DataSet.FieldByName('IfMoveToOtherSo').AsString);


    if (sL_IfMoveToOtherSo='Y') and (sL_EmcCustID='')then //表示此筆資料為跨系統台移機, 且沒有 CATV 客編
    begin
      sL_RealEmcCompCode := getEmcCompCode(sL_EbtCatvID);//新的SO 代碼
      sL_RealEmcCustID:='null';
      sL_RealOldEmcCompCode := IntToStr(StrToIntDef(Copy(sL_OldEmcCustID,1,3),VALIDNESS_COMPCODE));//舊的SO 代碼
      sL_RealOldEmcCustID := IntToStr(StrToIntDef(Copy(sL_OldEmcCustID,4,8),0)); //舊SO的 EMC 客編
      aActionType := 1;
    end
    else if (sL_IfMoveToOtherSo='Y') and (sL_EmcCustID<>'')then //表示此筆資料為跨系統台移機, 且已經有 CATV 客編
    begin
      { EBT 給的資料有問題, 必須先判斷 }
      if IsEmcCompCodeExitis( Copy( sL_EmcCustID, 1, 3 ) ) then
      begin
        sL_RealEmcCompCode := IntToStr(StrToIntDef(Copy(sL_EmcCustID,1,3),VALIDNESS_COMPCODE));//新的SO 代碼
        sL_RealEmcCustID := IntToStr(StrToIntDef(Copy(sL_EmcCustID,4,8),0)); //新SO的 EMC 客編
        sL_RealOldEmcCompCode := IntToStr(StrToIntDef(Copy(sL_OldEmcCustID,1,3),VALIDNESS_COMPCODE));//舊的SO 代碼
        sL_RealOldEmcCustID := IntToStr(StrToIntDef(Copy(sL_OldEmcCustID,4,8),0)); //舊SO的 EMC 客編
        aActionType := 2;
      end else
      begin
        aActionType := -1;
        aErrCode := IntToStr( ERR_CODE_20 );
        aErrMsg := ERR_MSG_20;
      end;
    end
    else if (sL_IfMoveToOtherSo='X') and (sL_EmcCustID<>'')then//表示此筆資料只是一般異動資料,不是移機資料
    begin
      { EBT 給的資料有問題, 必須先判斷 }
      if IsEmcCompCodeExitis( Copy( sL_EmcCustID, 1, 3 ) ) then
      begin
        sL_RealEmcCompCode := IntToStr(StrToIntDef(Copy(sL_EmcCustID,1,3),VALIDNESS_COMPCODE));
        sL_RealEmcCustID := IntToStr(StrToIntDef(Copy(sL_EmcCustID,4,8),0));
        aActionType := 3;
      end else
      begin
        aActionType := -1;
        aErrCode := IntToStr( ERR_CODE_20 );
        aErrMsg := ERR_MSG_20;
      end;  
    end
    else if (sL_IfMoveToOtherSo='N') and ( sL_EmcCustID<>'')then//表示此筆資料為同一系統台移機
    begin
      { EBT 給的資料有問題, 必須先判斷 }
      if IsEmcCompCodeExitis( Copy( sL_EmcCustID, 1, 3 ) ) then
      begin
        sL_RealEmcCompCode := IntToStr(StrToIntDef(Copy(sL_EmcCustID,1,3),VALIDNESS_COMPCODE));
        sL_RealEmcCustID := IntToStr(StrToIntDef(Copy(sL_EmcCustID,4,8),0));
        aActionType := 4;
      end else
      begin
        aActionType := -1;
        aErrCode := IntToStr( ERR_CODE_20 );
        aErrMsg := ERR_MSG_20;
      end;
    end;

    sL_Tag := I_DataSet.FieldByName('TAG').AsString;

    sL_EbtContractNo := I_DataSet.FieldByName('EbtContractNo').AsString;

    sL_EbtCustID := I_DataSet.FieldByName('EbtCustID').AsString;


    sL_EbtContractBDate := I_DataSet.FieldByName('EbtContractBDate').AsString;
    if sL_EbtContractBDate<>'' then
      dL_EbtContractBDate := I_DataSet.FieldByName('EbtContractBDate').AsDateTime
    else
      dL_EbtContractBDate := 0;

    sL_EbtContractEDate := I_DataSet.FieldByName('EbtContractEDate').AsString;
    if sL_EbtContractEDate<>'' then
      dL_EbtContractEDate := I_DataSet.FieldByName('EbtContractEDate').AsDateTime
    else
      dL_EbtContractEDate := 0;



    sL_EbtCustCName := I_DataSet.FieldByName('EbtCustCName').AsString;
    sL_EbtCustContactPhone := I_DataSet.FieldByName('EbtCustContactPhone').AsString;
    sL_EbtCustContactMobile := I_DataSet.FieldByName('EbtCustContactMobile').AsString;
    sL_EbtCompOwnerName := I_DataSet.FieldByName('EbtCompOwnerName').AsString;
    sL_EbtContactPhone := I_DataSet.FieldByName('EbtContactPhone').AsString;
    sL_EbtItContactName := I_DataSet.FieldByName('EbtItContactName').AsString;
    sL_EbtItContactPhone := I_DataSet.FieldByName('EbtItContactPhone').AsString;
    sL_EbtItContactMobile := I_DataSet.FieldByName('EbtItContactMobile').AsString;
    sL_EbtItEMail := I_DataSet.FieldByName('EbtItEMail').AsString;
    sL_EbtInstAddr := I_DataSet.FieldByName('EbtInstAddr').AsString;
    sL_EbtCustAddr := I_DataSet.FieldByName('EbtCustAddr').AsString;

    sL_EbtBillAddr := I_DataSet.FieldByName('EbtBillAddr').AsString;
    sL_EbtContractStatusCode := I_DataSet.FieldByName('EbtContractStatusCode').AsString;
    sL_EbtContractStatusDesc := I_DataSet.FieldByName('EbtContractStatusDesc').AsString;
    sL_EbtFeePeriodCode := I_DataSet.FieldByName('EbtFeePeriodCode').AsString;
    sL_EbtFeePeriodDesc := I_DataSet.FieldByName('EbtFeePeriodDesc').AsString;
    sL_EbtServiceType := I_DataSet.FieldByName('EbtServiceType').AsString;
    sL_EbtAgentName := I_DataSet.FieldByName('EbtAgentName').AsString;
    sL_EbtAgentPhone := I_DataSet.FieldByName('EbtAgentPhone').AsString;
    sL_EbtAgentID := I_DataSet.FieldByName('EbtAgentID').AsString;
    sL_EbtAgentAddress := I_DataSet.FieldByName('EbtAgentAddress').AsString;

    sL_EbtIdCardId := I_DataSet.FieldByName('EbtIdCardId').AsString;
    sL_EbtCompanyOwnerId := I_DataSet.FieldByName('EbtCompanyOwnerId').AsString;
    sL_EbtNonProfitCompanyId := I_DataSet.FieldByName('EbtNonProfitCompanyId').AsString;
    sL_EbtCompanyId := I_DataSet.FieldByName('EbtCompanyId').AsString;
    sL_EbtNotes := I_DataSet.FieldByName('EbtNotes').AsString;
    sL_UpdEn := I_DataSet.FieldByName('UpdEn').AsString;
    sL_UpdTime := I_DataSet.FieldByName('UpdTime').AsString;
    sL_EbtModifySerialNo := I_DataSet.FieldByName('EbtModifySerialNo').AsString;
    sL_EbtReqId := I_DataSet.FieldByName('EbtReqId').AsString;
    sL_EbtReqType := I_DataSet.FieldByName('EbtReqType').AsString;

    if aActionType <> -1 then
    begin
      aErrCode := EmptyStr;
      aErrMsg := EmptyStr;
    end;
      

try
    case aActionType of
      -1:  // 資料有誤
       begin
        frmMain.write007Error( I_DataSet, aErrCode, aErrMsg,
          GetMappingUISeq( EBT007_UIDataSet,
            I_DataSet.FieldByName( 'ROWIDS' ).AsString )  );
       end;
      1://跨系統台移機, 且沒有 CATV 客編
       begin
         aCanProcs :=
           CheckCanProcs( sL_RealEmcCompCode ) and
           CheckCanProcs( sL_RealOldEmcCompCode );
         if aCanProcs then
         begin
           insertSo153(sL_RealEmcCompCode);
           updateSO152MoveToOtherSo(sL_RealOldEmcCompCode);
           sL_ProcessFlag := PROCESS_FLAG_2;
         end else
         begin
           aErrCode := '-1001';
           aErrMsg := '此系統台不須處理';
           sL_ProcessFlag := PROCESS_FLAG_4;
         end;
       end;
      2://跨系統台移機, 且已經有 CATV 客編
       begin
         aCanProcs :=
           CheckCanProcs( sL_RealEmcCompCode ) and
           CheckCanProcs( sL_RealOldEmcCompCode );
         if aCanProcs then
         begin
           if hasSo152Data( sL_RealEmcCompCode,sL_EbtContractNo ) then
             UpdateSO152( sL_RealEmcCompCode )
           else
             InsertSO152( sL_RealEmcCompCode );
           UpdateSO152MoveToOtherSo( sL_RealOldEmcCompCode );
           sL_ProcessFlag := PROCESS_FLAG_5;
         end else
         begin
           aErrCode := '-1001';
           aErrMsg := '此系統台不須處理';
           sL_ProcessFlag := PROCESS_FLAG_4;
         end;
       end;
      3://一般異動資料,不是移機資料
       begin
         aCanProcs := CheckCanProcs( sL_RealEmcCompCode );
         if aCanProcs then
         begin
           if hasSo152Data(sL_RealEmcCompCode,sL_EbtContractNo) then
             updateSO152(sL_RealEmcCompCode)
           else
             insertSO152(sL_RealEmcCompCode);
           sL_ProcessFlag := PROCESS_FLAG_5;
         end else
         begin
           aErrCode := '-1001';
           aErrMsg := '此系統台不須處理';
           sL_ProcessFlag := PROCESS_FLAG_4;
         end;
       end;
      4://同一系統台移機
       begin
         aCanProcs := CheckCanProcs( sL_RealEmcCompCode );
         if aCanProcs then
         begin
           if hasSo152Data(sL_RealEmcCompCode,sL_EbtContractNo) then
             updateSO152(sL_RealEmcCompCode)
           else
             insertSO152(sL_RealEmcCompCode);
           sL_ProcessFlag := PROCESS_FLAG_5;
         end else
         begin
           aErrCode := '-1001';
           aErrMsg := '此系統台不須處理';
           sL_ProcessFlag := PROCESS_FLAG_4;
         end;
       end;
    end;

    if aActionType <> -1 then
      updateEmcEbt007( sL_EbtContractNo, sL_EbtModifySerialNo, aErrCode, aErrMsg,
        sL_ProcessFlag );

    { 一律只處理新的客編, 若是跨系統台移機, 則舊的客編及吊牌號碼不處理 }    
    if ( sL_Tag <> EmptyStr ) and ( aErrCode = EmptyStr ) then
    begin
      if aActionType in [2,3,4] then
      begin
        ProcessSo004( sL_RealEmcCompCode, sL_RealEmcCustID, sL_EbtContractNo,
         sL_EbtModifySerialNo, sL_Tag  );
      end;
    end;


except
  on E: Exception do
  begin
    frmMain.ShowErrorOnMemo( Format(
      'EMC_EBT007資料有誤, 系統台:%s, 合約編號:%s, 異動序號:%s, 客戶名稱:%s ' +
      '錯誤訊息:%s ',
       [sL_RealEmcCompCode, sL_EbtContractNo, sL_EbtModifySerialNo,
        sL_EbtCustCName, E.Message] ) );
  end;
end;

end;

function TdtmMain.getSoAdoQueryForUpdate(sI_CompCode: String): TADOQuery;
var
    nL_Ndx : Integer;
    L_AdoQuery : TADOQuery;
begin
    L_AdoQuery := nil;
    nL_Ndx := G_CompCodeList.IndexOf(sI_CompCode);
    if nL_Ndx<>-1 then
    begin
      L_AdoQuery := G_SoAdoQueryForUpdate[nL_Ndx];
    end;

    result := L_AdoQuery;
end;

procedure TdtmMain.getCompMappingData;
var
    sL_EmcCompCode, sL_EbtCompCode : String;
    ii : Integer;
    L_Obj :TCompMappingObj;
    function getEbtCompCode(sI_EmcCompCode:String):String;
    var
        sL_Result : String;
    begin
      case StrToInt(sI_EmcCompCode) of
        1:
          sL_Result := '080101';
        2:
          sL_Result := '060205';
        3:
          sL_Result := '060101';
        5:
          sL_Result := '040101';
        6:
          sL_Result := '040201';
        7:
          sL_Result := '040103';
        8:
          sL_Result := '030210';
        9:
          sL_Result := '020510';
        10:
          sL_Result := '020410';
        11:
          sL_Result := '020210';
        12:
          sL_Result := '020110';
        13:
          sL_Result := '030430';
        14:
          sL_Result := '030440';
        16:
          sL_Result := '032614';
        else
          sL_Result := '';
      end;
      result := sL_Result;
    end;
begin
{
020110	大安文山    12
020210	金頻道      11
020410	新台北      10
020510	陽明山       9
030210	全聯         8
030430	新唐城      13
030440	大新店民主  14
032614	北桃園      16
040101	北彰         5
040103	振道         7
040201	豐盟         6
060101	南天         3
060205	屏南         2
080101	觀昇         1

}
    for ii:=0 to G_CompCodeList.Count -1 do
    begin
      sL_EmcCompCode := G_CompCodeList[ii];
      sL_EbtCompCode := getEbtCompCode(sL_EmcCompCode);
      if sL_EbtCompCode<>'' then
      begin
        L_Obj := TCompMappingObj.Create;
        L_Obj.EmcCompCode := sL_EmcCompCode;
        G_CompMappingData.AddObject(sL_EbtCompCode,L_Obj);
      end;
    end;



end;

function TdtmMain.getEmcCompCode(sI_EbtCompCode: String): String;
var
    sL_Result : String;
    nL_Ndx : Integer;
begin
    sL_Result := IntToStr(VALIDNESS_COMPCODE);
    nL_Ndx := G_CompMappingData.IndexOf(sI_EbtCompCode);
    if nL_Ndx<>-1 then
      sL_Result := (G_CompMappingData.Objects[nL_Ndx] as TCompMappingObj).EmcCompCode;

    result := sL_Result;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMain.IsEmcCompCodeExitis(const aEmcCompCode: String): Boolean;
var
  aIndex: Integer;
  aCompare: String;
  aConvert: string; 
begin
  Result := False;
  try
    aConvert := IntToStr( StrToIntDef( aEmcCompCode, 0 ) );
  except
    { ... }
  end;
  for aIndex := 0 to G_CompMappingData.Count - 1 do
  begin
    aCompare := TCompMappingObj( G_CompMappingData.Objects[aIndex] ).EmcCompCode;
    Result := ( aCompare = aConvert );
    if Result then Break;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMain.getEmcEbt006Data: TDataSet;
var
  aSQL : String;
begin
  //此function 傳回 EBT 完工寫入之資料...
  aSQL := 'select EmcCustID,EbtCatvID,EbtContractNo,EbtCustID,';
  aSQL := aSQL +  ' to_char(EbtContractBDate,' + ORACLE_STR_SEP + 'YYYY/MM/DD' + ORACLE_STR_SEP + ') EbtContractBDate,';
  aSQL := aSQL + ' to_char(EbtContractEDate,' + ORACLE_STR_SEP + 'YYYY/MM/DD' + ORACLE_STR_SEP + ') EbtContractEDate,';
  aSQL := aSQL +  'EbtCustCName,EbtCustContactPhone,EbtCustContactMobile,EbtCompOwnerName,';
  aSQL := aSQL + ' EbtContactPhone,EbtItContactName,EbtItContactPhone,EbtItContactMobile,EbtItEMail,';
  aSQL := aSQL +  ' EbtInstAddr,EbtCustAddr,EbtBillAddr,EbtContractStatusCode,EbtContractStatusDesc,';
  aSQL := aSQL + ' EbtFeePeriodCode,EbtFeePeriodDesc,';
  aSQL := aSQL + ' EbtServiceType,EbtAgentName,EbtAgentPhone,EbtAgentID,EbtAgentAddress,';
  aSQL := aSQL +  ' EbtIdCardId,EbtCompanyOwnerId,';
  aSQL := aSQL +  ' EbtNonProfitCompanyId,EbtCompanyId,ProcessFlag,ErrorCode,ErrorMsg,EbtNotes,EmcNotes, CATVVALID,';
  aSQL := aSQL + ' to_char(UpdTime,' + ORACLE_STR_SEP + 'YYYY/MM/DD' + ORACLE_STR_SEP + ') UpdTime, UpdEn, TAG ';
  aSQL := aSQL + ' from EMC_EBT006 where ProcessFlag=' + ORACLE_STR_SEP + PROCESS_FLAG_1 + ORACLE_STR_SEP;
  Result := RunEbtSQL( SELECT_MODE, aSQL );
end;

{ ---------------------------------------------------------------------------- }

function TdtmMain.getEmcEbt007Data: TDataSet;
var
  aSQL: String;
begin
    //此function 傳回 EBT 之異動資料...
{
      sL_SQL := 'select EbtModifySerialNo, EbtReqId, EbtReqType, OldEmcCustID, IfMoveToOtherSo, EmcCustID,EbtCatvID,EbtContractNo,EbtCustID,';
      sL_SQL := sL_SQL +  ' to_char(EbtContractBDate,' + ORACLE_STR_SEP + 'YYYY/MM/DD' + ORACLE_STR_SEP + ') EbtContractBDate,';
      sL_SQL := sL_SQL + ' to_char(EbtContractEDate,' + ORACLE_STR_SEP + 'YYYY/MM/DD' + ORACLE_STR_SEP + ') EbtContractEDate,';
      sL_SQL := sL_SQL +  'EbtCustCName,EbtCustContactPhone,EbtCustContactMobile,EbtCompOwnerName,';
      sL_SQL := sL_SQL + ' EbtContactPhone,EbtItContactName,EbtItContactPhone,EbtItContactMobile,EbtItEMail,';
      sL_SQL := sL_SQL +  ' EbtInstAddr,EbtCustAddr,EbtBillAddr,EbtContractStatusCode,EbtContractStatusDesc,';
      sL_SQL := sL_SQL + ' EbtFeePeriodCode,EbtFeePeriodDesc,';
      sL_SQL := sL_SQL + ' EbtServiceType,EbtAgentName,EbtAgentPhone,EbtAgentID,EbtAgentAddress,';
      sL_SQL := sL_SQL +  ' EbtIdCardId,EbtCompanyOwnerId,';
      sL_SQL := sL_SQL +  ' EbtNonProfitCompanyId,EbtCompanyId,ProcessFlag,ErrorCode,ErrorMsg,EbtNotes, CATVVALID,';
      sL_SQL := sL_SQL + ' to_char(UpdTime,' + ORACLE_STR_SEP + 'YYYY/MM/DD' + ORACLE_STR_SEP + ') UpdTime,UpdEn ';
      sL_SQL := sL_SQL + ' from EMC_EBT007 where ProcessFlag=' + ORACLE_STR_SEP + PROCESS_FLAG_1 + ORACLE_STR_SEP;
      L_AdoQry := runEbtSQL(SELECT_MODE,sL_SQL);
}
   aSQL :=
     ' SELECT ROWIDTOCHAR( ROWID ) ROWIDS,     ' +
     '        EBTMODIFYSERIALNO,              ' +
     '        EBTREQID,                       ' +
     '        EBTREQTYPE,                     ' +
     '        OLDEMCCUSTID,                   ' +
     '        IFMOVETOOTHERSO,                ' +
     '        EMCCUSTID,                      ' +
     '        EBTCATVID,                      ' +
     '        EBTCONTRACTNO,                  ' +
     '        EBTCUSTID,                      ' +
     '        TO_CHAR( EBTCONTRACTBDATE,''YYYY/MM/DD'' ) EBTCONTRACTBDATE, ' +
     '        TO_CHAR( EBTCONTRACTEDATE,''YYYY/MM/DD'' ) EBTCONTRACTEDATE, ' +
     '        EBTCUSTCNAME,                   ' +
     '        EBTCUSTCONTACTPHONE,            ' +
     '        EBTCUSTCONTACTMOBILE,           ' +
     '        EBTCOMPOWNERNAME,               ' +
     '        EBTCONTACTPHONE,                ' +
     '        EBTITCONTACTNAME,               ' +
     '        EBTITCONTACTPHONE,              ' +
     '        EBTITCONTACTMOBILE,             ' +
     '        EBTITEMAIL,                     ' +
     '        EBTINSTADDR,                    ' +
     '        EBTCUSTADDR,                    ' +
     '        EBTBILLADDR,                    ' +
     '        EBTCONTRACTSTATUSCODE,          ' +
     '        EBTCONTRACTSTATUSDESC,          ' +
     '        EBTFEEPERIODCODE,               ' +
     '        EBTFEEPERIODDESC,               ' +
     '        EBTSERVICETYPE,                 ' +
     '        EBTAGENTNAME,                   ' +
     '        EBTAGENTPHONE,                  ' +
     '        EBTAGENTID,                     ' +
     '        EBTAGENTADDRESS,                ' +
     '        EBTIDCARDID,                    ' +
     '        EBTCOMPANYOWNERID,              ' +
     '        EBTNONPROFITCOMPANYID,          ' +
     '        EBTCOMPANYID,                   ' +
     '        PROCESSFLAG,                    ' +
     '        ERRORCODE,                      ' +
     '        ERRORMSG,                       ' +
     '        EBTNOTES,                       ' +
     '        CATVVALID,                      ' +
     '        TO_CHAR( UPDTIME,''YYYY/MM/DD'' ) UPDTIME, ' +
     '        UPDEN,                          ' +
     '        TAG                             ' +  
     '   FROM EMC_EBT007                      ' +
     '  WHERE PROCESSFLAG = ''W''             ';

    Result := runEbtSQL( SELECT_MODE, aSQL );

end;

function TdtmMain.getIfLogData: boolean;
var
    bL_Result : boolean;
begin
    bL_Result := false;
    if cdsParam.State <> dsInactive then
      bL_Result := cdsParam.fieldByName('bLogData').AsBoolean;

    result := bL_Result;
end;

function TdtmMain.getIfShowUI: boolean;
var
    bL_Result : boolean;
begin
    bL_Result := false;
    if cdsParam.State <> dsInactive then
      bL_Result := cdsParam.fieldByName('bShowUI').AsBoolean;

    result := bL_Result;
end;

function TdtmMain.getProcessPeriod: Integer;
var
    nL_Result : Integer;
begin
    nL_Result := 0;
    if cdsParam.State <> dsInactive then
      nL_Result := cdsParam.fieldByName('nProcessPeriod').AsInteger;

    result := nL_Result;
end;

function TdtmMain.getSysName: String;
var
    sL_Result : String;
begin
    sL_Result := '';
    if cdsParam.State <> dsInactive then
      sL_Result := cdsParam.fieldByName('sSysName').AsString;

    result := sL_Result;
end;

procedure TdtmMain.processSo151(sI_CompCode:String);
var
    L_AdoQueryForUpdate : TAdoQuery;
    nL_EmcCompCode, nL_EmcCustID : Integer;
    sL_SQL, sL_InsertSql, sL_SeqNo : String;
    sL_EmcNewCustStatusCode,sL_EmcNewCustStatusDesc : String;
    sL_EmcOldCustStatusCode, sL_EmcOldCustStatusDesc, sL_EbtCustID, sL_EbtContractNo, sL_EbtServiceType, sL_UpdEn : STring;
    sL_CombineEmcCustID, sL_UpdateSQL : String;
    L_AdoQry : TADOQuery;
begin
    sL_SQL := ' SELECT * FROM SO151 WHERE PROCESSFLAG = ''W'' ';

    L_AdoQry := runSoSQL( SELECT_MODE, sI_CompCode, sL_SQL );
    with L_AdoQry do
    begin

      while not Eof do
      begin
        frmMain.setSo151Ui( L_AdoQry );
        Application.ProcessMessages;
        sL_SeqNo := FieldByName('SeqNo').asString;
        nL_EmcCompCode := FieldByName('EmcCompCode').asInteger;
        nL_EmcCustID := FieldByName('EmcCustID').asInteger;


        sL_CombineEmcCustID := Format('%.3d', [nL_EmcCompCode]) + Format('%.8d',[nL_EmcCustID]);

        sL_EmcNewCustStatusCode := FieldByName('EmcNewCustStatusCode').asString;
        sL_EmcNewCustStatusDesc := FieldByName('EmcNewCustStatusDesc').asString;
        sL_EmcOldCustStatusCode := FieldByName('EmcOldCustStatusCode').asString;
        sL_EmcOldCustStatusDesc := FieldByName('EmcOldCustStatusDesc').asString;
        sL_EbtCustID := FieldByName('EbtCustID').asString;
        sL_EbtContractNo := FieldByName('EbtContractNo').asString;
        sL_EbtServiceType := FieldByName('EbtServiceType').asString;

        sL_UpdEn := FieldByName('UpdEn').asString;

        sL_InsertSql := ' INSERT INTO EMC_EBT008(EMCCUSTID,EMCNEWCUSTSTATUSCODE, EMCNEWCUSTSTATUSDESC, EMCOLDCUSTSTATUSCODE, EMCOLDCUSTSTATUSDESC, EBTCUSTID, EBTCONTRACTNO, EBTSERVICETYPE, PROCESSFLAG, UPDTIME, UPDEN)';
        sL_InsertSql := sL_InsertSql + ' values(' + ORACLE_STR_SEP + sL_CombineEmcCustID + ORACLE_STR_SEP;
        sL_InsertSql := sL_InsertSql +',' + sL_EmcNewCustStatusCode + ',' + ORACLE_STR_SEP + sL_EmcNewCustStatusDesc + ORACLE_STR_SEP;
        sL_InsertSql := sL_InsertSql +',' + sL_EmcOldCustStatusCode + ',' + ORACLE_STR_SEP + sL_EmcOldCustStatusDesc + ORACLE_STR_SEP;
        sL_InsertSql := sL_InsertSql + ',' + ORACLE_STR_SEP + sL_EbtCustID + ORACLE_STR_SEP ;
        sL_InsertSql := sL_InsertSql + ',' +  ORACLE_STR_SEP + sL_EbtContractNo + ORACLE_STR_SEP ;
        sL_InsertSql := sL_InsertSql + ',' +  ORACLE_STR_SEP + sL_EbtServiceType + ORACLE_STR_SEP ;
        sL_InsertSql := sL_InsertSql + ',' +  ORACLE_STR_SEP + 'W' + ORACLE_STR_SEP ;
        sL_InsertSql := sL_InsertSql + ',' + TUstr.getOracleSQLDateTimeStr(now);
        sL_InsertSql := sL_InsertSql + ',' +  ORACLE_STR_SEP + sL_UpdEn + ORACLE_STR_SEP + ')';

        runEbtSQLForUpdate(IUD_MODE, sL_InsertSql);


          sL_UpdateSQL := 'update SO151 set ProcessFlag=' + ORACLE_STR_SEP + 'F' + ORACLE_STR_SEP ;
          sL_UpdateSQL := sL_UpdateSQL +' where SeqNo=' + sL_SeqNo;
          L_AdoQueryForUpdate := runSoSQLForUpdate(IUD_MODE,sI_CompCode,sL_UpdateSQL);

        Next;
      end;
    end;

end;

procedure TdtmMain.processSoData;
var
  aIndex: Integer;
  aCompCode: String;
  aCanProcs: Boolean;
begin

  for aIndex := 0 to G_CompCodeList.Count-1 do
  begin
    aCompCode := G_CompCodeList.Strings[aIndex];
    aCanProcs := CheckCanProcs( aCompCode );
    if aCanProcs then
    begin
      try
        processSo151( aCompCode );
      except
        on E: Exception do
          frmMain.showErrorOnMemo(
            '處理SO151時發生錯誤!(公司別:' + aCompCode + ') ' +
            'HelpContext:' + IntToStr( E.HelpContext )+ ' Error Message:' + E.Message );
      end;
    end; 
    if aCanProcs then
    begin
      try
        processSo153( aCompCode );
      except
        on E: Exception do
          frmMain.showErrorOnMemo(
            '處理SO153時發生錯誤!(公司別:' + aCompCode + ') ' +
            'HelpContext:' + IntToStr( E.HelpContext)+ ' Error Message:' + E.Message );
      end;
    end;  
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMain.SaveCDS(aCDS: TClientDataSet; aRefFileName: String);
var
 aPath: String;
begin
  aCDS.CheckBrowseMode;
  if ( aCDS.ChangeCount > 0 ) then
  begin
    aPath := IncludeTrailingPathDelimiter(
       ExtractFilePath( Application.ExeName ) ) + aRefFileName;
    aCDS.SaveToFile( aPath );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMain.updateEmcEbt006(sI_EbtContractNo, sI_ErrCode, sI_ErrMsg,
  sI_ProcessFlag: String);
var
  aSQL : String;
  aErrorCode : String;
begin
  if sI_ErrCode = '' then
    aErrorCode := 'null'
  else
    aErrorCode := sI_ErrCode;

  aSQL :=
    ' UPDATE EMC_EBT006     ' +
    '    SET PROCESSFLAG =  ' + QuotedStr( sI_ProcessFlag ) + ',' +
    '        ERRORCODE =    ' + aErrorCode + ',' +
    '        ERRORMSG =     ' + QuotedStr( sI_ErrMsg ) +
    ' WHERE EBTCONTRACTNO = ' + QuotedStr( sI_EbtContractNo );

  RunEbtSQLForUpdate( IUD_MODE, aSQL );
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMain.UpdateEmcEbt007(aContractNo, aEbtModifySerialNo,
  aErrCode, aErrMsg, aProcessFlag: string);
var
  aSQL : String;
begin

  if aErrCode = '' then aErrCode := 'null';

  aSQL := Format(
          ' UPDATE EMC_EBT007                  ' +
          '    SET PROCESSFLAG = ''%s'',       ' +
          '        ERRORCODE = %s,             ' +
          '        ERRORMSG = ''%s''           ' +
          '  WHERE EBTCONTRACTNO = ''%s''      ' +
          '    AND EBTMODIFYSERIALNO = ''%s''  ',
          [aProcessFlag, aErrCode, aErrMsg, aContractNo, aEbtModifySerialNo] );

  RunEbtSQLForUpdate( IUD_MODE, aSQL );
end;

{ ---------------------------------------------------------------------------- }

function TdtmMain.getSoAdoQuery(sI_CompCode: String): TADOQuery;
var
    nL_Ndx : Integer;
    L_AdoQuery : TADOQuery;
begin
    L_AdoQuery := nil;
    nL_Ndx := G_CompCodeList.IndexOf(sI_CompCode);
    if nL_Ndx<>-1 then
    begin
      L_AdoQuery := G_SoAdoQuery[nL_Ndx];
    end;

    result := L_AdoQuery;
end;

procedure TdtmMain.processSo153( sI_CompCode: String);
var
    L_AdoQueryForUpdate : TAdoQuery;
    nL_EmcCompCode, nL_EmcCustID : Integer;
    aSQL, sL_InsertSql, sL_SeqNo : String;
    sL_EmcNewCustStatusCode,sL_EmcNewCustStatusDesc : String;
    sL_EmcOldCustStatusCode, sL_EmcOldCustStatusDesc, sL_EbtCustID, sL_EbtContractNo, sL_EbtServiceType, sL_UpdEn : STring;
    sL_CombineEmcCustID, sL_UpdateSQL : String;

    dL_EbtContractBDate, dL_EbtContractEDate : TDate;
    sL_EbtContractBDate, sL_EbtContractEDate, sL_EbtCustCName : String;
    sL_EbtCustContactPhone, sL_EbtCustContactMobile, sL_EbtCompOwnerName : String;
    sL_EbtContactPhone, sL_EbtItContactName, sL_EbtItContactPhone : String;
    sL_EbtItContactMobile, sL_EbtItEMail, sL_EbtInstAddr, sL_EbtCustAddr : String;

    sL_EbtBillAddr, sL_EbtContractStatusCode,sL_EbtContractStatusDesc : String;
    sL_EbtFeePeriodCode, sL_EbtFeePeriodDesc,sL_EbtAgentName : String;
    sL_EbtAgentPhone, sL_EbtAgentID,sL_EbtAgentAddress, sL_EbtIdCardId : String;
    sL_EbtCompanyOwnerId, sL_EbtNonProfitCompanyId, sL_EbtCompanyId, sL_EbtNotes : String;
    sL_UpdTime,sL_EbtModifySerialNo, sL_ErrorCode, sL_ErrorMsg, sL_Tag: String;

    aQuery, aEbtQuery: TADOQuery;

    aSQL2, sL_EmcNotes, sL_ProcessFlag : String;

    aEbtReqId, aEbtReqType: String;

begin

    aSQL := ' SELECT * FROM SO153                      ' +
            '  WHERE PROCESSFLAG IN ( ''C'', ''E'' )   ' +
            '    AND IFSYNC = ''N''                    ';


    aQuery := RunSoSQL( SELECT_MODE, sI_CompCode, aSQL );

    with aQuery do
    begin
      while not Eof do
      begin
        frmMain.SetSo153Ui( aQuery );
        Application.ProcessMessages;

        sL_SeqNo := FieldByName('SeqNo').asString;

        nL_EmcCompCode := FieldByName('EmcCompCode').asInteger;
        nL_EmcCustID := FieldByName('EmcCustID').asInteger;


        sL_CombineEmcCustID := Format('%.3d', [nL_EmcCompCode]) + Format('%.8d',[nL_EmcCustID]);

        sL_EbtContractNo := FieldByName('EbtContractNo').AsString;

        sL_EbtCustID := FieldByName('EbtCustID').AsString;


        sL_EbtContractBDate := FieldByName('EbtContractBDate').AsString;
        if sL_EbtContractBDate<>'' then
          dL_EbtContractBDate := FieldByName('EbtContractBDate').AsDateTime
        else
          dL_EbtContractBDate := 0;

        sL_EbtContractEDate := FieldByName('EbtContractEDate').AsString;
        if sL_EbtContractEDate<>'' then
          dL_EbtContractEDate := FieldByName('EbtContractEDate').AsDateTime
        else
          dL_EbtContractEDate := 0;



        sL_EbtCustCName := FieldByName('EbtCustCName').AsString;
        sL_EbtCustContactPhone := FieldByName('EbtCustContactPhone').AsString;
        sL_EbtCustContactMobile := FieldByName('EbtCustContactMobile').AsString;
        sL_EbtCompOwnerName := FieldByName('EbtCompOwnerName').AsString;
        sL_EbtContactPhone := FieldByName('EbtContactPhone').AsString;
        sL_EbtItContactName := FieldByName('EbtItContactName').AsString;
        sL_EbtItContactPhone := FieldByName('EbtItContactPhone').AsString;
        sL_EbtItContactMobile := FieldByName('EbtItContactMobile').AsString;
        sL_EbtItEMail := FieldByName('EbtItEMail').AsString;
        sL_EbtInstAddr := FieldByName('EbtInstAddr').AsString;
        sL_EbtCustAddr := FieldByName('EbtCustAddr').AsString;

        sL_EbtBillAddr := FieldByName('EbtBillAddr').AsString;
        sL_EbtContractStatusCode := FieldByName('EbtContractStatusCode').AsString;
        sL_EbtContractStatusDesc := FieldByName('EbtContractStatusDesc').AsString;
        sL_EbtFeePeriodCode := FieldByName('EbtFeePeriodCode').AsString;
        sL_EbtFeePeriodDesc := FieldByName('EbtFeePeriodDesc').AsString;
        sL_EbtServiceType := FieldByName('EbtServiceType').AsString;
        sL_EbtAgentName := FieldByName('EbtAgentName').AsString;
        sL_EbtAgentPhone := FieldByName('EbtAgentPhone').AsString;
        sL_EbtAgentID := FieldByName('EbtAgentID').AsString;
        sL_EbtAgentAddress := FieldByName('EbtAgentAddress').AsString;

        sL_EbtIdCardId := FieldByName('EbtIdCardId').AsString;
        sL_EbtCompanyOwnerId := FieldByName('EbtCompanyOwnerId').AsString;
        sL_EbtNonProfitCompanyId := FieldByName('EbtNonProfitCompanyId').AsString;
        sL_EbtCompanyId := FieldByName('EbtCompanyId').AsString;
        sL_EbtNotes := FieldByName('EbtNotes').AsString;
        sL_UpdEn := FieldByName('UpdEn').AsString;
        sL_UpdTime := FieldByName('UpdTime').AsString;
        sL_EbtModifySerialNo := FieldByName('EbtModifySerialNo').AsString;
        sL_ErrorCode := FieldByName('ErrorCode').AsString;
        if sL_ErrorCode='' then
          sL_ErrorCode := 'null';
        sL_ErrorMsg := FieldByName('ErrorMsg').AsString;
        sL_EmcNotes := FieldByName('EmcNotes').AsString;
        sL_ProcessFlag := FieldByName('ProcessFlag').AsString;
        sL_Tag := FieldByName('TAG').AsString;

        if sL_EbtModifySerialNo<>'' then
        begin

          sL_UpdateSQL := 'update EMC_EBT007 set EmcCustID=' + ORACLE_STR_SEP + sL_CombineEmcCustID + ORACLE_STR_SEP;
          sL_UpdateSQL := sL_UpdateSQL + ',ProcessFlag=' + ORACLE_STR_SEP + sL_ProcessFlag + ORACLE_STR_SEP;
          sL_UpdateSQL := sL_UpdateSQL + ',ErrorCode=' + sL_ErrorCode + ',ErrorMsg=' + ORACLE_STR_SEP + sL_ErrorMsg + ORACLE_STR_SEP;
          sL_UpdateSQL := sL_UpdateSQL + ',UpdTime=' + TUstr.getOracleSQLDateTimeStr(now);
          sL_UpdateSQL := sL_UpdateSQL + ',UpdEn='+ ORACLE_STR_SEP + sL_UpdEn + ORACLE_STR_SEP;
          sL_UpdateSQL := sL_UpdateSQL + ',EmcNotes=' + ORACLE_STR_SEP + sL_EmcNotes + ORACLE_STR_SEP;
          sL_UpdateSQL := sL_UpdateSQL + ' where EbtModifySerialNo='+ ORACLE_STR_SEP + sL_EbtModifySerialNo + ORACLE_STR_SEP;
        end
        else
        begin
          sL_UpdateSQL := 'update EMC_EBT006 set EmcCustID=' + ORACLE_STR_SEP + sL_CombineEmcCustID + ORACLE_STR_SEP;
          sL_UpdateSQL := sL_UpdateSQL + ',ProcessFlag=' + ORACLE_STR_SEP + sL_ProcessFlag + ORACLE_STR_SEP;
          sL_UpdateSQL := sL_UpdateSQL + ',ErrorCode=' + sL_ErrorCode + ',ErrorMsg=' + ORACLE_STR_SEP + sL_ErrorMsg + ORACLE_STR_SEP;
          sL_UpdateSQL := sL_UpdateSQL + ',UpdTime=' + TUstr.getOracleSQLDateTimeStr(now);
          sL_UpdateSQL := sL_UpdateSQL + ',UpdEn='+ ORACLE_STR_SEP + sL_UpdEn + ORACLE_STR_SEP;
          sL_UpdateSQL := sL_UpdateSQL + ',EmcNotes=' + ORACLE_STR_SEP + sL_EmcNotes + ORACLE_STR_SEP;
          sL_UpdateSQL := sL_UpdateSQL + ' where EbtContractNo='+ ORACLE_STR_SEP + sL_EbtContractNo + ORACLE_STR_SEP;
        end;

        RunEbtSQLForUpdate(IUD_MODE, sL_UpdateSQL);
        L_AdoQueryForUpdate := runSoSQLForUpdate(IUD_MODE,sI_CompCode, aSQL );

        //down, 將已經賦予客編之資料insert 到 SO152
        if UpperCase(sL_ProcessFlag) = 'C' then
        begin
          if not hasSo152Data(IntToStr(nL_EmcCompCode),sL_EbtContractNo) then
          begin
            aEbtReqId := '';
            aEbtReqType := '';
            if sL_EbtModifySerialNo <> '' then
            begin
              aSQL2 :=
                Format( ' SELECT EBTREQID, EBTREQTYPE        ' +
                        '   FROM EMC_EBT007                  ' +
                        '  WHERE EBTMODIFYSERIALNO = ''%s''  ', [sL_EbtModifySerialNo] );
              aEbtQuery := runEbtSQL( SELECT_MODE, aSQL2 );
              aEbtReqId := aEbtQuery.FieldByName( 'EBTREQID' ).AsString;
              aEbtReqType := aEbtQuery.FieldByName( 'EBTREQTYPE' ).AsString;
              aEbtQuery.Close;
            end;
            sL_InsertSQL := 'insert into SO152(EMCCOMPCODE,EMCCUSTID,EBTCONTRACTNO,EBTCUSTID,EBTCONTRACTBDATE,EBTCONTRACTEDATE,EBTCUSTCNAME,EBTCUSTCONTACTPHONE,';
            sL_InsertSQL := sL_InsertSQL +'EBTCUSTCONTACTMOBILE,EBTCOMPOWNERNAME,EBTCONTACTPHONE,EBTITCONTACTNAME,EBTITCONTACTPHONE,EBTITCONTACTMOBILE,';
            sL_InsertSQL := sL_InsertSQL +'EBTITEMAIL,EBTINSTADDR,EBTCUSTADDR,EBTBILLADDR,EBTCONTRACTSTATUSCODE,EBTCONTRACTSTATUSDESC,';
            sL_InsertSQL := sL_InsertSQL +'EBTFEEPERIODCODE,EBTFEEPERIODDESC,EBTSERVICETYPE,EBTAGENTNAME,EBTAGENTPHONE,EBTAGENTID,EBTAGENTADDRESS,';
            sL_InsertSQL := sL_InsertSQL +'EBTIDCARDID,EBTCOMPANYOWNERID,EBTNONPROFITCOMPANYID,EBTCOMPANYID,EBTNOTES,UPDEN, UPDTIME, EBTREQID, EBTREQTYPE, TAG ) values(';
            sL_InsertSQL := sL_InsertSQL + IntToStr(nL_EmcCompCode) + ',' +  IntToStr(nL_EmcCustID) + ',' + ORACLE_STR_SEP + sL_EbtContractNo + ORACLE_STR_SEP ;
            sL_InsertSQL := sL_InsertSQL + ',' + ORACLE_STR_SEP + sL_EbtCustID + ORACLE_STR_SEP +',' + TUstr.getOracleSQLDateStr(dL_EbtContractBDate);
            sL_InsertSQL := sL_InsertSQL + ',' + TUstr.getOracleSQLDateStr(dL_EbtContractEDate) + ',' +  ORACLE_STR_SEP + sL_EbtCustCName + ORACLE_STR_SEP;
            sL_InsertSQL := sL_InsertSQL + ',' + ORACLE_STR_SEP + sL_EbtCustContactPhone + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtCustContactMobile + ORACLE_STR_SEP;
            sL_InsertSQL := sL_InsertSQL + ',' + ORACLE_STR_SEP + sL_EbtCompOwnerName + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtContactPhone + ORACLE_STR_SEP;
            sL_InsertSQL := sL_InsertSQL + ',' + ORACLE_STR_SEP + sL_EbtItContactName + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtItContactPhone + ORACLE_STR_SEP;
            sL_InsertSQL := sL_InsertSQL + ',' + ORACLE_STR_SEP + sL_EbtItContactMobile + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtItEMail + ORACLE_STR_SEP;
            sL_InsertSQL := sL_InsertSQL + ',' + ORACLE_STR_SEP + sL_EbtInstAddr + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtCustAddr + ORACLE_STR_SEP;

            sL_InsertSQL := sL_InsertSQL + ',' + ORACLE_STR_SEP + sL_EbtBillAddr + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtContractStatusCode + ORACLE_STR_SEP;
            sL_InsertSQL := sL_InsertSQL + ',' + ORACLE_STR_SEP + sL_EbtContractStatusDesc + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtFeePeriodCode + ORACLE_STR_SEP;
            sL_InsertSQL := sL_InsertSQL + ',' + ORACLE_STR_SEP + sL_EbtFeePeriodDesc + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtServiceType + ORACLE_STR_SEP;
            sL_InsertSQL := sL_InsertSQL + ',' + ORACLE_STR_SEP + sL_EbtAgentName + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtAgentPhone + ORACLE_STR_SEP;
            sL_InsertSQL := sL_InsertSQL + ',' + ORACLE_STR_SEP + sL_EbtAgentID + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtAgentAddress + ORACLE_STR_SEP;
            sL_InsertSQL := sL_InsertSQL + ',' + ORACLE_STR_SEP + sL_EbtIdCardId + ORACLE_STR_SEP;
            sL_InsertSQL := sL_InsertSQL + ',' + ORACLE_STR_SEP + sL_EbtCompanyOwnerId + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtNonProfitCompanyId + ORACLE_STR_SEP;

            sL_InsertSQL := sL_InsertSQL + ',' + ORACLE_STR_SEP + sL_EbtCompanyId + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + sL_EbtNotes + ORACLE_STR_SEP;
            sL_InsertSQL := sL_InsertSQL + ',' + ORACLE_STR_SEP + sL_UpdEn + ORACLE_STR_SEP + ',' + TUstr.getOracleSQLDateTimeStr(now); 
            sL_InsertSQL := sL_InsertSQL + ',' + ORACLE_STR_SEP + aEbtReqId + ORACLE_STR_SEP + ',' + ORACLE_STR_SEP + aEbtReqType + ORACLE_STR_SEP + ','  + ORACLE_STR_SEP + sL_Tag + ORACLE_STR_SEP + ')';            

            L_AdoQueryForUpdate := runSoSQLForUpdate(IUD_MODE,sI_CompCode,sL_InsertSQL);

          end;
          { 吊牌號碼處理 }
          if sL_Tag <> EmptyStr then
            ProcessSo004( sI_CompCode, IntToStr( nL_EmcCustID ), sL_EbtContractNo, sL_EbtModifySerialNo, sL_Tag );
        end;


        sL_UpdateSQL := ' UPDATE SO153          ' +
                        '    SET IFSYNC = ''Y'' ' +
                        '  WHERE SEQNO = ' + sL_SeqNo;

        L_AdoQueryForUpdate := RunSoSQLForUpdate( IUD_MODE, sI_CompCode,
          sL_UpdateSQL );


        Next;
      end;
    end;

end;

function TdtmMain.getIfAcceptNullData: boolean;
var
    bL_Result : boolean;
begin
    bL_Result := false;
    if cdsParam.State <> dsInactive then
      bL_Result := cdsParam.fieldByName('bAcceptNullData').AsBoolean;

    result := bL_Result;
end;

procedure TdtmMain.setTirggerStatus;
var
    bL_EnableTrigger : boolean;
    sL_SQL, sL_EnableKeyWord : STring;
begin
    bL_EnableTrigger := cdsParam.FieldByName('bLogData').AsBoolean;

    if bL_EnableTrigger then
      sL_EnableKeyWord := 'enable'
    else
      sL_EnableKeyWord := 'disable';
    sL_SQL := 'alter trigger TR_EMCEBT006_INSERT ' + sL_EnableKeyWord;
    runEbtSQL(IUD_MODE, sL_SQL);
{
    G_AdoQueryForEbtData.SQL.Clear;
    G_AdoQueryForEbtData.SQL.Add(sL_SQL);
    G_AdoQueryForEbtData.ExecSQL;
}    
end;

function TdtmMain.runEbtSQL(nI_Mode: Integer; sI_SQL: String): TADOQuery;
var
  L_AdoQuery : TADOQuery;
begin
  L_AdoQuery := G_AdoQueryForEbtData;
  if L_AdoQuery<>nil then
  begin
    L_AdoQuery.SQL.Clear;
    L_AdoQuery.SQL.Add(sI_SQL);
    case nI_Mode of
      SELECT_MODE://select
        L_AdoQuery.Open;
      IUD_MODE://DML
        L_AdoQuery.ExecSQL;
    end;
  end;
  result := L_AdoQuery;

end;

function TdtmMain.runSoSQL(nI_Mode: Integer; sI_CompCode,
  sI_SQL: String): TADOQuery;
var
  aQuery : TADOQuery;
  aSQL: String;
begin
  aQuery := getSoAdoQuery(sI_CompCode);
  if Assigned( aQuery ) then
  begin
    aSQL := TrimInvalidateChar( sI_SQL );
    aQuery.SQL.Text := aSQL;
    case nI_Mode of
      SELECT_MODE: aQuery.Open;
      IUD_MODE: aQuery.ExecSQL;
    end;
  end;
  Result := aQuery;
end;

function TdtmMain.runSoSQLForUpdate(nI_Mode: Integer; sI_CompCode,
  sI_SQL: String): TADOQuery;
var
  aQuery : TADOQuery;
  aSQL: String;
begin
  aQuery := getSoAdoQueryForUpdate( sI_CompCode );
  if Assigned( aQuery ) then
  begin
    aSQL := TrimInvalidateChar( sI_SQL );
    aQuery.SQL.Text := aSQL;
    case nI_Mode of
      SELECT_MODE: aQuery.Open;
      IUD_MODE: aQuery.ExecSQL;
    end;
  end;
  Result := aQuery;
end;

function TdtmMain.runEbtSQLForUpdate(nI_Mode: Integer; sI_SQL: String): TADOQuery;
var
  aQuery : TADOQuery;
  aSQL: String;
begin
  aQuery := G_AdoQueryForEbtDataUpdate;
  if Assigned( aQuery ) then
  begin
    aSQL := TrimInvalidateChar( sI_SQL );
    aQuery.SQL.Text := aSQL;
    case nI_Mode of
      SELECT_MODE: aQuery.Open;
      IUD_MODE: aQuery.ExecSQL;
    end;
  end;
  Result := aQuery;
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

    L_StrList.LoadFromFile(sL_ExePath + '\' + CSIS_INI_FILE_NAME);
    for ii:=0 to L_StrList.Count-1 do
    begin
       if (Copy(L_StrList.Strings[ii],1,2)<>'//') then
         L_TmpStrList.Add(L_Intf.Decrypt(L_StrList.Strings[ii], sL_EncKey))
       else
         L_TmpStrList.Add(L_StrList.Strings[ii]);
    end;
    L_TmpStrList.SaveToFile(sL_ExePath + '\' + TMP_CSIS_INI_FILE_NAME);
    L_TmpStrList.Free;

    L_Intf := nil;


end;

    function TdtmMain.hasSo152Data(sI_CompCode, sI_EbtContractNo: String): boolean;
var
    bL_Result : boolean;
    L_AdoQuery : TAdoQuery;
    sL_SQL : String;
begin
  bL_Result := true;
  sL_SQL := 'select count(*) RECCOUNT from SO152 where EBTCONTRACTNO=' + ORACLE_STR_SEP + sI_EbtContractNo + ORACLE_STR_SEP;
  L_AdoQuery := runSoSQLForUpdate(SELECT_MODE,sI_CompCode,sL_SQL);
  if (L_AdoQuery<>nil) then
  begin
    if (L_AdoQuery.FieldByName('RECCOUNT').AsInteger=0) then
      bL_Result := false
    else
      bL_Result := true;

  end;

  result := bL_Result;

end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMain.GetCompanyCodeToStringList;
var
  AFileName: String;
  AIniObj: TIniFile;
  ADBCount, AIndex: Integer;
begin
  { 先將加密過的 IniFile 解來後, 存成暫存檔 }
  TransTmpIniFile;
  try
   G_CompCodeList.Clear;
    AFileName := IncludeTrailingPathDelimiter( ExtractFilePath(
      Application.ExeName ) ) + TMP_CSIS_INI_FILE_NAME;
    if not FileExists( AFileName ) then
    begin
      Application.MessageBox(
        PChar( '讀取資料庫連線參數檔<' + AFileName + '>失敗' ),
        PChar( '錯誤' ), MB_OK + MB_ICONERROR );
      Exit;
    end;
    AIniObj := TIniFile.Create( AFileName );
    try
      ADbCount := AIniObj.ReadInteger( 'DBINFO', 'DB_COUNT', 0 );
      for AIndex := 1 to ADBCount do
      begin
        G_CompCodeList.Add( AIniObj.ReadString( 'DBINFO','COMPCODE_' +
          IntToStr( AIndex ), '' ) );
      end;
    finally
      AIniObj.Free;
    end;
  finally
    DeleteFile( AFileName );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMain.ReCreateUIDataSet;
begin
  EBT007_UIDataSet.Data := Null;
  EBT007_UIDataSet.Fields.Clear;
  EBT007_UIDataSet.FieldDefs.Clear;
  EBT007_UIDataSet.FieldDefs.Add( 'ROWID', ftString, 18 );
  EBT007_UIDataSet.FieldDefs.Add( 'SEQ', ftString, 10 );
  EBT007_UIDataSet.CreateDataSet;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMain.GetMappingUISeq(const aDataSet: TClientDataSet;
  const aRowId: string ): string;
begin
  Result := '';
  if aDataSet.Locate( 'ROWID', aRowId, [] ) then
  begin
    Result := aDataSet.FieldByName( 'SEQ' ).AsString;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMain.TrimInvalidateChar(const aSource: String): String;
var
  aOut: String;
  aIndex: Integer;
begin
  aOut := aSource;
  while AnsiPos( '`', aOut ) > 0 do
  begin
    aIndex := AnsiPos( '`', aOut );
    Delete( aOut, aIndex, 1 );
  end;
  Result := aOut;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMain.ProcessSo004(aCompCode, aCustId, aEbtConstractNo,
  aEbtModifySerialNo, aTag: String);
var
  aSQL2, aInstDate, aUpdEn, aPrefix: String;
  aEbtQuery: TADOQuery;
  aMustUpdate: Boolean;
  aCount: Integer;
begin
  { 處理吊牌號碼到 SO004 }
  { 原則是 1. 用 EBT 的合約號碼先去查 So004否已有該合約編號 }
  {        2. 若沒有該合約編號, 表示沒有舊的吊牌號碼 }
  {        3. 若有合約編號, 則必須先把舊的合約編號(含該筆)標成拆除 }
  {        4. 然後寫一筆新的進去 }
  if aTag = EmptyStr then Exit;
  if ( aEbtModifySerialNo = EmptyStr ) then
  begin
     aSQL2 :=
       Format( ' SELECT TO_CHAR( UPDTIME, ''YYYYMMDD HH24MISS'' ) UPDTIME, ' +
               '        UPDEN,                                             ' +
               '        TO_CHAR( SYSDATE, ''YYYYMMDD'' ) PREFIX            ' +
               '        FROM EMC_EBT006                                    ' +
               '  WHERE EBTCONTRACTNO = ''%s''  ', [aEbtConstractNo] );
     aEbtQuery := runEbtSQL( SELECT_MODE, aSQL2 );
     aInstDate := aEbtQuery.FieldByName( 'UPDTIME' ).AsString;
     aUpdEn := aEbtQuery.FieldByName( 'UPDEN' ).AsString;
     aPrefix := aEbtQuery.FieldByName( 'PREFIX' ).AsString;
  end else
  begin
     aSQL2 :=
       Format( ' SELECT TO_CHAR( UPDTIME, ''YYYYMMDD HH24MISS'' ) UPDTIME,  ' +
               '        UPDEN,                                              ' +
               '        TO_CHAR( SYSDATE, ''YYYYMMDD'' ) PREFIX             ' +
               '   FROM EMC_EBT007                                          ' +
               '  WHERE EBTCONTRACTNO = ''%s''                              ' +
               '    AND EBTMODIFYSERIALNO = ''%s''  ', [aEbtConstractNo, aEbtModifySerialNo] );
     aEbtQuery := runEbtSQL( SELECT_MODE, aSQL2 );
     aInstDate := aEbtQuery.FieldByName( 'UPDTIME' ).AsString;
     aUpdEn := aEbtQuery.FieldByName( 'UPDEN' ).AsString;
     aPrefix := aEbtQuery.FieldByName( 'PREFIX' ).AsString;
  end;
  aEbtQuery.Close;
  aMustUpdate := False;
  { 先取原 Table 的 UpdTime, 做為安裝日期 }
  { 先看此合約編號及吊牌號碼是否存在 SO004 }
  aSQL2 := Format(
    ' SELECT * FROM SO004               ' +
    '  WHERE EBTCONTNO = ''%s''         ',
    [aEbtConstractNo, aTag] );
  aEbtQuery := RunSoSQLForUpdate( SELECT_MODE, aCompCode, aSQL2 );
  aEbtQuery.First;
  while not aEbtQuery.Eof do
  begin
    { 將舊的標成拆除 }
    if ( aEbtQuery.FieldByName( 'FACISNO' ).AsString <> aTag ) then
    begin
      aMustUpdate := True;
    end else
    begin
      aMustUpdate := ( aEbtQuery.FieldByName( 'CUSTID' ).AsString <> aCustId );
    end;
    if aMustUpdate then
    begin
      aSQL2 := Format(
        ' UPDATE SO004                     ' +
        '    SET PRDATE = TO_DATE( ''%s'', ''YYYYMMDD HH24MISS'' ), ' +
        '        UPDEN = ''%s''            ' +
        '  WHERE SEQNO = ''%s''            ' +
        '    AND PRDATE IS NULL            ',
        [aInstDate, aUpdEn,
         aEbtQuery.FieldByName( 'SEQNO' ).AsString] );
      RunSoSQLForUpdate( IUD_MODE, aCompCode, aSQL2 );
    end;
    aEbtQuery.Next;
  end;
  aEbtQuery.Close;
  { 再檢查一次 SO004, 確定此 Tag 號碼是否可以新增進去 }
  aSQL2 := Format(
      ' SELECT COUNT(1) COUNTS FROM SO004   ' +
      '  WHERE EBTCONTNO = ''%s''           ' +
      '    AND PRDATE IS NULL               ',
      [aEbtConstractNo] );

  aEbtQuery := RunSoSQLForUpdate( SELECT_MODE, aCompCode, aSQL2 );
  aCount := aEbtQuery.FieldByName( 'COUNTS' ).AsInteger;
  aEbtQuery.Close;
  if aCount <= 0 then
  begin
    aSQL2 := Format(
      ' INSERT INTO SO004 ( CUSTID, FACICODE, FACINAME, FACISNO, BUYCODE, BUYNAME,    ' +
      '    INSTDATE, PRDATE, UPDTIME, UPDEN, SERVICETYPE, COMPCODE, EBTCONTNO, SEQNO  ) ' +
      ' VALUES ( ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                      ' +
      '    TO_DATE( ''%s'', ''YYYYMMDD HH24MISS'' ), NULL, '  +
      '    TO_DATE( ''%s'', ''YYYYMMDD HH24MISS'' ), ''%s'', ''C'', ''%s'', ''%s'', ''%s'' || S_SO004.NEXTVAL ) ',
      [aCustId, '347', '設備-CM吊牌', aTag, '1', '買斷',
       aInstDate, aInstDate, aUpdEn, aCompCode, aEbtConstractNo, aPrefix ] );
    RunSoSQLForUpdate( IUD_MODE, aCompCode, aSQL2 );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMain.CheckCanProcs(aCompCode: String): Boolean;
begin
 if aCompCode = EmptyStr then
 begin
   Result := False;
   Exit;
 end;
 if ( StrToInt( aCompCode ) = 5 ) or ( StrToInt( aCompCode ) = 6 ) or
    ( StrToInt( aCompCode ) = 7 ) or ( StrToInt( aCompCode ) = 16 ) then
    Result := False
  else
    Result := True;
end;

{ ---------------------------------------------------------------------------- }

end.
