unit dtmMainU;

interface

uses
  SysUtils, Classes, DB, DBClient,Forms,IniFiles,Dialogs,Variants ,ADODB,
  Provider;

type
  TSGSParamIniData = class(TObject)
    nDataAreaNo : Integer;
    sDataArea : String;
    sAlias : String;
    sUserID : String;
    sPassword : String;
    sCompCode : String;
    sCompName : String;
    sDataPath1 : String;
    sDataPath2 : String;
    sFirstDate : String;
    sAdministratorMail1 : String;
    sAdministratorMail2 : String;
    sVersion : String;
    sSrcCode : String;
    sSrcIP : String;
    sSrcType : String;
    sDstCode : String;
    sDstIP : String;
    sDstType : String;
    sCasID : String;
  end;

  TdtmMain = class(TDataModule)
    cdsUser: TClientDataSet;
    cdsUsersID: TStringField;
    cdsUsersName: TStringField;
    cdsUsersPassword: TStringField;
    cdsSGSParam: TClientDataSet;
    cdsSGSParamGroup: TStringField;
    cdsSGSParamDataAreaNo: TIntegerField;
    cdsSGSParamDataArea: TStringField;
    cdsSGSParamAlias: TStringField;
    cdsSGSParamUserID: TStringField;
    cdsSGSParamPassword: TStringField;
    cdsSGSParamCompCode: TStringField;
    cdsSGSParamCompName: TStringField;
    cdsSGSParamDataPath1: TStringField;
    cdsSGSParamDataPath2: TStringField;
    cdsSGSParamFirstDate: TStringField;
    cdsSGSParamAdministratorMail1: TStringField;
    cdsSGSParamAdministratorMail2: TStringField;
    cdsSGSParamVersion: TStringField;
    cdsSGSParamCasID: TStringField;
    cdsSGSParamSrcCode: TStringField;
    cdsSGSParamSrcIP: TStringField;
    cdsSGSParamSrcType: TStringField;
    cdsSGSParamDstCode: TStringField;
    cdsSGSParamDstIP: TStringField;
    cdsSGSParamDstType: TStringField;
    adoQryQueryData: TADOQuery;
    adoQryNightRunLog: TADOQuery;
    adoQryQueryLogData: TADOQuery;
    cdsQueryLogData: TClientDataSet;
    cdsQueryLogDataDstMsgID: TStringField;
    cdsQueryLogDataModeType: TStringField;
    cdsQueryLogDataInstQuery: TStringField;
    cdsQueryLogDataDstCode: TStringField;
    cdsQueryLogDataSrcCode: TStringField;
    cdsQueryLogDataHandleEDateTime: TStringField;
    cdsQueryLogDataExecDateTime: TStringField;
    cdsQueryLogDataSrcMsgID: TStringField;
    cdsQueryLogDataErrorCode: TStringField;
    cdsQueryLogDataErrorMsg: TStringField;
    cdsChannelID: TClientDataSet;
    cdsChannelIDProductCode: TStringField;
    cdsChannelIDProvider: TStringField;
    cdsChannelIDChannelName: TStringField;
    adoQryLoadExecCounts: TADOQuery;
    procedure DataModuleCreate(Sender: TObject);
    procedure DataModuleDestroy(Sender: TObject);
  private
    { Private declarations }
    G_AdoQueryAry : array of TADOQuery;
    G_AdoStoredProcAry1 : array of TADOStoredProc; //SF_ICCardQuery
    G_AdoStoredProcAry2 : array of TADOStoredProc; //SF_CAProductQuery
    G_AdoStoredProcAry3 : array of TADOStoredProc; //SF_ProdPurchaseQuery
    G_AdoStoredProcAry4 : array of TADOStoredProc; //SF_EntitlementQuery
    G_AdoStoredProcAry5 : array of TADOStoredProc; //SP_InsertSO453.sql
    function connectToDB(sI_CompCode : String) : WideString;

  public
    { Public declarations }
    G_DbConnAry : array of TADOConnection;
    G_UserIDStrList : TStringList;
    sG_ExePath,sG_TcOrSc : String;
    G_LanguageSettingIni : TIniFile;
    G_SGSParamIniStrList : TStringList;
    nG_NightRunExecTime,nG_HttpPort : Integer;
    sG_CancelChannelStr,SG_WinZipExePath : String;
    G_CompCodeAndNameStrList : TStringList;

    function getExePath : String;
    procedure activeCDS;
    function checkUserIdIsOnly(sI_UserID : String) : Boolean;
    function getAllUserID : TStringList;

    function initLanguageIniFile : Boolean;
    function getLanguageSettingInfo(sI_Key:String):String;

    procedure TransTmpIniFile;
    procedure DeleteTmpIniFile;
    function getSGSParamIniInfo : WideString;

    procedure initialcdsSGSParamSet(I_StrList : TStringList);


    procedure connectAllCompanyDB;
    function connectToDBAgain(sI_CompCode : String) : String;
    function getCompName(sI_CompCode : String) : String;
    function getFirstDate(sI_CompCode : String) : String;
    function getSrcIP(sI_CompCode : String) : String;
    function getDstIP(sI_CompCode : String) : String;
    function getDataPath(sI_CompCode : String;var sI_DataPath1,sI_DataPath2 : String) : String;
    function getCompCode(sI_SrcCode : String;var nI_Index : Integer) : String;
    function getActiveQuery(sI_CompCode : String) : TADOQuery;
    function getAreaDataNo(sI_CompCode : String) : Integer;
    function runSQL(nI_Mode:Integer;sI_CompCode,sI_SQL : String) : TADOQuery;

    function checkIsConnection(sI_CompCode : String) : Boolean;

    procedure saveToFile(vI_Content:OleVariant; sI_FileName:String);

    function runICCardQuerySF(sI_CompCode,sI_SrcCode,sI_DstCode,sI_DstMsgID,
             sI_CasID,sI_InstQuery,sI_FirstDate,sI_HandleEDateTime : String
             ;var sI_RetCode,sI_Msg,sI_SrcMsgID,sI_FirstFlag : String;var nI_TotalExecCounts,nI_TotalICCardNum,nI_TotalSubNum : Integer) : String;

    function runCAProductQuerySF(sI_CompCode,sI_SrcCode,sI_DstCode,sI_DstMsgID,
             sI_CasID,sI_InstQuery,sI_FirstDate,sI_HandleEDateTime : String
             ;var sI_RetCode,sI_Msg,sI_SrcMsgID,sL_FirstFlag : String;var nI_TotalExecCounts,nI_TotalProductNum,nI_TotalChannelNum : Integer) : String;

    function runProdPurchaseQuerySF(sI_CompCode,sI_SrcCode,sI_DstCode,sI_DstMsgID,
             sI_CasID,sI_InstQuery,sI_FirstDate,sI_HandleEDateTime,sI_CancelChannelStr : String
             ;var sI_RetCode,sI_Msg,sI_SrcMsgID,sI_FirstFlag : String;var nI_TotalExecCounts : Integer) : String;


    function runEntitlementQuerySF(sI_CompCode,sI_SrcCode,sI_DstCode,sI_DstMsgID,
             sI_CasID,sI_InstQuery,sI_FirstDate,sI_HandleEDateTime,sI_CancelChannelStr : String
             ;var sI_RetCode,sI_Msg,sI_SrcMsgID,sI_FirstFlag : String;var nI_TotalExecCounts : Integer) : String;

    function runTestDataInsertSO453SP(sI_CompCode : String) : String;


    function getSrcMsgID(sI_CompCode,sI_ActionMsg : String) : String;
    function getQueryData(sI_QueryType,sI_Table,sI_CompCode,sI_HandleEDateTime : String) : TADOQuery;

    function getICCardNumAndSubNum(bI_IsTestData : Boolean;sI_CompCode,sI_HandleEDateTime : String;var nI_TotalICCardNum,nI_TotalSubNum : Integer) : String;
    function getProductNumAndChannelNum(sI_CompCode,sI_HandleEDateTime : String;var nI_TotalProductNum,nI_TotalChannelNum : Integer) : String;

    function getFileSize(sI_XmlFilePath : String) : Double;
    function getRealFilePath(sI_CompCode,sI_InstQuery,sI_QueryType,sI_HandleEDateTime : String;bI_IsZipFile : Boolean) : String;
    function getFileName(sI_HandleEDateTime,sI_QueryType : String;bI_IsZipFile : Boolean) : String;
    function logToSO460(sI_CompCode,sI_ModeType,sI_HandleEDateTime,sI_InstQuery,sI_SrcCode,sI_SrcMsgID,sI_DstCode,sI_DstMsgID,sI_RetCode,sI_Msg : String) : String;
    function logToSO461(sI_CompCode,sI_ModeType,sI_HandleEDateTime,sI_QueryDateTime,sI_InstQuery,sI_SrcCode,sI_DstCode,sI_DstMsgID,sI_RequestXml : String) : String;
    function logToSO459(sI_CompCode,sI_ModeType,sI_FirstFlag,sI_HandleEDateTime,sI_RetCode,sI_Msg,sI_TotalICCardNum,sI_TotalSubNum,sI_TotalProductNum,sI_TotalChannelNum : String) : String;
    procedure getSFRealMsg(sI_SrcRetCode,sI_SrcMsg : String;var sI_RealRetCode,sI_RealMsg : String);

    function getNightRunLogData(sI_CompCode,sI_ModeType,sI_SDateTime,sI_EDateTime : String) : TADOQuery;
    function getQueryLogData(sI_CompCode,sI_ModeType,sI_SDateTime,sI_EDateTime : String) : TADOQuery;
    function getSMSBasicData(sI_CompCode : String;var sI_STBFaciCode,sI_STBFaciName,sI_ICCFaciCode,sI_ICCFaciName,sI_BuyCode,sI_BuyName : String;var I_CustStrLis : TStringList) : String;
    function getSMSSO005TransCustID(sI_CompCode,sI_ICCardNo : String;var sI_CustId : String) : String;
    function getYesterdayTotalICCardNum(sI_Table,sI_CompCode,sI_StartDate : String) : Integer;
  end;

var
  dtmMain: TdtmMain;

implementation

uses ConstU, DevPower_TLB, Ustru, frmMainU;

{$R *.dfm}

{ TdtmMain }

function TdtmMain.runICCardQuerySF(sI_CompCode, sI_SrcCode, sI_DstCode,
  sI_DstMsgID, sI_CasID, sI_InstQuery, sI_FirstDate,
  sI_HandleEDateTime: String;var sI_RetCode,sI_Msg,sI_SrcMsgID,sI_FirstFlag : String;
  var nI_TotalExecCounts,nI_TotalICCardNum,nI_TotalSubNum : Integer): String;
var
    nL_ActiveConn : Integer;
begin
  try
    nL_ActiveConn := getAreaDataNo(sI_CompCode);

    with G_AdoStoredProcAry1[nL_ActiveConn] do
    begin
      Parameters.ParamByName('p_CompCode').Value := sI_CompCode;
      Parameters.ParamByName('P_SrcCode').Value := sI_SrcCode;
      Parameters.ParamByName('P_DstCode').Value := sI_DstCode;
      Parameters.ParamByName('P_DstMsgID').Value := sI_DstMsgID;

      Parameters.ParamByName('P_CasID').Value := sI_CasID;
      Parameters.ParamByName('P_InstQuery').Value := sI_InstQuery;
      Parameters.ParamByName('P_FirstDate').Value := sI_FirstDate;
      Parameters.ParamByName('P_HandleEDateTime').Value := sI_HandleEDateTime;

      ExecProc;

      //sL_FunReturn := Parameters.Items[0].Value; //用ParamByName會有問題
      //sL_RetCode := Parameters.Items[1].Value; //用ParamByName會有問題
      //sL_Msg := Parameters.Items[2].Value; //用ParamByName會有問題

      sI_RetCode := Parameters.ParamByName('p_ErrCode').Value;
      sI_Msg := Parameters.ParamByName('p_ErrMsg').Value;
      sI_SrcMsgID := Parameters.ParamByName('p_SrcMsgID').Value;
      nI_TotalExecCounts := Parameters.ParamByName('p_TotalExecCounts').Value;
      nI_TotalICCardNum := Parameters.ParamByName('p_TotalICCardNum').Value;
      nI_TotalSubNum := Parameters.ParamByName('p_TotalSubNum').Value;
      sI_FirstFlag := Parameters.ParamByName('p_FirstFlag').Value;
    end;

  except

  end;
end;

procedure TdtmMain.activeCDS;
var
    sL_FileName, sL_ExePath : String;

begin

    sL_ExePath := getExePath;

    sL_FileName := sL_ExePath + '\' + USER_INFO_FILENAME;
    cdsUser.LoadFromFile(sL_FileName);
    if not cdsUser.Active then
      cdsUser.CreateDataSet;

{
    sL_FileName := sL_ExePath + '\' + DISPATCHER_SYS_PARAM_FILENAME;
    cdsParam.LoadFromFile(sL_FileName);
    if not cdsParam.Active then
      cdsParam.CreateDataSet;

    if not cdsDbLinkSet.Active then
      cdsDbLinkSet.CreateDataSet; 
}
end;

function TdtmMain.checkIsConnection(sI_CompCode: String): Boolean;
var
    nL_DataAreaNo : Integer;
begin
    //檢核是否已連上該公司資料庫
    nL_DataAreaNo := getAreaDataNo(sI_CompCode);

    if not G_DbConnAry[nL_DataAreaNo].Connected then
      Result := false
    else
      Result := true;
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

procedure TdtmMain.connectAllCompanyDB;
var
    ii : Integer;
    sL_CompCodeAndName,sL_CompCode,sL_Result : String;
begin
  //連結至所有公司的資料庫
  for ii:=0 to G_SGSParamIniStrList.Count-1 do
  begin
    sL_CompCode := G_SGSParamIniStrList[ii];

    sL_Result := connectToDB(sL_CompCode);
    //sL_Result := '';
    if sL_Result<>'' then
      frmMain.showLogMsg(sL_Result);
  end;
end;

function TdtmMain.connectToDB(sI_CompCode: String): WideString;
var
    ii,nL_DataAreaNo : Integer;
    sL_DbConnStr,sL_Temp : String;
    sL_DbPassword, sL_DbAlias, sL_DbUserName,sL_CompCode,sL_CompName : String;
begin
    try
      setlength(G_DbConnAry, G_SGSParamIniStrList.Count);
      setlength(G_AdoQueryAry, G_SGSParamIniStrList.Count);

      setlength(G_AdoStoredProcAry1, G_SGSParamIniStrList.Count);
      setlength(G_AdoStoredProcAry2, G_SGSParamIniStrList.Count);
      setlength(G_AdoStoredProcAry3, G_SGSParamIniStrList.Count);
      setlength(G_AdoStoredProcAry4, G_SGSParamIniStrList.Count);
      setlength(G_AdoStoredProcAry5, G_SGSParamIniStrList.Count);

      for ii:=0 to G_SGSParamIniStrList.Count -1 do
      begin
        nL_DataAreaNo := (G_SGSParamIniStrList.Objects[ii] as TSGSParamIniData).nDataAreaNo;

        sL_Temp := IntToStr(nL_DataAreaNo);
        sL_DbPassword := (G_SGSParamIniStrList.Objects[ii] as TSGSParamIniData).sPassword;
        sL_DbAlias := (G_SGSParamIniStrList.Objects[ii] as TSGSParamIniData).sAlias;
        sL_DbUserName := (G_SGSParamIniStrList.Objects[ii] as TSGSParamIniData).sUserID;
        sL_CompCode := (G_SGSParamIniStrList.Objects[ii] as TSGSParamIniData).sCompCode;
        sL_CompName := (G_SGSParamIniStrList.Objects[ii] as TSGSParamIniData).sCompName;

//    sL_DbConnStr := 'Provider=OraOLEDB.Oracle.1;Password=' + sL_DbPassword +';Persist Security Info=True;User ID='+sL_DbUserName+';Data Source='+sL_DbAlias; //=> 適用於my nb
        sL_DbConnStr := 'Provider=MSDAORA.1;Password=' + sL_DbPassword + ';User ID='+sL_DbUserName+';Data Source='+sL_DbAlias+';Persist Security Info=True'; //適用於其他 PC



        //連線該公司所屬之資料庫,且若該資料庫已開啟則不再重複連線
        if sI_CompCode = sL_CompCode then
        begin
          G_DbConnAry[nL_DataAreaNo] := TADOConnection.Create(nil);
          if not G_DbConnAry[nL_DataAreaNo].Connected then
          begin
            //ADOConnection 資料庫連線設定
            G_DbConnAry[nL_DataAreaNo].ConnectionString := sL_DbConnStr;
            G_DbConnAry[nL_DataAreaNo].LoginPrompt := false;
            G_DbConnAry[nL_DataAreaNo].Connected := true;

            //ADOQuery 連ADOConnection 設定
            G_AdoQueryAry[nL_DataAreaNo] :=  TADOQuery.Create(nil);
            G_AdoQueryAry[nL_DataAreaNo].Connection := G_DbConnAry[nL_DataAreaNo];


            //ADOStoredProc 連 ADOConnection 設定
            G_AdoStoredProcAry1[nL_DataAreaNo] := TADOStoredProc.Create(nil);
            G_AdoStoredProcAry1[nL_DataAreaNo].Connection := G_DbConnAry[nL_DataAreaNo];
            G_AdoStoredProcAry1[nL_DataAreaNo].ProcedureName := 'SF_ICCARDQUERY';
            G_AdoStoredProcAry1[nL_DataAreaNo].Parameters.Refresh;

            G_AdoStoredProcAry2[nL_DataAreaNo] := TADOStoredProc.Create(nil);
            G_AdoStoredProcAry2[nL_DataAreaNo].Connection := G_DbConnAry[nL_DataAreaNo];
            G_AdoStoredProcAry2[nL_DataAreaNo].ProcedureName := 'SF_CAPRODUCTQUERY';
            G_AdoStoredProcAry2[nL_DataAreaNo].Parameters.Refresh;

            G_AdoStoredProcAry3[nL_DataAreaNo] := TADOStoredProc.Create(nil);
            G_AdoStoredProcAry3[nL_DataAreaNo].Connection := G_DbConnAry[nL_DataAreaNo];
            G_AdoStoredProcAry3[nL_DataAreaNo].ProcedureName := 'SF_ProdPurchaseQuery';
            G_AdoStoredProcAry3[nL_DataAreaNo].Parameters.Refresh;

            G_AdoStoredProcAry4[nL_DataAreaNo] := TADOStoredProc.Create(nil);
            G_AdoStoredProcAry4[nL_DataAreaNo].Connection := G_DbConnAry[nL_DataAreaNo];
            G_AdoStoredProcAry4[nL_DataAreaNo].ProcedureName := 'SF_EntitlementQuery';
            G_AdoStoredProcAry4[nL_DataAreaNo].Parameters.Refresh;

            G_AdoStoredProcAry5[nL_DataAreaNo] := TADOStoredProc.Create(nil);
            G_AdoStoredProcAry5[nL_DataAreaNo].Connection := G_DbConnAry[nL_DataAreaNo];
            G_AdoStoredProcAry5[nL_DataAreaNo].ProcedureName := 'SP_InsertSO453';
            G_AdoStoredProcAry5[nL_DataAreaNo].Parameters.Refresh;
          end;
        end;

      end;

    except
      Result := getLanguageSettingInfo('dtmMain_Msg_4_Content') + '=<' + sL_CompName + '>';
      exit;
    end;
//    savetofile('ok','c:\aa4.txt');
    result := '';

end;


function TdtmMain.connectToDBAgain(sI_CompCode: String): String;
var
    nL_DataAreaNo,nL_Ndx : Integer;
    sL_DbConnStr,sL_CompName : String;
    sL_DbPassword, sL_DbAlias, sL_DbUserName,sL_CompCode : String;
begin
    try
      nL_Ndx := G_SGSParamIniStrList.IndexOf(sI_CompCode);
      if nL_Ndx<>-1 then
      begin

        nL_DataAreaNo := (G_SGSParamIniStrList.Objects[nL_Ndx] as TSGSParamIniData).nDataAreaNo;
        sL_DbPassword := (G_SGSParamIniStrList.Objects[nL_Ndx] as TSGSParamIniData).sPassword;
        sL_DbAlias := (G_SGSParamIniStrList.Objects[nL_Ndx] as TSGSParamIniData).sAlias;
        sL_DbUserName := (G_SGSParamIniStrList.Objects[nL_Ndx] as TSGSParamIniData).sUserID;
        sL_CompCode := (G_SGSParamIniStrList.Objects[nL_Ndx] as TSGSParamIniData).sCompCode;
        sL_CompName := (G_SGSParamIniStrList.Objects[nL_Ndx] as TSGSParamIniData).sCompName;

        sL_DbConnStr := 'Provider=MSDAORA.1;Password=' + sL_DbPassword + ';User ID='+sL_DbUserName+';Data Source='+sL_DbAlias+';Persist Security Info=True'; //適用於其他 PC
        G_DbConnAry[nL_DataAreaNo].Connected := false;

        G_DbConnAry[nL_DataAreaNo].ConnectionString := sL_DbConnStr;
        G_DbConnAry[nL_DataAreaNo].LoginPrompt := false;

        G_DbConnAry[nL_DataAreaNo].Connected := true;

        G_AdoQueryAry[nL_DataAreaNo].Connection := G_DbConnAry[nL_DataAreaNo];
        G_AdoQueryAry[nL_DataAreaNo].Active := true;

      end;
    except
      Result := getLanguageSettingInfo('dtmMain_Msg_4_Content') + '=<' + sL_CompName + '>';
      exit;
    end;

    result := '';

end;

procedure TdtmMain.DataModuleCreate(Sender: TObject);
var
    sL_ExeFileName,sL_Result : String;
begin
    G_UserIDStrList := TStringList.Create;
    G_UserIDStrList.Clear;

    G_SGSParamIniStrList := TStringList.Create;
    G_CompCodeAndNameStrList := TStringList.Create;



    sL_ExeFileName := Application.ExeName;
    sG_ExePath := ExtractFileDir(sL_ExeFileName);
    

    //開啟畫面參數ini
    if not dtmmain.initLanguageIniFile then
      Application.Terminate;


    //down, 讀取GateWayParam.ini並存成變數使用
    sL_Result := getSGSParamIniInfo;
    if sL_Result<>'' then
    begin
      MessageDlg(sL_Result,mtError, [mbOK],0);
      Application.Terminate;
    end;
    //up, 讀取GateWayParam.ini並存成變數使用

    //刪除TempIni檔案
    DeleteTmpIniFile;
end;

procedure TdtmMain.DataModuleDestroy(Sender: TObject);
begin
    G_UserIDStrList.Free;
    G_SGSParamIniStrList.Free;
end;

procedure TdtmMain.DeleteTmpIniFile;
var
    sL_IniFileName : String;
begin
    //刪除解密後之 ini 檔
    sL_IniFileName := sG_ExePath + '\' + TMP_SYS_INI_FILE_NAME;
    DeleteFile(sL_IniFileName);
end;

function TdtmMain.getActiveQuery(sI_CompCode: String): TADOQuery;
var
    ii, nL_Pos,nL_AreaDataNo : Integer;
begin
    nL_AreaDataNo := getAreaDataNo(sI_CompCode);

    nL_Pos := -1;
    for ii:=0 to G_SGSParamIniStrList.Count -1 do
    begin
      if (G_SGSParamIniStrList.Objects[ii] as TSGSParamIniData).nDataAreaNo = nL_AreaDataNo then
      //if (DATA_AREA_HEADER + sL_AreaDataNo)=G_SGSParamIniStrList.Strings[ii] then
      begin
        nL_Pos := nL_AreaDataNo;
        break;
      end;
    end;


    if nL_Pos<>-1 then
    begin
      result := G_AdoQueryAry[nL_Pos];
      //savetofile('get query '+ IntToStr(nL_AreaDataNo), 'c:\bb4.txt');
    end
    else
    begin
      result := nil;
      //savetofile('can not get query '+ IntToStr(nL_AreaDataNo), 'c:\bb5.txt');
    end;

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

function TdtmMain.getAreaDataNo(sI_CompCode: String): Integer;
var
    nL_Ndx,nL_AreaDatNo : Integer;

begin
    nL_Ndx := G_SGSParamIniStrList.IndexOf(sI_CompCode);
    if nL_Ndx<>-1 then
    begin
       nL_AreaDatNo := (G_SGSParamIniStrList.Objects[nL_Ndx] as TSGSParamIniData).nDataAreaNo;

      Result := nL_AreaDatNo;
    end;
end;

function TdtmMain.getCompName(sI_CompCode: String): String;
var
    nL_Ndx : Integer;
    sL_CompName : String;
begin
    nL_Ndx := G_SGSParamIniStrList.IndexOf(sI_CompCode);
    if nL_Ndx<>-1 then
    begin
       sL_CompName := (G_SGSParamIniStrList.Objects[nL_Ndx] as TSGSParamIniData).sCompName;

      Result := sL_CompName;
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




function TdtmMain.getLanguageSettingInfo(sI_Key: String): String;
var
    sL_Result, sL_LanguageType : String;
begin
    if Assigned(G_LanguageSettingIni) then
    begin
      sL_LanguageType := G_LanguageSettingIni.ReadString('CURRENT_LANGUAGE_TYPE','CURRENT_LANGUAGE_TYPE','TC');// default is Traditional Chinese

      sL_Result := G_LanguageSettingIni.ReadString(sL_LanguageType,sI_Key,'');
    end
    else
      sL_Result := sI_Key;
    result := sL_Result;
end;

function TdtmMain.getSGSParamIniInfo: WideString;
var
    L_IniFile : TIniFile;
    ii, nL_TotalCompCounts : Integer;
    sL_CompFlag,sL_Temp : String;
    sL_IniFileName,sL_DbAlias, sL_DbUserName, sL_DbPassword, sL_Tag : String;
    sL_CompCodeAndName,sL_CompCode,sL_CompName : String;
    L_Obj : TSGSParamIniData;
begin
    //將加密的ini檔案解密
    TransTmpIniFile;

    sL_IniFileName := sG_ExePath + '\' + TMP_SYS_INI_FILE_NAME;

    if not FileExists(sL_IniFileName) then
    begin
      result := getLanguageSettingInfo('dtmMain_Msg_1_Content') + '<'+ sL_IniFileName + '>';
      exit;
    end;

    L_IniFile := TIniFile.Create(sL_IniFileName);

    //判別簡體繁體
    sG_TcOrSc := L_IniFile.ReadString('CURRENT_LANGUAGE_TYPE','CURRENT_LANGUAGE_TYPE','');

    nL_TotalCompCounts := L_IniFile.ReadInteger(sG_TcOrSc,'COMP_COUNT',0);
    if nL_TotalCompCounts=0 then
    begin
      result := getLanguageSettingInfo('dtmMain_Msg_2_Content') + '<'+sL_IniFileName +'>';
      exit;
    end;

    nG_NightRunExecTime := L_IniFile.ReadInteger('SYSINFO','NIGHT_RUN_EXEC_TIME',30);
    if nG_NightRunExecTime < 5 then
      nG_NightRunExecTime := 5
    else if nG_NightRunExecTime > 30 then
      nG_NightRunExecTime := 30;

    sG_CancelChannelStr := L_IniFile.ReadString('SYSINFO','CANCEL_CHANNEL','');

    nG_HttpPort := L_IniFile.ReadInteger('SYSINFO','HTTP_PORT',80);

    SG_WinZipExePath := L_IniFile.ReadString('SYSINFO','WINZIP_PATH','');

    try
      for ii:=0 to nL_TotalCompCounts-1 do
      begin
        sL_Tag := DATA_AREA_HEADER + IntToStr(ii);
        sL_CompFlag := '_' + IntToStr(ii + 1);

        L_Obj := TSGSParamIniData.Create;

        L_Obj.nDataAreaNo := ii;
        L_Obj.sDataArea := sG_TcOrSc;
        L_Obj.sAlias := L_IniFile.ReadString(sG_TcOrSc,'ALIAS' + sL_CompFlag,'');
        L_Obj.sUserID := L_IniFile.ReadString(sG_TcOrSc,'USERID' + sL_CompFlag,'');
        L_Obj.sPassword := L_IniFile.ReadString(sG_TcOrSc,'PASSWORD' + sL_CompFlag,'');

        sL_CompCode := L_IniFile.ReadString(sG_TcOrSc,'COMPCODE' + sL_CompFlag,'');
        L_Obj.sCompCode := sL_CompCode;
        sL_CompName := L_IniFile.ReadString(sG_TcOrSc,'COMPNAME' + sL_CompFlag,'');
        L_Obj.sCompName := sL_CompName;

        sL_CompCodeAndName := sL_CompCode + '-' + sL_CompName;
        G_CompCodeAndNameStrList.Add(sL_CompCodeAndName);


        L_Obj.sDataPath1 := L_IniFile.ReadString(sG_TcOrSc,'DATAPATH1' + sL_CompFlag,'');
        L_Obj.sDataPath2 := L_IniFile.ReadString(sG_TcOrSc,'DATAPATH2' + sL_CompFlag,'');

        L_Obj.sFirstDate := L_IniFile.ReadString(sG_TcOrSc,'FIRSTDATE' + sL_CompFlag,'');
        L_Obj.sAdministratorMail1 := L_IniFile.ReadString(sG_TcOrSc,'ADMINISTRATORMAIL1' + sL_CompFlag,'');
        L_Obj.sAdministratorMail2 := L_IniFile.ReadString(sG_TcOrSc,'ADMINISTRATORMAIL2' + sL_CompFlag,'');
        L_Obj.sVersion := L_IniFile.ReadString(sG_TcOrSc,'VERSION' + sL_CompFlag,'');
        L_Obj.sSrcCode := L_IniFile.ReadString(sG_TcOrSc,'SRCCODE' + sL_CompFlag,'');
        L_Obj.sSrcIP := L_IniFile.ReadString(sG_TcOrSc,'SRCIP' + sL_CompFlag,'');
        L_Obj.sSrcType := L_IniFile.ReadString(sG_TcOrSc,'SRCTYPE' + sL_CompFlag,'');
        L_Obj.sDstCode := L_IniFile.ReadString(sG_TcOrSc,'DSTCODE' + sL_CompFlag,'');
        L_Obj.sDstIP := L_IniFile.ReadString(sG_TcOrSc,'DSTIP' + sL_CompFlag,'');
        L_Obj.sDstType := L_IniFile.ReadString(sG_TcOrSc,'DSTTYPE' + sL_CompFlag,'');
        L_Obj.sCasID := L_IniFile.ReadString(sG_TcOrSc,'CASID' + sL_CompFlag,'');

        G_SGSParamIniStrList.AddObject(sL_CompCode, L_Obj);
      end;

      L_IniFile.Free;
    except
      result := getLanguageSettingInfo('dtmMain_Msg_3_Content') + '<'+sL_IniFileName +'>';
      exit;
    end;
    result := '';
end;


procedure TdtmMain.initialcdsSGSParamSet(I_StrList: TStringList);
var
    ii : Integer;
    L_IniStrList1,L_IniStrList2 : TStringList;
    sL_Value,sL_Group,sL_ColumnName : String;
begin
    L_IniStrList1 := TStringList.Create;
    L_IniStrList2 := TStringList.Create;

    for ii:=1 to I_StrList.Count-1 do
    begin
      //將 ALIAS_1=mis 拆解成  [ALIAS_1], [mis]=Value, [ALIAS]=ColumnName,[1]=Group
      L_IniStrList1 := TUstr.ParseStrings(I_StrList[ii],'=');
      sL_Value := L_IniStrList1[1];

      L_IniStrList2 := TUstr.ParseStrings(L_IniStrList1[0],'_');
      sL_ColumnName := L_IniStrList2[0];
      sL_Group := L_IniStrList2[1];

      with cdsSGSParam do
      begin
        if Locate('Group',VarArrayOf([sL_Group]) , []) then
        begin
          Edit;
          FieldByName(sL_ColumnName).AsString := sL_Value;
          Post;
        end
        else
        begin
          Append;
          FieldByName('Group').AsString := sL_Group;
          FieldByName(sL_ColumnName).AsString := sL_Value;
          Post;
        end;
      end;
      //showmessage(IntToStr(ii) + ' = ' + I_StrList[ii] + '---' + sL_Group + '/' + sL_Value);
    end;

    cdsSGSParam.First;

    L_IniStrList1.Free;
    L_IniStrList2.Free;
end;

function TdtmMain.initLanguageIniFile : boolean;
var
    sL_IniFileName : String;

begin
    sL_IniFileName := sG_ExePath + '\' +  FORM_INI_FILE_NAME;

    if not FileExists(sL_IniFileName) then
    begin
      MessageDlg(dtmMain.getLanguageSettingInfo('dtmMain_Msg_5_Content') + '<'+sL_IniFileName +'>',mtError, [mbOK],0);//讀取畫面參數檔失敗
      Result := false;   
      exit;
    end;

    G_LanguageSettingIni := TIniFile.Create(sL_IniFileName);
    Result := true;
end;



function TdtmMain.runSQL(nI_Mode: Integer; sI_CompCode,
  sI_SQL: String): TADOQuery;
var
  L_AdoQuery : TADOQuery;
begin
  L_AdoQuery := getActiveQuery(sI_CompCode);
  if L_AdoQuery<>nil then
  begin
    L_AdoQuery.SQL.Clear;
    L_AdoQuery.SQL.Add(sI_SQL);
    case nI_Mode of
      SELECT_MODE: //select
        L_AdoQuery.Open;
      IUD_MODE:    //DML
        L_AdoQuery.ExecSQL;
    end;
  end;
  result := L_AdoQuery;
end;

procedure TdtmMain.saveToFile(vI_Content: OleVariant; sI_FileName: String);
var
    L_StrList : TStringList;
begin
    L_StrList := TStringList.Create;;
    L_StrList.Add(vI_Content);
    L_StrList.SaveToFile(sI_FileName);
    L_StrList.Free;
end;

procedure TdtmMain.TransTmpIniFile;
var
    L_StrList, L_TmpStrList : TStringList;
    //sL_ExeFileName, sL_ExePath ,sL_IniFileName,sL_TempIniFileName : String;
    sL_IniFileName,sL_TempIniFileName : String;
    ii : Integer;
    L_Intf : _Encrypt;
    sL_EncKey : WideString;
    sL_Temp : WideString;
begin
    sL_IniFileName := sG_ExePath + '\' + SYS_INI_FILE_NAME;

    if not FileExists(sL_IniFileName) then
    begin
      MessageDlg(getLanguageSettingInfo('dtmMain_Msg_1_Content') + '<' + sL_IniFileName + '>',mtError, [mbOK],0);
      Application.Terminate;
    end
    else
    begin
      sL_EncKey := 'CS';
      L_Intf := CoEncrypt.Create;

      L_StrList := TStringList.Create;
      L_TmpStrList := TStringList.Create;

      L_StrList.LoadFromFile(sL_IniFileName);
      for ii:=0 to L_StrList.Count-1 do
      begin
         if (Copy(L_StrList.Strings[ii],1,1)=';') or (L_StrList.Strings[ii]='')
            or (Copy(L_StrList.Strings[ii],1,8)='COMPCODE') or (Copy(L_StrList.Strings[ii],1,8)='COMPNAME')
            or (Copy(L_StrList.Strings[ii],1,8)='DATAPATH') or (Copy(L_StrList.Strings[ii],1,23)='[CURRENT_LANGUAGE_TYPE]')
            or (Copy(L_StrList.Strings[ii],1,21)='CURRENT_LANGUAGE_TYPE') or (Copy(L_StrList.Strings[ii],1,9)='[SYSINFO]')
            or (Copy(L_StrList.Strings[ii],1,19)='NIGHT_RUN_EXEC_TIME') or (Copy(L_StrList.Strings[ii],1,14)='CANCEL_CHANNEL')
            or (Copy(L_StrList.Strings[ii],1,9)='HTTP_PORT') or (Copy(L_StrList.Strings[ii],1,11)='WINZIP_PATH') then
         begin
           L_TmpStrList.Add(L_StrList.Strings[ii]);
         end
         else
         begin
           sL_Temp := L_StrList.Strings[ii];
           L_TmpStrList.Add(L_Intf.Decrypt(sL_Temp));
         end;

      end;
      L_TmpStrList.SaveToFile(sG_ExePath + '\' + TMP_SYS_INI_FILE_NAME);
      L_TmpStrList.Free;

      L_Intf := nil;

    end;

end;

function TdtmMain.getCompCode(sI_SrcCode: String;var nI_Index : Integer): String;
var
    sL_TempSrcCode,sL_CompCode : String;
    ii : Integer;
begin
    for ii:=0 to G_SGSParamIniStrList.count-1 do
    begin
      sL_TempSrcCode := (G_SGSParamIniStrList.Objects[ii] as TSGSParamIniData).sSrcCode;
      if sI_SrcCode = sL_TempSrcCode then
      begin
        nI_Index := ii;
        Result := (G_SGSParamIniStrList.Objects[ii] as TSGSParamIniData).sCompCode;
      end;
    end;
end;

function TdtmMain.runCAProductQuerySF(sI_CompCode, sI_SrcCode, sI_DstCode,
  sI_DstMsgID, sI_CasID, sI_InstQuery, sI_FirstDate,
  sI_HandleEDateTime: String; var sI_RetCode, sI_Msg,
  sI_SrcMsgID,sL_FirstFlag: String;var nI_TotalExecCounts,nI_TotalProductNum,nI_TotalChannelNum : Integer): String;
var
    nL_ActiveConn : Integer;
begin
  try
    nL_ActiveConn := getAreaDataNo(sI_CompCode);

    with G_AdoStoredProcAry2[nL_ActiveConn] do
    begin
      Parameters.ParamByName('p_CompCode').Value := sI_CompCode;
      Parameters.ParamByName('P_SrcCode').Value := sI_SrcCode;
      Parameters.ParamByName('P_DstCode').Value := sI_DstCode;
      Parameters.ParamByName('P_DstMsgID').Value := sI_DstMsgID;

      Parameters.ParamByName('P_CasID').Value := sI_CasID;
      Parameters.ParamByName('P_InstQuery').Value := sI_InstQuery;
      Parameters.ParamByName('P_FirstDate').Value := sI_FirstDate;
      Parameters.ParamByName('P_HandleEDateTime').Value := sI_HandleEDateTime;

      ExecProc;

      //sL_FunReturn := Parameters.Items[0].Value; //用ParamByName會有問題
      //sL_RetCode := Parameters.Items[1].Value; //用ParamByName會有問題
      //sL_Msg := Parameters.Items[2].Value; //用ParamByName會有問題

      sI_RetCode := Parameters.ParamByName('p_ErrCode').Value;
      sI_Msg := Parameters.ParamByName('p_ErrMsg').Value;
      sI_SrcMsgID := Parameters.ParamByName('p_SrcMsgID').Value;
      nI_TotalExecCounts := Parameters.ParamByName('p_TotalExecCounts').Value;
      sL_FirstFlag := Parameters.ParamByName('p_FirstFlag').Value;
      nI_TotalProductNum := Parameters.ParamByName('p_TotalProductNum').Value;
      nI_TotalChannelNum := Parameters.ParamByName('p_TotalChannelNum').Value;

    end;

  except

  end;
end;

function TdtmMain.runProdPurchaseQuerySF(sI_CompCode, sI_SrcCode,
  sI_DstCode, sI_DstMsgID, sI_CasID, sI_InstQuery, sI_FirstDate,
  sI_HandleEDateTime, sI_CancelChannelStr: String; var sI_RetCode, sI_Msg,
  sI_SrcMsgID,sI_FirstFlag: String;var nI_TotalExecCounts : Integer): String;
var
    nL_ActiveConn : Integer;
    sL_AAA : String;
begin
  try
    nL_ActiveConn := getAreaDataNo(sI_CompCode);

    with G_AdoStoredProcAry3[nL_ActiveConn] do
    begin
      Parameters.ParamByName('p_CompCode').Value := sI_CompCode;
      Parameters.ParamByName('P_SrcCode').Value := sI_SrcCode;
      Parameters.ParamByName('P_DstCode').Value := sI_DstCode;
      Parameters.ParamByName('P_DstMsgID').Value := sI_DstMsgID;

      Parameters.ParamByName('P_CasID').Value := sI_CasID;
      Parameters.ParamByName('P_InstQuery').Value := sI_InstQuery;
      Parameters.ParamByName('P_FirstDate').Value := sI_FirstDate;
      Parameters.ParamByName('P_HandleEDateTime').Value := sI_HandleEDateTime;

      //sL_AAA := '''B1'',''B2''';
      Parameters.ParamByName('P_CancelChStr').Value := sI_CancelChannelStr;

      ExecProc;

      //sL_FunReturn := Parameters.Items[0].Value; //用ParamByName會有問題
      //sL_RetCode := Parameters.Items[1].Value; //用ParamByName會有問題
      //sL_Msg := Parameters.Items[2].Value; //用ParamByName會有問題

      sI_RetCode := Parameters.ParamByName('p_ErrCode').Value;
      sI_Msg := Parameters.ParamByName('p_ErrMsg').Value;
      sI_SrcMsgID := Parameters.ParamByName('p_SrcMsgID').Value;
      nI_TotalExecCounts := Parameters.ParamByName('p_TotalExecCounts').Value;
      sI_FirstFlag := Parameters.ParamByName('p_FirstFlag').Value;

    end;

  except

  end;
end;

function TdtmMain.runEntitlementQuerySF(sI_CompCode, sI_SrcCode,
  sI_DstCode, sI_DstMsgID, sI_CasID, sI_InstQuery, sI_FirstDate,
  sI_HandleEDateTime, sI_CancelChannelStr: String; var sI_RetCode, sI_Msg,
  sI_SrcMsgID,sI_FirstFlag: String;var nI_TotalExecCounts : Integer): String;
var
    nL_ActiveConn : Integer;
begin
  try
    nL_ActiveConn := getAreaDataNo(sI_CompCode);

    with G_AdoStoredProcAry4[nL_ActiveConn] do
    begin
      Parameters.ParamByName('p_CompCode').Value := sI_CompCode;
      Parameters.ParamByName('P_SrcCode').Value := sI_SrcCode;
      Parameters.ParamByName('P_DstCode').Value := sI_DstCode;
      Parameters.ParamByName('P_DstMsgID').Value := sI_DstMsgID;

      Parameters.ParamByName('P_CasID').Value := sI_CasID;
      Parameters.ParamByName('P_InstQuery').Value := sI_InstQuery;
      Parameters.ParamByName('P_FirstDate').Value := sI_FirstDate;
      Parameters.ParamByName('P_HandleEDateTime').Value := sI_HandleEDateTime;
      Parameters.ParamByName('P_CancelChStr').Value := sI_CancelChannelStr;

      ExecProc;

      //sL_FunReturn := Parameters.Items[0].Value; //用ParamByName會有問題
      //sL_RetCode := Parameters.Items[1].Value; //用ParamByName會有問題
      //sL_Msg := Parameters.Items[2].Value; //用ParamByName會有問題

      sI_RetCode := Parameters.ParamByName('p_ErrCode').Value;
      sI_Msg := Parameters.ParamByName('p_ErrMsg').Value;
      sI_SrcMsgID := Parameters.ParamByName('p_SrcMsgID').Value;
      nI_TotalExecCounts := Parameters.ParamByName('p_TotalExecCounts').Value;
      sI_FirstFlag := Parameters.ParamByName('p_FirstFlag').Value;

    end;

  except

  end;
end;

function TdtmMain.getSrcMsgID(sI_CompCode, sI_ActionMsg: String): String;
var
    sL_SQL,sL_CompName : String;
    L_AdoQry : TADOQuery;
begin
    //取的SrcMsgID
    sL_SQL := 'select S_SGSMSG_SeqNo.NextVal AS SrcMsgID from dual';
    L_AdoQry := dtmMain.runSQL(SELECT_MODE,sI_CompCode,sL_SQL);

    if L_AdoQry=nil then
    begin
      sL_CompName := dtmMain.getCompName(sI_CompCode);
      sI_ActionMsg := dtmMain.getLanguageSettingInfo('Return_Msg_100_Content') + '=<' + sL_CompName + '>'; //資料庫不連線, 公司名稱
      Result := '';
    end
    else
    begin
      sI_ActionMsg := '';
      Result := L_AdoQry.FieldByName('SrcMsgID').AsString;
    end;
end;

function TdtmMain.getQueryData(sI_QueryType,sI_Table,sI_CompCode,
  sI_HandleEDateTime: String): TADOQuery;
var
    sL_SQL : String;
begin
    if UpperCase(sI_QueryType) = UpperCase(CA_PRODUCT_QUERY) then
    begin
      sL_SQL := 'select * from ' + sI_Table + ' where CompCode=' + sI_CompCode +
                ' and HandleEDateTime=' + TUstr.getOracleSQLDateTimeStr(StrToDateTime(sI_HandleEDateTime)) +
                ' ORDER BY ProductCode';
    end
    else
    begin
      sL_SQL := 'select * from ' + sI_Table + ' where CompCode=' + sI_CompCode +
                ' and HandleEDateTime=' + TUstr.getOracleSQLDateTimeStr(StrToDateTime(sI_HandleEDateTime));
    end;

    Result := dtmMain.runSQL(SELECT_MODE,sI_CompCode,sL_SQL);
end;

function TdtmMain.getICCardNumAndSubNum(bI_IsTestData : Boolean;sI_CompCode,sI_HandleEDateTime: String;
  var nI_TotalICCardNum, nI_TotalSubNum: Integer): String;
var
    sL_SQL : String;
    L_AdoQry : TADOQuery;
begin
    if not bI_IsTestData then   //僅正式資料要計算
    begin
      //計算當前整個SMS內IC卡總數
      sL_SQL := 'Select Count(CustID) AS TotalICCardNum from SO004 ' +
                ' Where CompCode=' + sI_CompCode +
                ' And (PRDate Is Null or PRDate>=' + TUstr.getOracleSQLDateTimeStr(StrToDateTime(sI_HandleEDateTime)) + ') ' +
                ' And InstDate Is Not Null ' +
                ' And InstDate<=' + TUstr.getOracleSQLDateTimeStr(StrToDateTime(sI_HandleEDateTime)) +
                ' And SmartCardNo Is Not Null';

      L_AdoQry := dtmMain.runSQL(SELECT_MODE,sI_CompCode,sL_SQL);
      if L_AdoQry = nil  then
        nI_TotalICCardNum := 0
      else
        nI_TotalICCardNum := L_AdoQry.FieldByName('TotalICCardNum').AsInteger;
    end
    else
      nI_TotalICCardNum := 0;


    //計算當前整個SMS內用戶總數
    sL_SQL := 'Select Count(Distinct CustID) AS TotalSubNum from SO004 ' +
              ' Where CompCode=' + sI_CompCode +
              ' And (PRDate Is Null or PRDate>=' + TUstr.getOracleSQLDateTimeStr(StrToDateTime(sI_HandleEDateTime)) + ') ' +
              ' And InstDate Is Not Null ' +
              ' And InstDate<=' + TUstr.getOracleSQLDateTimeStr(StrToDateTime(sI_HandleEDateTime)) +
              ' And SmartCardNo Is Not Null';

    L_AdoQry := dtmMain.runSQL(SELECT_MODE,sI_CompCode,sL_SQL);
    if L_AdoQry = nil  then
      nI_TotalSubNum := 0
    else
      nI_TotalSubNum := L_AdoQry.FieldByName('TotalSubNum').AsInteger;
end;

function TdtmMain.getDataPath(sI_CompCode: String;var sI_DataPath1,sI_DataPath2 : String): String;
var
    nL_Ndx : Integer;
begin
    nL_Ndx := G_SGSParamIniStrList.IndexOf(sI_CompCode);
    if nL_Ndx<>-1 then
    begin
       sI_DataPath1 := (G_SGSParamIniStrList.Objects[nL_Ndx] as TSGSParamIniData).sDataPath1;
       sI_DataPath2 := (G_SGSParamIniStrList.Objects[nL_Ndx] as TSGSParamIniData).sDataPath2;
    end;
end;

function TdtmMain.getFileSize(sI_XmlFilePath: String): Double;
var
    F : File of Byte;
    nL_Size : LongInt;
begin
    AssignFile(F,sI_XmlFilePath);
    Reset(F);
    nL_Size := FileSize(F);

    Close(F);

    Result := nL_Size/1024;
end;


function TdtmMain.getRealFilePath(sI_CompCode, sI_InstQuery,
  sI_QueryType,sI_HandleEDateTime: String;bI_IsZipFile: Boolean): String;
var
    sL_DataPath1,sL_DataPath2,sL_SunFileName,sL_FileName,sL_XmlFilePath : String;

begin
    dtmMain.getDataPath(sI_CompCode,sL_DataPath1,sL_DataPath2);

    sL_FileName := getFileName(sI_HandleEDateTime,sI_QueryType,bI_IsZipFile);

    if sI_InstQuery = '0' then
      sL_XmlFilePath := sL_DataPath1 + '\' + sL_FileName
    else
      sL_XmlFilePath := sL_DataPath2 + '\' + sL_FileName;

    Result := sL_XmlFilePath;
end;

function TdtmMain.getSrcIP(sI_CompCode: String): String;
var
    nL_Ndx : Integer;
begin
    nL_Ndx := G_SGSParamIniStrList.IndexOf(sI_CompCode);
    if nL_Ndx<>-1 then
    begin
       Result := (G_SGSParamIniStrList.Objects[nL_Ndx] as TSGSParamIniData).sSrcIP;
    end;
end;

function TdtmMain.getFileName(sI_HandleEDateTime, sI_QueryType: String;
  bI_IsZipFile: Boolean): String;
var
    sL_SunFileName : String;
begin
    if bI_IsZipFile then
      sL_SunFileName := '.zip'
    else
      sL_SunFileName := '.xml';

    Result := Copy(Trim(TUstr.replaceStr(TUstr.replaceStr(sI_HandleEDateTime,'/',''),'-','')),0,8) + '_SMS_' + sI_QueryType + 'R' + sL_SunFileName;

end;

function TdtmMain.logToSO460(sI_CompCode, sI_ModeType, sI_HandleEDateTime,
  sI_InstQuery, sI_SrcCode, sI_SrcMsgID, sI_DstCode, sI_DstMsgID,
  sI_RetCode, sI_Msg: String): String;
var
    sL_SQL : String;
begin
    sL_SQL := 'INSERT INTO SO460(CompCode,ModeType,HandleEDateTime,ExecDateTime,' +
              'InstQuery,SrcCode,SrcMsgID,DstCode,DstMsgID,ErrorCode,ErrorMsg) Values(' +
              sI_CompCode + ',' +  sI_ModeType + ','  + TUstr.getOracleSQLDateTimeStr(StrToDateTime(sI_HandleEDateTime)) +
              ',' + TUstr.getOracleSQLDateTimeStr(now) + ',' + sI_InstQuery + ',''' +
              sI_SrcCode + ''',''' +  sI_SrcMsgID + ''',''' + sI_DstCode + ''',''' +
              sI_DstMsgID + ''',''' + sI_RetCode + ''',''' + sI_Msg + ''')';
    dtmMain.runSQL(IUD_MODE,sI_CompCode,sL_SQL);
end;

function TdtmMain.getProductNumAndChannelNum(sI_CompCode,sI_HandleEDateTime: String;
  var nI_TotalProductNum,nI_TotalChannelNum : Integer): String;
var
    sL_SQL : String;
    L_AdoQry : TADOQuery;
begin
    //當前整個SMS內CA產品總數
    sL_SQL := 'Select Count(Distinct ChannelID) AS TotalProductNum from CD024 ' +
              'Where CompCode=' + sI_CompCode + ' And StopFlag=0 And ChannelID Is Not Null';

    L_AdoQry := dtmMain.runSQL(SELECT_MODE,sI_CompCode,sL_SQL);
    if L_AdoQry = nil  then
      nI_TotalProductNum := 0
    else
      nI_TotalProductNum := L_AdoQry.FieldByName('TotalProductNum').AsInteger;



    //當前整個SMS內數位電視頻道總數
    sL_SQL := 'Select Count(CodeNo) AS TotalChannelNum from CD024A ' +
              'Where CompCode=' + sI_CompCode + ' And StopFlag=0';

    L_AdoQry := dtmMain.runSQL(SELECT_MODE,sI_CompCode,sL_SQL);
    if L_AdoQry = nil  then
      nI_TotalChannelNum := 0
    else
      nI_TotalChannelNum := L_AdoQry.FieldByName('TotalChannelNum').AsInteger;
end;

function TdtmMain.logToSO461(sI_CompCode, sI_ModeType, sI_HandleEDateTime,
  sI_QueryDateTime, sI_InstQuery, sI_SrcCode, sI_DstCode, sI_DstMsgID,
  sI_RequestXml: String): String;
var
    sL_SQL : String;
begin
    sL_SQL := 'INSERT INTO SO461(CompCode,ModeType,HandleEDateTime,QueryDateTime,ExecDateTime,' +
              'InstQuery,SrcCode,DstCode,DstMsgID,XmlContent) Values(' +
              sI_CompCode + ',' +  sI_ModeType + ','  + TUstr.getOracleSQLDateTimeStr(StrToDateTime(sI_HandleEDateTime)) +
              ',' + TUstr.getOracleSQLDateTimeStr(StrToDateTime(TUstr.replaceStr(sI_QueryDateTime,'-','/'))) +
              ',' + TUstr.getOracleSQLDateTimeStr(now) + ',' + sI_InstQuery + ',''' +
              sI_SrcCode + ''',''' +   sI_DstCode + ''',''' +
              sI_DstMsgID + ''',''' + sI_RequestXml + ''')';
    dtmMain.runSQL(IUD_MODE,sI_CompCode,sL_SQL);
end;

function TdtmMain.logToSO459(sI_CompCode, sI_ModeType, sI_FirstFlag,
  sI_HandleEDateTime, sI_RetCode, sI_Msg,sI_TotalICCardNum,
  sI_TotalSubNum, sI_TotalProductNum, sI_TotalChannelNum : String): String;
var
    sL_SQL,sL_Title,sL_ICCardNumSQL,sL_ProductNumSQL,sL_MsgSQL,sL_Msg : String;
begin
    sL_Title := 'INSERT INTO SO459(CompCode,ModeType,FirstFlag,HandleEDateTime,ExecDateTime,' +
              'TotalICCardNum,TotalSubNum,TotalProductNum,TotalChannelNum,ErrorCode,ErrorMsg) VALUES(' +
              sI_CompCode + ',' + sI_ModeType + ',' + sI_FirstFlag + ',' +
              TUstr.getOracleSQLDateTimeStr(StrToDateTime(sI_HandleEDateTime)) + ',' +
              TUstr.getOracleSQLDateTimeStr(now) + ',';

    if sI_TotalICCardNum='' then
      sL_ICCardNumSQL := 'null,null,'
    else
      sL_ICCardNumSQL :=  '''' + sI_TotalICCardNum + ''',''' + sI_TotalSubNum + ''',';

    if sI_TotalProductNum = '' then
      sL_ProductNumSQL := 'null,null,'
    else
      sL_ProductNumSQL :=  '''' + sI_TotalProductNum + ''',''' + sI_TotalChannelNum + ''',';

    if (sI_Msg = '') or (sI_Msg = ' ') then
      sL_MsgSQL := 'null,null)'
    else
      sL_MsgSQL :=  '''' + sI_RetCode + ''',''' + sI_Msg + ''')';

    sL_SQL := sL_Title + sL_ICCardNumSQL +  sL_ProductNumSQL + sL_MsgSQL;
    dtmMain.runSQL(IUD_MODE,sI_CompCode,sL_SQL);
end;

procedure TdtmMain.getSFRealMsg(sI_SrcRetCode, sI_SrcMsg: String;
  var sI_RealRetCode, sI_RealMsg: String);

begin
    if sI_SrcRetCode='-1' then  //-1 代表後端程式傳來的自訂中文錯誤訊息(資料庫錯誤)
    begin
      sI_RealRetCode := '105';
      sI_RealMsg := dtmMain.getLanguageSettingInfo('Return_Msg_105_Content');//資料庫處理異常
    end
    else if sI_SrcRetCode='-2' then  //-2 代表後端程式傳來的自訂中文錯誤訊息
    begin
      sI_RealRetCode := Copy(TUstr.replaceStr(sI_SrcMsg,'Return_Msg_',''),0,3);
      sI_RealMsg := dtmMain.getLanguageSettingInfo(sI_SrcMsg);
    end
    else
    begin
      sI_RealRetCode := sI_SrcRetCode;
      sI_RealMsg := sI_SrcMsg;
    end;
end;

function TdtmMain.getDstIP(sI_CompCode: String): String;
var
    nL_Ndx : Integer;
begin
    nL_Ndx := G_SGSParamIniStrList.IndexOf(sI_CompCode);
    if nL_Ndx<>-1 then
    begin
       Result := (G_SGSParamIniStrList.Objects[nL_Ndx] as TSGSParamIniData).sDstIP;
    end;
end;

function TdtmMain.getNightRunLogData(sI_CompCode, sI_ModeType,
  sI_SDateTime, sI_EDateTime: String): TADOQuery;
var
    sL_SQL : String;
begin
    sL_SQL := 'select HandleEDateTime,ExecDateTime,ErrorCode,ErrorMsg from SO459 Where CompCode=' + sI_CompCode +
              ' and ModeType=' + sI_ModeType + ' and HandleEDateTime between ' +
              TUstr.getOracleSQLDateTimeStr(StrToDateTime(sI_SDateTime)) + ' and ' +
              TUstr.getOracleSQLDateTimeStr(StrToDateTime(sI_EDateTime)) +
              ' ORDER BY HandleEDateTime,ExecDateTime';

    Result := dtmMain.runSQL(SELECT_MODE,sI_CompCode,sL_SQL);
end;

function TdtmMain.getQueryLogData(sI_CompCode, sI_ModeType, sI_SDateTime,
  sI_EDateTime: String): TADOQuery;
var
    sL_SQL : String;
begin
    sL_SQL := 'select a.DstMsgID,a.ModeType,a.InstQuery,a.DstCode,a.SrcCode,a.HandleEDateTime,' +
              'b.ExecDateTime,b.SrcMsgID,b.ErrorCode,b.ErrorMsg' +
              ' from so461 a,so460 b where a.DstMsgID=b.DstMsgID(+) and ((a.CompCode=' + sI_CompCode +
              ' and a.ModeType=' + sI_ModeType + ' and a.HandleEDateTime between ' +
              TUstr.getOracleSQLDateTimeStr(StrToDateTime(sI_SDateTime)) + ' and ' +
              TUstr.getOracleSQLDateTimeStr(StrToDateTime(sI_EDateTime)) +
              ') OR ((a.CompCode=' + sI_CompCode +
              ' and a.ModeType=' + sI_ModeType + ' and a.HandleEDateTime between ' +
              TUstr.getOracleSQLDateTimeStr(StrToDateTime(sI_SDateTime)) + ' and ' +
              TUstr.getOracleSQLDateTimeStr(StrToDateTime(sI_EDateTime)) +
              ') and (b.HandleEDateTime is null)))' +
              ' order by a.HandleEDateTime,to_number(a.DstMsgID)';

    Result := dtmMain.runSQL(SELECT_MODE,sI_CompCode,sL_SQL);
end;

function TdtmMain.getFirstDate(sI_CompCode: String): String;
var
    nL_Ndx : Integer;
    sL_FirstDate : String;
begin
    nL_Ndx := G_SGSParamIniStrList.IndexOf(sI_CompCode);
    if nL_Ndx<>-1 then
    begin
       sL_FirstDate := (G_SGSParamIniStrList.Objects[nL_Ndx] as TSGSParamIniData).sFirstDate;

      Result := sL_FirstDate;
    end;
end;

function TdtmMain.getSMSBasicData(sI_CompCode: String; var sI_STBFaciCode,
  sI_STBFaciName, sI_ICCFaciCode, sI_ICCFaciName, sI_BuyCode,
  sI_BuyName: String;var I_CustStrLis : TStringList): String;
var
    sL_SQL,sL_CompName,sL_Msg,sL_Counts : String;
    L_AdoQry : TADOQuery;
    //L_CustStrLis : TStringList;
    nL_Counts : Integer;
begin
    sL_SQL := 'select CodeNo, Description from CD022 where RefNo=3 and RowNum<=1';
    L_AdoQry := dtmMain.runSQL(SELECT_MODE,sI_CompCode,sL_SQL);
    if L_AdoQry=nil then
    begin
      sL_CompName := dtmMain.getCompName(sI_CompCode);
      sL_Msg := dtmMain.getLanguageSettingInfo('Return_Msg_100_Content') + '=<' + sL_CompName + '>'; //資料庫不連線, 公司名稱
      Result := sL_Msg;
      Exit;
    end
    else
    begin
      if L_AdoQry.RecordCount = 0 then
      begin   
        sL_Msg := 'No STB record in CD022';
        Result := sL_Msg;
        Exit;
      end
      else
      begin
        sI_STBFaciCode := L_AdoQry.FieldByName('CodeNo').AsString;
        sI_STBFaciName := L_AdoQry.FieldByName('Description').AsString;
      end;
    end;


    sL_SQL := 'select CodeNo, Description from CD022 where RefNo=4 and RowNum<=1';
    L_AdoQry := dtmMain.runSQL(SELECT_MODE,sI_CompCode,sL_SQL);
    if L_AdoQry=nil then
    begin
      sL_CompName := dtmMain.getCompName(sI_CompCode);
      sL_Msg := dtmMain.getLanguageSettingInfo('Return_Msg_100_Content') + '=<' + sL_CompName + '>'; //資料庫不連線, 公司名稱
      Result := sL_Msg;
      Exit;
    end
    else
    begin
      if L_AdoQry.RecordCount = 0 then
      begin   
        sL_Msg := 'No ICC record in CD022';
        Result := sL_Msg;
        Exit;
      end
      else
      begin
        sI_ICCFaciCode := L_AdoQry.FieldByName('CodeNo').AsString;
        sI_ICCFaciName := L_AdoQry.FieldByName('Description').AsString;
      end;
    end;


    sL_SQL := 'select CodeNo, Description from CD034 where CodeNo=1';
    L_AdoQry := dtmMain.runSQL(SELECT_MODE,sI_CompCode,sL_SQL);
    if L_AdoQry=nil then
    begin
      sL_CompName := dtmMain.getCompName(sI_CompCode);
      sL_Msg := dtmMain.getLanguageSettingInfo('Return_Msg_100_Content') + '=<' + sL_CompName + '>'; //資料庫不連線, 公司名稱
      Result := sL_Msg;
      Exit;
    end
    else
    begin
      if L_AdoQry.RecordCount = 0 then
      begin
        sL_Msg := 'No ICC record in CD022';
        Result := sL_Msg;
        Exit;
      end
      else
      begin
        sI_BuyCode := L_AdoQry.FieldByName('CodeNo').AsString;
        sI_BuyName := L_AdoQry.FieldByName('Description').AsString;
      end;
    end;


    sL_SQL := 'select CustId from SO002 where CustStatusCode=1 and ServiceType=''C''';
    L_AdoQry := dtmMain.runSQL(SELECT_MODE,sI_CompCode,sL_SQL);
    if L_AdoQry=nil then
    begin
      sL_CompName := dtmMain.getCompName(sI_CompCode);
      sL_Msg := dtmMain.getLanguageSettingInfo('Return_Msg_100_Content') + '=<' + sL_CompName + '>'; //資料庫不連線, 公司名稱
      Result := sL_Msg;
      Exit;
    end
    else
    begin
      if L_AdoQry.RecordCount = 0 then
      begin
        sL_Msg := 'No CustId in SO002';
        Result := sL_Msg;
        Exit;
      end
      else
      begin
        //L_CustStrLis := TStringList.Create;

        nL_Counts := 0;
        with L_AdoQry do
        begin
          First;
          while not Eof do
          begin
            I_CustStrLis.Add(FieldByName('CustId').AsString);
            Next;
          end;
        end;

        //I_CustStrLis := L_CustStrLis;
        //L_CustStrLis.Free;
      end;
    end;
    Result := '';
end;

function TdtmMain.getSMSSO005TransCustID(sI_CompCode,
  sI_ICCardNo: String;var sI_CustId : String): String;
var
    sL_SQL,sL_CompName,sL_Msg,sL_Counts : String;
    L_AdoQry : TADOQuery;
    nL_Counts : Integer;
begin
    sL_SQL := 'select CustId from SO004 where SmartCardNo=''' + sI_ICCardNo +
              ''' and InstDate is not null and PrDate is null';
    L_AdoQry := dtmMain.runSQL(SELECT_MODE,sI_CompCode,sL_SQL);

    if L_AdoQry=nil then
    begin
      sL_CompName := dtmMain.getCompName(sI_CompCode);
      sL_Msg := dtmMain.getLanguageSettingInfo('Return_Msg_100_Content') + '=<' + sL_CompName + '>'; //資料庫不連線, 公司名稱
      Result := sL_Msg;
      Exit;
    end
    else
    begin
      sI_CustId := L_AdoQry.FieldByName('CustId').AsString;
    end;
end;

function TdtmMain.getYesterdayTotalICCardNum(sI_Table, sI_CompCode,
  sI_StartDate: String): Integer;
var
    sL_SQL : String;
    L_AdoQry : TADOQuery;    
begin
    sL_SQL := 'SELECT COUNT(*) ICCardNum FROM ' + sI_Table +
              ' where CompCode=' + sI_CompCode + ' and  HANDLEEDATETIME<' +
              TUstr.getOracleSQLDateTimeStr(StrToDateTime(sI_StartDate));

    L_AdoQry := dtmMain.runSQL(SELECT_MODE,sI_CompCode,sL_SQL);

    if L_AdoQry=nil then
    begin
      Result := 0;
    end
    else
    begin
      Result := L_AdoQry.FieldByName('ICCardNum').AsInteger;
    end;
end;

function TdtmMain.runTestDataInsertSO453SP(sI_CompCode : String): String;
var
    nL_ActiveConn : Integer;
begin
  try
    nL_ActiveConn := getAreaDataNo(sI_CompCode);

    G_AdoStoredProcAry5[nL_ActiveConn].ExecProc;
    {
    with G_AdoStoredProcAry5[nL_ActiveConn] do
    begin
      ExecProc;
    end;
    }
  except

  end;

end;

end.
