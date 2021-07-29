unit dtmMainU;

interface

uses
  SysUtils, Classes,IniFiles,Forms,Dialogs;

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
    procedure DataModuleCreate(Sender: TObject);
    procedure DataModuleDestroy(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    G_LanguageSettingIni : TIniFile;
    G_SGSParamIniStrList : TStringList;
    G_CompCodeAndNameStrList : TStringList;
    sG_LocalIP, sG_ExePath,sG_TcOrSc,sG_CancelChannelStr : String;
    nG_NightRunExecTime,nG_HttpPort,nG_DstHttpPort : Integer;
    function getSGSParamIniInfo : WideString;
    procedure TransTmpIniFile;
    procedure DeleteTmpIniFile;
    function getLanguageSettingInfo(sI_Key:String):String;
    function initLanguageIniFile : Boolean;
    function getSrcIP(sI_CompCode : String) : String;
  end;

var
  dtmMain: TdtmMain;

implementation

uses ConstU, DevPower_TLB;

{$R *.dfm}

procedure TdtmMain.DataModuleCreate(Sender: TObject);
var
    sL_ExeFileName,sL_Result : String;
begin
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
    if (nG_NightRunExecTime<5) Or (nG_NightRunExecTime>30) then
      nG_NightRunExecTime := 30;

    sG_CancelChannelStr := L_IniFile.ReadString('SYSINFO','CANCEL_CHANNEL','');

    nG_HttpPort := L_IniFile.ReadInteger('SYSINFO','HTTP_PORT',80);
    nG_DstHttpPort := L_IniFile.ReadInteger('SYSINFO','HTTP_DST_PORT',80);
    sG_LocalIP := L_IniFile.ReadString('SYSINFO','LOCAL_IP','127.0.0.1');    

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
            or (Copy(L_StrList.Strings[ii],1,9)='HTTP_PORT') or (Copy(L_StrList.Strings[ii],1,11)='WINZIP_PATH')
            or (Copy(L_StrList.Strings[ii],1,13)='HTTP_DST_PORT') then
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

procedure TdtmMain.DataModuleDestroy(Sender: TObject);
begin
    G_SGSParamIniStrList.Free;
    G_CompCodeAndNameStrList.Free;
end;

function TdtmMain.initLanguageIniFile: Boolean;
var
    sL_IniFileName : String;

begin
    sL_IniFileName := sG_ExePath + '\' +  FORM_INI_FILE_NAME;

    if not FileExists(sL_IniFileName) then
    begin
      MessageDlg('-1: 讀取畫面參數檔<'+sL_IniFileName +'>失敗',mtError, [mbOK],0);
      Result := false;
      exit;
    end;

    G_LanguageSettingIni := TIniFile.Create(sL_IniFileName);
    Result := true;
end;

procedure TdtmMain.DeleteTmpIniFile;
var
    sL_IniFileName : String;
begin
    //刪除解密後之 ini 檔
    sL_IniFileName := sG_ExePath + '\' + TMP_SYS_INI_FILE_NAME;
    DeleteFile(sL_IniFileName);
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

end.
