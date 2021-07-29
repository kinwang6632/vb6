unit dtmCommU;

interface

uses
  SysUtils, Classes,Windows, Dialogs,IniFiles, DB, ADODB, DBClient;

type
  PSMSDb = ^TSMSDb;
  TSMSDb = record
    ConnSeq: Integer;
    OracleSid: String;
    Description: String;
    UserId: String;
    Password: String;
    ReportPasth: String;
    MisDbOwner: String;
    MisDbPassword: String;
    MisDbCompCode: String;
    MisDbCompName: String;
    MisDbSid: String;
    MisDbLink: String;
    SendSid : String;
    SendUserId:String;
    SendPassword:String;
  end;


type
  TdtmComm = class(TDataModule)
    InvConnection: TADOConnection;
    adoINV008: TADOQuery;
    LogConnection: TADOConnection;
    adoLog: TADOQuery;
    procedure DataModuleCreate(Sender: TObject);
    procedure DataModuleDestroy(Sender: TObject);
  private
    { Private declarations }
    FSMSDbList: TList;
    FMisDbOwner: String;     // 客服系統的 Owner
    FMisDbCompCode : String;
    FMisDbSID:String;
    sG_DbSID: String;
    sG_DbUserID: String;
    sG_DbPassword: String;

    sG_CanUseSend: Boolean;
  public
    FInvDefFile : String;
    function GetSMSDb(const aFileName: String): Boolean;
    function ConnectToSO(aObj: PSMSDb): Boolean;
    function GetLogData(aInvId,aCompId : String) : Boolean;
    function GetCompID : String;
    function GetIniField(const aFileName,aSection:String;const aReadType:Integer;const aValue : String):String;
    property GetOwner : String read FMisDbOwner;
    property GetSID : string read   FMisDbSID;
    property CanUseSend:Boolean read sG_CanUseSend;

    { Public declarations }
  end;

var
  dtmComm: TdtmComm;

implementation


{$R *.dfm}

{ TdtmComm }

function TdtmComm.GetSMSDb(const aFileName: String): Boolean ;
var
  aIndex: Integer;
  aIniFile: TIniFile;
  aStrList: TStringList;
  aSMSDbId: String;
  aObj: PSMSDb;
begin
  aStrList := TStringList.Create;
  try
    aIniFile := TIniFile.Create( aFileName );
    FSMSDbList.Clear;
    try

      aIniFile.ReadSection( ParamStr(4) + ParamStr(1), aStrList );
      for aIndex :=0 to aStrList.Count-1 do
      begin

        //aSMSDbId := aStrList.Strings[aIndex];
        aSMSDbId := ParamStr(4) + ParamStr(1);
        New( aObj );
        //aObj.ConnSeq := StrToInt( aSMSDbId );
        aObj.OracleSid := aIniFile.ReadString( aSMSDbId, 'SID', EmptyStr );
        //aObj.Description := aIniFile.ReadString( 'DATAAREA', aSMSDbId, EmptyStr );
        aObj.UserId := aIniFile.ReadString( aSMSDbId, 'SID_USER', EmptyStr );
        aObj.Password := aIniFile.ReadString( aSMSDbId, 'SID_PASSWORD', EmptyStr );
        //aObj.ReportPasth := aIniFile.ReadString( aSMSDbId, 'REPORT_PATH', EmptyStr );
        aObj.MisDbOwner := aIniFile.ReadString( aSMSDbId, 'MISDB_OWNER', EmptyStr );
        //aObj.MisDbPassword := aIniFile.ReadString( aSMSDbId, 'MISDB_PASSWORD', EmptyStr );
        aObj.MisDbCompCode := aIniFile.ReadString( aSMSDbId, 'MISDB_COMPCODE', EmptyStr );
        //aObj.MisDbCompName := aIniFile.ReadString( aSMSDbId, 'MISDB_COMPNAME', EmptyStr );
        //aObj.MisDbSid := aIniFile.ReadString( aSMSDbId, 'MISDB_SID', EmptyStr );
        //aObj.MisDbLink := aIniFile.ReadString( aSMSDbId, 'MISDB_DBLINK', EmptyStr );
        //#5922 SEND_COM增加服務別區別 By Kin 2011/02/17
        aObj.SendSid := aIniFile.ReadString('SEND_COM_' + ParamStr(1),'SID',EmptyStr);
        aObj.SendUserId := aIniFile.ReadString('SEND_COM_' + ParamStr(1),'SID_USER',EmptyStr);
        aObj.SendPassword := aIniFile.ReadString('SEND_COM_' + ParamStr(1) , 'SID_PASSWORD',EmptyStr);
        FSMSDbList.Add(aObj);

      end;
    finally
      aIniFile.Free;
    end;
  finally
    aStrList.Free;
  end;
  if ConnectToSO(aObj) then
    Result := True
  else
    Result := False;

end;

procedure TdtmComm.DataModuleCreate(Sender: TObject);
begin
  FSMSDbList := TList.Create;
end;

procedure TdtmComm.DataModuleDestroy(Sender: TObject);
begin
  FSMSDbList.Free;
end;

function TdtmComm.ConnectToSO(aObj: PSMSDb): Boolean;
begin
  Result := False;
  try
   if not Assigned( aObj ) then
     raise Exception.Create( '對應不到系統台資料區。' );
   InvConnection.Connected := False;
   InvConnection.ConnectionString :=Format(
     'Provider=MSDAORA.1;Password=%s;User ID=%s;' +
     'Data Source=%s;Persist Security Info=True',
     [aObj.Password, aObj.UserId, aObj.OracleSid] );
   InvConnection.Connected := True;
   FMisDbOwner := aObj.MisDbOwner;
//   FMisDbOwner := aObj.UserId;
   FMisDbCompCode := aObj.MisDbCompCode;
   FMisDbSID := aobj.OracleSid;
   if ( aObj.SendSid <> EmptyStr ) and ( aObj.SendUserId <> EmptyStr )
    and (aObj.SendPassword <> EmptyStr ) then
   begin
     LogConnection.Connected := False;
     LogConnection.ConnectionString := Format(
       'Provider=MSDAORA.1;Password=%s;User ID=%s;' +
      'Data Source=%s;Persist Security Info=True',
      [aObj.SendPassword, aObj.SendUserId , aObj.SendSid] );
      sG_CanUseSend := True;

   end else
   begin
     sG_CanUseSend := False;
   end;
   Result := True;
  except
    on E: Exception do
    begin
      MessageDlg( Format( '【客服系統】資料區連結失敗, 訊息:%s', [E.Message] ),mtWarning,[mbOK],0);
      Exit;
    end;
  end;
end;

function TdtmComm.GetLogData(aInvId,aCompId : String): Boolean;
  var aSQL : String;
begin
  Result := False;
  try
    LogConnection.Connected := True;
  except
    on E: Exception do
    begin
      MessageDlg( Format( '【傳送通知】資料區連結失敗, 訊息:%s', [E.Message] ),mtWarning,[mbOK],0);
      Exit;
    end;
  end;
  {INV036}
  {#5800 INV039 增加顯示READTIME欄位,所以其它Table要用Null}
  aSQL := aSQL + Format('SELECT 1 TYPE,MAIL TYPE2,BOOKINGTIME SENDTIME,        ' +
      ' NVL(SENDSTATE,0) SENDSTATE,SENDCNT,CURRENTSTATENAME,                   ' +
      ' RETURNCODE,RETURNMSG, INV_SENDSTATUS.CURRENTSTATE,NULL READTIME        ' +
      ' FROM INV036ERR , INV_SENDSTATUS                                        ' +
      ' WHERE INVID(+) = ''%s''                                                ' +
      ' AND INV036ERR.CURRENTSTATE = INV_SENDSTATUS.CURRENTSTATE(+)            ' +
      ' AND INV_SENDSTATUS.SENDTYPE(+) = 1                                     ' ,
      [ aInvId ] );
  aSQL := aSQL + ' UNION ALL ';
  aSQL := aSQL + Format('SELECT 1 TYPE,MAIL TYPE2, BOOKINGTIME SENDTIME,       ' +
      ' NVL(SENDSTATE,0) SENDSTATE,SENDCNT,CURRENTSTATENAME,                   ' +
      ' RETURNCODE,RETURNMSG,INV_SENDSTATUS.CURRENTSTATE,NULL READTIME         ' +
      ' FROM INV036LOG , INV_SENDSTATUS                                        ' +
      ' WHERE INVID(+) = ''%s''                                                ' +
      ' AND INV036LOG.CURRENTSTATE = INV_SENDSTATUS.CURRENTSTATE(+)            ' +
      ' AND INV_SENDSTATUS.SENDTYPE(+) = 1                                     ' ,
      [ aInvId ] );

  {INV037}

  aSQL := aSQL + ' UNION ALL ';
  aSQL := aSQL + Format('SELECT 2 TYPE,CONTMOBILE TYPE2, BOOKINGTIME SENDTIME,  ' +
      ' NVL(SENDSTATE,0) SENDSTATE,SENDCNT,CURRENTSTATENAME,                    ' +
      ' RETURNCODE,RETURNMSG,INV_SENDSTATUS.CURRENTSTATE,NULL READTIME          ' +
      ' FROM INV037ERR , INV_SENDSTATUS                                         ' +
      ' WHERE INVID(+) = ''%s''                                                 ' +
      ' AND INV037ERR.CURRENTSTATE = INV_SENDSTATUS.CURRENTSTATE(+)             ' +
      ' AND INV_SENDSTATUS.SENDTYPE(+) = 2                                      ' ,
      [ aInvId ] );

  aSQL := aSQL + ' UNION ALL ';
  aSQL := aSQL + Format('SELECT 2 TYPE,CONTMOBILE TYPE2, BOOKINGTIME SENDTIME,    ' +
      ' NVL(SENDSTATE,0) SENDSTATE,SENDCNT,CURRENTSTATENAME,                      ' +
      ' RETURNCODE,RETURNMSG, INV_SENDSTATUS.CURRENTSTATE,NULL READTIME           ' +
      ' FROM INV037LOG , INV_SENDSTATUS                                           ' +
      ' WHERE INVID(+) = ''%s''                                                   ' +
      ' AND INV037LOG.CURRENTSTATE = INV_SENDSTATUS.CURRENTSTATE(+)               ' +
      ' AND INV_SENDSTATUS.SENDTYPE(+) = 2                                        ' ,
      [ aInvId ] );

  {INV038}
  aSQL := aSQL + ' UNION ALL ';
  aSQL := aSQL + Format('SELECT 3 TYPE,FACISNO TYPE2, BOOKINGTIME SENDTIME,       ' +
      ' NVL(SENDSTATE,0) SENDSTATE,SENDCNT, NULL CURRENTSTATENAME,                ' +
      ' RETURNCODE,RETURNMSG,NULL CURRENTSTATE,NULL READTIME                      ' +
      ' FROM INV038ERR WHERE INVID = ''%s''                                       ' ,
      [ aInvId ] );

  aSQL := aSQL + ' UNION ALL ';
  aSQL := aSQL + Format('SELECT 3 TYPE,FACISNO TYPE2, BOOKINGTIME SENDTIME,       ' +
      ' NVL(SENDSTATE,0) SENDSTATE,SENDCNT, NULL CURRENTSTATENAME,                ' +
      ' RETURNCODE,RETURNMSG,NULL CURRENTSTATE,NULL READTIME                      ' +
      ' FROM INV038LOG WHERE INVID = ''%s''                                       ' ,
      [ aInvId ] );

  {INV039}
  {#5800 增加INV039增加顯示CURRENTSTATE By Kin 2012/07/06}
  aSQL := aSQL + ' UNION ALL ';
  aSQL := aSQL + Format('SELECT 4 TYPE,FACISNO TYPE2,BOOKINGTIME SENDTIME,        ' +
      ' NVL(SENDSTATE,0) SENDSTATE,SENDCNT,CURRENTSTATENAME,                      ' +
      ' RETURNCODE,RETURNMSG,INV_SENDSTATUS.CURRENTSTATE,READTIME                 ' +
      ' FROM INV039ERR , INV_SENDSTATUS                                           ' +
      ' WHERE INVID = ''%s''                                                      ' +
      ' AND INV039ERR.CURRENTSTATE = INV_SENDSTATUS.CURRENTSTATE(+)               ' +
      ' AND INV_SENDSTATUS.SENDTYPE(+) = 4                                        ',
      [ aInvId ] );

  aSQL := aSQL + ' UNION ALL ';
  aSQL := aSQL + Format('SELECT 4 TYPE,FACISNO TYPE2,BOOKINGTIME SENDTIME,        ' +
      ' NVL(SENDSTATE,0) SENDSTATE,SENDCNT, CURRENTSTATENAME,                     ' +
      ' RETURNCODE,RETURNMSG,INV_SENDSTATUS.CURRENTSTATE,READTIME                 ' +
      ' FROM INV039LOG , INV_SENDSTATUS                                           ' +
      ' WHERE INVID = ''%s''                                                      ' +
      ' AND INV039LOG.CURRENTSTATE = INV_SENDSTATUS.CURRENTSTATE(+)               ' +
      ' AND INV_SENDSTATUS.SENDTYPE(+) = 4                                        ',
      [ aInvId ] );
  aSQL := 'SELECT A.* FROM ( ' + aSQL + ') A ORDER BY TYPE ASC,SENDTIME DESC';
  adoLog.Close;
  adoLog.SQL.Clear;
  adoLog.SQL.Text := aSQL;
  adoLog.Open;
  Result := True;
end;

function TdtmComm.GetCompID: String;
begin
  Result := FMisDbCompCode;
end;

function TdtmComm.GetIniField(const aFileName,aSection: String;
  const aReadType: Integer;const aValue:String): String;
var
  aIniFile: TIniFile;
  aStrList,aValueList : TStringList;
  aIndex : Integer;
  aRet : String;

begin
  try
    Result := EmptyStr;
    aRet := EmptyStr;
    aStrList := TStringList.Create;
    aValueList := TStringList.Create;
    try
      aIniFile := TIniFile.Create( aFileName );
      if aReadType = 2 then
      begin
        aIniFile.ReadSection('InvFields',aStrList);

        for aIndex := 0 to aStrList.Count-1 do
        begin
          if aRet = EmptyStr then
          begin
            aRet:=aStrList[aIndex];
          end else
          begin
            aRet := aRet + ',' + aStrList[aIndex];
          end;
        end;
        Result := aRet;
        Exit;
      end;
      aIniFile.ReadSection('InvFields',aStrList);
      if aReadType = 0 then
      begin
        if aStrList.IndexOf( aSection ) >= 0 then
        begin
          Result := aIniFile.ReadString('InvFields',aSection,EmptyStr);
        end;
      end else
      begin
         if aIniFile.SectionExists( aSection ) then
         begin
           if aValue <> EmptyStr then
           begin
             Result := aIniFile.ReadString( aSection,aValue,'其它');
           end else
           begin
             Result := aIniFile.ReadString( aSection,'NULL',EmptyStr);
           end;
         end;

      end;
    except
      Result := EmptyStr;
    end;
  finally
    aIniFile.Free;
    aStrList.Free;
    aValueList.Free;
  end;


end;

end.
