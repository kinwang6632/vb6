unit cbSoModule;

interface

uses
  Windows, SysUtils, Classes, Variants, DB, ADODB, DBClient, cbAppClass,
  Encryption_TLB;

type
  TSoDataModule = class(TDataModule)
    CAConnection: TADOConnection;
    CADataReader: TADOQuery;
    SoConnection: TADOConnection;
    SoDataReader: TADOQuery;
    cdsInput: TClientDataSet;
    CADataWriter: TADOQuery;
    AccessConnection: TADOConnection;
    AccessDataReader: TADOQuery;
    cdsNoCall: TClientDataSet;
    cdsCall: TClientDataSet;
    AccessDataWriter: TADOQuery;
    procedure DataModuleCreate(Sender: TObject);
    procedure DataModuleDestroy(Sender: TObject);
  private
    { Private declarations }
    FUser: TUser;
    FSoList: TList;
    FPassObj: Password;
    function ConnectToAccess(const aFileName: String; var aErrMsg: String): Boolean;
    function GetSoList(var aErrMsg: String): Boolean;
    procedure DisconnectFromAccess;
    procedure CleanupSoList;
    procedure SetSwitchComp(aCompStr: String);
  public
    { Public declarations }
    function DbLogin(const aCompCode: String; var aErrMsg: String): Boolean;
    function DbLoginSo(const aCompCode: String; var aErrMsg: String): Boolean;
    function DbLoginWareHouse(const aCompCode: String; var aErrMsg: String): Boolean;
    procedure DbLogout(const aCompCode: String);
    {}
    function IsWareHouse(const Args: array of string): Boolean;
    function UserAuthorize(var aErrMsg: String): Boolean;
    function UserAuthorizeSo(var aErrMsg: String): Boolean;
    function UserAuthorizeWareHouse(var aErrMsg: String): Boolean;
    {}
    procedure PrepareInputDataSet;
    procedure UnPrepareInputDataSet;
    procedure PrepareNoCallDataSet;
    procedure UnPrepareNoCallDataSet;
    procedure PrepareCallDataSet;
    procedure UnPrepareCallDataSet;
    {}
    procedure RecordToObject(aDataSet: TDataSet; aObj: TCARecycle);
    function ObjectToInsertRecord(aDataSet: TDataSet; aObj: TCARecycle; var aErrMsg: String): Boolean;
    function ObjectToEditRecord(aDataSet: TDataSet; aObj: TCARecycle; var aErrMsg: String): Boolean;
    {}
    function InseretInputData(var aErrMsg: String): Boolean;
    {}
    function VdInputData(aObj: TCARecycle; var aErrMsg: String): Boolean;
    function VdInputDataSO(aObj: TCARecycle; var aErrMsg: String): Boolean;
    function VdInputDataWareHouse(aObj: TCARecycle; var aErrMsg: String): Boolean;
    {}
    function VdDuplicateIcc(aObj: TCARecycle; aDataSet: TDataSet; var aErrMsg: String): Boolean;
    {}
    function GetOtherInfo(aObj: TCARecycle; var aErrMsg: String): Boolean;
    function GetOtherInfoSo(aObj: TCARecycle; var aErrMsg: String): Boolean;
    function GetOtherInfoWareHouse(aObj: TCARecycle; var aErrMsg: String): Boolean;

    {}
    function ConfigLoadFromFile(var aErrMsg: String): Boolean;
    property SoList: TList read FSoList;
    property User: TUser read FUser;
    {}
    function AddToRecycleTable(aDataSet: TDataSet; var aErrMsg: String): Boolean;
    function QueryProcessRecord(const aCondition: Integer; var aErrMsg: String;
      aConText1, aConText2: String): Boolean;
    function ConfirmData(aDataSet: TDataSet; var aCount: Integer; var aErrMsg: String): Boolean;
    {}
    function SaveDefaultSo(aCompCode: String; var aErrMsg: String): Boolean;
  end;

var
  SoDataModule: TSoDataModule;

implementation

uses cbMain, cbUtilis;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

{ TSoDataModule }

procedure TSoDataModule.DataModuleCreate(Sender: TObject);
begin
  FUser := TUser.Create;
  FSoList := TList.Create;
  FPassObj := CoPassword.Create;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDataModule.DataModuleDestroy(Sender: TObject);
begin
  FUser.Free;
  CleanupSoList;
  FPassObj := nil;
end;

{ ---------------------------------------------------------------------------- }

function TSoDataModule.DbLogin(const aCompCode: String; var aErrMsg: String): Boolean;
begin
  { 倉管暫時跟一般系統台一樣, 設定檔設定 --> 為新台北 }
  if IsWareHouse( [aCompCode] )  then
    Result := DbLoginSo( aCompCode, aErrMsg )
  else
    Result := DbLoginSo( aCompCode, aErrMsg )
end;

{ ---------------------------------------------------------------------------- }

function TSoDataModule.DbLoginSo(const aCompCode: String; var aErrMsg: String): Boolean;
var
  aSoInfo: TSo;
  aIndex: Integer;
  aFound: Boolean;
  aConnection: TADOConnection;
begin
  aErrMsg := EmptyStr;
  aSoInfo := nil;
  aFound := False;
  for aIndex := 0 to FSoList.Count - 1 do
  begin
    aSoInfo := TSo( FSoList[aIndex] );
    aFound := ( aSoInfo.CompCode = aCompCode ) ;
    if aFound then Break;
  end;
  if ( not aFound ) then
  begin
    aErrMsg := Format( '查無此系統台代碼(%s)。', [aCompCode] );
    Result := False;
    Exit;
  end;
  aConnection := SoConnection;
  if ( Pos( 'COM', UpperCase( aSoInfo.CompName ) ) > 0 ) then
    aConnection := CAConnection;
  aConnection.Connected := False;
  aConnection.ConnectionString := Format(
    'Provider=MSDAORA.1;Password=%s;User ID=%s;Data Source=%s;Persist Security Info=True',
    [aSoInfo.Password, aSoInfo.UserId, aSoInfo.Aliase] );
  try
    aConnection.Open;
  except
    on E: Exception do aErrMsg := E.Message;
  end;
  Result := aConnection.Connected;
end;

{ ---------------------------------------------------------------------------- }

function TSoDataModule.DbLoginWareHouse(const aCompCode: String;
  var aErrMsg: String): Boolean;
begin
  aErrMsg := EmptyStr;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDataModule.DbLogout(const aCompCode: String);
var
  aSoInfo: TSo;
  aIndex: Integer;
  aFound: Boolean;
  aConnection: TADOConnection;
begin
  aSoInfo := nil;
  aFound := False;
  for aIndex := 0 to FSoList.Count - 1 do
  begin
    aSoInfo := TSo( FSoList[aIndex] );
    aFound := ( aSoInfo.CompCode = aCompCode ) ;
    if aFound then Break;
  end;
  if not aFound then Exit;
  aConnection := SoConnection;
  if ( Pos( 'COM', UpperCase( aSoInfo.CompName ) ) > 0 ) then
    aConnection := CAConnection;
  if aConnection.Connected then aConnection.Close;  
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDataModule.PrepareInputDataSet;
begin
  if VarIsNull( cdsInput.Data ) then cdsInput.CreateDataSet;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDataModule.UnPrepareInputDataSet;
begin
  if not VarIsNull( cdsInput.Data ) then cdsInput.EmptyDataSet;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDataModule.PrepareNoCallDataSet;
begin
  if VarIsNull( cdsNoCall.Data ) then cdsNoCall.CreateDataSet;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDataModule.UnPrepareNoCallDataSet;
begin
  if not VarIsNull( cdsNoCall.Data ) then cdsNoCall.EmptyDataSet;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDataModule.PrepareCallDataSet;
begin
  if VarIsNull( cdsCall.Data ) then cdsCall.CreateDataSet;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDataModule.UnPrepareCallDataSet;
begin
  if not VarIsNull( cdsCall.Data ) then cdsCall.EmptyDataSet;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDataModule.RecordToObject(aDataSet: TDataSet;
  aObj: TCARecycle);
begin
  if ( not aDataSet.Active ) or ( aDataSet.IsEmpty ) then Exit;
  {}
  aObj.SeqNo := aDataSet.FieldByName( 'SEQNO' ).AsString;
  {}
  aObj.IccNo := aDataSet.FieldByName( 'ICCNO' ).AsString;
  {}
  aObj.StbNo := aDataSet.FieldByName( 'STBNO' ).AsString;
  {}
  aObj.StbFlag := aDataSet.FieldByName( 'STBFLAG' ).AsInteger;
  {}
  aObj.StbAutoFlag := aDataSet.FieldByName( 'STBAUTOCB' ).AsInteger;
  {}
  aObj.CompCode := aDataSet.FieldByName( 'COMPCODE' ).AsInteger;
  {}
  aObj.RCStattDate := 0;
  if not VarIsNull( aDataSet.FieldByName( 'RCSTARTDATE' ).Value ) then
    aObj.RCStattDate := aDataSet.FieldByName( 'RCSTARTDATE' ).AsDateTime;
  {}
  aObj.RCEndDate := 0;  
  if not VarIsNull( aDataSet.FieldByName( 'RCENDDATE' ).Value ) then
    aObj.RCEndDate := aDataSet.FieldByName( 'RCENDDATE' ).AsDateTime;
  {}  
  aObj.Operator := aDataSet.FieldByName( 'OPERATOR' ).AsString;
  {}
  aObj.UpdTime := Now;
  if not VarIsNull( aDataSet.FieldByName( 'UPDTIME' ).Value ) then
    aObj.UpdTime := aDataSet.FieldByName( 'UPDTIME' ).AsDateTime;
  {}
  aObj.LastDownloadTime := 0;  
  if not VarIsNull( aDataSet.FieldByName( 'LASTDOWNLOADTIME' ).Value ) then
    aObj.LastDownloadTime := aDataSet.FieldByName( 'LASTDOWNLOADTIME' ).AsDateTime;
  {}  
  aObj.CMDFlag := aDataSet.FieldByName( 'CMDFLAG' ).AsInteger;
  {}
  aObj.RecycleText := aDataSet.FieldByName( 'RECYCLETEXT' ).AsString;
  {}
  aObj.TransNum := aDataSet.FieldByName( 'TRANSNUM' ).AsString;
  {}
  aObj.Confirm := aDataSet.FieldByName( 'CONFIRM' ).AsString;
  {}
  aObj.ErrMsg := aDataSet.FieldByName( 'ERR_MSG' ).AsString;
end;

{ ---------------------------------------------------------------------------- }

function TSoDataModule.InseretInputData(var aErrMsg: String): Boolean;
var
  aBookMark: TBookmark;
  aSql, aDownloadTime, aSeqNo: String;

  { ------------------------------------------------ }

  function GetSeqNo: String;
  begin
    CADataReader.Close;
    CADataReader.SQL.Text := ' select s_carecycle_seqno.nextval from dual ';
    CADataReader.Open;
    Result := CADataReader.Fields[0].AsString;
    CADataReader.Close;
  end;

  { ------------------------------------------------ }

begin
  aErrMsg := EmptyStr;
  aSql :=
    ' insert into carecycle ( seqno,                                    ' +
    '   icc_no, stb_no, stb_flag, stbautocb,                            ' +
    '   lastdownloadtime, compcode, rcstartdate, rcenddate, operator,   ' +
    '   updtime, cmdflag, recycletext, transnum, confirm, err_msg )     ' +
    ' values (  ''%s'', ''%s'', ''%s'', ''%s'',                         ' +
    '   to_date( ''%s'', ''YYYYMMDD HH24MISS'' ), compcode,             ' +
    '   sysdate, null, ''%s'', sysdate, ''%s'', ''%s'', ''%s'', ''N'',  ' +
    '   ''%s'' )                                                        ';
  {}
  cdsInput.DisableControls;
  try
    aBookMark := cdsInput.GetBookmark;
    try
      CAConnection.BeginTrans;
      try
        CADataWriter.Close;
        cdsInput.First;
        while not cdsInput.Eof do
        begin
          aDownloadTime := EmptyStr;
          if not VarIsNull( cdsInput.FieldByName( 'RCSTARTDATE' ).Value ) then
            aDownloadTime := FormatDateTime( 'YYYYMMDD hhnnss', cdsInput.FieldByName( 'RCSTARTDATE' ).AsDateTime );
          {}
          aSeqNo := GetSeqNo;
          aSeqNo :=
            Lpad( cdsInput.FieldByName( 'COMPCODE' ).AsString, 2, '0' ) +
            Lpad( aSeqNo, 10, '0' );
          {}
          CADataWriter.SQL.Text := Format( aSql, [
            aSeqNo,
            cdsInput.FieldByName( 'ICC_NO' ).AsString,
            cdsInput.FieldByName( 'STB_NO' ).AsString,
            cdsInput.FieldByName( 'STB_FLAG' ).AsString,
            cdsInput.FieldByName( 'STBAUTOCB' ).AsString,
            aDownloadTime,
            cdsInput.FieldByName( 'COMPCODE' ).AsString,
            cdsInput.FieldByName( 'OPERATOR' ).AsString,
            cdsInput.FieldByName( 'CMDFLAG' ).AsString,
            cdsInput.FieldByName( 'RECYCLETEXT' ).AsString,
            cdsInput.FieldByName( 'TRANSNUM' ).AsString,
            cdsInput.FieldByName( 'ERRMSG' ).AsString] );
          CADataWriter.ExecSQL;
          cdsInput.Next;
        end;
      except
        on E: Exception do
        begin
          aErrMsg := E.Message;
          if CAConnection.InTransaction then CAConnection.RollbackTrans;
        end;
      end;  
      cdsInput.GotoBookmark( aBookMark );
    finally
      cdsInput.FreeBookmark( aBookMark );
    end;
  finally
    cdsInput.EnableControls;
  end;
  Result := ( aErrMsg = EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

function TSoDataModule.VdInputData(aObj: TCARecycle;
  var aErrMsg: String): Boolean;
begin
  if FUser.IsWareHouse then
    Result := VdInputDataWareHouse( aObj, aErrMsg )
  else
    Result := VdInputDataSO( aObj, aErrMsg );
end;

{ ---------------------------------------------------------------------------- }

function TSoDataModule.VdInputDataSO(aObj: TCARecycle; var aErrMsg: String): Boolean;
begin
  aErrMsg := EmptyStr;
  SoDataReader.Close;
  try
    SoDataReader.SQL.Text := Format(
      ' SELECT USEFLAG FROM SOAC0201B WHERE          ' +
      '   FACISNO = ''%s'' AND COMPCODE = ''%s''     ',
      [aObj.IccNo, aObj.CompCode] );
    SoDataReader.Open;
    Result := ( SoDataReader.RecordCount > 0 );
    if ( not Result ) then
    begin
      aErrMsg := '設備檔查無此ICC卡號,請確認是否輸入正確。';
      Exit;
    end;
    SoDataReader.Close;
    {}
    SoDataReader.SQL.Text := Format(
      ' SELECT COUNT(1) COUNTS FROM SO004 WHERE        ' +
      '   FACISNO = ''%s'' AND COMPCODE = ''%d''       ' +
      '   AND INSTDATE IS NOT NULL                     ' +
      '   AND PRDATE IS NULL                           ',
      [aObj.IccNo, aObj.CompCode] );
    SoDataReader.Open;
    Result := ( SoDataReader.FieldByName( 'COUNTS' ).AsInteger <= 0 );
    SoDataReader.Close;
    if ( not Result ) then
    begin
      aErrMsg := '此ICC卡尚在客戶端使用中,請確認是否輸入正確。';
      Exit;
    end;
    {}
    SoDataReader.SQL.Text := Format(
      ' SELECT MODEFLAG FROM SOAC0201A                  ' +
      '  WHERE FACISNO = ''%s'' AND COMPCODE = ''%d''   ',
      [aObj.StbNo, aObj.CompCode] );
    SoDataReader.Open;
    Result := ( SoDataReader.RecordCount > 0 );
    if ( not Result ) then
    begin
      aErrMsg := '設備檔查無此STB序號,請確認是否輸入正確。';
      Exit;
    end;
    SoDataReader.Close;
    {}
    SoDataReader.SQL.Text := Format(
      ' SELECT COUNT(1) COUNTS FROM SO004 WHERE        ' +
      '   FACISNO = ''%s'' AND COMPCODE = ''%d''       ' +
      '   AND INSTDATE IS NOT NULL                     ' +
      '   AND PRDATE IS NULL                           ',
      [aObj.StbNo, aObj.CompCode] );
    SoDataReader.Open;
    Result := ( SoDataReader.FieldByName( 'COUNTS' ).AsInteger <= 0 );
    if ( not Result ) then
    begin
      aErrMsg := '此STB尚在客戶端使用中,請確認是否輸入正確。';
      Exit;
    end;
    SoDataReader.Close;
    {}
  finally
    SoDataReader.Close;
    Result := ( aErrMsg = EmptyStr );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TSoDataModule.VdInputDataWareHouse(aObj: TCARecycle;
  var aErrMsg: String): Boolean;
var
  aCompCode: String;
begin
  aErrMsg := EmptyStr;
  try
    CADataReader.SQL.Text := Format(
      ' SELECT USEFLAG FROM SOAC0201B WHERE          ' +
      '   FACISNO = ''%s''                           ',
      [aObj.IccNo] );
    CADataReader.Open;
    Result := ( CADataReader.RecordCount > 0 );
    CADataReader.Close;
    if ( not Result ) then
    begin
      aErrMsg := '設備檔查無此ICC卡號,請確認是否輸入正確。';
      Exit;
    end;
    CADataReader.Close;
    {}
    CADataReader.SQL.Text := Format(
      ' SELECT COMPCODE FROM SOAC0201B WHERE         ' +
      '   FACISNO = ''%s'' AND COMPCODE = ''0''      ',
      [aObj.IccNo] );
    CADataReader.Open;
    aCompCode := CADataReader.FieldByName( 'COMPCODE' ).AsString;
    CADataReader.Close;
    Result := ( aCompCode = '0' );
    if not Result then
    begin
      aErrMsg := '此ICC卡號已在其它系統台使用中。';
      Exit;
    end;
    {}
    CADataReader.SQL.Text := Format(
      ' SELECT USEFLAG FROM SOAC0201A WHERE          ' +
      '   FACISNO = ''%s''                           ',
      [aObj.StbNo] );
    CADataReader.Open;
    Result := ( CADataReader.RecordCount > 0 );
    CADataReader.Close;
    if ( not Result ) then
    begin
      aErrMsg := '設備檔查無此STB序號,請確認是否輸入正確。';
      Exit;
    end;
    CADataReader.Close;
    {}
    CADataReader.SQL.Text := Format(
      ' SELECT COMPCODE FROM SOAC0201A WHERE         ' +
      '   FACISNO = ''%s'' AND COMPCODE = ''0''      ',
      [aObj.StbNo] );
    CADataReader.Open;
    aCompCode := CADataReader.FieldByName( 'COMPCODE' ).AsString;
    CADataReader.Close;
    Result := ( aCompCode = '0' );
    if not Result then
    begin
      aErrMsg := '此STB已在其它系統使用中。';
      Exit;
    end;
  finally
    Result := ( aErrMsg = EmptyStr );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TSoDataModule.VdDuplicateIcc(aObj: TCARecycle;  aDataSet: TDataSet;
  var aErrMsg: String): Boolean;
begin
  aErrMsg := EmptyStr;
  try
    aDataSet.DisableControls;
    try
      if ( aDataSet.Locate( 'ICCNO', aObj.IccNo, [] ) ) then
      begin
        aErrMsg := '此ICC卡號重覆輸入, 請確認是否輸入正確。';
        Exit;
      end;
      if ( aDataSet.Locate( 'STBNO', aObj.StbNo, [] ) ) then
      begin
        aErrMsg := '此STB序號重覆輸入, 請確認是否輸入正確。';
        Exit;
      end;
      CADataReader.Close;
      CADataReader.SQL.Text := Format(
        ' SELECT COUNT(1) COUNTS FROM CARECYCLE   '  +
        '  WHERE COMPCODE = ''%d''                '  +
        '    AND ICC_NO = ''%s''                  '  +
        '    AND CMDFLAG IN ( 0, 1, 2, 4 )           ',
        [aObj.CompCode, aObj.IccNo] );
      CADataReader.Open;
      try
        if ( CADataReader.FieldByName( 'COUNTS' ).AsInteger > 0 ) then
        begin
          aErrMsg := '此ICC卡號已輸入且正在處理中,  請確認是否輸入正確。';
          Exit;
        end;
        {}
        CADataReader.Close;
        CADataReader.SQL.Text := Format(
          ' SELECT COUNT(1) COUNTS FROM CARECYCLE   '  +
          '  WHERE COMPCODE = ''%d''                '  +
          '    AND STB_NO = ''%s''                  '  +
          '    AND CMDFLAG IN ( 0, 1, 2, 4 )        ',
          [aObj.CompCode, aObj.StbNo] );
        CADataReader.Open;
        if ( CADataReader.FieldByName( 'COUNTS' ).AsInteger > 0 ) then
        begin
          aErrMsg := '此STB卡號已輸入且正在處理中,  請確認是否輸入正確。';
          Exit;
        end;
      finally
        CADataReader.Close;
      end;
    finally
      aDataSet.EnableControls;
    end;  
  finally
    Result := ( aErrMsg = EmptyStr );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TSoDataModule.GetOtherInfo(aObj: TCARecycle; var aErrMsg: String): Boolean;
begin
  if ( FUser.IsWareHouse ) then
    Result := GetOtherInfoWareHouse( aObj, aErrMsg )
  else
    Result := GetOtherInfoSo( aObj, aErrMsg );
end;

{ ---------------------------------------------------------------------------- }

function TSoDataModule.GetOtherInfoSo(aObj: TCARecycle; var aErrMsg: String): Boolean;
var
  aDate1, aDate2: TDateTime;
begin
  aErrMsg := EmptyStr;
  SoDataReader.Close;
  {}
  try
    SoDataReader.SQL.Text := Format(
      ' SELECT NVL( STBAUTOCB, 0 ) STBAUTOCB,   ' +
      '        LASTDOWNLOADTIME                 ' +
      '   FROM SO004                            ' +
      '  WHERE FACISNO = ''%s''                 ' +
      '    AND COMPCODE = ''%d''                ',
      [aObj.IccNo, aObj.CompCode] );
    SoDataReader.Open;
    aObj.StbAutoFlag := SoDataReader.FieldByName( 'STBAUTOCB' ).AsInteger;
    aDate1 := 0;
    if not VarIsNull( SoDataReader.FieldByName( 'LASTDOWNLOADTIME' ).Value ) then
      aDate1 := SoDataReader.FieldByName( 'LASTDOWNLOADTIME' ).AsDateTime;
    SoDataReader.Close;
    { 取曾經清除過點數最後日期 }
    CADataReader.Close;
    CADataReader.SQL.Text := Format(
      ' SELECT MAX( LASTDOWNLOADTIME ) LASTDOWNLOADTIME ' +
      '   FROM CARECYCLE                                ' +
      '  WHERE ICC_NO = ''%s''                          ', [aObj.IccNo] );
    CADataReader.Open;
    aDate2 := 0;
    if not VarIsNull( CADataReader.FieldByName( 'LASTDOWNLOADTIME' ).Value ) then
      aDate2 := CADataReader.FieldByName( 'LASTDOWNLOADTIME' ).AsDateTime;
    CADataReader.Close;
    { 取最後一個清除點數日 }
    if ( aDate1 <> 0 ) and ( aDate2 <> 0 ) then
      aObj.LastDownloadTime := IIF( aDate1 > aDate2, aDate1, aDate2 );
    {}
    SoDataReader.SQL.Text := Format(
      ' SELECT NVL( MODEFLAG, 0 ) MODEFLAG FROM SOAC0201A   ' +
      '  WHERE FACISNO = ''%s'' AND COMPCODE = ''%d''       ',
      [aObj.StbNo, aObj.CompCode] );
    SoDataReader.Open;
    aObj.StbFlag := SoDataReader.FieldByName( 'MODEFLAG' ).AsInteger;
    SoDataReader.Close;
  finally
    SoDataReader.Close;
    CADataReader.Close;
    Result := ( aErrMsg = EmptyStr );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TSoDataModule.GetOtherInfoWareHouse(aObj: TCARecycle; var aErrMsg: String): Boolean;
var
  aDate2: TDateTime;
begin
  aErrMsg := EmptyStr;
  {}
  try
    { 取曾經清除過點數最後日期 }
    CADataReader.Close;
    CADataReader.SQL.Text := Format(
      ' SELECT MAX( LASTDOWNLOADTIME ) LASTDOWNLOADTIME ' +
      '   FROM CARECYCLE                                ' +
      '  WHERE ICC_NO = ''%s''                          ', [aObj.IccNo] );
    CADataReader.Open;
    aDate2 := 0;
    if not VarIsNull( CADataReader.FieldByName( 'LASTDOWNLOADTIME' ).Value ) then
      aDate2 := CADataReader.FieldByName( 'LASTDOWNLOADTIME' ).AsDateTime;
    CADataReader.Close;
    aObj.LastDownloadTime := aDate2;
    {}
    CADataReader.SQL.Text := Format(
      ' SELECT NVL( MODEFLAG, 0 ) MODEFLAG FROM SOAC0201A   ' +
      '  WHERE FACISNO = ''%s'' AND COMPCODE = ''%d''       ',
      [aObj.StbNo, aObj.CompCode] );
    CADataReader.Open;
    aObj.StbFlag := CADataReader.FieldByName( 'MODEFLAG' ).AsInteger;
  finally
    CADataReader.Close;
    Result := ( aErrMsg = EmptyStr );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TSoDataModule.ObjectToInsertRecord(aDataSet: TDataSet;
  aObj: TCARecycle; var aErrMsg: String): Boolean;
begin
  aDataSet.DisableControls;
  try
    aErrMsg := EmptyStr;
    try
      aDataSet.Append;
      aDataSet.FieldByName( 'ICCNO' ).AsString := aObj.IccNo;
      {}
      aDataSet.FieldByName( 'STBNO' ).AsString := aObj.StbNo;
      {}
      aDataSet.FieldByName( 'STBFLAG' ).AsInteger := aObj.StbFlag;
      {}
      aDataSet.FieldByName( 'STBAUTOCB' ).AsInteger := aObj.StbAutoFlag;
      {}
      aDataSet.FieldByName( 'COMPCODE' ).AsInteger := aObj.CompCode;
      {}
      aDataSet.FieldByName( 'RCSTARTDATE' ).Value := Null;
      if ( aOBj.RCStattDate > 0 ) then
        aDataSet.FieldByName( 'RCSTARTDATE' ).AsDateTime := aObj.RCStattDate;
      {}
      aDataSet.FieldByName( 'RCENDDATE' ).Value := Null;
      if ( aObj.RCEndDate > 0 ) then
        aDataSet.FieldByName( 'RCENDDATE' ).AsDateTime := aObj.RCEndDate;
      {}
      aDataSet.FieldByName( 'OPERATOR' ).AsString := aObj.Operator;
      {}
      if ( aObj.UpdTime > 0 ) then
        aDataSet.FieldByName( 'UPDTIME' ).AsDateTime := aObj.UpdTime
      else
        aDataSet.FieldByName( 'UPDTIME' ).AsDateTime := Now;
      {}
      aDataSet.FieldByName( 'LASTDOWNLOADTIME' ).Value := Null;
      if ( aObj.LastDownloadTime > 0 ) then
        aDataSet.FieldByName( 'LASTDOWNLOADTIME' ).AsDateTime := aObj.LastDownloadTime;
      {}
      aDataSet.FieldByName( 'CMDFLAG' ).AsInteger := aObj.CMDFlag;
      aDataSet.FieldByName( 'RECYCLETEXT' ).AsString := aObj.RecycleText;
      aDataSet.FieldByName( 'TRANSNUM' ).AsString := aObj.TransNum;
      aDataSet.FieldByName( 'CONFIRM' ).AsString := aObj.Confirm;
      aDataSet.Post;
    except
      on E: Exception do aErrMsg := E.Message;
    end;
  finally
    aDataSet.EnableControls;
  end;
  Result := ( aErrMsg = EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

function TSoDataModule.ObjectToEditRecord(aDataSet: TDataSet;
  aObj: TCARecycle; var aErrMsg: String): Boolean;
begin
  aDataSet.DisableControls;
  try
    aErrMsg := EmptyStr;
    try
      if ( aDataSet.Locate( 'ICCNO', aObj.IccNo, [] ) ) then
      begin
        aDataSet.Edit;
        aDataSet.FieldByName( 'STBFLAG' ).AsInteger := aObj.StbFlag;
        {}
        aDataSet.FieldByName( 'STBAUTOCB' ).AsInteger := aObj.StbAutoFlag;
        {}
        aDataSet.FieldByName( 'COMPCODE' ).AsInteger := aObj.CompCode;
        {}
        aDataSet.FieldByName( 'RCSTARTDATE' ).Value := Null;
        if ( aOBj.RCStattDate > 0 ) then
          aDataSet.FieldByName( 'RCSTARTDATE' ).AsDateTime := aObj.RCStattDate;
        {}
        aDataSet.FieldByName( 'RCENDDATE' ).Value := Null;
        if ( aObj.RCEndDate > 0 ) then
          aDataSet.FieldByName( 'RCENDDATE' ).AsDateTime := aObj.RCEndDate;
        {}
        aDataSet.FieldByName( 'OPERATOR' ).AsString := aObj.Operator;
        {}
        if ( aObj.UpdTime > 0 ) then
          aDataSet.FieldByName( 'UPDTIME' ).AsDateTime := aObj.UpdTime
        else
          aDataSet.FieldByName( 'UPDTIME' ).AsDateTime := Now;
        {}
        aDataSet.FieldByName( 'LASTDOWNLOADTIME' ).Value := Null;
        if ( aObj.LastDownloadTime > 0 ) then
          aDataSet.FieldByName( 'LASTDOWNLOADTIME' ).AsDateTime := aObj.LastDownloadTime;
        {}
        aDataSet.FieldByName( 'CMDFLAG' ).AsInteger := aObj.CMDFlag;
        aDataSet.FieldByName( 'RECYCLETEXT' ).AsString := aObj.RecycleText;
        aDataSet.FieldByName( 'TRANSNUM' ).AsString := aObj.TransNum;
        aDataSet.FieldByName( 'CONFIRM' ).AsString := aObj.Confirm;
        aDataSet.FieldByName( 'ERRMSG' ).AsString := aObj.ErrMsg;
        aDataSet.Post;
      end;
    except
      on E: Exception do aErrMsg := EmptyStr;
    end;
  finally
    aDataSet.EnableControls;
  end;
  Result := ( aErrMsg = EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

function TSoDataModule.ConnectToAccess(const aFileName: String; var aErrMsg: String): Boolean;
const
  aDbPassword = 'cyc84177282';
begin
  aErrMsg := EmptyStr;
  AccessConnection.Connected := False;
  AccessConnection.ConnectionString := Format(
    'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=%s;Jet OLEDB:Database Password=%s;'+
    'Mode=Share Deny Read|Share Deny Write;', [aFileName, aDbPassword] );
  try
    AccessConnection.Connected := True;
  except
    on E: Exception do aErrMsg := E.Message;
  end;
  Result := AccessConnection.Connected;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDataModule.DisconnectFromAccess;
begin
  if AccessDataReader.Active then
    AccessDataReader.Close;
  if AccessConnection.Connected then
    AccessConnection.Close;
end;

{ ---------------------------------------------------------------------------- }

function TSoDataModule.ConfigLoadFromFile(var aErrMsg: String): Boolean;
begin
  Result := ConnectToAccess( fmMain.ConfigFileName, aErrMsg );
  if Result then
  begin
    try
      CleanupSoList;
      Result := GetSoList( aErrMsg );
    finally
      DisconnectFromAccess;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TSoDataModule.GetSoList(var aErrMsg: String): Boolean;
var
  aSoInfo: TSo;
begin
  aErrMsg := EmptyStr;
  AccessDataReader.Close;
  AccessDataReader.SQL.Text := ' SELECT * FROM SO_COMP ';
  try
    AccessDataReader.Open;
    AccessDataReader.First;
    while not AccessDataReader.Eof do
    begin
      aSoInfo := TSo.Create;
      aSoInfo.UserId := AccessDataReader.FieldByName( 'SO_LOGINUSER' ).AsString;
      aSoInfo.Password := AccessDataReader.FieldByName( 'SO_LOGINPASS' ).AsString;
      aSoInfo.CompCode := AccessDataReader.FieldByName( 'SO_COMP_CODE' ).AsString;
      aSoInfo.CompName := AccessDataReader.FieldByName( 'SO_COMP_NAME' ).AsString;
      aSoInfo.Aliase := AccessDataReader.FieldByName( 'SO_DBALIASE' ).AsString;
      aSoInfo.Active := ( AccessDataReader.FieldByName( 'ACTIVE' ).AsInteger = 1 );
      aSoInfo.CanSwitch := False;
      FSoList.Add( aSoInfo );
      AccessDataReader.Next;
    end;
  except
    on E: Exception do aErrMsg := E.Message;
  end;
  AccessDataReader.Close;
  Result := ( aErrMsg = EmptyStr ); 
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDataModule.CleanupSoList;
var
  aIndex: Integer;
begin
  for aIndex := 0 to FSoList.Count - 1 do
    TSo( FSoList[aIndex] ).Free;
  FSoList.Clear;
end;

{ ---------------------------------------------------------------------------- }

function TSoDataModule.IsWareHouse(const Args: array of string): Boolean;
var
  aIndex: Integer;
begin
  Result := False;
  for aIndex := Low( Args ) to High( Args ) do
  begin
    Result :=
     ( Args[aIndex] = '20' ) or
     ( Args[aIndex] = '21' ) or
     ( Args[aIndex] = '22' );
    if Result then Break; 
  end;
end;

{ ---------------------------------------------------------------------------- }

function TSoDataModule.UserAuthorize(var aErrMsg: String): Boolean;
begin
  if IsWareHouse( [FUser.CompCode] ) then 
    Result := UserAuthorizeWareHouse( aErrMsg )
  else
    Result := UserAuthorizeSo( aErrMsg );
end;

{ ---------------------------------------------------------------------------- }

function TSoDataModule.UserAuthorizeSo(var aErrMsg: String): Boolean;
var
  aEncyText, aGroupId, aUserName, aCompStr: String;
  aKey: WideString;
begin
  aErrMsg := EmptyStr;
  try
    SoDataReader.Close;
    try
      SoDataReader.SQL.Text := Format(
        ' SELECT USERID, USERNAME, PASSWORD, GROUPID, COMPSTR  ' +
        '   FROM SO026                                         ' +
        '  WHERE USERID = ''%s'' AND COMPCODE = ''%s''         ',
        [FUser.UserId, FUser.CompCode] );
      SoDataReader.Open;
      if ( SoDataReader.IsEmpty ) then
      begin
        aErrMsg := '無此登入帳號。';
        Exit;
      end;
      aKey := 'CS';
      aEncyText := FPassObj.Encrypt( FUser.Password, aKey );
      if ( SoDataReader.FieldByName( 'PASSWORD' ).AsString <> aEncyText ) then
      begin
        aErrMsg := '密碼錯誤。';
        Exit;
      end;
      aUserName := SoDataReader.FieldByName( 'USERNAME' ).AsString;
      aGroupId := SoDataReader.FieldByName( 'GROUPID' ).AsString;
      aCompStr := SoDataReader.FieldByName( 'COMPSTR' ).AsString;
      FUser.Admin := False; 
      FUser.IsWareHouse := IsWareHouse( [FUser.CompCode] );
      SoDataReader.Close;
      {}
      SoDataReader.SQL.Text := Format(
          ' SELECT MID FROM SO029 WHERE MID = ''%s''    ' +
          '    AND GROUP%s = ''1''                      ', ['SO7X00',aGroupId] );
      SoDataReader.Open;    
      if ( SoDataReader.IsEmpty ) then
      begin
        aErrMsg := '無權限可執行回收作業系統功能。';
        Exit;
      end;
      SetSwitchComp( aCompStr );
      {}
      FUser.UserName := aUserName;
      FUser.ComputerName := cbUtilis.GetComputerName;
    finally
      SoDataReader.Close;
    end;
  finally
    Result := ( aErrMsg = EmptyStr );
  end;
end;

{ ---------------------------------------------------------------------------- }


function TSoDataModule.UserAuthorizeWareHouse(var aErrMsg: String): Boolean;
var
  aEncyText, aGroupId, aUserName, aCompStr: String;
  aKey: WideString;
begin
  aErrMsg := EmptyStr;
  try
    SoDataReader.Close;
    try
      SoDataReader.SQL.Text := Format(
        ' SELECT USERID, USERNAME, PASSWORD, GROUPID, COMPSTR  ' +
        '   FROM SO026                                         ' +
        '  WHERE USERID = ''%s''                               ',
        [FUser.UserId] );
      SoDataReader.Open;
      if ( SoDataReader.IsEmpty ) then
      begin
        aErrMsg := '無此登入帳號。';
        Exit;
      end;
      aKey := 'CS';
      aEncyText := FPassObj.Encrypt( FUser.Password, aKey );
      if ( SoDataReader.FieldByName( 'PASSWORD' ).AsString <> aEncyText ) then
      begin
        aErrMsg := '密碼錯誤。';
        Exit;
      end;
      { 系統管理員 GroupId = 200 死的 }
      aUserName := SoDataReader.FieldByName( 'USERNAME' ).AsString;
      aGroupId := SoDataReader.FieldByName( 'GROUPID' ).AsString;
      aCompStr := SoDataReader.FieldByName( 'COMPSTR' ).AsString;
      FUser.Admin := False;
      FUser.IsWareHouse := IsWareHouse( [FUser.CompCode] );
      SoDataReader.Close;
      {}
      SoDataReader.SQL.Text := Format(
          ' SELECT MID FROM SO029 WHERE MID = ''%s''    ' +
          '    AND GROUP%s = ''1''                      ', ['SO7X00',aGroupId] );
      SoDataReader.Open;    
      if ( SoDataReader.IsEmpty ) then
      begin
        aErrMsg := '無權限可執行回收作業系統功能。';
        Exit;
      end;
      SetSwitchComp( aCompStr );
      {}
      FUser.UserName := aUserName;
      FUser.ComputerName := cbUtilis.GetComputerName;
    finally
      SoDataReader.Close;
    end;
  finally
    Result := ( aErrMsg = EmptyStr );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDataModule.SetSwitchComp(aCompStr: String);
var
  aCompId: String;
  aFound: Boolean;
  aIndex: Integer;
begin
  if ( Trim( aCompStr ) = EmptyStr ) then Exit;
  repeat
     aCompId := ExtractValue( aCompStr );
     for aIndex := 0 to FSoList.Count - 1 do
     begin
       aFound := ( TSo( FSoList[aIndex] ).CompCode = aCompId );
       if aFound then
       begin
         TSo( FSoList[aIndex] ).CanSwitch := True;
         Break;
       end;
     end;
  until ( aCompStr = EmptyStr )
end;

{ ---------------------------------------------------------------------------- }

function TSoDataModule.AddToRecycleTable(aDataSet: TDataSet;
  var aErrMsg: String): Boolean;
var
  aRcDate, aDownloadTime: String;

  { ------------------------------------------------ }

  function GetSeqNo: String;
  begin
    CADataReader.Close;
    CADataReader.SQL.Text := ' select s_carecycle_seqno.nextval from dual ';
    CADataReader.Open;
    Result := CADataReader.Fields[0].AsString;
    CADataReader.Close;
  end;

  { ------------------------------------------------ }

var
  aSeqNo: String;
begin
  aErrMsg := EmptyStr;
  CAConnection.BeginTrans;
  try
    aDataSet.DisableControls;
    try
      aDataSet.First;
      while not aDataSet.Eof do
      begin
        aRcDate := EmptyStr;
        if not VarIsNull( aDataSet.FieldByName( 'RCSTARTDATE' ).Value ) then
          aRcDate := FormatDateTime( 'YYYYMMDD', aDataSet.FieldByName( 'RCSTARTDATE' ).AsDateTime );
        {}
        aDownloadTime := EmptyStr;
        if not VarIsNull( aDataSet.FieldByName( 'LASTDOWNLOADTIME' ).Value ) then
           aDownloadTime := FormatDateTime( 'YYYYMMDD', aDataSet.FieldByName( 'LASTDOWNLOADTIME' ).AsDateTime );
        {}
        aSeqNo := GetSeqNo;
        aSeqNo :=
          Lpad( cdsInput.FieldByName( 'COMPCODE' ).AsString, 2, '0' ) +
          Lpad( aSeqNo, 10, '0' );
        {}
        CADataWriter.SQL.Text := Format(
          ' INSERT INTO CARECYCLE ( SEQNO,                               ' +
          '   ICC_NO, STB_NO, STB_FLAG,                                  ' +
          '   STBAUTOCB, COMPCODE, RCSTARTDATE, RCENDDATE,               ' +
          '   OPERATOR, UPDTIME, LASTDOWNLOADTIME, CMDFLAG,              ' +
          '   RECYCLETEXT, TRANSNUM, CONFIRM )                           ' +
          ' VALUES ( ''%s'',                                             ' +
          '   ''%s'', ''%s'', ''%s'',                                    ' +
          '   ''%s'', ''%s'', TO_DATE( ''%s'', ''YYYYMMDD'' ), NULL,     ' +
          '   ''%s'', SYSDATE, TO_DATE( ''%s'', ''YYYYMMDD'' ), ''%s'',  ' +
          '   ''%s'', ''%s'', ''N''  )                                   ',
          [aSeqNo,
           aDataSet.FieldByName( 'ICCNO' ).AsString,
           aDataSet.FieldByName( 'STBNO' ).AsString,
           aDataSet.FieldByName( 'STBFLAG' ).AsString,
           aDataSet.FieldByName( 'STBAUTOCB' ).AsString,
           aDataSet.FieldByName( 'COMPCODE' ).AsString,
           aRcDate,
           aDataSet.FieldByName( 'OPERATOR' ).AsString,
           aDownloadTime,
           aDataSet.FieldByName( 'CMDFLAG' ).AsString,
           aDataSet.FieldByName( 'RECYCLETEXT' ).AsString,
           aDataSet.FieldByName( 'TRANSNUM' ).AsString] );
        CADataWriter.ExecSQL;
        aDataSet.Next;
      end;
      CAConnection.CommitTrans;
    finally
      aDataSet.EnableControls;
    end;
  except
    on E: Exception do
    begin
      if CAConnection.InTransaction then CAConnection.RollbackTrans;
      aErrMsg := E.Message;
    end;
  end;
  Result := ( aErrMsg = EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

function TSoDataModule.QueryProcessRecord(const aCondition: Integer;
  var aErrMsg: String; aConText1, aConText2: String): Boolean;
var
  aSql, aWhere, aOrder: String;
  aCmd: TRecycleCommand;
begin
  aSql :=
    ' select icc_no, stb_no, stb_flag, stbautocb, compcode,    ' +
    '        rcstartdate, rcenddate, operator, updtime,        ' +
    '        lastdownloadtime, cmdflag, recycletext, transnum, ' +
    '        confirm, err_msg, seqno                           ' +
    '   from carecycle                                         ' +
    '  where 1 = 1                                             ';
  {}
  aOrder := ' order by rcstartdate ';
  {}
  case aCondition of
    1: //已完成, 待確認
      begin
          aWhere := Format( ' and compcode = ''%s'' and cmdflag = 2 and confirm = ''N'' ',
          [FUser.CompCode] );
      end;
    2: //處理中
      begin
        aWhere := Format( ' and compcode = ''%s'' and cmdflag in ( 0,1 ) and confirm = ''N'' ',
          [FUser.CompCode] );
      end;
    3: //處理失敗
      begin
        aWhere := Format( ' and compcode = ''%s'' and cmdflag = 4 and confirm = ''N'' ', [FUser.CompCode] );
      end;
    4: //ICC
      begin
        aWhere := Format( ' and compcode = ''%s'' and icc_no = ''%s'' ', [FUser.CompCode, aConText1] );
      end;
    5: //處理日期
      begin
        if ( aConText1 <> EmptyStr ) and ( aConText2 <> EmptyStr ) then
        begin
          aWhere := Format(
            ' and compcode = ''%s'' and rcstartdate between to_date( ''%s'', ''YYYYMMDD'' ) and to_date( ''%s'', ''YYYYMMDD'' ) ',
            [FUser.CompCode, aConText1, aConText2] );
        end else
        begin
          aConText1 := Nvl( aConText1, aConText2 );
          aWhere := Format(
            ' and compcode = ''%s'' and rcstartdate = to_date( ''%s'', ''YYYYMMDD'' ) ',
            [FUser.CompCode, aConText1] );
        end;
      end;
  else
    aWhere := Format( ' and compcode = ''%s'' and cmdflag in ( 0,1) ', [FUser.CompCode] );
  end;
  {}
  aCmd := TRecycleCommand.Create( 0 );
  {}
  CADataReader.Close;
  try
    CADataReader.SQL.Text := ( aSql + aWhere + aOrder );
    CADataReader.Open;
    CADataReader.First;
    while not CADataReader.Eof do
    begin
      aCmd.AutoCallBack := CADataReader.FieldByName( 'STBAUTOCB' ).AsInteger;
      aCmd.CmdText := CADataReader.FieldByName( 'RECYCLETEXT' ).AsString;
      if ( aCmd.AutoCallBack = 0 ) then
      begin
        cdsNoCall.Append;
        cdsNoCall.FieldByName( 'SEQNO' ).AsString := CADataReader.FieldByName( 'SEQNO' ).AsString;
        cdsNoCall.FieldByName( 'ICCNO' ).AsString := CADataReader.FieldByName( 'ICC_NO' ).AsString;
        cdsNoCall.FieldByName( 'STBNO' ).AsString := CADataReader.FieldByName( 'STB_NO' ).AsString;
        cdsNoCall.FieldByName( 'STB_FLAG' ).AsString := CADataReader.FieldByName( 'STB_FLAG' ).AsString;
        cdsNoCall.FieldByName( 'STBAUTOCB' ).AsString := CADataReader.FieldByName( 'STBAUTOCB' ).AsString;
        cdsNoCall.FieldByName( 'COMPCODE' ).AsString := CADataReader.FieldByName( 'COMPCODE' ).AsString;
        cdsNoCall.FieldByName( 'RCSTARTDATE' ).Value := CADataReader.FieldByName( 'RCSTARTDATE' ).Value;
        cdsNoCall.FieldByName( 'RCENDDATE' ).Value := CADataReader.FieldByName( 'RCENDDATE' ).Value;
        cdsNoCall.FieldByName( 'OPERATOR' ).AsString := CADataReader.FieldByName( 'OPERATOR' ).AsString;
        cdsNoCall.FieldByName( 'UPDTIME' ).Value := CADataReader.FieldByName( 'UPDTIME' ).Value;
        cdsNoCall.FieldByName( 'LASTDOWNLOADTIME' ).Value := CADataReader.FieldByName( 'LASTDOWNLOADTIME' ).Value;
        cdsNoCall.FieldByName( 'CMDFLAG' ).AsString := CADataReader.FieldByName( 'CMDFLAG' ).AsString;
        cdsNoCall.FieldByName( 'RECYCLETEXT' ).AsString := CADataReader.FieldByName( 'RECYCLETEXT' ).AsString;
        cdsNoCall.FieldByName( 'TRANSNUM' ).AsString := CADataReader.FieldByName( 'TRANSNUM' ).AsString;
        cdsNoCall.FieldByName( 'CONFIRM' ).AsString := CADataReader.FieldByName( 'CONFIRM' ).AsString;
        cdsNoCall.FieldByName( 'ERRMSG' ).AsString := CADataReader.FieldByName( 'ERR_MSG' ).AsString;
        { STATE 的欄位是用來判斷, 是否從 Table 取出時就已經被確認過 }
        cdsNoCall.FieldByName( 'STATE' ).AsString := IIF( ( CADataReader.FieldByName( 'CONFIRM' ).AsString = 'Y' ), '1', '0' );
        cdsNoCall.FieldByName( 'CMD52_1' ).AsString := aCmd.FindTextByCmd( '52_1' );
        cdsNoCall.FieldByName( 'CMD8' ).AsString := aCmd.FindTextByCmd( '8' );
        cdsNoCall.FieldByName( 'CMD97_96' ).AsString := aCmd.FindTextByCmd( '97_96' );
        cdsNoCall.FieldByName( 'CMD7' ).AsString := aCmd.FindTextByCmd( '7' );
        cdsNoCall.FieldByName( 'CMD48' ).AsString := aCmd.FindTextByCmd( '48' );
        cdsNoCall.FieldByName( 'CMD52_2' ).AsString := aCmd.FindTextByCmd( '52_2' );
        cdsNoCall.Post;
      end else
      begin
        cdsCall.Append;
        cdsCall.FieldByName( 'SEQNO' ).AsString := CADataReader.FieldByName( 'SEQNO' ).AsString;
        cdsCall.FieldByName( 'ICCNO' ).AsString := CADataReader.FieldByName( 'ICC_NO' ).AsString;
        cdsCall.FieldByName( 'STBNO' ).AsString := CADataReader.FieldByName( 'STB_NO' ).AsString;
        cdsCall.FieldByName( 'STB_FLAG' ).AsString := CADataReader.FieldByName( 'STB_FLAG' ).AsString;
        cdsCall.FieldByName( 'STBAUTOCB' ).AsString := CADataReader.FieldByName( 'STBAUTOCB' ).AsString;
        cdsCall.FieldByName( 'COMPCODE' ).AsString := CADataReader.FieldByName( 'COMPCODE' ).AsString;
        cdsCall.FieldByName( 'RCSTARTDATE' ).Value := CADataReader.FieldByName( 'RCSTARTDATE' ).Value;
        cdsCall.FieldByName( 'RCENDDATE' ).Value := CADataReader.FieldByName( 'RCENDDATE' ).Value;
        cdsCall.FieldByName( 'OPERATOR' ).AsString := CADataReader.FieldByName( 'OPERATOR' ).AsString;
        cdsCall.FieldByName( 'UPDTIME' ).Value := CADataReader.FieldByName( 'UPDTIME' ).Value;
        cdsCall.FieldByName( 'LASTDOWNLOADTIME' ).Value := CADataReader.FieldByName( 'LASTDOWNLOADTIME' ).Value;
        cdsCall.FieldByName( 'CMDFLAG' ).AsString := CADataReader.FieldByName( 'CMDFLAG' ).AsString;
        cdsCall.FieldByName( 'RECYCLETEXT' ).AsString := CADataReader.FieldByName( 'RECYCLETEXT' ).AsString;
        cdsCall.FieldByName( 'TRANSNUM' ).AsString := CADataReader.FieldByName( 'TRANSNUM' ).AsString;
        cdsCall.FieldByName( 'CONFIRM' ).AsString := CADataReader.FieldByName( 'CONFIRM' ).AsString;
        cdsCall.FieldByName( 'ERRMSG' ).AsString := CADataReader.FieldByName( 'ERR_MSG' ).AsString;
        { STATE 的欄位是用來判斷, 是否從 Table 取出時就已經被確認過 }
        cdsCall.FieldByName( 'STATE' ).AsString := IIF( ( CADataReader.FieldByName( 'CONFIRM' ).AsString = 'Y' ), '1', '0' );
        cdsCall.FieldByName( 'CMD52_1' ).AsString := aCmd.FindTextByCmd( '52_1' );
        cdsCall.FieldByName( 'CMD62' ).AsString := aCmd.FindTextByCmd( '62' );
        cdsCall.FieldByName( 'CMD8' ).AsString := aCmd.FindTextByCmd( '8' );
        cdsCall.FieldByName( 'CMD60' ).AsString := aCmd.FindTextByCmd( '60' );
        cdsCall.FieldByName( 'CMD7' ).AsString := aCmd.FindTextByCmd( '7' );
        cdsCall.FieldByName( 'CMD48' ).AsString := aCmd.FindTextByCmd( '48' );
        cdsCall.FieldByName( 'CMD99' ).AsString := aCmd.FindTextByCmd( '99' );
        cdsCall.FieldByName( 'CMD52_2' ).AsString := aCmd.FindTextByCmd( '52_2' );
        cdsCall.Post;
      end;
      CADataReader.Next;
    end;
    cdsNoCall.First;
    cdsCall.First;
  except
    on E: Exception do aErrMsg := E.Message;
  end;
  CADataReader.Close;
  aCmd.Free;  
  Result := ( aErrMsg = EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

function TSoDataModule.ConfirmData(aDataSet: TDataSet; var aCount: Integer;
  var aErrMsg: String): Boolean;
var
  aSql, aWriteFlag: String;
begin
  aErrMsg := EmptyStr;
  aCount := 0;
  aSql :=
    ' update carecycle set cmdflag = ''%s'',     ' +
    '    operator = ''%s'', updtime = sysdate,   ' +
    '    rcenddate = trunc( sysdate ),           ' +
    '    confirm = ''Y''                         ' +
    '   where seqno = ''%s''                     ' +
    '     and icc_no = ''%s''                    ';
  CAConnection.BeginTrans;
  try
    aDataSet.DisableControls;
    try
      aDataSet.First;
      while not aDataSet.Eof do
      begin
        aWriteFlag := EmptyStr;
        { 成功/確認 }
        if ( ( aDataSet.FieldByName( 'CMDFLAG' ).AsString = '2' ) or
             ( aDataSet.FieldByName( 'CMDFLAG' ).AsString = '1' ) ) and
           ( aDataSet.FieldByName( 'CONFIRM' ).AsString = 'Y' ) then
        begin
          aWriteFlag := '3';
        end else
        { 失敗/取消 }
        if ( aDataSet.FieldByName( 'CMDFLAG' ).AsString = '4' ) and
           ( aDataSet.FieldByName( 'CONFIRM' ).AsString = 'Y' ) then
        begin
          aWriteFlag := '5';
        end;
        { STATE 的欄位是用來判斷, 是否從 Table 取出時就已經被確認過 }
        { 如果已經是備確認過, 就不須要下 SQL 更新資料 } 
        if ( aWriteFlag <> EmptyStr ) and
           ( aDataSet.FieldByName( 'STATE' ).AsString = '0' ) then
        begin
          CADataWriter.Close;
          CADataWriter.SQL.Text := Format( aSql, [
            aWriteFlag, FUser.UserName,
            aDataSet.FieldByName( 'SEQNO' ).AsString,
            aDataSet.FieldByName( 'ICCNO' ).AsString] );
          CADataWriter.ExecSQL;
          Inc( aCount );
        end;
        aDataSet.Next;
      end;
      { 刪掉確認 OK 的 }
      aDataSet.First;
      while not aDataSet.Eof do
      begin
        if ( aDataSet.FieldByName( 'CONFIRM' ).AsString = 'Y' ) and
           ( aDataSet.FieldByName( 'STATE' ).AsString = '0' ) then
          aDataSet.Delete
        else
          aDataSet.Next;
      end;
      CAConnection.CommitTrans;
    finally
      aDataSet.EnableControls;
    end;
  except
    on E: Exception do
    begin
      if CAConnection.InTransaction then CAConnection.RollbackTrans;
      aErrMsg := E.Message;
    end;
  end;
  Result := ( aErrMsg = EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

function TSoDataModule.SaveDefaultSo(aCompCode: String;
  var aErrMsg: String): Boolean;
begin
  Result := ConnectToAccess( fmMain.ConfigFileName, aErrMsg );
  if Result then
  begin
    try
      try
        AccessDataWriter.SQL.Text := ' update so_comp set active = 0 ';
        AccessDataWriter.ExecSQL;
        AccessDataWriter.SQL.Text := Format(
          ' update so_comp set active = 1 where so_comp_code = ''%s'' ',
          [aCompCode] );
        AccessDataWriter.ExecSQL;
      except
        on E: Exception do aErrMsg := E.Message;
      end;
    finally
      DisconnectFromAccess;
    end;
    Result := ( aErrMsg = EmptyStr );
  end;
end;

{ ---------------------------------------------------------------------------- }

end.
