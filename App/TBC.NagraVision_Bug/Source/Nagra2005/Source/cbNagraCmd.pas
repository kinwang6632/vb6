unit cbNagraCmd;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, cbClass, DB, DBClient,
  IdGlobal;

type

  { TNagraCommandBuilder }

  TNagraCommandBuilder = class(TObject)
  private
    FPrimalObj: PSendNagra;
    FCommandList: TStringList;
    FTemporaryList: TStringList;
    FTransactionNumberList: TStringList;
    FCASEnv: TCASEnv;
    FHighCmdEnv: TClientDataSet;
    FCmdTransEnv: TCmdTransEnv;
    procedure SetCASEnv(const aValue: TCASEnv);
    procedure SetHighCmdEnv(aValue: TClientDataSet);
    procedure SetCmdTransEnv(const aValue: TCmdTransEnv);
    procedure ClearList(aList: TStrings);
    procedure BuildCommandSetpByStep;
    procedure CloneObj(aSource, aTarget: PSendNagra);
    procedure HighObjExtractToLowObj(aCompositText: String);
    { 舊的 SetopBox 盒子用的 IRD 指令 }
    function GetIrdStructure(aObj: PSendNagra; aIrdObj: PIrdInfo): Boolean;
    { 新的 SetopBox 盒子用的 IRD 指令 }
    function GetIrdStructure2(aObj: PSendNagra; aIrdObj: PIrdInfo): Boolean;
    { 依不同的指令, 處理 Notes 欄位 }
    procedure SpecialLowCommandRule(aObj: PSendNagra);
    procedure SpecialHighCommandRule;
    procedure NormalHighCommandRule;
    function GetValue(const aFindField, aResultField: String;
      const Args: array of Variant): String;
    function CaclMailSegment(aMsg: String): Integer;
    function CaclMailText(var aMsg: String): String;
    function MailTextToUtf8(aMsg: String): String;
  protected
    function SplitToCommandList: Integer; virtual;
    function GetCommandHeader(aObj: PSendNagra): String; virtual;
    function GetCommandSection(aObj: PSendNagra): String; virtual;
    function GetCommandBody(aObj: PSendNagra): String; virtual;
    function GetCommandCount: Integer; virtual;
    function GetTransactionNumberGenerator(const aIdentify: String):
      TTransactionNumberGenerator;
    function GetMailIdGenerator(const aIdentify: String):
      TTransactionNumberGenerator;
    function GetCommandText(Index: Integer): String; virtual;
    function GetCommandObject(Index: Integer): PSendNagra; virtual;
  public
    constructor Create;
    destructor Destroy; override;
    procedure BuildCommand(aObj: PSendNagra); overload;
    procedure BuildCommand(aObj: PRecvNagra); overload;
    function BuildCommunicationCheck(aObj: PSendNagra): String; overload;
    function BuildCommunicationCheck(aObj: PRecvNagra): String; overload;
    procedure Clear;
    procedure ClearCommand; virtual;
    property CASEnv: TCASEnv read FCASEnv write SetCASEnv;
    property HighCmdEnv: TClientDataSet read FHighCmdEnv write SetHighCmdEnv;
    property CmdTransEnv: TCmdTransEnv read FCmdTransEnv write SetCmdTransEnv;
    property CommandCount: Integer read GetCommandCount;
    property CommandText[Index: Integer]: String read GetCommandText;
    property CommandObject[Index: Integer]: PSendNagra read GetCommandObject;
  end;

  { TNagraAckParser }

  TNagraAckParser = class
  private
    class function InternalParseFeedbackData(aObj: PRecvNagra; aCmdBody: String): Boolean;
  public
    class function ParserAcknowledge(aAckData: String): PSendNagra;
    class function ParseFeedbackData(aAckData: String): PRecvNagra;
  end;


  function HexToLength(const AHex: String): Integer;
  function LengthToHex(const ALen: Integer): String;
  function TranslateCommandType(ACmdId: Integer): String;


implementation

uses cbNagraDevice, cbUtilis;


{ ---------------------------------------------------------------------------- }

function HexToLength(const aHex: String): Integer;
var
  aTmp: String;
begin
  aTmp := '$' +
    IntToHex( Ord( AHex[1] ), 2 ) +
    IntToHex( Ord( AHex[2] ), 2 );
  Result := StrToInt( aTmp );
end;

{ ---------------------------------------------------------------------------- }

function LengthToHex(const ALen: Integer): String;
var
  aHex, aVal1, aVal2: String;
begin
  { 長度轉成 4 碼 16 進制 }
  aHex := IntToHex( ALen, 4 );
  { 每 2 碼一個單位, 轉成 10 進制後, 再取其 ASCII,
    Ex: Len --> 36 --> 轉 16 進制 = 0024
        0024 -->
        00 --> 轉 10 進制 = 0, 取 ASCII 為 #0,
        24 --> 轉 10 進制 = 36, 取 ASCII 為 $ }
  aVal1 := Chr( StrToInt( '$' + Copy( aHex, 1, 2 ) ) );
  aVal2 := Chr( StrToInt( '$' + Copy( aHex, 3, 2 ) ) );
  Result := ( aVal1 + aVal2 );
end;

{ ---------------------------------------------------------------------------- }

function TranslateCommandType(ACmdId: Integer): String;
begin
  { 請參考 EMC SasGwyStdSpe020705.pdf,
    第 15頁 --> 4.5.1 Root Header }
  case ACmdId of
    0..99: Result := '01';         { EMM }
    100..199: Result := '02';      { Control }
    200..299: Result := '04';      { Feedback }
    1000..1002: Result := '05';    { Operation }
  else
    Result := EmptyStr;
  end;
end;

{ ---------------------------------------------------------------------------- }

{ 將中文的 OPPV Event Name Big-5 轉成 UTF-8 碼 }
{ 因為東森 STB 顯示文字一律只吃 UTF-8 碼 }

function TranslateEventName(const AEventName: String): String;
begin
  Result := AEventName;
  if Result <> EmptyStr then
    Result := AnsiToUtf8( AEventName );
end;

{ ---------------------------------------------------------------------------- }

function TranslateSubBeginDate(ADateText: String): String;
var
  aConvertDate: TDateTime;
begin
  Result := ADateText;
  if ( Result = EmptyStr ) then Exit;
  ADateText := cbUtilis.FormatMaskText( '####/##/##', ADateText );
  if TryStrToDate( ADateText, aConvertDate ) then
  begin
    aConvertDate := cbUtilis.IIF( ( aConvertDate < Date ), Date, aConvertDate );
    Result := FormatDateTime( 'yyyymmdd', aConvertDate );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TranslateSubEndDate(AProductId, ADateText: String): String;
begin
  Result := ADateText;
  if ( Result = EmptyStr ) and ( AProductId <> EmptyStr ) then
    Result := '20201231';
end;

{ ---------------------------------------------------------------------------- }

{ Set Callback IP Address, 必須是完整的 IP 字串 }
{ Ex: 192.168.001.001 }

function TranslateIpAddr(const AIpAddr: String): String;
var
  aPos, aIndex, aCurrent: Integer;
  aIpSpec: array [1..4] of String;
begin
  for aIndex := Low( aIpSpec ) to High( aIpSpec ) do
    aIpSpec[aIndex] := '000';
  aIndex := 1;
  aCurrent := 0;
  aPos := 0;
  while ( aIndex <= Length( AIpAddr ) ) do
  begin
    if ( AIpAddr[aIndex] <> '.' ) and ( aIndex < Length( AIpAddr ) ) then
      Inc( aIndex )
    else begin
      Inc( aCurrent );
      Inc( aPos );
      aIpSpec[aCurrent] := Lpad( Copy( AIpAddr, aPos, ( aIndex - aPos ) ), 3, '0' );
      aPos := ( aIndex );
      Inc( aIndex );
    end;
  end;
  Result := Format( '%s.%s.%s.%s', [
    StringReplace( aIpSpec[1], #32, '0', [rfReplaceAll] ),
    StringReplace( aIpSpec[2], #32, '0', [rfReplaceAll] ),
    StringReplace( aIpSpec[3], #32, '0', [rfReplaceAll] ),
    StringReplace( aIpSpec[4], #32, '0', [rfReplaceAll] )] );
end;

{ ---------------------------------------------------------------------------- }

function TranslateCredit(ACredit: String): String;
var
  aPos: Integer;
  aSign, aDigit: String;
begin
  Result := Lpad( EmptyStr, 7, '0' );
  if ACredit = EmptyStr then Exit;
  aPos := Pos( '.', ACredit );
  if aPos <= 0 then
  begin
    { 假設 Credit 給的長度大於 5, 則 Replace 掉成字母 X, 讓 command 送出去後,
      變成是錯的, 因為 credit 長度大於 5 表示有問題 }
    if Length( ACredit ) > 5 then
      Result := Lpad( EmptyStr, 7, 'X' )
    else
      Result := Rpad( Lpad( ACredit, 5, '0' ), 7, '0' );
  end
  else begin
    aSign := Copy( ACredit, 1, ( aPos - 1 ) );
    aDigit := Copy( ACredit, ( aPos + 1 ), 2 );
    { 長度有錯, Replace 掉成字母 X }
    if Length( aSign ) > 5 then aSign := Lpad( EmptyStr, 5, 'X' );
    if Length( aDigit ) > 2 then aDigit := Lpad( EmptyStr, 2, 'X' );
    Result := ( Lpad( aSign, 5, '0' ) + Rpad( aDigit, 2, '0' ) );
  end;  
end;

{ ---------------------------------------------------------------------------- }

function TranslateSTBNoBySpecialRule(AObj: PSendNagra): String;
begin
  Result := Lpad( Copy( AObj.StbNo, 1, 10 ), 10, '0' );
  { 配對跟解配對都是用低階指令 Command 52 }
  { 配對的話 STB NO 給值 }
  { 解配對 STB NO 全部補 0 }
  if ( AObj.LowCmdId = '52' ) then
  begin
    { 以下為 Unpaire 的高階指令 }
    if ( AObj.HighCmdId = 'A2' ) or
       ( AObj.HighCmdId = 'C1' ) then
    begin
      Result := Lpad( EmptyStr, 10, '0' );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

{ 將 Purge 指令會用到的 Condition Date 做轉換 }
function TranslateConditionDateBySpecialRule(AObj: PSendNagra): String;
begin
  { 沒有 Return Path 的話, 一律是系統日 }
  Result := Nvl( AObj.ConditionDate, FormatDateTime( 'yyyymmdd', Date ) );
  { 若有 Return Path, 則 Condition Date 一律為 19920101 }
  if ( AObj.LowCmdId = '96' ) and ( AObj.HighCmdId = 'P3' ) then
  begin
    Result := '19920101';
  end;
end;

{ ---------------------------------------------------------------------------- }

{ TCommandBuilder }

constructor TNagraCommandBuilder.Create;
begin
  inherited Create;
  FHighCmdEnv := TClientDataSet.Create( nil );
  FCommandList := TStringList.Create;
  FTemporaryList := TStringList.Create;
  FTransactionNumberList := TStringList.Create;
  FTransactionNumberList.Sorted := True;
end;

{ ---------------------------------------------------------------------------- }

destructor TNagraCommandBuilder.Destroy;
var
  aFileName: String;
  aIndex: Integer;
begin
  FCommandList.Free;
  FTemporaryList.Free;
  aFileName :=  IncludeTrailingPathDelimiter( ExtractFilePath(
    ParamStr( 0 ) ) ) + DEFAULT_TRANSACTIONFILE;
  for aIndex := 0 to FTransactionNumberList.Count - 1 do
  begin
    TTransactionNumberGenerator(
      FTransactionNumberList.Objects[aIndex] ).SaveTo( aFileName );
    TTransactionNumberGenerator( FTransactionNumberList.Objects[aIndex] ).Free;
  end;
  FTransactionNumberList.Free;
  FHighCmdEnv.Free;
  inherited Destroy;
end;

{ ---------------------------------------------------------------------------- }

procedure TNagraCommandBuilder.SetCASEnv(const aValue: TCASEnv);
begin
  FCASEnv.IP := aValue.IP;
  FCASEnv.ControlPort := aValue.ControlPort;
  FCASEnv.FeedbackPort := aValue.FeedbackPort;
  FCASEnv.ControlSourceId := aValue.ControlSourceId;
  FCASEnv.ControlDestId := aValue.ControlDestId;
  FCASEnv.ControlMopPPId := aValue.ControlMopPPId;
  FCASEnv.FeedbackSourceId := aValue.FeedbackSourceId;
  FCASEnv.FeedbackDestId := aValue.FeedbackDestId;
  FCASEnv.FeedbackMopPPId := aValue.FeedbackMopPPId;
  FCASEnv.BoardCastMode := aValue.BoardCastMode;
  FCASEnv.AddressType := aValue.AddressType;
  FCASEnv.ControlTransId := aValue.ControlTransId;
  FCASEnv.FeedbackTransId := aValue.FeedbackTransId;
  FCASEnv.MailId := aValue.MailId;
end;

{ ---------------------------------------------------------------------------- }

procedure TNagraCommandBuilder.SetHighCmdEnv(aValue: TClientDataSet);
begin
  FHighCmdEnv.Data := aValue.Data;
end;

{ ---------------------------------------------------------------------------- }

procedure TNagraCommandBuilder.SetCmdTransEnv(const aValue: TCmdTransEnv);
begin
  FCmdTransEnv.UseSimulator := aValue.UseSimulator;
end;

{ ---------------------------------------------------------------------------- }

procedure TNagraCommandBuilder.ClearList(aList: TStrings);
var
  aIndex: Integer;
begin
  if not Assigned( aList ) then Exit;
  for aIndex := 0 to aList.Count - 1 do
  begin
    if Assigned( aList.Objects[aIndex] ) then
      Dispose( PSendNagra( aList.Objects[aIndex] ) );
  end;
  aList.Clear;
end;

{ ---------------------------------------------------------------------------- }

procedure TNagraCommandBuilder.BuildCommand(aObj: PSendNagra);
begin
  FPrimalObj := aObj;
  SplitToCommandList;
  BuildCommandSetpByStep;
end;

{ ---------------------------------------------------------------------------- }

procedure TNagraCommandBuilder.BuildCommand(aObj: PRecvNagra);
var
  aObj2: PSendNagra;
begin
  New( aObj2 );
  try
    aObj2.CompCode := aObj.CompCode;
    aObj2.HighCmdId := 'Z2';
    aObj2.LowCmdId := EmptyStr;
    { 要用 CmdStatus 判斷, 因為確定寫入資料庫了, 就可以 ACK 回去 }
    if aObj.CmdStatus <> 'C' then
    begin
      aObj2.HighCmdId := 'Z3';
      aObj2.LowCmdId := aObj.LowLevelCmd;
    end;
    aObj2.CaAckTransactionNumber := aObj.ResponseTransactionNumber;
    aObj.UpdTime := Now;
    BuildCommand( aObj2 );
    aObj.TransactionNumber := CommandObject[0].SmsTransactionNumber;
    aObj.CommandText := CommandObject[0].SmsCommandText;
    aObj.FullCommandText := CommandObject[0].SmsFullCommandText;
  finally
    Dispose( aObj2 );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TNagraCommandBuilder.SplitToCommandList: Integer;
var
  aIndex: Integer;
  aDetailCmdId: String;
begin
  ClearList( FTemporaryList );
  aDetailCmdId := GetValue( 'HIGH_LEVEL_CMD', 'LOW_LEVEL_CMD', [FPrimalObj.HighCmdId] );
  FPrimalObj.VodMoppId := GetValue( 'HIGH_LEVEL_CMD', 'VOD_MOPP_ID', [FPrimalObj.HighCmdId] );
  { 先將高階轉成低階指令 }
  HighObjExtractToLowObj( aDetailCmdId );
  { B7 收視權限複製,
    E2 Mail 傳送訊息
    E5 母子機配對
    E7 臨時驗證子機
    E11 Pop up 訊息
    E21 Mail 傳送訊息(全域廣播)
    這幾個高階指令特殊格式, 須要其它規則拆解 }
  if ( FPrimalObj.HighCmdId = 'A6' ) or
     ( FPrimalObj.HighCmdId = 'A8' ) or
     ( FPrimalObj.HighCmdId = 'A9' ) or
     ( FPrimalObj.HighCmdId = 'B7' ) or
     ( FPrimalObj.HighCmdId = 'E2' ) or
     ( FPrimalObj.HighCmdId = 'E5' ) or
     ( FPrimalObj.HighCmdId = 'E7' ) or
     ( FPrimalObj.HighCmdId = 'E11' ) or
     ( FPrimalObj.HighCmdId = 'E21' ) then
    SpecialHighCommandRule
  else
    NormalHighCommandRule;
  { 將 Temp List 轉寫到 Command List }
  ClearList( FCommandList );
  for aIndex := 0 to FTemporaryList.Count - 1 do
  begin
    FCommandList.AddObject( FTemporaryList.Strings[aIndex],
      FTemporaryList.Objects[aIndex] );
  end;
  { 不可以呼叫 ClearList 這個 Procedure, 因為會將 aObj Free 掉 }
  FTemporaryList.Clear;
  Result := FCommandList.Count;
end;

{ ---------------------------------------------------------------------------- }

function TNagraCommandBuilder.GetCommandHeader(aObj: PSendNagra): String;
var
  aTransNum, aCmdType, aSourceId, aDestId, aMoppId, aCreateDate: String;
  aTransGenerator: TTransactionNumberGenerator;
begin
  { 請參考 EMC SasGwyStdSpe020705.pdf,
    第 15頁 --> 4.5.1 Root Header }
  { CommandRoot Header =
     transaction number ( 9 碼 ) +
     command type       ( 2 碼 ) +
     source_id          ( 4 碼 ) +
     dest_id            ( 4 碼 ) +
     mop_ppid           ( 5 碼 ) +
     creation_date      { 8 碼 )    }
  if ( aObj.CompCode = FEEDBACK_COMPCODE ) then
  begin
    { Feedback 用的 TransactionNumbe }
    aTransGenerator := GetTransactionNumberGenerator( FCASEnv.FeedbackTransId );
    { transaction_number = FeedbackId(2碼) + 序號(7碼,不足補零, 右靠), 總長度 9 碼
      ex: 110000001 }
    aTransNum := Rpad( FCASEnv.FeedbackTransId, 2, '0' ) + Lpad( IntToStr(
      aTransGenerator.NextValue ), 7, '0' );
    aSourceId := Lpad( FCASEnv.FeedbackSourceId, 4, '0' );
    aDestId := Lpad( FCASEnv.FeedbackDestId, 4, '0' );
    aMoppId := Lpad( FCASEnv.FeedbackMopPPId, 5, '0' );
  end else
  begin
    aTransGenerator := GetTransactionNumberGenerator( FCASEnv.ControlTransId );
    { transaction_number = SourceId(2碼) + 序號(7碼,不足補零, 右靠), 總長度 9 碼
      ex: 120000001 }
    aTransNum := Lpad( FCASEnv.ControlTransId, 2, '0' ) + Lpad( IntToStr(
      aTransGenerator.NextValue ), 7, '0' );
    aSourceId := Lpad( FCASEnv.ControlSourceId, 4, '0' );
    aDestId := Lpad( FCASEnv.ControlDestId, 4, '0' );
    aMoppId := Lpad( FCASEnv.ControlMopPPId, 5, '0' );
    { CNS 假如是有關 VOD 的 指令, 則 MoppId 要讀取每個指令設定的 MoppId }
    if ( Trim( aObj.VodMoppId ) <> EmptyStr ) then
      aMoppId := Lpad( aObj.VodMoppId, 5, '0' );
  end;
  aObj.SmsTransactionNumber := aTransNum;
  aCmdType := TranslateCommandType( StrToInt( aObj.LowCmdId ) );
  aCreateDate := FormatDateTime( 'yyyymmdd', Date );
  Result := ( aTransNum + aCmdType + aSourceId + aDestId + aMoppId + aCreateDate );
end;

{ ---------------------------------------------------------------------------- }

function TNagraCommandBuilder.GetCommandSection(aObj: PSendNagra): String;
var
  aCmdType, aSection: String;
  aBMode, aBStartDate, aBEndDate, aAdrType: String;
begin
  { 請參考 EMC SasGwyStdSpe020705.pdf,
    第 16 頁 --> 4.5.2 ~ 4.5.4 Address Header = Command Section }
  { BoardCaseMode --> N=Normal, B=Batch
    BoardCast Start Date ~ End Date , 預設為昨天至明天
    Address Type --> U=Unique(for IccNo), G=Global }
  aCmdType := TranslateCommandType( StrToInt( aObj.LowCmdId ) );
  aBMode := 'N';
  aBStartDate := FormatDateTime( 'yyyymmdd', Date - 1 );
  aBEndDate := FormatDateTime( 'yyyymmdd', Date + 1 );
  { EMM }
  if ( aCmdType = '01' ) then
  begin
    { 假如是全域廣播指令 }
    if ( aObj.IsGlobalCommand ) then
    begin
      { ICC 號碼不用填, 補 10 碼 0 }
      aAdrType := 'G';
      aSection := aBMode + aBStartDate + aBEndDate + aAdrType + Rpad( EmptyStr, 10, '0' );
    end else
    begin
      { ICC 號碼必填, 取前10碼, 後2碼為 check sum }
      aAdrType := 'U';
      aSection := aBMode + aBStartDate + aBEndDate + aAdrType + Rpad( Copy( aObj.IccNo, 1, 10 ), 10, '0' );
    end;
  end else
  { Control }
  if ( aCmdType = '02' ) then
  begin
    aAdrType := 'U';
    aSection := aBMode + aBStartDate + aBEndDate + aAdrType +
      Rpad( Copy( aObj.IccNo, 1, 10 ), 10, '0' );
  end else
  { Feedback --> 傳指令不會用到 Feedback 的 Command Section, 回傳才有 }
  if ( aCmdType = '04' ) then
  begin
    aSection := Rpad( Copy( aObj.IccNo, 1, 10 ), 10, '0' );
  end else
  { Operation }
  if ( aCmdType = '05' ) then
  begin
    aSection := EmptyStr;
  end;
  Result := aSection;
end;

{ ---------------------------------------------------------------------------- }

function TNagraCommandBuilder.GetCommandBody(aObj: PSendNagra): String;
var
  aIrdObj: PIrdInfo;
  aFullBody: String;
  aRealCmdId, aProductId, aSubBeginDate, aSubEndDate, aZipCode: String;
  aCreditMode, aCredit, aCreditThreshold: String;
  aEventName, aEventNameLen, aEventPrice: String;
  aCCPhoneNum, aPhoneNum1, aPhoneNum2, aPhoneNum3: String;
  aSTBNo: String;
  aIpAddr, aIpPort: String;
  aCallbackDate, aCallbackTime, aCallFrequence, aDateFirstCall: String;
  aIrdCmdId, aIrdOperation, aIrdDataLength, aIrdData: String;
  aCleanDate, aConditionDate, aCollectDate: String;
  aCreditLimit: String;
begin
  SpecialLowCommandRule( aObj );
  { 是否是 IRD Command ID }
  if aObj.LowCmdId = '69' then
  begin
    New( aIrdObj );
    try
      if aObj.STBFlag = '0' then
        GetIrdStructure( aObj, aIrdObj )   { KBRO 舊的盒子, ADB }
      else
        GetIrdStructure2( aObj, aIrdObj ); { 新盒子, HDT }
      aIrdCmdId := aIrdObj.IrdCmdId;
      aIrdOperation := aIrdObj.IrdOperation;
      aIrdDataLength := aIrdObj.IrdDataLength;
      { IrdData 必須補足 96 個 Byte, 不足補 0 }
      aIrdData := Rpad( aIrdObj.IrdData, 96, '0' );
    finally
      Dispose( aIrdObj );
    end;
  end;
  aRealCmdId := Lpad( aObj.LowCmdId, 4, '0' );
  {}
  aProductId := Lpad( aObj.ImdProductId, 12, '0' );
  aSubBeginDate := TranslateSubBeginDate( aObj.SubBeginDate );
  aSubEndDate := TranslateSubEndDate( aObj.ImdProductId, aObj.SubEndDate );
  {}
  aCreditMode := Lpad( aObj.CreditMode, 2, '0' );
  aCredit := TranslateCredit( aObj.Credit );
  aCreditThreshold := TranslateCredit( aObj.ThreholdCredit );
  { Zip Code format: 09 + 000, 09 --> CompCode, 000 ---> default }
  aZipCode := aObj.ZipCode;
  { For OPPV Command, 轉 UTF-8, 因為名稱是由 CAS 控制, 現在就算把名稱轉成 UTF-8,
    指令送下去也沒用, BOX 會自動抓 SI Table 名稱來秀 }
  aEventName := TranslateEventName( aObj.EventName );
  { EventName 最長 30, 後面 2 碼固定補 0 }
  if Length( aEventName ) > 30 then
    aEventName := Copy( aEventName, 1, 30 );
  aEventName := Rpad( aEventName, 32, '0' );
  aEventNameLen := Lpad( IntToStr( Length( aObj.EventName ) ), 2, '0' );
  { EventPrice 0.00 ~ 999.99 }
  { 等 EMC 確認超過 999 的金額, 如何顯示 }
  aEventPrice := Lpad( Copy( Nvl( aObj.Price, '0' ), 1, 3 ), 3, '0' ) + '00';
  { Callback 認證, 授權電話 }
  aCCPhoneNum := Rpad( aObj.CCNumber1, 16, #32 );
  aPhoneNum1 := Rpad( aObj.PhoneNum1, 16, #32 );
  aPhoneNum2 := Rpad( aObj.PhoneNum2, 16, #32 );
  aPhoneNum3 := Rpad( aObj.PhoneNum3, 16, #32 );
  { BOX }
  aSTBNo := TranslateSTBNoBySpecialRule( aObj );
  { 使用模擬器 }
  if FCmdTransEnv.UseSimulator then
    aSTBNo := Rpad( aSTBNo, 14, '0' )
  else
    aSTBNo := Rpad( aSTBNo, 14, #32 );
  { Callback IP Address }
  aIpAddr := TranslateIpAddr( aObj.IPAddr );
  aIpPort := Lpad( aObj.CCPort, 5, '0' );
  { Callback Date and Time }
  aCallbackDate := aObj.CallbackDate;
  aCallbackTime := Nvl( aObj.CallbackTime, FormatDateTime( 'hhnnss', Now ) );
  { Enabled Automatic Callback }
  aCallFrequence := Lpad( aObj.CallbackFrequency, 2, '0' );
  aDateFirstCall := Nvl( aObj.FirstCallbackDate,
    FormatDateTime( 'yyyymmdd', Date + 1 ) );
  { Purge }
  aCleanDate := Nvl( aObj.CleanupDate, FormatDateTime( 'yyyymmdd', Date ) );
  aConditionDate := TranslateConditionDateBySpecialRule( aObj );
  { Set IPPV Record AS Reported }
  aCollectDate := Nvl( aObj.CollectDate, FormatDateTime( 'yyyymmdd', Date ) );
  { Redefine Credit Limit }
  aCreditLimit := Lpad( Nvl( aObj.CreditLimit, '0' ), 7, '0' );
  case StrToInt( aObj.LowCmdId ) of
    2 :
      aFullBody := aRealCmdId + aProductId + aSubBeginDate + aSubEndDate;
    3 :
      aFullBody := aRealCmdId + aProductId + aSubEndDate;
    4, 5, 6 :
      aFullBody := aRealCmdId + aProductId;
    7, 14, 15, 20, 21, 50, 51, 53, 62, 71, 105, 110, 111:
      aFullBody := aRealCmdId;
    8 :
      aFullBody := aRealCmdId + aCreditMode + aCredit;
    9 :
      aFullBody := aRealCmdId + aCreditThreshold;
    10 :
      aFullBody :=
        aRealCmdId + aProductId + aEventNameLen + aEventName + aEventPrice ;
    13 :
      aFullBody := aRealCmdId + aCredit + aCreditThreshold;
    48 :
      aFullBody := aRealCmdId + aZipCode;
    49 :
      aFullBody := aRealCmdId + aCCPhoneNum;
    52, 104 :
      aFullBody := aRealCmdId + aSTBNo;
    54 :
      aFullBody := aRealCmdId + aIpAddr + aIpPort;
    60 :
      aFullBody := aRealCmdId + aCallbackDate + aCallbackTime;
    61 :
      aFullBody := aRealCmdId + aCallFrequence + aDateFirstCall + aCallbackTime;
    69 :
      aFullBody := aRealCmdId + aIrdCmdId + aIrdOperation + aIrdDataLength + aIrdData;
    96:
      aFullBody := aRealCmdId + aCleanDate + aConditionDate;
    97:
      aFullBody := aRealCmdId + aCollectDate;
    100 :
      aFullBody := aRealCmdId + aCreditLimit;
    101 :
      aFullBody := aRealCmdId + aPhoneNum1 + aPhoneNum2 + aPhoneNum3;
    { Ack, 請參考 EMC SasGwyStdSpe020705.pdf 第 66 頁 4.9.1 }  
    1000 :
      aFullBody := aRealCmdId + aObj.CaAckTransactionNumber +
        Lpad( EmptyStr, 12, '0' )  + Lpad( EmptyStr, 12, '0' );
    { Nack, 請參考 EMC SasGwyStdSpe020705.pdf 第 66 頁 4.9.2 }
    1001 :
      { 如果執行 NACK 回去, 則 nack_status = 1, 代表 reject }
      { ErrorCode, ErrorMessage 代碼, 請參考 69~73頁 }
      { 這邊一律給 0x0000 }
      aFullBody := aRealCmdId + aObj.CaAckTransactionNumber + '2' +
        '0000' + '0000' + '004' + aObj.LowCmdId;
    { Communication Check }
    1002 :
      aFullBody := aRealCmdId;
  end;
  Result := aFullBody;
end;

{ ---------------------------------------------------------------------------- }

procedure TNagraCommandBuilder.Clear;
begin
  { 不釋放 CommandObject }
  FCommandList.Clear;
end;

{ ---------------------------------------------------------------------------- }

procedure TNagraCommandBuilder.ClearCommand;
begin
  { 須要釋放 CommandObject }
  ClearList( FCommandList );
end;

{ ---------------------------------------------------------------------------- }

function TNagraCommandBuilder.GetCommandCount: Integer;
begin
  Result := FCommandList.Count;
end;

{ ---------------------------------------------------------------------------- }

function TNagraCommandBuilder.GetCommandText(Index: Integer): String;
begin
  Result := PSendNagra( FCommandList.Objects[Index] ).SmsFullCommandText;
end;

{ ---------------------------------------------------------------------------- }

function TNagraCommandBuilder.GetCommandObject(Index: Integer): PSendNagra;
begin
  Result := PSendNagra( FCommandList.Objects[Index] );
end;

{ ---------------------------------------------------------------------------- }

function TNagraCommandBuilder.GetTransactionNumberGenerator(
  const aIdentify: String): TTransactionNumberGenerator;
var
  aFileName: String;
  aIndex: Integer;
begin
  aIndex := FTransactionNumberList.IndexOf( Format( 'Trans%s', [aIdentify] ) );
  if aIndex < 0 then
  begin
    Result := TTransactionNumberGenerator.Create( Format( 'Trans%s', [aIdentify] ),
      False, 1, 999999 );
    aFileName := IncludeTrailingPathDelimiter( ExtractFilePath(
      ParamStr( 0 ) ) ) + DEFAULT_TRANSACTIONFILE;
    if FileExists( aFileName ) then Result.Restore( aFileName );
    FTransactionNumberList.AddObject( Format( 'Trans%s', [aIdentify] ), Result );
  end else
    Result := TTransactionNumberGenerator( FTransactionNumberList.Objects[aIndex] );
end;

{ ---------------------------------------------------------------------------- }

function TNagraCommandBuilder.GetMailIdGenerator(
  const aIdentify: String): TTransactionNumberGenerator;
var
  aFileName: String;
  aIndex: Integer;
begin
  aIndex := FTransactionNumberList.IndexOf( Format( 'Mail%s', [aIdentify] ) );
  if aIndex < 0 then
  begin
    Result := TTransactionNumberGenerator.Create( Format( 'Mail%s', [aIdentify] ),
      False, 1, 99 );
    aFileName := IncludeTrailingPathDelimiter( ExtractFilePath(
      ParamStr( 0 ) ) ) + DEFAULT_TRANSACTIONFILE;
    if FileExists( aFileName ) then Result.Restore( aFileName );
    FTransactionNumberList.AddObject( Format( 'Mail%s', [aIdentify] ), Result );
  end else
    Result := TTransactionNumberGenerator( FTransactionNumberList.Objects[aIndex] );
end;

{ ---------------------------------------------------------------------------- }

procedure TNagraCommandBuilder.BuildCommandSetpByStep;
var
  aObj: PSendNagra;
  aIndex: Integer;
  aHeader, aSection, aBody, aFullText, aHexLen: String;
begin
  for aIndex := 0 to FCommandList.Count - 1 do
  begin
    aObj := PSendNagra( FCommandList.Objects[aIndex] );
    try
      aBody := GetCommandBody( aObj );
      aSection := GetCommandSection( aObj );
      aHeader := GetCommandHeader( aObj );
      aFullText := aHeader + aSection + aBody;
      aHexLen := LengthToHex( Length( aFullText ) );
    except
      { ... }
    end;
    { 請參考 EMC SasGwyStdSpe020705.pdf,
      第 8 頁 --> 2.3.10 Message_5 , len + data }
    aObj.SmsCommandText := aFullText;
    aObj.SmsFullCommandText := aHexLen + aFullText;
  end;
end;

{ ---------------------------------------------------------------------------- }

{ 舊的盒子用到的 IRD 指令 }

function TNagraCommandBuilder.GetIrdStructure(aObj: PSendNagra;
  aIrdObj: PIrdInfo): Boolean;
var
  aIndex, aIrdCmdId: Integer;
  aIrdData, aRealIrdCmdId, aRealIrdOperation, aRealIrdData, aRealIrdLength: String;

  { --------------------------------------------------- }

  function ConvertToRealIrdCmdId(const aMisId: Integer): String;
  begin
    Result := EmptyStr;
    case aMisId of
      1, 8: Result := '018';      { 設定親子密碼, 設定 IPPV 訂購密碼, 0x12 }
      2: Result := EmptyStr;      { 傳訊息, 不支援 }
      3: Result := EmptyStr;      { 強制換台, 不支援 }
      4: Result := '194';         { SetNetwork Id, 舊盒子用的是 0xC2 }
      5: Result := EmptyStr;      { 母子機配對 Master/Slave Continuous Mode Initialisation, 不支援 }
      6: Result := EmptyStr;      { 母子機解配對 Master/Slave Cancellation, 不支援 }
      7: Result := EmptyStr;      { 母子機臨時驗證子機 Master/Slave Single Shot, 不支援 }
    end;
  end;

  { --------------------------------------------------- }

begin
  aRealIrdCmdId := EmptyStr;
  aRealIrdOperation := EmptyStr;
  aRealIrdData := EmptyStr;
  aRealIrdLength := '00';
  aIrdCmdId := StrToInt( aObj.MisIrdCmd );
  aIrdData := aObj.MisIrdData;
  aRealIrdCmdId := ConvertToRealIrdCmdId( aIrdCmdId );
  case aIrdCmdId of
     1,8: { 親子密碼, IPPV 訂購密碼, 在舊盒子上, 2 組指的是同一個密碼 }
       begin
         aRealIrdOperation := '001';
         aRealIrdLength := '04';
         aIrdData := Nvl( Copy( aIrdData, 1, 4 ), '1234' );
         { 取 ASCII 值後, 轉 16 進制 }
         for aIndex := 1 to Length( aIrdData ) do
         begin
           aRealIrdData := ( aRealIrdData +
             IntToHex( Ord( aIrdData[aIndex] ), 0 ) );
         end;
       end;
     4: { Set Network Id }
       begin
         aRealIrdOperation := '001';
         aIrdData := IntToHex( StrToInt( aObj.NetWorkId ), 4 );
         aRealIrdData := aIrdData;
         aRealIrdLength := '04';
       end;
  end;
  aIrdObj.IrdCmdId := aRealIrdCmdId;
  aIrdObj.IrdOperation := aRealIrdOperation;
  aIrdObj.IrdDataLength := aRealIrdLength;
  aIrdObj.IrdData := aRealIrdData;
  Result := ( aRealIrdCmdId <> EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

{ 新的盒子用到的 IRD 指令 }

function TNagraCommandBuilder.GetIrdStructure2(aObj: PSendNagra;
  aIrdObj: PIrdInfo): Boolean;
var
  aIndex, aIrdCmdId: Integer;
  aIrdData, aRealIrdCmdId, aRealIrdOperation, aRealIrdData: String;
  aValidationPeriod, aRandomPeriod, aTimeOut: String;
  aMailId, aSegmentId, aTotalSegnment, aPriority, aMessage: String;
  aOrignalNetworkId, aTransportId, aServiceId: String;
  aTemp, aMasterIcc, aConfig: String; // , aPattern

  { --------------------------------------------------- }

  function ConvertToRealIrdCmdId(const aMisId: Integer): String;
  begin
    Result := EmptyStr;
    case aMisId of
      1: Result := '200';  { 0xC8, 設定親子密碼 }
      2: Result := '192';  { 0xC0, 傳訊息 }
      3: Result := '193';  { 0xC1, 強制換台 }
      4: Result := '198';  { 0xC6, SetNetwork Id --> CNS 用來設 BAT ID }
      5: Result := '199';  { 0xC7, 母子機配對 Master/Slave Continuous Mode Initialisation }
      6: Result := '199';  { 0xC7, 母子機解配對 Master/Slave Cancellation }
      7: Result := '199';  { 0xC7, 母子機臨時驗證子機 Master/Slave Single Shot }
      8: Result := '200';  { 0xC8, 設定 IPPV 訂購密碼 }
      9: Result := '197';  { 0xC5, Set BAT Id, Configure STB  --> TBC 用來設 BAT ID }
      10: Result := '018'; { 0x12, Reset Pin Code }
      11: Result := '207'; { 0xCF, Pop up }
      15: Result := '210'; { 0xD2, Force software download }
      16: Result := '210'; { 0xD2, Force software download interactive }
      17: Result := '197'; { 0xC5, Set DVR quota value, Configure STB  --> TBC 用來設定DVR使用容量 }
      18: Result := '197'; { 0xC5, Pair STB with HDD S/N no., Configure STB  --> TBC 用來DVR配對 }
      20: Result := '018'; { 0x12, Reset VOD Pin Code --> CNS VOD; 19:已被前端制定命令使用 }
    end;
  end;

  { --------------------------------------------------- }

begin
  {}
  aIrdCmdId := StrToInt( aObj.MisIrdCmd );
  aRealIrdCmdId := ConvertToRealIrdCmdId( aIrdCmdId );
  {}
  case aIrdCmdId of
     1: { 親子密碼 }
       begin
         aRealIrdOperation := '001';
         aIrdData := Nvl( Copy( aObj.PinCode, 1, 4 ), '0000' );
         { 將密碼轉 ASCII 再轉 16 進制 }
         for aIndex := 1 to Length( aIrdData ) do
           aRealIrdData := aRealIrdData + IntToHex( Ord( aIrdData[aIndex] ), 2 );
         { 密碼長度 + 密碼: Ex: 04 + 轉換過的值 }
         aRealIrdData := ( Lpad( IntToStr( Length( aIrdData ) ), 2, '0' ) +
           aRealIrdData );
       end;
     2, 11:{ Mail 傳送訊息, Pop up 傳送訊息 }
       begin
         aRealIrdOperation := '001';
         if ( aIrdCmdId = 11 ) then aRealIrdOperation := '000';
         aIrdData := aObj.MisIrdData;
         { 取出 Mail_id, 轉成 2 進制, 固定取 10 碼 }
         aMailId := ExtractValue( aIrdData );
         aMailId := IntToBin( StrToInt( aMailId ) );
         aMailId := Copy( aMailId, Length( aMailId ) - 9, 10 );
         { 截出 Total_Segnment, 轉成 2 進制, 取 6 碼 }
         aTotalSegnment := ExtractValue( aIrdData );
         aTotalSegnment := IntToBin( StrToInt( aTotalSegnment ) );
         aTotalSegnment := Copy( aTotalSegnment, Length( aTotalSegnment ) - 5, 6  );
         { 截出 Segnment_Id, 轉 2 進制, 取 6 碼 }
         aSegmentId := ExtractValue( aIrdData );
         aSegmentId := IntToBin( StrToInt( aSegmentId ) );
         aSegmentId := Copy( aSegmentId, Length( aSegmentId ) - 5, 6 );
         { 取出緊急度, 轉 2 進制, 取 2 碼 }
         aIrdData := Nvl( aIrdData, '0' );
         aPriority := IntToBin( StrToInt( aIrdData ) );
         aPriority := Copy( aPriority, Length( aPriority ) - 1, 2 );
         aObj.MisIrdData := aIrdData;
         { 將所有值相加之後, 轉 16 進制,
           Ex: 00000000 10000001 00000000 = 0x008100  }
         aIrdData := IntToHex( BinToInt(
           aMailId + aTotalSegnment + aPriority + aSegmentId ), 6 );
         { 取出訊息, 將每一個字元取 ASCII 後, 轉 16 進制 }
         aMessage := EmptyStr;
         { 先將中文轉UTF-8, 英數字及符號不用轉 }
         aTemp := MailTextToUtf8( aObj.Notes );
         for aIndex := 1 to Length( aTemp ) do
           aMessage := ( aMessage + IntToHex( Ord( aTemp[aIndex] ), 2 ) );
         aRealIrdData := ( aIrdData + aMessage );
       end;
     3: { Force Tune 強制換台 }
       begin
         aRealIrdOperation := '001';
         aIrdData := aObj.MisIrdData;
         if ( aIrdData <> EmptyStr ) then
         begin
           aOrignalNetworkId := Copy( aIrdData, 1, Pos( '~', aIrdData ) - 1 );
           Delete( aIrdData, 1, Pos( '~', aIrdData ) );
           aTransportId := Copy( aIrdData, 1, Pos( '~', aIrdData ) - 1 );
           Delete( aIrdData, 1, Pos( '~', aIrdData ) );
           aServiceId := aIrdData;
           { network_id + transport_id + service_id }
           aRealIrdData := (
             IntToHex( StrToInt( aOrignalNetworkId ), 4 ) +
             IntToHex( StrToInt( aTransportId ), 4 ) +
             IntToHex( StrToInt( aServiceId ), 4 ) );
         end;
       end;
     4: { Set Network Id --> CNS 用來設定 BAT ID }
       begin
         aRealIrdOperation := '001';
         { 10 進制數字 --> 轉 16 進制
           Ex: NetworkId = 10, 轉 16 進制 = 000A }
         aIrdData := IntToHex( StrToInt( aObj.NetWorkId ), 4 );
         aOrignalNetworkId := IntToHex( StrToInt( Nvl( aObj.MisIrdData, '0' ) ), 4 );
         aRealIrdData := ( aIrdData + aOrignalNetworkId );
       end;
     5: { 母子機配對 }
       begin
         aRealIrdOperation := '001';
         { 先算出 Master Icc 卡的 16 進制 }
         aMasterIcc := Copy( ExtractValue( aObj.MisIrdData ), 1, 10 );
         aIrdData := IntToHex( StrToInt( aMasterIcc ), 0 );
         { 取出設定值 }
         aTemp := aObj.MisIrdData;
         { 取出須要傳入的參數, validationPeriod, randomPeriod, timeout }
         aValidationPeriod := ExtractValue( aTemp, '~' );
         aRandomPeriod := ExtractValue( aTemp, '~' );
         aTimeOut := ExtractValue( aTemp, '~' );
         if ( aValidationPeriod <> EmptyStr ) then
           aValidationPeriod := IntToHex( StrToInt( aValidationPeriod ), 2 );
         if ( aRandomPeriod <> EmptyStr ) then
           aRandomPeriod := IntToHex( StrToInt( aRandomPeriod ), 2 );
         if ( aTimeOut <> EmptyStr ) then
           aTimeOut := IntToHex( StrToInt( aTimeOut ), 2 );
         aRealIrdData := ( aIrdData + aValidationPeriod + aRandomPeriod + aTimeOut );
       end;
     6: { 母子機解配對 }
       begin
         aRealIrdOperation := '002';
         aIrdData := EmptyStr;
         aRealIrdData := EmptyStr;
         { No Data Value Need }
       end;
     7: { 臨時驗證子機 }
       begin
         aRealIrdOperation := '003';
         aMasterIcc := Copy( ExtractValue( aObj.MisIrdData ), 1, 10 ); 
         { 先算出 Master Icc 卡的 16 進制 }
         aIrdData := IntToHex( StrToInt( aMasterIcc ), 0 );
         { Timeout 參數的 16 進制 }
         aTimeOut := IntToHex( StrToInt( Nvl( aObj.MisIrdData, '1' ) ), 2 );
         aRealIrdData := ( aIrdData + aTimeOut );
         { 轉 ASCII }
       end;
     8: { IPPV 訂購密碼 }
       begin
         aRealIrdOperation := '002';
         aIrdData := Nvl( Copy( aObj.PinCode, 1, 4 ), '1234' );
         { 將密碼轉 ASCII 再轉 16 進制 }
         for aIndex := 1 to Length( aIrdData ) do
           aRealIrdData := aRealIrdData + IntToHex( Ord( aIrdData[aIndex] ), 2 );
         { 密碼長度 + 密碼: Ex: 04 + 轉換過的值 }
         aRealIrdData := ( Lpad( IntToStr( Length( aIrdData ) ), 2, '0' ) +
           aRealIrdData );
       end;
     9:  { Set BAT ID }
       begin
         aRealIrdOperation := '001';
         aIrdData := Lpad( Copy( aObj.MisIrdData, 1, 5 ), 5, '0' );
         { BAT ID 直接轉 16 進制 }
         aRealIrdData := IntToHex( StrToInt( aIrdData ), 4 );
       end;
     10, 20: { Reset Pin Code, Reset VOD Pin Code }
       begin
          aRealIrdOperation := '001';
          if ( aIrdCmdId = 20 ) then aRealIrdOperation := '002';
          aIrdData := EmptyStr;
          aRealIrdData := EmptyStr; 
       end;
     15: { Force software download }
       begin
          aRealIrdOperation := '000';
          aIrdData := EmptyStr;
          aRealIrdData := EmptyStr;
       end;
     16: { Force software download interactive }
       begin
          aRealIrdOperation := '001';
          aIrdData := EmptyStr;
          aRealIrdData := EmptyStr;
       end;
     17: { Set DVR quota value }
       begin
          aRealIrdOperation := '011';
          aIrdData := aObj.DVRQuota;
          { DVR quota value 直接轉 16 進制 }
          // aRealIrdData := '01' + IntToHex( StrToInt( aIrdData ), 2 );
          { 2009.06.17 by Lawrence Change the parameters of the Nagra specifications}
          aRealIrdData := IntToHex( StrToInt( aIrdData ), 4 );
       end;
     18: { Pair STB with HDD S/N no. }
       begin
          aRealIrdOperation := '010';
          aIrdData := aObj.MisIrdData;
         if ( aIrdData <> EmptyStr ) then
         begin
           { 2009.06.17 by Lawrence Change the parameters of the Nagra specifications
             aConfig := Lpad(Copy( aIrdData, 1, Pos( '~', aIrdData ) - 1 ), 4, '0');
             Delete( aIrdData, 1, Pos( '~', aIrdData ) );
             aPattern := aIrdData;
             // para config
             aIrdData := aConfig;
             for aIndex := 1 to Length( aIrdData ) do
               aRealIrdData := aRealIrdData + IntToHex( Ord( aIrdData[aIndex] ), 2 );
             // para pattern
             aIrdData := aPattern;
           }
           if aIrdData='00' then // for CNS unpair
              begin
                aConfig :='00';  // 16 bit
                aIrdData :='';
              end
           else  // for TBC reserved
              aConfig :='0000';  // 16 bit for config reserved
           aRealIrdData :=aRealIrdData + aConfig;
           for aIndex := 1 to Length( aIrdData ) do
             aRealIrdData := aRealIrdData + IntToHex( Ord( aIrdData[aIndex] ), 2 );
           { 2009.06.17 by Lawrence Change the parameters of the Nagra specifications
             // Length + aConfig + aPattern
             aRealIrdData := ( Lpad( IntToStr( Length( aConfig + aPattern ) ), 2, '0' ) +
                aRealIrdData );
           }
         end;
       end;
  end;
  aIrdObj.IrdCmdId := aRealIrdCmdId;
  aIrdObj.IrdOperation := aRealIrdOperation;
  aIrdObj.IrdDataLength := '00';
  { KBRO, TBC , IrdData 取16進制後, 全部總長度除以 2 }
  if ( aRealIrdData <> EmptyStr ) then
    aIrdObj.IrdDataLength := Lpad( IntToStr( Length( aRealIrdData ) div 2 ), 2, '0' );
  { CNS Spec for Unpair STB with HDD S/N no. }
  if (aIrdCmdId=18) and (aConfig='00') then
    begin
       aRealIrdData := aRealIrdData + '00';    
       aIrdObj.IrdDataLength := '01';
    end;
  { TBC Spec for Pair STB with HDD S/N no. }
  if (aIrdCmdId=18) and (aRealIrdData=EmptyStr) then
    aIrdObj.IrdDataLength := '30';
  aIrdObj.IrdData := aRealIrdData;
  Result := ( aRealIrdCmdId <> EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

procedure TNagraCommandBuilder.SpecialLowCommandRule(aObj: PSendNagra);
const
  aCommon = '~';
var
  aPrefix: String;
  aIndex: Integer;
begin
  if not Assigned( aObj ) then Exit;
  { 開頻道, 開機+頻道, 恢復頻道, 收視權複製, 延展更改期限
    維修換拆中的開頻道 }
  if ( aObj.HighCmdId = 'B1' ) or
     ( aObj.HighCmdId = 'B5' ) or
     ( aObj.HighCmdId = 'B6' ) or
     ( aObj.HighCmdId = 'B7' ) or
     ( ( aObj.HighCmdId = 'A1' ) and ( aObj.LowCmdId = '2' ) ) or
     ( ( aObj.HighCmdId = 'A6' ) and ( aObj.LowCmdId = '2' ) ) or
     ( ( aObj.HighCmdId = 'A9' ) and ( aObj.LowCmdId = '2' ) ) or
     ( ( aObj.HighCmdId = 'B9' ) and ( aObj.LowCmdId = '2' ) ) or
     ( ( aObj.HighCmdId = 'B10' ) and ( aObj.LowCmdId = '2' ) ) then
  begin
    { A~12345 或是 B~12345~20050101~20051231 }
    aPrefix := Copy( aObj.Notes, 1, 1 );
    Delete( aObj.Notes, 1, 2 );
    if ( aPrefix = 'A' ) then
      aObj.ImdProductId := aObj.Notes
    else begin
      aIndex := Pos( aCommon, aObj.Notes );
      aObj.ImdProductId := Copy( aObj.Notes, 1, aIndex - 1 );
      Delete( aObj.Notes, 1, aIndex );
      aIndex := Pos( aCommon, aObj.Notes );
      aObj.SubBeginDate := Copy( aObj.Notes, 1, aIndex - 1 );
      Delete( aObj.Notes, 1, aIndex );
      aObj.SubEndDate := aObj.Notes;
    end;
    aObj.Notes := EmptyStr;
  end else
  { 關頻道, 暫停瀕道, 取消 PPV 訂購節目 }
  if ( aObj.HighCmdId = 'B2' ) or
     ( aObj.HighCmdId = 'B4' ) or
     ( aObj.HighCmdId = 'B5' ) or
     ( aObj.HighCmdId = 'P31' ) then
  begin
    aObj.ImdProductId := aObj.Notes;
    aObj.Notes := EmptyStr;
  end else
  { 解配對 }
  if ( aObj.HighCmdId = 'C1' ) then
  begin
    aObj.IccNo := aObj.Notes;
    aObj.Notes := EmptyStr;
  end else
  { 母子機配對, 臨時驗證子機 }
  if ( aObj.HighCmdId = 'E5' ) or
     ( aObj.HighCmdId = 'E7' ) then
  begin
    { 母機卡號放在 IccNo 裏, 把它移到 MisIrdData 裏 }
    aObj.MisIrdData := Format( '%s,%s', [aObj.IccNo,aObj.MisIrdData] );
    { 把真正的子機卡移到 IccNo ( Command Header ) 裏  }
    aObj.IccNo := aObj.Notes;
  end else
  { 停用 STB 回報異常 }
  if ( aObj.HighCmdId = 'P15' ) then
  begin
    aObj.IccNo := aObj.Notes;
    aObj.Notes := EmptyStr;
  end else
  { 訂購 OPPV 節目 }
  if ( aObj.HighCmdId = 'P30' ) then
  begin
    aIndex := AnsiPos( aCommon, aObj.Notes );
    aObj.ImdProductId := Copy( aObj.Notes, 1, aIndex - 1 );
    Delete( aObj.Notes, 1, aIndex );
    aIndex := AnsiPos( aCommon, aObj.Notes );
    aObj.EventName := Copy( aObj.Notes, 1, aIndex - 1 );
    Delete( aObj.Notes, 1, aIndex );
    aObj.Price := aObj.Notes;
  end else
  if ( aObj.LowCmdId = '48' ) then
  begin
    aObj.ZipCode := Lpad( aObj.ZipCode, 5, '0' );
  end else
  { ICC 回收清除作業的 Set ZipCode, 給 0 }
  if ( aObj.HighCmdId = 'X1' ) or
     ( aObj.HighCmdId = 'X2' ) then
  begin
    aObj.ZipCode := Rpad( EmptyStr, 5, '0' );
  end;
end;


{ ---------------------------------------------------------------------------- }

function TNagraCommandBuilder.BuildCommunicationCheck(aObj: PSendNagra): String;
var
  aHeader, aSection, aBody, aText, aHexLen: String;
begin
  aHeader := GetCommandHeader( aObj );
  aSection := GetCommandSection( aObj );
  aBody := GetCommandBody( aObj );
  aText := aHeader + aSection + aBody;
  aHexLen := LengthToHex( Length( aText ) );
  { 請參考 EMC SasGwyStdSpe020705.pdf,
    第 8 頁 --> 2.3.10 Message_5 , len + data }
  aObj.SmsCommandText := aText;
  aObj.SmsFullCommandText := aHexLen + aText;
  Result := aObj.SmsFullCommandText;
end;

{ ---------------------------------------------------------------------------- }

function TNagraCommandBuilder.BuildCommunicationCheck(aObj: PRecvNagra): String;
var
  aObj2: PSendNagra;
begin
  New( aObj2 );
  try
    aObj2.CompCode :=  aObj.CompCode;
    aObj2.HighCmdId := aObj.HighLevelCmd;
    aObj2.LowCmdId :=  aObj.LowLevelCmd;
    Result := BuildCommunicationCheck( aObj2 );
    aObj.TransactionNumber := aObj2.SmsTransactionNumber;
    aObj.CommandText := aObj2.SmsCommandText;
    aObj.FullCommandText := aObj2.SmsFullCommandText;
  finally
    Dispose( aObj2 );
  end;
end;

{ ---------------------------------------------------------------------------- }

{ TNagraAckParser }

class function TNagraAckParser.ParserAcknowledge(aAckData: String): PSendNagra;
var
  aCmdId: String;
  aObj: PSendNagra;
begin
  Result := nil;
  if ( aAckData = EmptyStr ) then Exit;
  { 先判斷是不是 Communication Check 的 Command }
  aCmdId := Copy( aAckData, 33, 4 );
  if aCmdId = '1002' then Exit;
  New( aObj );
  { Ack 回來指令別 Ack --> 1000 or NAck --> 1001 }
  aObj.CaAckCommandId := aCmdId;
  aObj.LowCmdId := aCmdId;
  { Ack 回來的時間 }
  aObj.CaAckTime := Now;
  { Ack 回來指令的 Transaction Number }
  aObj.CaAckTransactionNumber := Copy( aAckData, 1, 9 );
  { Ack 回來的是那一筆指令? 原指令的 Transaction Number }
  aObj.SmsTransactionNumber := Copy( aAckData, 37, 9 );
  { Ack 回來的 Full Command Text }
  aObj.CaAckCommandText := aAckData;
  { 只要不是 1000 一律當成 Nack, 取 ErrorCode, ErrorMsg  }
  if ( aObj.CaAckCommandId <> '1000' ) then
  begin
    aObj.CaAckStatus := Copy( aAckData, 46, 1 );
    aObj.ErrCode :=  Copy( aAckData, 47, 4 );
    aObj.ErrMsg := Copy( aAckData, 51, 4 );
    aObj.SmsSendStatus := 'E';
  end else
  begin
    aObj.CaAckStatus := EmptyStr;
    aObj.ErrCode :=  EmptyStr;
    aObj.ErrMsg := EmptyStr;
    aObj.SmsSendStatus := 'C';
  end;
  Result := aObj;
end;

{ ---------------------------------------------------------------------------- }

class function TNagraAckParser.InternalParseFeedbackData(aObj: PRecvNagra;
   aCmdBody: String): Boolean;
var
  aCmdId: Integer;
begin
  Result := True;
  aCmdId := StrToIntDef( aObj.LowLevelCmd, 0 );
  case aCmdId of
    200, 201:
      begin
        if aCmdId = 200 then
          aObj.HighLevelCmd := 'R2'
        else
          aObj.HighLevelCmd := 'R3';
        aObj.StbNo := Copy( aCmdBody, 1, 14 );
        aObj.Credit := Format( '%s.%s', [Copy( aCmdBody, 15, 5 ),
          Copy( aCmdBody, 20, 2 )] );
        aObj.Debit := Format( '%s.%s', [Copy( aCmdBody, 22, 5 ),
          Copy( aCmdBody, 27, 2 )] );
      end;
    202:
      begin
        aObj.HighLevelCmd := 'R1';
        aObj.StbNo := Copy( aCmdBody, 1, 14 );
        aObj.ImdProductId := IntToStr( StrToIntDef( Copy( aCmdBody, 15, 12 ), 0 ) );
        aObj.PurchaseDate := Copy( aCmdBody, 27, 8 );
        aObj.WatchStatus := Copy( aCmdBody, 35, 1 );
      end;
    205:
      begin
        aObj.HighLevelCmd := 'R4';
        aObj.StbNo := Copy( aCmdBody, 1, 14 );
        aObj.PhoneNum1 := Trim( Copy( aCmdBody, 15, 16 ) );
        aObj.PhoneNum2 := Trim( Copy( aCmdBody, 31, 16 ) );
        aObj.PhoneNum3 := Trim( Copy( aCmdBody, 47, 16 ) );
        aObj.AbnormalPhone := Trim( Copy( aCmdBody, 63, 16 ) );
      end;
    206:
      begin
        aObj.HighLevelCmd := 'R5';
        aObj.StbNo := Copy( aCmdBody, 1, 14 );
        aObj.StbResponding := Copy( aCmdBody, 15, 1 );
      end;
    207:
      begin
        aObj.HighLevelCmd := 'R6';
        aObj.StbNo := Copy( aCmdBody, 1, 14 );
      end;
    211:
      begin
        aObj.HighLevelCmd := 'R10';
        aObj.CallbackDate := Copy( aCmdBody, 1, 8 );
        aObj.CallbackTime := Copy( aCmdBody, 9, 6 );
      end;
    212:
      begin
        aObj.HighLevelCmd := 'R11';
        aObj.NumberOfIppv := Copy( aCmdBody, 1, 2 );
      end;
    215:
      begin
        Result := False;
      end;  
  end;
end;

{ ---------------------------------------------------------------------------- }

class function TNagraAckParser.ParseFeedbackData(aAckData: String): PRecvNagra;
var
  aCmdId: String;
  aObj: PRecvNagra;
  aIsAcknowledge: Boolean;
  aCmdHeader, aCmdBody, aIccNo: String;
begin
  Result := nil;
  if ( aAckData = EmptyStr ) then Exit;
  New( aObj );
  aObj.CompCode := EmptyStr;
  aObj.TransactionNumber := EmptyStr;
  aObj.ResponseTransactionNumber := EmptyStr;
  aObj.LowLevelCmd := EmptyStr;
  aObj.IccNo := EmptyStr;
  aObj.CommandText := aAckData;
  aObj.CallbackDate := EmptyStr;
  aObj.CallbackTime := EmptyStr;
  aObj.NumberOfIppv := EmptyStr;
  aObj.Credit := EmptyStr;
  aObj.Debit := EmptyStr;
  aObj.ImdProductId := EmptyStr;
  aObj.PurchaseDate := EmptyStr;
  aObj.WatchStatus := EmptyStr;
  aObj.ProductSuspended := EmptyStr;
  aObj.IccSuspended := EmptyStr;
  aObj.PhoneNum1 := EmptyStr;
  aObj.PhoneNum2 := EmptyStr;
  aObj.PhoneNum3 := EmptyStr;
  aObj.AbnormalPhone := EmptyStr;
  aObj.StbResponding := EmptyStr;
  aObj.CmdStatus := 'W';
  aObj.SendStatus := 'W';
  aObj.CommandText := EmptyStr;
  aObj.NackStatus := EmptyStr;
  aObj.ErrCode := EmptyStr;
  aObj.ErrMsg := EmptyStr;
  aObj.UpdTime := Now;
  aObj.Data := nil;
  aObj.ParentData := nil;
  aObj.Index := -1;
  aObj.ParentIndex := -1;
  { 先判斷是 Communication Check 的 Command 或是真正的 Feedback Data }
  aIsAcknowledge := ( Copy( aAckData, 10, 2 ) = '05' );
  if ( aIsAcknowledge ) then
  begin
    aCmdId := Copy( aAckData, 33, 4 );
    aObj.TransactionNumber := Copy( aAckData, 37, 9 );
    { Ack 回來指令的 Transaction Number }
    aObj.ResponseTransactionNumber := Copy( aAckData, 1, 9 );
    aObj.HighLevelCmd := aCmdId;
    aObj.LowLevelCmd := aCmdId;
    { 只要不是 1000 一律當成 Nack, 取 ErrorCode, ErrorMsg  }
    if ( aCmdId <> '1000' ) then
    begin
      aObj.NackStatus := Copy( aAckData, 46, 1 );
      aObj.ErrCode :=  Copy( aAckData, 47, 4 );
      aObj.ErrMsg := Copy( aAckData, 51, 4 );
      aObj.SendStatus := 'E';
    end else
    begin
      aObj.NackStatus := EmptyStr;
      aObj.ErrCode :=  EmptyStr;
      aObj.ErrMsg := EmptyStr;
      aObj.SendStatus := 'C';
    end;
  end
  else begin
    aCmdHeader := Copy( aAckData, 1, 32 );
    aIccNo := Copy( aAckData, 33, 10 );
    aCmdId := Copy( aAckData, 43, 4 );
    aCmdBody := Copy( aAckData, 47, Length( aAckData ) - 46 );
    aObj.CompCode := FEEDBACK_COMPCODE;
    aObj.LowLevelCmd := aCmdId;
    aObj.ResponseTransactionNumber := Copy( aCmdHeader, 1, 9 );
    aObj.ResponseCommandText := aAckData;
    aObj.IccNo := aIccNo;
    if not InternalParseFeedbackData( aObj, aCmdBody ) then
     Dispose( aObj );
  end;
  Result := aObj;
end;

{ ---------------------------------------------------------------------------- }

procedure TNagraCommandBuilder.CloneObj(aSource, aTarget: PSendNagra);
begin
  aTarget.HighCmdId := aSource.HighCmdId;
  aTarget.IccNo := aSource.IccNo;
  aTarget.StbNo := aSource.StbNo;
  aTarget.SubBeginDate := aSource.SubBeginDate;
  aTarget.SubEndDate := aSource.SubEndDate;
  aTarget.Notes := aSource.Notes;
  aTarget.CmdStatus := aSource.CmdStatus;
  aTarget.Operator := aSource.Operator;
  { 低階指令的處理時間或日期, 在這裏先指定跟高階日期一樣 }
  { 後續傳送出去後, 再指定為傳送時的日期時間 }
  aTarget.UpdTime := aSource.UpdTime;
  aTarget.ErrCode := EmptyStr;
  aTarget.ErrMsg := EmptyStr;
  aTarget.ImdProductId := aSource.ImdProductId;
  aTarget.ZipCode := aSource.ZipCode;
  aTarget.CreditMode := aSource.CreditMode;
  aTarget.Credit := aSource.Credit;
  aTarget.CreditLimit := aSource.CreditLimit;
  aTarget.ThreholdCredit := aSource.ThreholdCredit;
  aTarget.EventName := aSource.ThreholdCredit;
  aTarget.Price := aSource.Price;
  aTarget.CCNumber1 := aSource.CCNumber1;
  aTarget.IPAddr := aSource.IPAddr;
  aTarget.CCPort := aSource.CCPort;
  aTarget.CallbackDate := aSource.CallbackDate;
  aTarget.CallbackTime := aSource.CallbackTime;
  aTarget.CallbackFrequency := aSource.CallbackFrequency;
  aTarget.FirstCallbackDate := aSource.FirstCallbackDate;
  aTarget.PhoneNum1 := aSource.PhoneNum1;
  aTarget.PhoneNum2 := aSource.PhoneNum2;
  aTarget.PhoneNum3 := aSource.PhoneNum3;
  aTarget.MisIrdCmd := aSource.MisIrdCmd;
  aTarget.MisIrdData := aSource.MisIrdData;
  aTarget.SeqNo := aSource.SeqNo;
  aTarget.CompCode := aSource.CompCode;
  aTarget.ResentTimes := aSource.ResentTimes;
  aTarget.CleanupDate := aSource.CleanupDate;
  aTarget.ConditionDate := aSource.ConditionDate;
  aTarget.CollectDate := aSource.CollectDate;
  aTarget.STBFlag := aSource.STBFlag;
  aTarget.NetWorkId := aSource.NetWorkId;
  aTarget.PinCode := aSource.PinCode;
  aTarget.DVRQuota := aSource.DVRQuota;
  aTarget.VodMoppId := aSource.VodMoppId;
  { 判斷是否全域廣播指令 }
  aTarget.IsGlobalCommand := aSource.IsGlobalCommand;
  { 不用跟原始指令一樣的欄位 }
  aTarget.Data := nil;
  aTarget.ParentData := nil;
  aTarget.Index := -1;
  aTarget.ParentIndex := -1;
  { 此低階指令尚未傳送 }
  aTarget.SmsSendStatus := 'W';
  aTarget.SmsSendTime := -1;
  { 給高階指令用的, 每一個低階指令用不到此值, 所以填 0 }
  aTarget.ReferenceLowCmdCount := 0;
  { 此低階指令對應的各系統台 List Index 值, 方便後續 Ack 回來後快速找到原高階指令 }
  aTarget.SoCompListIndex := aSource.SoCompListIndex;
  { 此低階指令尚未寫入 Log }
  aTarget.HasbeenLog := False;
  {}
  aTarget.SmsTransactionNumber := EmptyStr;
  aTarget.SmsCommandText := EmptyStr;
  aTarget.SmsFullCommandText := EmptyStr;
  aTarget.LowCmdId := aSource.LowCmdId;
  aTarget.CaAckStatus := EmptyStr;
  aTarget.CaAckTime := -1;
  aTarget.CaAckTransactionNumber := EmptyStr;
  aTarget.CaAckCommandId := EmptyStr;
  aTarget.CaAckCommandText := EmptyStr;
  {}
end;

{ ---------------------------------------------------------------------------- }

procedure TNagraCommandBuilder.HighObjExtractToLowObj(aCompositText: String);
var
  aObj: PSendNagra;
  aSequence: Integer;
  aLowCmdId, aIrdMappingCmd: String;
begin
  { 將所有的高階指令, 轉成實際對應的低階指令
    Ex: A1 --> 51,52,21,48,69,69 }
  aSequence := 0;
  aIrdMappingCmd := EmptyStr;
  repeat
    aLowCmdId := ExtractValue( aCompositText );
    New( aObj );
    { 將所有值對抄一份 }
    CloneObj( FPrimalObj, aObj );
    { 指定該實際的 Command Id }
    aObj.LowCmdId := aLowCmdId;
    { 如果是 IRD Command, 找出實際的 IRD Command Id }
    if ( SameText( aLowCmdId, '69' ) )  then
    begin
      if ( aIrdMappingCmd = EmptyStr ) then
        aIrdMappingCmd := GetValue( 'HIGH_LEVEL_CMD', 'IRD_CMD', [aObj.HighCmdId] );
      aObj.MisIrdCmd := ExtractValue( aIrdMappingCmd );  
    end;
    { 先加到 Temp List, 稍後再處理, Temp List 先將每個低階指令以 1000 做間隔,
      先編號碼, Ex : 1000, 2000, 3000 .....  }
    Inc( aSequence, 1000 );
    FTemporaryList.AddObject( Lpad( IntToStr( aSequence ), 6, '0' ),
      TObject( aObj ) );
  until ( aCompositText = EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

procedure TNagraCommandBuilder.NormalHighCommandRule;
var
  aIndex, aShift, aBigNum: Integer;
  aObj, aObj2: PSendNagra;
  aFirstEnter: Boolean;
  aNotes, aExtract: String;
begin
  { 將每一個低階指令, 檢查 Notes 或其它規則是否再須要拆成更多指令,
       Ex: B1 ( 頻道 -- 多組 ), 每一個頻道寫在 Notes 裏, 用逗號區隔, 或
           E2 ( 傳送訊息 -- Mail ), 訊息在 Notes 欄裏, 可是超過一定數量的
           的訊息, 一樣要拆成第2個指令 }
  { 不用 for 迴圈, 因為會用 insert 方式, 在應該再拆解新的指令地方,
    寫進新的低階指令, Ex: B1 ( 頻道 -- 多組 ) --> 再拆成新的 }
  aIndex := 0;
  while ( aIndex < FTemporaryList.Count ) do
  begin
    aShift := 0;
    aBigNum := StrToInt( FTemporaryList.Strings[aIndex] );
    aObj := PSendNagra( FTemporaryList.Objects[aIndex] );
    { Notes 欄位有填值, 就是要再多拆低階指令 }
    if ( aObj.Notes <> EmptyStr ) then
    begin
      { Notes 欄位是否是這個低階指令會去用到 }
      if GetValue( 'HIGH_LEVEL_CMD;NOTES_USE', 'NOTES_USE',
        [aObj.HighCmdId, aObj.LowCmdId] ) <> EmptyStr then
      begin
        { 取出 Notes 值 }
        aNotes := aObj.Notes;
        aFirstEnter := True;
        repeat
          { 解出每一個 Notes 值 }
          aExtract := ExtractValue( aNotes );
          { 本身就是第一個使用 Notes 的低階指令 }
          if aFirstEnter then
          begin
            { 重新給一予第一個拆解出來的 Notes 值 }
            aFirstEnter := False;
            aObj.Notes := aExtract;
          end else
          begin
            { 第 2 個以後, 就必須要產生新的低階指令 }
            New( aObj2 );
            CloneObj( aObj, aObj2 );
            aObj2.LowCmdId := aObj.LowCmdId;
            aObj2.Notes := aExtract;
            { 編號 }
            Inc( aBigNum );
            FTemporaryList.InsertObject( aIndex + 1,
              Lpad( IntToStr( aBigNum ), 6, '0' ), TObject( aObj2 ) );
            { 位移 --> 在原本指令後插入了幾個新的低階指令 }
            Inc( aShift );
          end;
        until ( aNotes = EmptyStr );
      end else
      begin
        aObj.Notes := EmptyStr;
      end;
    end;
    { 換下一個指令, 必須跳過從中間插入的 }
    Inc( aIndex, aShift + 1 );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TNagraCommandBuilder.SpecialHighCommandRule;
var
  aBigNum, aUpperBounds, aBaseIndex, aIndex, aAddProductIndex: Integer;
  aCardFirstEnter, aProductFirstEnter, aMailFirstEnter: Boolean;
  aObj, aObj2, aObj3, aTmpObj: PSendNagra;
  aNotes, aProduct, aCard, aIccNo, aMisIrdData: String;
  aTemp, aSaveProduct, aOldIccAndStb, aNewIccAndStb: String;
  aMailGenerator: TTransactionNumberGenerator;
  aMailId, aTotalSegment, aSegmentNumber: Integer;
  aMailMsg, aMailPriotity: String;
begin
  aBaseIndex := 0;
  aObj := PSendNagra( FTemporaryList.Objects[aBaseIndex] );
  { A6:維修換拆-整組
    A9:維修換拆-只換 ICC }
  { 10001/00001@10002/00002;A~1,A~2,A~3,B~5~20020201~20020205 }
  { 舊機卡為 10001及00001 會下關機指令
    新機卡為 10002及00002 會下開機指令
    後面新機卡會跟著下開頻道指令 }
  if ( aObj.HighCmdId = 'A6' ) or
     ( aObj.HighCmdId = 'A9' ) then
  begin
    aNotes := aObj.Notes;
    { 先將整組機卡解出 }
    aTemp := ExtractValue( aNotes, ';' );
    { 舊機卡 }
    aOldIccAndStb := ExtractValue( aTemp, '@' );
    { 新機卡 }
    aNewIccAndStb := aTemp;
    { 取出舊機卡指令數 }
    aUpperBounds := StrToIntDef( GetValue(
      'HIGH_LEVEL_CMD', 'OLD_CMD_COUNT', [aObj.HighCmdId] ), 0 );
    for aIndex := 1 to aUpperBounds do
    begin
      aObj := PSendNagra( FTemporaryList.Objects[aBaseIndex] );
      { 這裏要將 STB NO 清掉, 因為這個高階指令裏面有2個 cmd 52
        後面程式不知道那一個才是 pair 跟 unpair  }
      if ( aObj.LowCmdId <> '52' ) then
        aObj.StbNo := Copy( aOldIccAndStb, 1, Pos( '/', aOldIccAndStb ) - 1 );
      aObj.IccNo := Copy( aOldIccAndStb, Pos( '/', aOldIccAndStb ) + 1, Length( aOldIccAndStb ) );
      Inc( aBaseIndex );
    end;
    { 取新機卡的指令數 }
    aAddProductIndex := -1;
    aUpperBounds := StrToIntDef( GetValue(
      'HIGH_LEVEL_CMD', 'NEW_CMD_COUNT', [aObj.HighCmdId] ), 0 );
    for aIndex := 1 to aUpperBounds do
    begin
      aObj := PSendNagra( FTemporaryList.Objects[aBaseIndex] );
      aObj.StbNo := Copy( aNewIccAndStb, 1, Pos( '/', aNewIccAndStb ) - 1 );
      aObj.IccNo := Copy( aNewIccAndStb, Pos( '/', aNewIccAndStb ) + 1, Length( aNewIccAndStb ) );
      if ( aObj.LowCmdId = '2' ) then aAddProductIndex := aBaseIndex;
      if ( aIndex < aUpperBounds ) then Inc( aBaseIndex );
    end;
    if ( aAddProductIndex >= 0 ) then
    begin
      { 新機卡的 2: Add Product }
      aObj := PSendNagra( FTemporaryList.Objects[aAddProductIndex] );
      aBigNum := StrToInt( FTemporaryList.Strings[aAddProductIndex] );
      aProductFirstEnter := True;
      repeat
        if ( aProductFirstEnter ) then
        begin
          aObj.Notes := ExtractValue( aNotes );
          aProductFirstEnter := False;
        end else
        begin
          New( aObj2 );
          CloneObj( aObj, aObj2 );
          aObj2.Notes := ExtractValue( aNotes );
          Inc( aBigNum );
          FTemporaryList.InsertObject( aAddProductIndex,
            Lpad( IntToStr( aBigNum ), 6, '0' ), TObject( aObj2 ) );
        end;
      until ( aNotes = EmptyStr );
    end;  
  end else
  { A8: 維修換拆-只換 STB }
  if ( aObj.HighCmdId = 'A8' ) then
  begin
    aNotes := aObj.Notes;
    { 舊機卡 --> 不用下指令 }
    aOldIccAndStb := ExtractValue( aNotes, '@' );
    { 新機卡 }
    aNewIccAndStb := aNotes;
    aUpperBounds := 3;
    for aIndex := 1 to aUpperBounds do
    begin
      aObj := PSendNagra( FTemporaryList.Objects[aBaseIndex] );
      aObj.StbNo := Copy( aNewIccAndStb, 1, Pos( '/', aNewIccAndStb ) - 1 );
      aObj.IccNo := Copy( aNewIccAndStb, Pos( '/', aNewIccAndStb ) + 1, Length( aNewIccAndStb ) );
      if ( aIndex < aUpperBounds )  then
        Inc( aBaseIndex );
    end;
  end else
  { B7 收視權限複製, 只會有一個低階指令, Command 2 --> Add Product  }
  if ( aObj.HighCmdId = 'B7' ) then
  begin
    aBigNum := StrToInt( FTemporaryList.Strings[aBaseIndex] );
    { 傳入多卡號,多頻道(欄位Notes)=>將所有頻道 grant 給所有卡號.
      各 product 間以','區隔. 各卡號間以';'區隔.
         product data與卡號 data以'~'區隔.
         (先product data,後卡號 data)
      各 product 的info 包含三個部分=>product id, 收視起日,收視迄日.而這三個部分以'^'作為區隔
      Ex: 111111111111^20020801^20020901,222222222222^20020901^20021231~3333333333;4444444444 }
    aNotes := aObj.Notes;
    { 先將 Product Id 解出 }
    aSaveProduct := ExtractValue( aNotes, '~' );
    { 再將卡號解出 }
    aCard := aNotes;
    { 然後以卡號, 做出每一個低階 Obj, 就像是頻道一樣,
      後續低階指令在 Build Command Body 的時候, 再去拆解 }
    aCardFirstEnter := True;
    repeat
      aIccNo := ExtractValue( aCard, ';' );
      if ( aCardFirstEnter ) then
      begin
        aTmpObj := aObj;
        aCardFirstEnter := False;
      end
      else begin
        Inc( aBaseIndex );
        New( aObj2 );
        CloneObj( aObj, aObj2 );
        aTmpObj := aObj2;
      end;
      aTmpObj.IccNo := aIccNo;
      { Product 加進該 ICC 卡 }
      aProductFirstEnter := True;
      aProduct := aSaveProduct;
      repeat
         aTemp := ExtractValue( aProduct );
         if ( aProductFirstEnter ) then
         begin
           aTmpObj.Notes := 'B~' + StringReplace( aTemp, '^', '~', [rfReplaceAll] );
           aProductFirstEnter := False;
         end else
         begin
           Inc( aBaseIndex );
           New( aObj3 );
           CloneObj( aTmpObj, aObj3 );
           aObj3.Notes := 'B~' + StringReplace( aTemp, '^', '~', [rfReplaceAll] );
           aTmpObj := aObj3;
         end;
         Inc( aBigNum );
         if ( aBaseIndex > 0 ) then
         begin
           FTemporaryList.AddObject( Lpad( IntToStr( aBigNum ), 6, '0' ),
             TObject( aTmpObj ) );
         end;
      until ( aProduct = EmptyStr );
    until ( aCard = EmptyStr );
  end;
  { E2=傳送訊息, E11= Pop up 訊息, E21=傳送訊息(全域廣播) }
  if ( ( aObj.HighCmdId = 'E2' ) or ( aObj.HighCmdId = 'E11' ) or ( aObj.HighCmdId = 'E21' ) ) then
  begin
    aBigNum := StrToInt( FTemporaryList.Strings[aBaseIndex] );
    { 先算出傳送的訊息有多少個 Byte, 看是否須要拆成多個指令,
      每個訊息最多 45 個 Byte, 超過要拆成多筆指令 }
    aMailMsg := aObj.Notes;
    aMailPriotity := aObj.MisIrdData;
    aMailGenerator := GetMailIdGenerator( FCASEnv.MailId );
    aMailId := StrToInt( Format( '%s%.2d',  [FCASEnv.MailId, aMailGenerator.NextValue] ) );
    aTotalSegment := CaclMailSegment( aMailMsg );
    aSegmentNumber := 0;
    aMailFirstEnter := True;
    while ( aMailMsg <> EmptyStr ) do
    begin
      if aMailFirstEnter then
      begin
        aMailFirstEnter := False;
        aObj.Notes := CaclMailText( aMailMsg );
        aObj.MisIrdData := Format( '%d,%d,%d,%s', [aMailId, aTotalSegment,
          aSegmentNumber, aMailPriotity] );
      end else
      begin
        { 解出來訊息放在第1個以後每產生出來的指令 }
        New( aObj3 );
        CloneObj( aObj, aObj3 );
        Inc( aSegmentNumber );
        aObj3.Notes := CaclMailText( aMailMsg );
        aObj3.MisIrdData := Format( '%d,%d,%d,%s', [aMailId, aTotalSegment,
          aSegmentNumber, aMailPriotity] );
        Inc( aBigNum );
        FTemporaryList.AddObject( Lpad( IntToStr( aBigNum ), 6, '0' ),
          TObject( aObj3 ) );
      end;
    end;
  end;
  { 母子機 Master/Slave --> Single Shot or Continuous Mode }
  if ( aObj.HighCmdId = 'E5' ) or ( aObj.HighCmdId = 'E7' ) then
  begin
    aBigNum := StrToInt( FTemporaryList.Strings[aBaseIndex] );
    aCardFirstEnter := True;
    aNotes := aObj.Notes;
    aMisIrdData := aObj.MisIrdData;
    repeat
       { 先將所有子機的 ICC 解出來成多筆資料 }
       aIccNo := ExtractValue( aNotes );
       if aCardFirstEnter then
       begin
         aObj.Notes := aIccNo;
         aCardFirstEnter := False;
       end else
       begin
         New( aObj3 );
         CloneObj( aObj, aObj3 );
         { 將每一筆解出來的子機卡號, 放回 Notes }
         aObj3.Notes := aIccNo;
         Inc( aBigNum );
         FTemporaryList.AddObject( Lpad( IntToStr( aBigNum ), 6, '0' ),
           TObject( aObj3 ) );
       end;
    until ( aNotes = EmptyStr );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TNagraCommandBuilder.GetValue(const aFindField, aResultField: String;
  const Args: array of Variant): String;
begin
  Result := EmptyStr;
  FHighCmdEnv.First;
  if FHighCmdEnv.Locate( aFindField, VarArrayOf( Args ), [] ) then
    Result := FHighCmdEnv.FieldByName( aResultField ).AsString;
end;

{ ---------------------------------------------------------------------------- }

function TNagraCommandBuilder.CaclMailSegment(aMsg: String): Integer;
begin
  Result := 0;
  while ( aMsg <> EmptyStr ) do
  begin
    CaclMailText( aMsg );
    Inc( Result );
  end;
end;


{ ---------------------------------------------------------------------------- }

function TNagraCommandBuilder.CaclMailText(var aMsg: String): String;
var
  aIndex, aPos: Integer;
  aTemp: String;
begin
  Result := EmptyStr;
  aPos := 0;
  for aIndex := 1 to Length( aMsg ) do
  begin
    case ByteType( aMsg, aIndex ) of
      mbSingleByte:
        begin
          aTemp := ( aTemp + aMsg[aIndex] );
        end;
      mbTrailByte:
        begin
         aTemp := ( aTemp + AnsiToUtf8( Copy( aMsg, aIndex - 1, 2 ) ) );
        end;
      mbLeadByte:
        begin
          Continue;
        end;
    end;
    if ( Length( aTemp ) <= 45 ) then
    begin
      Result := Copy( aMsg, 1, aIndex );
      aPos := aIndex;
    end else
      Break;  
  end;
  Delete( aMsg, 1, aPos );
end;

{ ---------------------------------------------------------------------------- }

function TNagraCommandBuilder.MailTextToUtf8(aMsg: String): String;
var
  aIndex: Integer;
begin
  Result := EmptyStr;
  for aIndex := 1 to Length( aMsg ) do
  begin
    case ByteType( aMsg, aIndex ) of
      mbSingleByte:
        Result := ( Result + aMsg[aIndex] );
      mbTrailByte:
        Result := ( Result + AnsiToUtf8( Copy( aMsg, aIndex - 1, 2 ) ) );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

end.

