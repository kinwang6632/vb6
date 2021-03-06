unit cbDBController;

interface

uses
  Windows, SysUtils, Classes, Math, DB, ADODB, DBClient;

type

  PLoginInfo = ^TLoginInfo;
  TLoginInfo = record
    IsSuperVisor: Boolean;
    UserId: String;
    UserName: String;
    CompCode: String;
    CompName: String;
    DbAccount: String;
    DbPassword: String;
    DbAliase: String;
    ServiceType: String;
    Com_DbAccount: String;
    Com_DbPassword: String;
    Com_DbAliase: String;
    StarPVOD: Boolean;
  end;

  TCD078BRecord = record
    BPCode: String;
    InstCode: String;
    ServiceType: String;
    FaciItem: String;
  end;

  TDBController = class(TDataModule)
    DataConnection: TADOConnection;
    DataReader: TADOQuery;
    DataWriter: TADOQuery;
    CodeReader: TADOQuery;
    ComConnection: TADOConnection;
    ComReader: TADOQuery;
    procedure DataModuleCreate(Sender: TObject);
    procedure DataModuleDestroy(Sender: TObject);
  private
    { Private declarations }
    FLoginInfo: PLoginInfo;
    FGroupId : String;
    FCanUseEPG : Boolean;
    FCanModify: Boolean;
    function GetConnected: Boolean;
    function GetCanUseEPG: Boolean;
    procedure SetConnected(const Value: Boolean);
    procedure GetGropuId;
    {}
    procedure ResetDataSetFilter(const ADataSet: TClientDataSet);
    procedure SetCD078DDuplicateFilter(ADataSet: TClientDataSet; ABpCode, AFaciCode, AService: String);
    procedure SetCD078ADuplicateFilter(ADataSet: TClientDataSet; ABpCode, ACItemCode, AService: String);
    procedure SetCD078A1DuplicateFilter(ADataSet: TClientDataSet; ABpCode, ACItemCode, AService: String);
    procedure SetCD078CDuplicateFilter(ADataSet: TClientDataSet; ABpCode, ACItemCode, AService: String);
    procedure SetCD078JCDuplicateFilter(ADataSet: TClientDataSet; ABpCode, AProductCode, AService: String);
  public
    { Public declarations }
    property LoginInfo: PLoginInfo read FLoginInfo;
    property Connected: Boolean read GetConnected write SetConnected;
    property CanUseEPG: Boolean read FCanUseEPG;
    property CanModify: Boolean read FCanModify write FCanModify;
    function GetItemByHouse(const ACodeNo: String): String;
    function GetFaciItemSeqNo(const ADataSet: TClientDataSet): String;
    function AutoAppendCD078D(var ARecord: TCD078BRecord; const ADataSet: TClientDataSet; var AErrMsg: String): Boolean;
    function AutoAppendCD078A(var ARecord: TCD078BRecord; var AErrMsg: String): Boolean;
    function AutoAppendCD078C(var ARecord: TCD078BRecord; const ADataSet: TClientDataSet; var AErrMsg: String): Boolean;
    function AutoAppendCD078J(var ARecord: TCD078BRecord; const ADataSet: TClientDataSet; Var AerrMsg: String): Boolean;
    function VdInstCodeExitisInCD078B(AInstCodes: String; const ADataSet: TClientDataSet): Boolean;
    function GetInstName(AValues: String): String;
    function GetStepNo: String;
    function GetDiscountAmt(const ACodeNo: String): Double;
    function GetPeriod(const ACodeNo: String):Integer;
    function GetCitemName(const ACodeNo: String) : string;
    function GetUserPriv(const Mid : String ): Boolean;
  end;

var
  DBController: TDBController;

implementation

uses cbUtilis, frmCD078B1U;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TDBController.DataModuleCreate(Sender: TObject);
begin
  New( FLoginInfo );
end;

{ ---------------------------------------------------------------------------- }

procedure TDBController.DataModuleDestroy(Sender: TObject);
begin
  Dispose( FLoginInfo );
end;

{ ---------------------------------------------------------------------------- }

function TDBController.GetConnected: Boolean;
begin
  Result := DataConnection.Connected;
end;

{ ---------------------------------------------------------------------------- }

procedure TDBController.SetConnected(const Value: Boolean);
begin
  if ( Value ) then
  begin
    DataConnection.ConnectionString := Format(
     'Provider=MSDAORA.1;Password=%s;User ID=%s;Data Source=%s;Persist Security Info=True',
      [FLoginInfo.DbPassword, FLoginInfo.DbAccount, FLoginInfo.DbAliase] );
    try
      DataConnection.Connected := True;
    except
      on E: Exception do ErrorMsg( E.Message );
    end;
    {#6035 ??????COM???????? By Kin 2011/06/08}
    if FLoginInfo.StarPVOD  then
    begin
      ComConnection.ConnectionString := Format(
        'Provider=MSDAORA.1;Password=%s;User ID=%s;Data Source=%s;Persist Security Info=True',
          [FLoginInfo.Com_DbPassword, FLoginInfo.Com_DbAccount, FLoginInfo.Com_DbAliase] );
      try
         ComConnection.Connected := True;
      except
        on E: Exception do ErrorMsg( E.Message );
      end;
    end;
    {#6060 ???? GroupId By Kin 2013/08/15}
    GetGropuId;
    {?P?O?i?_??????EPG By Kin 2013/08/15}
    FCanUseEPG := GetCanUseEPG;
  end else
  begin
    DataConnection.Connected := False;
    ComConnection.Connected := False;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TDBController.GetItemByHouse(const ACodeNo: String): String;
begin
  DataReader.Close;
  DataReader.SQL.Text := Format(
    ' SELECT A.BYHOUSE FROM %s.CD019 A  ' +
    '  WHERE A.CODENO = ''%s''          ',
    [DBController.LoginInfo.DbAccount, ACodeNo] );
  DataReader.Open;
  Result := DataReader.FieldByName( 'BYHOUSE' ).AsString;
  DataReader.Close;
end;

{ ---------------------------------------------------------------------------- }

function TDBController.GetFaciItemSeqNo(const ADataSet: TClientDataSet): String;
var
  aOldFilters: String;

  { -------------------------------- }

  procedure SaveFilter;
  begin
    aOldFilters := EmptyStr;
    if ADataSet.Filtered then
    begin
      aOldFilters := ADataSet.Filter;
      ADataSet.Filtered := False;
    end;
  end;

  { -------------------------------- }

  procedure RestoreFilter;
  begin
    if ( aOldFilters <> EmptyStr ) then
    begin
      ADataSet.Filter := aOldFilters;
      ADataSet.Filtered := True;
    end;
  end;

  { -------------------------------- }

var
  aFaciItem: Integer;
begin
  Result := '1';
  SaveFilter;
  try
    if ADataSet.IsEmpty then Exit;
    aFaciItem := 0;
    ADataSet.First;
    while not ( ADataSet.Eof ) do
    begin
      aFaciItem := Max( aFaciItem, ADataSet.FieldByName( 'FACIITEM' ).AsInteger );
      ADataSet.Next;
    end;
    Result := IntToStr( aFaciItem + 1 );
  finally
    RestoreFilter;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TDBController.ResetDataSetFilter(const ADataSet: TClientDataSet);
begin
  ADataSet.Filtered := False;
  ADataSet.Filter := EmptyStr;
end;

{ ---------------------------------------------------------------------------- }

procedure TDBController.SetCD078DDuplicateFilter(ADataSet: TClientDataSet;
  ABpCode, AFaciCode, AService: String);
begin
  ADataSet.Filtered := False;
  ADataSet.Filter := Format( 'BPCODE=''%s'' AND FACICODE=''%s'' AND SERVICETYPE=''%s''',
    [ABpCode, AFaciCode, AService] );
  ADataSet.Filtered := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TDBController.SetCD078ADuplicateFilter(ADataSet: TClientDataSet;
  ABpCode, ACItemCode, AService: String);
begin
  ADataSet.Filtered := False;
  ADataSet.Filter := Format( 'BPCODE=''%s'' AND CItemCode=''%s'' AND SERVICETYPE=''%s''',
    [ABpCode, ACItemCode, AService] );
  ADataSet.Filtered := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TDBController.SetCD078A1DuplicateFilter(ADataSet: TClientDataSet;
  ABpCode, ACItemCode, AService: String);
begin
  ADataSet.Filtered := False;
  ADataSet.Filter := Format( 'BPCODE=''%s'' AND CItemCode=''%s'' AND SERVICETYPE=''%s''',
    [ABpCode, ACItemCode, AService] );
  ADataSet.Filtered := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TDBController.SetCD078CDuplicateFilter(ADataSet: TClientDataSet;
  ABpCode, ACItemCode, AService: String);
begin
  ADataSet.Filtered := False;
  ADataSet.Filter := Format( 'BPCODE=''%s'' AND CItemCode=''%s'' AND SERVICETYPE=''%s''',
    [ABpCode, ACItemCode, AService] );
  ADataSet.Filtered := True;
end;

{ ---------------------------------------------------------------------------- }

function TDBController.AutoAppendCD078D(var ARecord: TCD078BRecord;
  const ADataSet: TClientDataSet; var AErrMsg: String): Boolean;
var
  aIndex: Integer;
  aFacilCodes: array [1..3, 1..6] of String;
  aDefBuyCodes: array [1..3, 1..4] of String;
  aFaciItem, aInstCode: String;
begin
  AErrMsg := EmptyStr;
  ZeroMemory( @aFacilCodes, SizeOf( aFacilCodes ) );
  ZeroMemory( @aDefBuyCodes, SizeOf( aDefBuyCodes ) );
  { 1.???? CD005 ?????????}?C?? }
  DataReader.Close;
  DataReader.SQL.Text := Format(
    ' SELECT A.FACILCODE1, A.FACILCODE2, A.FACILCODE3,    ' +
    '        A.FACILNAME1, A.FACILNAME2, A.FACILNAME3,    ' +
    '        A.SERVICETYPE,                               ' +
    '        A.CMCODE, A.CMNAME, A.FIXIPCOUNT             ' +
    '   FROM %s.CD005 A                                   ' +
    '  WHERE CODENO = ''%s''                              ',
    [FLoginInfo.DbAccount, ARecord.InstCode] );
  DataReader.Open;
  for aIndex := Low( aFacilCodes ) to High( aFacilCodes ) do
  begin
    aFacilCodes[aIndex,1] := DataReader.FieldByName( Format( 'FACILCODE%d', [aIndex] ) ).AsString;
    aFacilCodes[aIndex,2] := DataReader.FieldByName( Format( 'FACILNAME%d', [aIndex] ) ).AsString;
    aFacilCodes[aIndex,3] := DataReader.FieldByName( 'SERVICETYPE' ).AsString;
    aFacilCodes[aIndex,4] := DataReader.FieldByName( 'CMCODE' ).AsString;
    aFacilCodes[aIndex,5] := DataReader.FieldByName( 'CMNAME' ).AsString;
    aFacilCodes[aIndex,6] := DataReader.FieldByName( 'FIXIPCOUNT' ).AsString;
  end;
  { 2. ???C?@???]???????R??????, CD022, CD022A--> ?]?????? }
  for aIndex := Low( aDefBuyCodes ) to High( aDefBuyCodes ) do
  begin
    DataReader.Close;
    DataReader.SQL.Text := Format(
      ' SELECT B.FACIMODELNO,             ' +
      '        B.FACIMODELNAME,           ' +
      '        A.DEFBUYCODE,              ' +
      '        A.DEFBUYNAME               ' +
      '   FROM %s.CD022 A,                ' +
      '       ( SELECT DISTINCT           ' +
      '                FACICODE,          ' +
      '                FACIMODELNO,       ' +
      '                FACIMODELNAME      ' +
      '           FROM %s.CD022A          ' +
      '          WHERE FACICODE = ''%s''  ' +
      '            AND ROWNUM <= 1 ) B    ' +
      '  WHERE A.CODENO = B.FACICODE(+)   ' +
      '    AND A.CODENO = ''%s''          ',
      [FLoginInfo.DbAccount, FLoginInfo.DbAccount,
       aFacilCodes[aIndex][1], aFacilCodes[aIndex][1]] );
    DataReader.Open;
    aDefBuyCodes[aIndex,1] := DataReader.FieldByName( 'FACIMODELNO' ).AsString;
    aDefBuyCodes[aIndex,2] := DataReader.FieldByName( 'FACIMODELNAME' ).AsString;
    aDefBuyCodes[aIndex,3] := DataReader.FieldByName( 'DEFBUYCODE' ).AsString;
    aDefBuyCodes[aIndex,4] := DataReader.FieldByName( 'DEFBUYNAME' ).AsString;
  end;
  DataReader.Close;
  { 3.???d?O?_????????????, ???????? Update InstCode ????, ?_?h?s?W
      ?@???? CD078D, ???K???????????????]???i?h }
  try
    for aIndex := Low( aFacilCodes ) to High( aFacilCodes ) do
    begin
      if ( aFacilCodes[aIndex,1] = EmptyStr ) then Continue;
      SetCD078DDuplicateFilter( ADataSet, ARecord.BPCode, aFacilCodes[aIndex,1],
        ARecord.ServiceType );
      if ( ADataSet.RecordCount <= 0 ) then
      begin
        { ???????? }
        aFaciItem := GetFaciItemSeqNo( ADataSet );
        ARecord.FaciItem := AddValue( aFaciItem, ARecord.FaciItem );
        ADataSet.Append;
        ADataSet.FieldByName( 'BpCode' ).AsString := ARecord.BPCode;
        ADataSet.FieldByName( 'FaciItem' ).AsString := aFaciItem;
        ADataSet.FieldByName( 'FaciType' ).AsString := '1';
        ADataSet.FieldByName( 'FaciCode' ).AsString := aFacilCodes[aIndex,1];
        ADataSet.FieldByName( 'FaciName' ).AsString := aFacilCodes[aIndex,2];
        ADataSet.FieldByName( 'ServiceType' ).AsString := aFacilCodes[aIndex,3];
        ADataSet.FieldByName( 'ModelCode' ).AsString := aDefBuyCodes[aIndex,1];
        ADataSet.FieldByName( 'ModelName' ).AsString := aDefBuyCodes[aIndex,2];
        ADataSet.FieldByName( 'BuyCode' ).AsString := aDefBuyCodes[aIndex,3];
        ADataSet.FieldByName( 'BuyName' ).AsString := aDefBuyCodes[aIndex,4];
        ADataSet.FieldByName( 'CMBaudNo' ).AsString := aFacilCodes[aIndex,4];
        ADataSet.FieldByName( 'CMBaudRate' ).AsString := aFacilCodes[aIndex,5];
        ADataSet.FieldByName( 'FixIpCount' ).AsString := aFacilCodes[aIndex,6];
        ADataSet.FieldByName( 'DynIpCount' ).AsString := EmptyStr;
        ADataSet.FieldByName( 'InstCodeStr' ).AsString := aRecord.InstCode;
        //#5859 ??????OK?A?n?????a1 By Kin 2011/05/30
        ADataSet.FieldByName( 'CBPChangeDynIP' ).AsString := '1';
        ADataSet.Post;
      end else
      begin
        aInstCode := ADataSet.FieldByName( 'InstCodeStr' ).AsString;
        if not ValueExists( ARecord.InstCode, aInstCode ) then
        begin
          ADataSet.Edit;
          ADataSet.FieldByName( 'InstCodeStr' ).AsString := AddValue( ARecord.InstCode, aInstCode );
          ADataSet.Post;
        end;
        aFaciItem := ADataSet.FieldByName( 'FaciItem' ).AsString;
        if not ValueExists( aFaciItem, ARecord.FaciItem ) then
          ARecord.FaciItem := AddValue( aFaciItem, ARecord.FaciItem );
      end;
    end;
  except
    on E: Exception do AErrMsg := E.Message;
  end;
  ResetDataSetFilter( ADataSet );
  Result := ( AErrMsg = EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

function TDBController.AutoAppendCD078A(var ARecord: TCD078BRecord;
  var AErrMsg: String): Boolean;
var
  aIndex: Integer;
  aCItemCodes: array [1..5, 1..11] of String;
  aInstCode, aFaciItem: String;
  aCD078A, aCD078A1: TClientDataSet;
  aMonthAmt, aDayAmt: Extended;
begin
  AErrMsg := EmptyStr;
  ZeroMemory( @aCItemCodes, SizeOf( aCItemCodes ) );
  {
   aCItemCodes[X,1] = CD005.CItemCode
   aCItemCodes[X,2] = CD005.CItemName
   aCItemCodes[X,3] = CD005.ServiceType
   aCItemCodes[X,4] = CD019.DeficiencyCode
   aCItemCodes[X,5] = CD019.DeficiencyName
   aCItemCodes[X,6] = CD019.Amount
   aCItemCodes[X,7] = CD019.Period
   aCItemCodes[X,8] = CD019.Sign
   aCItemCodes[X,9] = CD078A1.StepNo --> ???? CD078A1 ?t???s????????
   aCItemCodes[X,10] = CD019.CMCODE
   aCItemCodes[X,11] = CD019.CMNAME
  }
  { 1.???? CD005 ?????????}?C??, ???O???? }
  DataReader.Close;
  DataReader.SQL.Text := Format(
    ' SELECT A.CITEMCODE1, A.CITEMCODE2, A.CITEMCODE3,    ' +
    '        A.CITEMCODE4, A.CITEMCODE5,                  ' +
    '        A.CITEMNAME1, A.CITEMNAME2, A.CITEMNAME3,    ' +
    '        A.CITEMNAME4, A.CITEMNAME5,                  ' +
    '        A.SERVICETYPE                                ' +
    '   FROM %s.CD005 A                                   ' +
    '  WHERE A.CODENO = ''%s''                            ',
    [FLoginInfo.DbAccount, ARecord.InstCode] );
  DataReader.Open;
  while not DataReader.Eof do
  begin
    for aIndex := Low( aCItemCodes ) to High( aCItemCodes ) do
    begin
      aCItemCodes[aIndex,1] := DataReader.FieldByName( Format( 'CItemCode%d', [aIndex] ) ).AsString;
      aCItemCodes[aIndex,2] := DataReader.FieldByName( Format( 'CItemName%d', [aIndex] ) ).AsString;
      aCItemCodes[aIndex,3] := DataReader.FieldByName( 'ServiceType' ).AsString;
    end;
    DataReader.Next;
  end;
  { 2.?? CD019 ???????????????O }
  for aIndex := Low( aCItemCodes ) to High( aCItemCodes ) do
  begin
    DataReader.Close;
    DataReader.SQL.Text := Format(
      ' SELECT A.DEFICIENCYCODE, A.DEFICIENCYNAME,  ' +
      '        A.AMOUNT, A.PERIOD, A.SIGN,          ' +
      '        A.CMCODE, A.CMNAME                   ' +
      '   FROM %s.CD019 A                           ' +
      '  WHERE A.CODENO = ''%s''                    ' +
      '    AND A.PERIODFLAG = ''1''                 ' +
      '    AND A.STOPFLAG = ''0''                   ',
      [FLoginInfo.DbAccount, aCItemCodes[aIndex,1]] );
    DataReader.Open;
    if ( DataReader.IsEmpty ) then
    begin
      { ?S?????? CD019 ????, ?M?????????X?? CD005.CItemCode }
      aCItemCodes[aIndex,1] := EmptyStr;
      aCItemCodes[aIndex,2] := EmptyStr;
      aCItemCodes[aIndex,3] := EmptyStr;
    end else
    begin
      {--------------------------------------------------------------------------}
      { #4407 ?p?G?O???~?H???hCD019???????X REFNO=2 And MasterProduct <>1 ?~?i?x?s By Kin 2009/03/06 }
      { ???M???n?????o?????D By Kin 2009/03/06 }
      {
      if ( frmCD078B1.edtKindFunction.ItemIndex = 1 ) then
      begin
        CodeReader.Close;
        CodeReader.SQL.Text := Format(
          'SELECT COUNT(*) CNT FROM %s.CD019          ' +
          ' WHERE CODENO = %s                         ' +
          ' AND ( Nvl(REFNO,0) <> 2                   ' +
          ' OR Nvl(MasterProduct,0) = 1 )             ' ,
          [FLoginInfo.DbAccount, aCItemCodes[aIndex,1]] );
        CodeReader.Open;
        if ( CodeReader.Fields[0].AsInteger > 0 ) then
        begin
          CodeReader.Close;
          AErrMsg := '?????u???O?????O?????i???I?O?W?D?D???~,?????s?]?w?I';
          Result := False;
          Exit;
        end;
        CodeReader.Close;
      end;
      }
      {--------------------------------------------------------------------------}

      aCItemCodes[aIndex,4] := DataReader.FieldByName( 'DEFICIENCYCODE' ).AsString;
      aCItemCodes[aIndex,5] := DataReader.FieldByName( 'DEFICIENCYNAME' ).AsString;
      aCItemCodes[aIndex,6] := DataReader.FieldByName( 'AMOUNT' ).AsString;
      aCItemCodes[aIndex,7] := DataReader.FieldByName( 'PERIOD' ).AsString;
      aCItemCodes[aIndex,8] := DataReader.FieldByName( 'SIGN' ).AsString;
      aCItemCodes[aIndex,10] := DataReader.FieldByName( 'CMCODE' ).AsString;
      aCItemCodes[aIndex,11] := DataReader.FieldByName( 'CMNAME' ).AsString;
    end;
  end;
  DataReader.Close;
  {}
  aCD078A := frmCD078B1.CD078ADataSet;
  { ?o???????? CD078A1CloneDataSet }
  aCD078A1 := frmCD078B1.CD078A1CloneDataSet;
  { 3.???d?O?_????????????, ???????? Update InstCode ????, ?_?h?s?W
      ?@???? CD078A, ???K???????????????]???i?h }
  aFaciItem := ExtractValue( 1, ARecord.FaciItem );
  try
    for aIndex := Low( aCItemCodes ) to High( aCItemCodes ) do
    begin
      if ( aCItemCodes[aIndex,1] = EmptyStr ) then Continue;
      SetCD078ADuplicateFilter( aCD078A, ARecord.BPCode, aCItemCodes[aIndex,1], ARecord.ServiceType );
      aMonthAmt := 0;
      aDayAmt := 0;
      { ???p?? ???????B?? ???????B }
      if ( StrToFloatDef( aCItemCodes[aIndex, 7], 0 ) > 0 ) then
      begin
        aMonthAmt := CbRoundTo(
           ( StrToFloatDef( aCItemCodes[aIndex,6], 0 ) /
             StrToFloatDef( aCItemCodes[aIndex,7], 0 ) ), 3 );
        aDayAmt := CbRoundTo(
           ( ( StrToFloatDef( aCItemCodes[aIndex,6], 0 ) /
               StrToFloatDef( aCItemCodes[aIndex,7], 0 ) ) / 30 ), 3 );
        { ?t?? }
        if ( SameText( aCItemCodes[aIndex,8], '-' ) ) then
        begin
          aMonthAmt := ( 0 - aMonthAmt );
          aDayAmt := ( 0 - aDayAmt );
        end;
      end;
      if ( aCD078A.RecordCount <= 0 ) then
      begin
        aCD078A.Append;
        aCD078A.FieldByName( 'BpCode' ).AsString := ARecord.BPCode;
        aCD078A.FieldByName( 'CItemCode' ).AsString := aCItemCodes[aIndex,1];
        aCD078A.FieldByName( 'CItemName' ).AsString := aCItemCodes[aIndex,2];
        aCD078A.FieldByName( 'ServiceType' ).AsString := aCItemCodes[aIndex,3];
        aCD078A.FieldByName( 'FaciItem' ).AsString := aFaciItem;
        { ?w?]?? 2.?j?????? }
        aCD078A.FieldByName( 'PFLAG1' ).AsString := '2';
        aCD078A.FieldByName( 'PFlagCode' ).AsString := aCItemCodes[aIndex,4];
        aCD078A.FieldByName( 'PFlagName' ).AsString := aCItemCodes[aIndex,5];
        aCD078A.FieldByName( 'MONTHAMT' ).AsFloat := aMonthAmt;
        aCD078A.FieldByName( 'DAYAMT' ).AsFloat := aDayAmt;
        aCD078A.FieldByName( 'INSTCODESTR' ).AsString := ARecord.InstCode;
        aCD078A.FieldByName( 'SIGN' ).AsString := aCItemCodes[aIndex,8];
        aCD078A.FieldByName( 'CMBAUDNO' ).AsString := aCItemCodes[aIndex,10];
        aCD078A.FieldByName( 'CMBAUDRATE' ).AsString := aCItemCodes[aIndex,11];
        aCD078A.FieldByName( 'LONGPAYflag' ).AsString := '0';
        aCD078A.FieldByName( 'STOPFLAG' ).AsString := '0';
        { ?t????, ?Y?W?@????????, ???????????????O?????i MCItemCode }
        if ( SameText( aCItemCodes[aIndex,8], '-' ) ) then
        begin
          if ( aIndex - 1 in [Low( aCItemCodes )..High( aCItemCodes ) ] ) then
            if ( not SameText( aCItemCodes[aIndex-1, 8], '-' ) ) then
              aCD078A.FieldByName( 'MCItemCode' ).AsString := aCItemCodes[aIndex-1, 1];
        end;
        aCD078A.Post;
      end else
      begin
        aInstCode := aCD078A.FieldByName( 'INSTCODESTR' ).AsString;
        if not ValueExists( ARecord.InstCode, aInstCode ) then
        begin
          aCD078A.Edit;
          aCD078A.FieldByName( 'INSTCODESTR' ).AsString := AddValue( ARecord.InstCode, aInstCode );
          aCD078A.Post;
        end;
      end;
      { ?h???u?f }
      SetCD078A1DuplicateFilter( aCD078A1, ARecord.BPCode, aCItemCodes[aIndex,1], ARecord.ServiceType );
      if ( aCD078A1.RecordCount <= 0 ) then
      begin
        aCD078A1.Append;
        aCD078A1.FieldByName( 'BPCODE' ).AsString := ARecord.BPCode;
        aCD078A1.FieldByName( 'CITEMCODE' ).AsString := aCItemCodes[aIndex,1];
        aCD078A1.FieldByName( 'SERVICETYPE' ).AsString := aCItemCodes[aIndex,3];
        aCD078A1.FieldByName( 'FACIITEM' ).AsString := aFaciItem;
        aCD078A1.FieldByName( 'STEPNO' ).AsString := GetStepNo;
        aCD078A1.FieldByName( 'DESCRIPTION' ).AsString := '?w?]???O';
        aCD078A1.FieldByName( 'MASTERSALE' ).AsString := '1';
        {}
        aCD078A1.FieldByName( 'MON1' ).AsString := '12';
        aCD078A1.FieldByName( 'PERIOD1' ).AsString := aCItemCodes[aIndex,7];
        aCD078A1.FieldByName( 'RATETYPE1' ).AsString := '2';
        { ?t?? }
        if ( SameText( aCItemCodes[aIndex,8], '-' ) ) then
        begin
          aCD078A1.FieldByName( 'DISCOUNTAMT1' ).AsString := FloatToStr( 0 - StrToFloatDef( aCItemCodes[aIndex,6], 0 ) );
          if ( aIndex - 1 in [Low( aCItemCodes )..High( aCItemCodes ) ] ) then
            if ( not SameText( aCItemCodes[aIndex-1, 8], '-' ) ) then
              aCD078A1.FieldByName( 'LINKKEY' ).AsString := aCItemCodes[aIndex-1, 9];
        end else
        begin
          aCD078A1.FieldByName( 'DISCOUNTAMT1' ).AsString := aCItemCodes[aIndex,6];
          { ?Y??????????, ?N?? StepNo ?d?U??, ???U?@?????t????, ?????N?? StepNo ?a?i?h }
          aCItemCodes[aIndex, 9] := aCD078A1.FieldByName( 'STEPNO' ).AsString;
        end;
        {}  
        aCD078A1.FieldByName( 'MONTHAMT1' ).AsFloat := aMonthAmt;
        aCD078A1.FieldByName( 'DAYAMT1' ).AsFloat := aDayAmt;
        {}

        aCD078A1.Post;
      end;
    end;
  except
    on E: Exception do AErrMsg := E.Message;
  end;
  ResetDataSetFilter( aCD078A );
  ResetDataSetFilter( aCD078A1 );
  frmCD078B1.RestoreCD078A1Data;
  Result := ( AErrMsg = EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

function TDBController.AutoAppendCD078C(var ARecord: TCD078BRecord;
  const ADataSet: TClientDataSet; var AErrMsg: String): Boolean;
var
  aIndex: Integer;
  aCItemCodes: array [1..5, 1..5] of String;
  aInstCode, aFaciItem: String;
begin
  AErrMsg := EmptyStr;
  ZeroMemory( @aCItemCodes, SizeOf( aCItemCodes ) );
  { 1.???? CD005 ?????????}?C??, ???O???? }
  DataReader.Close;
  DataReader.SQL.Text := Format(
    ' SELECT A.CITEMCODE1, A.CITEMCODE2, A.CITEMCODE3,    ' +
    '        A.CITEMCODE4, A.CITEMCODE5,                  ' +
    '        A.CITEMNAME1, A.CITEMNAME2, A.CITEMNAME3,    ' +
    '        A.CITEMNAME4, A.CITEMNAME5,                  ' +
    '        A.SERVICETYPE                                ' +
    '   FROM %s.CD005 A                                   ' +
    '  WHERE A.CODENO = ''%s''                            ',
    [FLoginInfo.DbAccount, ARecord.InstCode] );
  DataReader.Open;
  while not DataReader.Eof do
  begin
    for aIndex := Low( aCItemCodes ) to High( aCItemCodes ) do
    begin
      aCItemCodes[aIndex,1] := DataReader.FieldByName( Format( 'CItemCode%d', [aIndex] ) ).AsString;
      aCItemCodes[aIndex,2] := DataReader.FieldByName( Format( 'CItemName%d', [aIndex] ) ).AsString;
      aCItemCodes[aIndex,3] := DataReader.FieldByName( 'ServiceType' ).AsString;
    end;
    DataReader.Next;
  end;
  { 2. ???? CD019 ?O?_???????? }
  for aIndex := Low( aCItemCodes ) to High( aCItemCodes ) do
  begin
    DataReader.Close;
    DataReader.SQL.Text := Format(
      ' SELECT A.SIGN, A.AMOUNT           ' +
      '   FROM %s.CD019 A                 ' +
      '  WHERE A.CODENO = ''%s''          ' +
      '    AND A.PERIODFLAG = ''0''       ' +
      '    AND A.STOPFLAG = ''0''         ',
      [FLoginInfo.DbAccount, aCItemCodes[aIndex,1]] );
    DataReader.Open;
    if ( DataReader.IsEmpty ) then
    begin
      { ?S?????? CD019 ????, ?M?????????X?? CD005.CItemCode }
      aCItemCodes[aIndex,1] := EmptyStr;
      aCItemCodes[aIndex,2] := EmptyStr;
      aCItemCodes[aIndex,3] := EmptyStr;
    end else
    begin
      aCItemCodes[aIndex,4] := DataReader.FieldByName( 'SIGN' ).AsString;
      aCItemCodes[aIndex,5] := Nvl( DataReader.FieldByName( 'AMOUNT' ).AsString, '0' );
    end;
  end;
  DataReader.Close;
  { 3.???d?O?_????????????, ???????? Update InstCode ????, ?_?h?s?W
      ?@???? CD078C, ???K???????????????]???i?h }
  aFaciItem := ExtractValue( 1, ARecord.FaciItem );
  try
    for aIndex := Low( aCItemCodes ) to High( aCItemCodes ) do
    begin
      if ( aCItemCodes[aIndex,1] = EmptyStr ) then Continue;
      SetCD078CDuplicateFilter( ADataSet, ARecord.BPCode, aCItemCodes[aIndex,1], ARecord.ServiceType );
      if ( ADataSet.RecordCount <= 0 ) then
      begin
        ADataSet.Append;
        ADataSet.FieldByName( 'BpCode' ).AsString := ARecord.BPCode;
        ADataSet.FieldByName( 'CItemCode' ).AsString := aCItemCodes[aIndex,1];
        ADataSet.FieldByName( 'CItemName' ).AsString := aCItemCodes[aIndex,2];
        ADataSet.FieldByName( 'ServiceType' ).AsString := aCItemCodes[aIndex,3];
        ADataSet.FieldByName( 'FaciItem' ).AsString := aFaciItem;
        ADataSet.FieldByName( 'InstCodeStr' ).AsString := ARecord.InstCode;
        ADataSet.FieldByName( 'Amount' ).AsString := aCItemCodes[aIndex,5];
        ADataSet.FieldByName( 'DiscountAmt' ).AsString := aCItemCodes[aIndex,5];
        { ?Y???t??, ???B?????t }
        if ( aCItemCodes[aIndex,4] = '-' ) then
        begin
          ADataSet.FieldByName( 'Amount' ).AsFloat := ( 0 - ADataSet.FieldByName( 'Amount' ).AsFloat );
          ADataSet.FieldByName( 'DiscountAmt' ).AsFloat := ( 0 - ADataSet.FieldByName( 'DiscountAmt' ).AsFloat );
        end;
        ADataSet.Post;
      end else
      begin
        aInstCode := ADataSet.FieldByName( 'InstCodeStr' ).AsString;
        if not ValueExists( ARecord.InstCode, aInstCode ) then
        begin
          ADataSet.Edit;
          ADataSet.FieldByName( 'InstCodeStr' ).AsString := AddValue( ARecord.InstCode, aInstCode );
          ADataSet.Post;
        end;
      end;
    end;
  except
    on E: Exception do AErrMsg := E.Message;
  end;
  ResetDataSetFilter( ADataSet );
  Result := ( AErrMsg = EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

function TDBController.VdInstCodeExitisInCD078B(AInstCodes: String;
  const ADataSet: TClientDataSet): Boolean;
var
  aExt: String;
begin
  Result := False;
  AInstCodes := TrimChar( AInstCodes, [''''] );
  while ( AInstCodes <> EmptyStr ) do
  begin
    aExt := ExtractValue( AInstCodes );
    Result := False;
    ADataSet.First;
    while not ADataSet.Eof do
    begin
      if ( ADataSet.FieldByName( 'INSTCODE' ).AsString = aExt ) then
      begin
        Result := True;
        Break;
      end;
      ADataSet.Next;
    end;
    { ?u?n?????@?? InstCode, ???????N?????????~?????? }
    if not Result then Break;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TDBController.GetInstName(AValues: String): String;
var
  aText: String;
begin
  Result := EmptyStr;
  repeat
    aText := ExtractValue( AValues );
    if ( aText <> EmptyStr ) then
    begin
      DataReader.Close;
      DataReader.SQL.Text := Format(
        ' SELECT A.DESCRIPTION FROM %s.CD005 A ' +
        '  WHERE A.CODENO = ''%s''             ',
        [DBController.LoginInfo.DbAccount, aText] );
      DataReader.Open;
      if ( not DataReader.IsEmpty ) then
      begin
        Result := ( Result +
          DataReader.FieldByName( 'DESCRIPTION' ).AsString + ',' );
      end;
      DataReader.Close;
    end;
  until ( AValues = EmptyStr );
  if IsDelimiter( ',', Result, Length( Result ) ) then
    Delete( Result, Length( Result ), 1 );
end;

{ ---------------------------------------------------------------------------- }

function TDBController.GetStepNo: String;
begin
  DataReader.Close;
  DataReader.SQL.Text := Format(
    ' SELECT %s.S_CD078A1.NEXTVAL FROM DUAL ',
    [DBController.LoginInfo.DbAccount] );
  DataReader.Open;
  Result := DataReader.Fields[0].AsString;
  DataReader.Close;
end;

{ ---------------------------------------------------------------------------- }

function TDBController.GetDiscountAmt(const ACodeNo: String): Double;
begin
  Result := 0;
  if ( ACodeNo <> EmptyStr ) then
  begin
    DataReader.Close;
    DataReader.SQL.Text := Format(
      ' SELECT A.AMOUNT FROM %s.CD019 A  ' +
      '  WHERE A.CODENO = ''%s''         ',
      [DBController.LoginInfo.DbAccount, ACodeNo] );
    DataReader.Open;
    Result := DataReader.Fields[0].AsFloat;
    DataReader.Close;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TDBController.GetPeriod(const ACodeNo: String): Integer;
begin
  Result:=0;
  if ( ACodeNo <> EmptyStr ) then
  begin
    DataReader.Close;
    DataReader.SQL.Text := Format(
      'SELECT Nvl(A.PERIOD,0) FROM %s.CD019 A ' +
      ' Where A.CODENO = ''%s''        ',
      [DBController.LoginInfo.DbAccount,ACodeNo] );
    DataReader.Open;
    Result := DataReader.Fields[0].AsInteger;
    DataReader.Close;
  end;
end;
{ ---------------------------------------------------------------------------- }
function TDBController.GetCitemName(const ACodeNo: String): string;
begin
  Result := EmptyStr;
  if ( ACodeNo <> EmptyStr ) then
  begin
    DataReader.Close;
    DataReader.SQL.Text := Format(
      'SELECT DESCRIPTION FROM %s.CD019 A  ' +
      ' WHERE A.CODENO = %s                ',
      [DBController.LoginInfo.DbAccount,ACodeNo] );
    DataReader.Open;
    if not DataReader.IsEmpty then
    begin
      Result := DataReader.Fields[0].AsString;
    end;
    DataReader.Close;
  end;

end;

function TDBController.AutoAppendCD078J(var ARecord: TCD078BRecord;
  const ADataSet: TClientDataSet; var AerrMsg: String): Boolean;
  var
  aIndex: Integer;
  aProductCodes: array [1..5, 1..5] of String;
  aInstCode, aFaciItem: String;
begin
  AErrMsg := EmptyStr;
  ZeroMemory( @aProductCodes, SizeOf( aProductCodes ) );
  { 1.???? CD005 ?????????}?C??, ???O???? }
  DataReader.Close;
  DataReader.SQL.Text := Format(
    ' SELECT A.PRODUCTCODE1, A.PRODUCTCODE2, A.PRODUCTCODE3,    ' +
    '        A.PRODUCTCODE4, A.PRODUCTCODE5,                    ' +
    '        A.PRODUCTNAME1, A.PRODUCTNAME2, A.PRODUCTNAME3,    ' +
    '        A.PRODUCTNAME4, A.PRODUCTNAME5,                    ' +
    '        A.SERVICETYPE                                      ' +
    '   FROM %s.CD005 A                                         ' +
    '  WHERE A.CODENO = ''%s''                                  ',
    [FLoginInfo.DbAccount, ARecord.InstCode] );
  DataReader.Open;
  while not DataReader.Eof do
  begin
    for aIndex := Low( aProductCodes ) to High( aProductCodes ) do
    begin
      aProductCodes[aIndex,1] := DataReader.FieldByName( Format( 'PRODUCTCODE%d', [aIndex] ) ).AsString;
      aProductCodes[aIndex,2] := DataReader.FieldByName( Format( 'PRODUCTNAME%d', [aIndex] ) ).AsString;
      aProductCodes[aIndex,3] := DataReader.FieldByName( 'ServiceType' ).AsString;
    end;
    DataReader.Next;
  end;

  DataReader.Close;
  { 3.???d?O?_????????????, ???????? Update InstCode ????, ?_?h?s?W
      ?@???? CD078J, ???K???????????????]???i?h }
  aFaciItem := ExtractValue( 1, ARecord.FaciItem );
  try
    for aIndex := Low( aProductCodes ) to High( aProductCodes ) do
    begin
      if ( aProductCodes[aIndex,1] = EmptyStr ) then Continue;
      SetCD078JCDuplicateFilter( ADataSet, ARecord.BPCode, aProductCodes[aIndex,1], ARecord.ServiceType );
      if ( ADataSet.RecordCount <= 0 ) then
      begin
        ADataSet.Append;
        ADataSet.FieldByName( 'BpCode' ).AsString := ARecord.BPCode;
        ADataSet.FieldByName( 'ProductCode' ).AsString := aProductCodes[aIndex,1];
        ADataSet.FieldByName( 'ProductName' ).AsString := aProductCodes[aIndex,2];
        ADataSet.FieldByName( 'ServiceType' ).AsString := aProductCodes[aIndex,3];
        ADataSet.FieldByName( 'FaciItem' ).AsString := aFaciItem;
        ADataSet.FieldByName( 'InstCodeStr' ).AsString := ARecord.InstCode;

        ADataSet.Post;
      end else
      begin
        aInstCode := ADataSet.FieldByName( 'InstCodeStr' ).AsString;
        if not ValueExists( ARecord.InstCode, aInstCode ) then
        begin
          ADataSet.Edit;
          ADataSet.FieldByName( 'InstCodeStr' ).AsString := AddValue( ARecord.InstCode, aInstCode );
          ADataSet.Post;
        end;
      end;
    end;
  except
    on E: Exception do AErrMsg := E.Message;
  end;
  ResetDataSetFilter( ADataSet );
  Result := ( AErrMsg = EmptyStr );
end;

procedure TDBController.SetCD078JCDuplicateFilter(ADataSet: TClientDataSet;
  ABpCode, AProductCode, AService: String);
begin
   ADataSet.Filtered := False;
  ADataSet.Filter := Format( 'BPCODE=''%s'' AND ProductCode=''%s'' AND SERVICETYPE=''%s''',
    [ABpCode, AProductCode, AService] );
  ADataSet.Filtered := True;
end;

function TDBController.GetUserPriv(const Mid: String): Boolean;
begin
  if FGroupId = '0' then
  begin
    Result := True;
    Exit;
  end;
  try
    try
      DataReader.Close;
      DataReader.SQL.Text := Format(
                        'SELECT COUNT(*) CNT FROM %s.SO029 ' +
                        ' WHERE MID = ''%s''               ' +
                        ' AND GROUP%s = 1                  ',
                        [DBController.LoginInfo.DbAccount,
                        Mid,
                        FGroupId]);

      DataReader.Open;
      Result := DataReader.FieldByName( 'CNT' ).AsInteger = 1;

    except
      on E: Exception do
      begin
        Result := False;
      end;
    end;
  finally
    DataReader.Close;
  end;
end;
procedure TDBController.GetGropuId;
begin
  FGroupId := '0';
  try
    if not LoginInfo.IsSuperVisor then
    begin
      DataReader.Close;
      DataReader.SQL.Text := format('SELECT GroupID FROM %s.SO026 ' +
                                    ' WHERE UserId = ''%s''        '             +
                                    ' AND COMPCODE = %s            ',
                                    [DBController.LoginInfo.DbAccount,
                                    DBController.LoginInfo.UserId,
                                    DBController.LoginInfo.CompCode]);
      DataReader.Open;
      if DataReader.RecordCount > 0 then
        FGroupId := DBController.DataReader.FieldByName( 'GroupID' ).AsString;
    end;
  finally
    DataReader.Close;
  end;
end;

function TDBController.GetCanUseEPG: Boolean;
begin
  Result := GetUserPriv('SO1100G');
end;

end.
