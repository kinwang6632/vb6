unit frmcd0421;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, Buttons, ComObj, DB, ADODB, ComCtrls;

type
  PDbInfo = ^TDbInfo;
  TDbInfo = record
    CompCode: Integer;
    UserId: String;
    Passowrd: String;
    TnsName: String;
    ConnectionString: String;
  end;

  TForm2 = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    Bevel1: TBevel;
    btnSave: TButton;
    btnCancel: TButton;
    btnCodeNo: TButton;
    btnClear1: TBitBtn;
    btnCompCode: TButton;
    btnClear2: TBitBtn;
    RadioButton1: TRadioButton;
    RadioButton2: TRadioButton;
    EDT_CodeNoStr: TEdit;
    EDT_CompCodeStr: TEdit;
    CopyConnection: TADOConnection;
    CopyWriter: TADOCommand;
    CopyReader: TADOQuery;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnCodeNoClick(Sender: TObject);
    procedure btnCompCodeClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure btnClear1Click(Sender: TObject);
    procedure btnClear2Click(Sender: TObject);
  private
    { Private declarations }
    EnCryptX: OleVariant;
    FDbInfoList: TStringList;
    FCodeNos: String;
    FCodeNames: string;
    FCompCodes: String;
    FCompNames: String;
    FCopyErrorList: TStringList;
    function GetCanSelectCompCode: String;
    function InternalCopy(aDbInfo: PDbInfo; const ACodeNo, ACodeName: String): Boolean;
    procedure DoCopy;
    function RemoteRecordExists(aDbInfo: PDbInfo; const aCodeNo: String): Boolean;
    procedure DeleteRemoteRecords(aDbInfo: PDbInfo; const aCodeNo: String);
  public
    { Public declarations }
    procedure LoadGICMIS2;
    procedure Cleanup;
  end;


var
  Form2: TForm2;

implementation

uses cbUtilis, frmcd042, frmMultiSelectU;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

function IsEqualCompCodeText(const aText: String; var aCompCode: Integer): Boolean;
var
  aIndex: Integer;
begin
  aCompCode := -1;
  for aIndex := 1 to 15 do
  begin
    Result := ( aText = Format( '[%d]', [aIndex] ) );
    if Result then
    begin
      aCompCode := aIndex;
      Break;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm2.FormCreate(Sender: TObject);
begin
  FCodeNos := EmptyStr;
  FCodeNos := EmptyStr;
  FCompCodes := EmptyStr;
  FCompNames := EmptyStr;
  EnCryptX := CreateOleObject('DevPower.Encrypt');
  FDbInfoList := TStringList.Create;
  FCopyErrorList := TStringList.Create;
  LoadGICMIS2;
  GetCanSelectCompCode;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm2.FormDestroy(Sender: TObject);
begin
  EnCryptX := Null;
  Cleanup;
  FCopyErrorList.Free;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm2.FormShow(Sender: TObject);
begin

end;

{ ---------------------------------------------------------------------------- }

procedure TForm2.LoadGICMIS2;
var
  aPath: String;
  aLine: WideString;
  aIndex, aCompCode: Integer;
  aTmpList: TStringList;
  aObj: PDbInfo;
begin
  aPath := IncludeTrailingPathDelimiter( ExtractFilePath(
    Application.ExeName ) ) + 'GICMIS2.ini';
  aTmpList := TStringList.Create;
  try
    aTmpList.LoadFromFile( aPath );
    for aIndex := 0 to aTmpList.Count - 1 do
    begin
      if IsEqualCompCodeText( aTmpList[aIndex], aCompCode ) then
      begin
        New( aObj );
        aObj.CompCode := aCompCode;
        {}
        aLine := Copy( aTmpList[aIndex + 1], 3, Length( aTmpList[aIndex + 1] ) - 2 );
        aObj.TnsName := EnCryptX.DeCrypt( aLine );
        {}
        aLine := Copy( aTmpList[aIndex + 2], 3, Length( aTmpList[aIndex + 2] ) - 2 );
        aObj.UserId := EnCryptX.DeCrypt( aLine );
        {}
        aLine := Copy( aTmpList[aIndex + 3], 3, Length( aTmpList[aIndex + 3] ) - 2 );
        aObj.Passowrd := EnCryptX.DeCrypt( aLine );
        {}
        aObj.ConnectionString := Format(
          'Provider=MSDAORA.1;Password=%s;User ID=%s;Data Source=%s;Persist Security Info=True',
          [aObj.Passowrd, aObj.UserId, aObj.TnsName] );
        FDbInfoList.AddObject( IntToStr( aObj.CompCode ), TObject( aObj ) );
      end;
    end;
  finally
    aTmpList.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm2.Cleanup;
var
  aIndex: Integer;
begin
  for aIndex := 0 to FDbInfoList.Count - 1 do
  begin
    Dispose( PDbInfo( FDbInfoList.Objects[aIndex] ) );
    FDbInfoList.Objects[aIndex] := nil;
  end;
  FDbInfoList.Clear;
end;

{ ---------------------------------------------------------------------------- }

function TForm2.GetCanSelectCompCode: String;
var
  aReader: TADOQuery;
  aIndex: Integer;
begin
  Result := EmptyStr;
  if ( Form1.sG_OperatorID = 'A' ) then
  begin
    for aIndex := 1 to 20 do
      Result := ( Result + Format( '%d,', [aIndex] ) );
    if IsDelimiter( ',', Result, Length( Result ) ) then
      Delete( Result, Length( Result ), 1 );
  end;
  if ( Form1.sG_IsSupervisor <> 'Y' ) and ( Form1.sG_OperatorID <> 'A' ) then
  begin
    aReader := Form1.ADOQuery2;
    aReader.Close;
    aReader.SQL.Text := Format(
      ' SELECT COMPSTR FROM %s.SO026 WHERE USERID=''%s'' ',
      [Form1.sG_DbUserID, Form1.sG_OperatorID] );
    aReader.Open;
    Result := aReader.Fields[0].AsString;
    aReader.Close;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm2.btnClear1Click(Sender: TObject);
begin
  FCodeNos := EmptyStr;
  FCodeNames := EmptyStr;
  EDT_CodeNoStr.Text := EmptyStr;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm2.btnClear2Click(Sender: TObject);
begin
  FCompCodes := EmptyStr;
  FCompNames := EmptyStr;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm2.btnCodeNoClick(Sender: TObject);
var
  aSql:String;
begin
  aSql := Format(
    ' SELECT CODENO, DESCRIPTION FROM %s.CD042 ' +
    '  WHERE STOPFLAG = 0 ORDER BY CODENO ', [Form1.sG_DbUserID] );
  frmMultiSelect := TfrmMultiSelect.Create( Application );
  try
    frmMultiSelect.Connection := Form1.ADOConnection1;
    frmMultiSelect.KeyFields := 'CODENO';
    frmMultiSelect.KeyValues := FCodeNos;
    frmMultiSelect.DisplayFields := 'CODENO,?P?P?????N?X,DESCRIPTION,?P?P?????W??';
    frmMultiSelect.ResultFields := 'DESCRIPTION';
    frmMultiSelect.SQL.Text := aSql;
    if frmMultiSelect.ShowModal = mrOk then
    begin
      FCodeNos := frmMultiSelect.SelectedValue;
      FCodeNames := frmMultiSelect.SelectedDisplay;
      EDT_CodeNoStr.Text := frmMultiSelect.SelectedDisplay;
    end;
  finally
    frmMultiSelect.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm2.btnCompCodeClick(Sender: TObject);
var
  aSql, aSelComps: String;
begin
  aSql := Format(
    ' SELECT CODENO, DESCRIPTION FROM %s.CD039 ' +
    '  WHERE CODENO <> ''%s'' ', [Form1.sG_DbUserID, Form1.sG_CompID] );
  aSelComps := GetCanSelectCompCode;
  if ( aSelComps <> EmptyStr ) then
    aSql := ( aSql + Format( ' AND CODENO IN ( %s )', [aSelComps] ) )
  else
    aSql := ( aSql + ' AND 1=2' );
  {}
  aSql := ( aSql + ' ORDER BY CODENO ' );
  {}
  frmMultiSelect := TfrmMultiSelect.Create( Application );
  try
    frmMultiSelect.Connection := Form1.ADOConnection1;
    frmMultiSelect.KeyFields := 'CODENO';
    frmMultiSelect.KeyValues := FCompCodes;
    frmMultiSelect.DisplayFields := 'CODENO,?N?X,DESCRIPTION,?W??';
    frmMultiSelect.ResultFields := 'DESCRIPTION';
    frmMultiSelect.SQL.Text := aSql;
    if frmMultiSelect.ShowModal = mrOk then
    begin
      FCompCodes := frmMultiSelect.SelectedValue;
      FCompNames := frmMultiSelect.SelectedDisplay;
      EDT_CompCodeStr.Text := frmMultiSelect.SelectedDisplay;
    end;
  finally
    frmMultiSelect.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TForm2.InternalCopy(aDbInfo: PDbInfo; const ACodeNo, ACodeName: String): Boolean;
var
  aReader: TADOQuery;

  { -------------------------------------------------- }

  function GenHeader(const ATableName: String): String;
  var
    aIndex: Integer;
    aSqlBody: String;
  begin
    aSqlBody := EmptyStr;
    for aIndex := 0 to aReader.Fields.Count - 1 do
      aSqlBody := ( aSqlBody + aReader.Fields[aIndex].FieldName + ',' );
    if IsDelimiter( ',', aSqlBody, Length( aSqlBody ) ) then
      Delete( aSqlBody, Length( aSqlBody ), 1 );
    Result := Format( ' INSERT INTO %s.%s ( ', [aDbInfo.UserId, ATableName] ) + aSqlBody + ' ) ';
  end;

  { -------------------------------------------------- }

  function GenFooter: String;
  var
    aIndex: Integer;
    aSqlFooter, aOraValue, aFmtText, aOraFmtText: String;
    aFieldType: TFieldType;
  begin
    aSqlFooter := EmptyStr;
    for aIndex := 0 to aReader.Fields.Count - 1 do
    begin
      aFieldType := aReader.Fields[aIndex].DataType;
      aOraValue := 'NULL';
      if ( not VarIsNull( aReader.Fields[aIndex].Value ) ) then
      begin
        case aFieldType of
          ftDate, ftDateTime:
            begin
              aFmtText := 'yyyy/mm/dd';
              aOraFmtText := 'YYYY/MM/DD';
              if Frac( aReader.Fields[aIndex].AsDateTime ) > 0 then
              begin
                aFmtText := 'yyyy/mm/dd hh:nn:ss';
                aOraFmtText := 'YYYY/MM/DD HH24:MI:SS';
              end;
              aOraValue := FormatDateTime( aFmtText, aReader.Fields[aIndex].AsDateTime );
              aOraValue := Format( 'TO_DATE( ''%s'', ''%s'' )', [aOraValue, aOraFmtText] );
            end;
        else
          aOraValue := QuotedStr( aReader.Fields[aIndex].AsString );
        end;
        if SameText( aReader.Fields[aIndex].FieldName, 'COMPCODE' ) then
          aOraValue := QuotedStr( IntToStr( aDbInfo.CompCode ) );
      end;
      aSqlFooter := ( aSqlFooter + aOraValue + ',' );
    end;
    if IsDelimiter( ',', aSqlFooter, Length( aSqlFooter ) ) then
      Delete( aSqlFooter, Length( aSqlFooter ), 1 );
    Result := ' VALUES ( ' + aSqlFooter + ' ) ';  
  end;

  { -------------------------------------------------- }

  function GenInsertSql(const ATableName: String): String;
  begin
    Result := GenHeader( ATableName ) + #13 + GenFooter;
  end;

  { -------------------------------------------------- }

var
  aCanInsert: Boolean;
begin
  {}
  Result := True;
  aReader := Form1.ADOQuery2;
  aCanInsert := True;
  if ( RadioButton1.Checked ) then
    aCanInsert := RemoteRecordExists( aDbInfo, aCodeNo );
  {}
  if not aCanInsert then Exit;
  {}
  CopyConnection.BeginTrans;
  try
    DeleteRemoteRecords( aDbInfo, aCodeNo );
    aReader.Close;
    aReader.SQL.Text := Format( ' SELECT * FROM %s.CD042 WHERE CODENO = ''%s'' ',
      [Form1.sG_DbUserID, aCodeNo] );
    aReader.Open;
    CopyWriter.CommandText := GenInsertSql( 'CD042' );
    CopyWriter.Execute;
    aReader.Close;
    {}
    aReader.SQL.Text := Format( ' SELECT * FROM %s.CD042A WHERE PROMCODE = ''%s'' ',
      [Form1.sG_DbUserID, aCodeNo] );
    aReader.Open;
    if ( not aReader.IsEmpty ) then
    begin
      CopyWriter.CommandText := GenInsertSql( 'CD042A' );
      CopyWriter.Execute;
    end;
    aReader.Close;
    CopyConnection.CommitTrans;
  except
    on E: Exception do
    begin
      CopyConnection.RollbackTrans;
      FCopyErrorList.Add( Lpad( EmptyStr, 6, #32 ) +
        Format( '?W??%s , ???]:%s?C', [ACodeName, E.Message] ) );
      Result := False;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm2.DoCopy;
var
  aCodeNo, aCodeName: String;
  aCompCode, aCompName: String;
  aTmpCompCode, aTmpCompName: String;
  aTmpCodeNo, aTmpCodeName: String;
  aMessage: string;
  aDbInfo: PDbInfo;
  aIdx, aErrCount, aTotalErrCount: Integer;
begin
  aMessage := EmptyStr;
  aTmpCompCode := FCompCodes;
  aTmpCompName := FCompNames;
  aTotalErrCount := 0;
  repeat
    aCompCode := StringReplace( ExtractValue( aTmpCompCode ), '''', EmptyStr, [rfReplaceAll] );
    aCompName := ExtractValue( aTmpCompName );
    if ( aCompCode = EmptyStr) then Continue;
    aIdx := FDbInfoList.IndexOf( aCompCode );
    if aIdx < 0 then Continue;
    aDbInfo := PDbInfo( FDbInfoList.Objects[aIdx] );
    CopyConnection.Close;
    CopyConnection.ConnectionString := aDbInfo.ConnectionString;
    CopyConnection.Open;
    try
      aTmpCodeNo := FCodeNos;
      aTmpCodeName := FCodeNames;
      FCopyErrorList.Clear;
      aErrCount := 0;
      while ( aTmpCodeNo <> EmptyStr ) do
      begin
        aCodeNo := StringReplace( ExtractValue( aTmpCodeNo ), '''', EmptyStr, [rfReplaceAll] );
        aCodeName := ExtractValue( aTmpCodeName );
        if ( not InternalCopy( aDbInfo, aCodeNo, aCodeName ) ) then
        begin
          Inc( aErrCount );
          Inc( aTotalErrCount );
        end;
      end;
      if ( aErrCount <= 0 ) then
        aMessage := ( aMessage + #13 + Format( '???q?O: %s, ???s?????C', [aCompName] ) )
      else begin
        aMessage := ( aMessage + #13 + Format( '???q?O: %s, ???s?????C', [aCompName] ) );
        aMessage := ( aMessage + FCopyErrorList.Text );
      end;
    finally
      CopyConnection.Close;
    end;
  until ( aTmpCompCode = EmptyStr );
  {}
  if ( aTotalErrCount <= 0 ) then
    InfoMsg( aMessage )
  else
    WarningMsg( aMessage );
end;

{ ---------------------------------------------------------------------------- }

procedure TForm2.btnSaveClick(Sender: TObject);
begin
  if ( FCodeNos = EmptyStr ) then
  begin
    WarningMsg( '???????P?P?????C' );
    if btnCodeNo.CanFocus then btnCodeNo.SetFocus;
    Exit;
  end;
  {}
  if ( FCompCodes = EmptyStr ) then
  begin
    WarningMsg( '?????????q?O?C' );
    if btnCompCode.CanFocus then btnCompCode.SetFocus;
    Exit;
  end;
  {}
  Screen.Cursor := crSQLWait;
  try
    DoCopy;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TForm2.RemoteRecordExists(aDbInfo: PDbInfo; const aCodeNo: String): Boolean;
begin
  CopyReader.Close;
  CopyReader.SQL.Text := Format(
    ' SELECT COUNT(1) FROM %s.CD042 WHERE CODENO = ''%s'' ',  [aDbInfo.UserId, aCodeNo] );
  CopyReader.Open;
  Result := ( CopyReader.Fields[0].AsInteger > 0 );
  CopyReader.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm2.DeleteRemoteRecords(aDbInfo: PDbInfo; const aCodeNo: String);
begin
  CopyWriter.CommandText := Format(
    ' DELETE FROM %s.CD042 WHERE CODENO = ''%s'' ', [aDbInfo.UserId, aCodeNo] );
  CopyWriter.Execute;
  CopyWriter.CommandText := Format(
    ' DELETE FROM %s.CD042A WHERE PROMCODE = ''%s'' ', [aDbInfo.UserId, aCodeNo] );
  CopyWriter.Execute;
end;

{ ---------------------------------------------------------------------------- }

end.
