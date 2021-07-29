unit cbParseThread;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, DateUtils, ActiveX, Math,
  cbClass, cbAppClass, cbParseController;

type
  TParseThread = class(TMessageQueueThread)
  private
    { Private declarations }
    FLastExecute: TDateTime;
    FLastChangeDetect: TDateTime;
    FLastDbErr: TDateTime;
    FMsgSubject: TMessageSubject;
    FParseSubject: TParseSubject;
    FParam: TAppParam;
    FList: TList;
    FDataBase: TAppDatabase;
    procedure SetParam(const Value: TAppParam);
    procedure SetDataBase(const Value: TAppDatabase);
    function XmlFileDetected: Boolean;
    function CanDataBaseRetryConnect: Boolean;
    function CanChangeCountDetect: Boolean;
    function DoDatabaseOpen(const aSlient: Boolean): Boolean;
    procedure DoDatabaseClose;
    function MappingXmlFileType(aPrefix: String): TXmlFileType;
    function ParseXmlFile: Boolean;
    function ChangeDetected: Boolean;
    function MergeChange: Boolean;
    procedure SortFileOrder;
    procedure XmlFileMove;
    procedure AddToList;
    procedure DeleteList;
    procedure CleanupList;
  protected
    procedure Execute; override;
    procedure WndProc(var Msg: TMessage); override;
  public
    constructor Create;
    destructor Destroy; override;
    property MessageSubject: TMessageSubject read FMsgSubject;
    property ParseSubject: TParseSubject read FParseSubject;
    property Param: TAppParam read FParam write SetParam;
    property DataBase: TAppDatabase read FDataBase write SetDataBase;
  end;

implementation


uses cbUtilis;

{ ---------------------------------------------------------------------------- }

{ TXmlParseThread }

constructor TParseThread.Create;
begin
  inherited Create( True );
  CoInitialize(nil);
  FMsgSubject := TMessageSubject.Create;
  FParseSubject := TParseSubject.Create;
  FParam := TAppParam.Create;
  FDataBase := TAppDatabase.Create;
  ParseController := TParseController.Create( nil );
  FList := TList.Create;
end;

{ ---------------------------------------------------------------------------- }

destructor TParseThread.Destroy;
begin
  FParam.Free;
  FDataBase.Free;
  FParseSubject.Free;
  FMsgSubject.Free;
  ParseController.Free;
  FList.Free;
  CoUninitialize;
  inherited;
end;

{ ---------------------------------------------------------------------------- }

function TParseThread.XmlFileDetected: Boolean;
var
  aMask:String;
  aSearch: TSearchRec;
  aIndex, aFileCount, aPreFileCount: Integer;
  aDir: TXmlDirectory;
begin
  Result := False;
  aPreFileCount := 0;
  { ���а��� 20 ��, �T�w�ɮ׼ƶq�ۦP�~�i�B�z }
  for aIndex := 1 to 20 do
  begin
    aFileCount := 0;
    for aDir := Low( TXmlDirectory ) to High( TXmlDirectory ) do
    begin
      case aDir of
        xdsWrapper:
          aMask := IncludeTrailingPathDelimiter( FParam.WrapperDest ) + 'Nagra_*.xml';
        xdsDex:
          aMask := IncludeTrailingPathDelimiter( FParam.DexDest ) + 'CA*.xml';
        xdsAsRun:
          aMask := IncludeTrailingPathDelimiter( FParam.AsRunDest ) + 'Export*.xml';
      else
        Continue;
      end;
      if FindFirst( aMask, faAnyFile, aSearch ) = 0 then
      begin
        try
          repeat
            if ( aSearch.Attr and faDirectory ) <> 0 then
            begin
              if ( aSearch.Name = '.' ) or ( aSearch.Name = '..' ) then Continue;
            end else
            if ( aSearch.Attr and faArchive ) <> 0 then
            begin
              Inc( aFileCount );
            end;
            if ( Self.Terminated ) then Break;
          until FindNext( aSearch ) <> 0;
        finally
          FindClose( aSearch );
        end;
      end;
    end;
    if ( Self.Terminated ) then Break;
    if ( aPreFileCount = 0 ) then aPreFileCount := aFileCount;
    Result := ( aPreFileCount = aFileCount ) and ( aFileCount > 0 );
    if not Result then Break;
    Sleep( 100 );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TParseThread.CanDataBaseRetryConnect: Boolean;
begin
  Result := ( FLastDbErr = 0 );
  if ( not Result ) then
    Result := ( SecondsBetween( Now, FLastDbErr ) >= FDataBase.DbRetrySec );
end;

{ ---------------------------------------------------------------------------- }

function TParseThread.CanChangeCountDetect: Boolean;
begin
  Result := ( FLastChangeDetect = 0 );
  if ( not Result ) then
    Result := ( SecondsBetween( Now, FLastChangeDetect ) >= 90 );
end;

{ ---------------------------------------------------------------------------- }

function TParseThread.DoDatabaseOpen(const aSlient: Boolean): Boolean;
begin
  DoDatabaseClose;
  if CanDataBaseRetryConnect then
  begin
    if not aSlient then FMsgSubject.Normal( '�i�ѪR�֤ߡj�ˬd��Ʈw�s�u��...' );
    Sleep( 300 );
    try
      ParseController.DataConnection.ConnectionString := Format(
        'Provider=MSDAORA.1;Persist Security Info=True;' +
        'Password=%s;User ID=%s;Data Source=%s;',
        [FDataBase.DbPassoword, FDataBase.DbAccount, FDataBase.DbSid] );
      ParseController.DataConnection.Open;
      if not aSlient then FMsgSubject.Normal( '�i�ѪR�֤ߡj��Ʈw�s�u���`�C' );
      FLastDbErr := 0;
    except
      on E: Exception do
      begin
        FMsgSubject.Error( Format( '�i�ѪR�֤ߡj��Ʈw�s������, ��]: %s �C', [E.Message] ) );
        Sleep( 300 );
        FMsgSubject.Warning( '�i�ѪR�֤ߡj�L�k�s����Ʈw, �Ȱ��ѪR�����Ʈw�s�u�Ƨ��C' );
        Sleep( 300 );
        FMsgSubject.Warning( Format( '�i�ѪR�֤ߡj%d���N���ճs����Ʈw�C',
          [FDataBase.DbRetrySec] ) );
        FLastDbErr := Now;
      end;
    end;
  end;
  Sleep( 300 );
  Result := ParseController.DataConnection.Connected;
end;

{ ---------------------------------------------------------------------------- }

procedure TParseThread.DoDatabaseClose;
begin
  if ( ParseController.DataConnection.Connected ) then
  try
    ParseController.DataConnection.Close;
  except
    {...}
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TParseThread.Execute;
var
  aResult: Boolean;
begin
  WaitForPlaySignal;
  FLastExecute := 0;
  FLastChangeDetect := 0;
  FLastDbErr := 0;
  while not Self.Terminated do
  begin
    Sleep( 1000 );
    try
      if ( Self.Terminated ) then Break;
      { Xml ����ɮ� }
      if ( XmlFileDetected ) then
      begin
        if not DoDatabaseOpen( False ) then Continue;
        FLastExecute := Now;
        { �Ƨǩ���ɮ׶��� }
        SortFileOrder;
        if Self.Terminated then Break;
        { �[��B�z�ǦC }
        AddToList;
        if Self.Terminated then Break;
        Sleep( 300 );
        { ��� }
        ParseXmlFile;
        DoDatabaseClose;
        if Self.Terminated then Break;
        { �h��XML�ɮ�, PS:�@�w�n�b DeletList ���e }
        XmlFileMove;
        { �R�� OK �άO Error }
        DeleteList;
        Sleep( 300 );
        if Self.Terminated then Break;
      end;
      { �O�_��F�ˬd��Ʈw��ƪ��ɶ�, �w�] 60 �� }
      if ( not CanChangeCountDetect ) then Continue;
      { ��Ʈw���O�_���|���X�֪����ʸ�� }
      if not DoDatabaseOpen( True ) then Continue;
      aResult := ChangeDetected;
      Sleep( 300 );
      if Self.Terminated then Break;
      { �B�z��Ʈw�X�ָ�� }
      if ( aResult ) then MergeChange;
      DoDatabaseClose;
      FLastChangeDetect := Now;
      if Self.Terminated then Break;
      { �o���R�����O��C���� "�X��" ���� }
      DeleteList;
      if Self.Terminated then Break;
    except
      on E: Exception do
      begin
        FMsgSubject.Error( Format( '�i�ѪR�֤ߡj������o�Ϳ��~, ��]:%s�C',
          [E.Message] ) );
      end;
    end;
    Sleep( 300 );
  end;
  WaitForStopSignal;
  CleanupList;
end;

{ ---------------------------------------------------------------------------- }

procedure TParseThread.WndProc(var Msg: TMessage);
begin
  inherited WndProc( Msg );
end;

{ ---------------------------------------------------------------------------- }

procedure TParseThread.SetParam(const Value: TAppParam);
begin
  FParam.Assign( Value );
end;

{ ---------------------------------------------------------------------------- }

procedure TParseThread.SetDataBase(const Value: TAppDatabase);
begin
  FDataBase.Assign( Value );
end;

{ ---------------------------------------------------------------------------- }

function TParseThread.MappingXmlFileType(aPrefix: String): TXmlFileType;
begin
  Result := xfMerge;
  aPrefix := UpperCase( aPrefix );
  if ( aPrefix = 'NAGRA' ) then
    Result := xfWrapper
  else if ( aPrefix = 'CAPPV' ) then
    Result := xfPpv
  else if ( aPrefix = 'CASUB' ) then
    Result := xfSub
  else if ( aPrefix = 'EXPOR' ) then
    Result := xfAsRun;
end;

{ ---------------------------------------------------------------------------- }

procedure TParseThread.SortFileOrder;
var
  aMask, aPrefix, aDateTime, aPath: String;
  aType: TXmlFileType;
  aSearchRec: TSearchRec;
  aIndex: TXmlDirectory;
begin
  ParseController.ParseOrderDataSet.EmptyDataSet;
  for aIndex := Low( TXmlDirectory ) to High( TXmlDirectory ) do
  begin
    if ( Self.Terminated ) then Break;
    case aIndex of
      xdsWrapper:
        aMask := IncludeTrailingPathDelimiter( FParam.WrapperDest ) + 'Nagra_*.xml';
      xdsDex:
        aMask := IncludeTrailingPathDelimiter( FParam.DexDest ) + 'CA*.xml';
      xdsAsRun:
        aMask := IncludeTrailingPathDelimiter( FParam.AsRunDest ) + 'Export*.xml';
    else
      Continue;
    end;
    if FindFirst( aMask, faAnyFile, aSearchRec ) = 0 then
    begin
      try
        repeat
          if ( aSearchRec.Attr and faDirectory ) <> 0 then
          begin
            if ( aSearchRec.Name = '.' ) or ( aSearchRec.Name = '..' ) then Continue;
          end else
          begin
            aPrefix := EmptyStr;
            { Wrapper, Dex File, AsRunLog ���e5�X�r���Y�i }
            if ( Pos( 'NAGRA_', UpperCase( aSearchRec.Name ) ) > 0 ) or
               ( Pos( 'CAPPV', UpperCase( aSearchRec.Name ) ) > 0 ) or
               ( Pos( 'CASUB', UpperCase( aSearchRec.Name ) ) > 0 ) or
               ( Pos( 'EXPORT-', UpperCase( aSearchRec.Name ) ) > 0 ) then
              aPrefix := Copy( aSearchRec.Name, 1, 5 );
            aType := MappingXmlFileType( aPrefix );
            if not ( aType in [xfWrapper..xfAsRun] ) then Continue;
            aDateTime := FormatDateTime( 'yyyymmddhhnnss', FileDateToDateTime(
              aSearchRec.Time ) );
            case aIndex of
              xdsWrapper:
                aPath := IncludeTrailingPathDelimiter( FParam.WrapperDest ) + aSearchRec.Name;
              xdsDex:
                aPath := IncludeTrailingPathDelimiter( FParam.DexDest ) + aSearchRec.Name;
              xdsAsRun:
                aPath := IncludeTrailingPathDelimiter( FParam.AsRunDest ) + aSearchRec.Name;
            end;
            ParseController.ParseOrderDataSet.AppendRecord( [Ord( aType ),
              aPath, aPrefix, aDateTime, 'N'] );
          end;
          if ( Self.Terminated ) then Break;
        until FindNext( aSearchRec ) <> 0;
      finally
        FindClose( aSearchRec );
      end;
    end;
  end;
  if ( Self.Terminated ) then
    ParseController.ParseOrderDataSet.EmptyDataSet;
end;

{ ---------------------------------------------------------------------------- }

procedure TParseThread.XmlFileMove;
var
  aSource, aBackupPath, aErrPath: String;
  aIndex: Integer;
  aObj: PActObject;
begin
  for aIndex := 0 to FList.Count - 1 do
  begin
    aObj := PActObject( FList[aIndex] );
    aSource := IncludeTrailingPathDelimiter( aObj.ActFilePath ) + aObj.ActFileName;
    case aObj.ActFileType of
      xfWrapper:
        begin
          aBackupPath := IncludeTrailingPathDelimiter( FParam.WrapperBackup );
          aErrPath := IncludeTrailingPathDelimiter( FParam.WrapperErr );
        end;
      xfPpv, xfSub:
        begin
          aBackupPath := IncludeTrailingPathDelimiter( FParam.DexBackup );
          aErrPath := IncludeTrailingPathDelimiter( FParam.DexErr );
        end;
      xfAsRun:
        begin
          aBackupPath := IncludeTrailingPathDelimiter( FParam.AsRunBackup );
          aErrPath := IncludeTrailingPathDelimiter( FParam.AsRunErr );
        end;
    else
      aBackupPath := EmptyStr;
      aErrPath := EmptyStr;
    end;
    if ( aObj.ActStatus = 'C' ) and ( aBackupPath <> EmptyStr ) then
    begin
      MoveFileEx( PChar( aSource ), PChar( aBackupPath + aObj.ActFileName ),
        MOVEFILE_REPLACE_EXISTING );
    end else
    if ( aObj.ActStatus = 'E' ) and ( aErrPath <> EmptyStr ) then
    begin
      MoveFileEx( PChar( aSource ), PChar( aErrPath + aObj.ActFileName ),
        MOVEFILE_REPLACE_EXISTING );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TParseThread.AddToList;

  { ----------------------------------------- }

  function GetFileSizeText(aOrignal: Double): String;
  begin
    if ( aOrignal > 0 ) and ( aOrignal <= 1024 ) then
      Result := '1K'
    else if ( aOrignal >= 1024 ) and ( aOrignal <= ( 1024 * 1024 * 10 ) ) then
      Result := Format( '%d KB', [Ceil( aOrignal / 1024 )] )
    else
      Result := Format( '%d MB', [Ceil( aOrignal / 1024 / 1024 )] );
  end;

  { ----------------------------------------- }

var
  aObj: PActObject;
  aFileSize: Double;
begin
  ParseController.ParseOrderDataSet.First;
  while not ParseController.ParseOrderDataSet.Eof do
  begin
    New( aObj );
    aObj.KeyId := EmptyStr;
    aObj.ActDate := FormatDateTime( 'yyyy/mm/dd', Date );
    aObj.ActTime := FormatDateTime( 'hh:nn:ss', Now );
    aObj.ActFileName := ExtractFileName(
      ParseController.ParseOrderDataSet.FieldByName( 'XmlFileName' ).AsString );
    aObj.ActFilePath := ExtractFilePath(
      ParseController.ParseOrderDataSet.FieldByName( 'XmlFileName' ).AsString );
    aObj.ActFileType := TXmlFileType(
      ParseController.ParseOrderDataSet.FieldByName( 'XmlFileType' ).AsInteger);
    aObj.ActSource := XmlFileTypeText[aObj.ActFileType];
    aObj.ActStatus := EmptyStr;
    aObj.ActCost := 0;
    aObj.ActProgress := 0;
    aObj.ActErrMsg := EmptyStr;
    aFileSize := CalcFileSizeInByte(
      ParseController.ParseOrderDataSet.FieldByName( 'XmlFileName' ).AsString );
    aObj.ActFileSize := GetFileSizeText( aFileSize );
    FList.Add( aObj );
    FParseSubject.ActionObject := aObj;
    FMsgSubject.Normal( Format( '�i�ѪR�֤ߡj�[�J�B�z�ɮ�: %s', [aObj.ActFileName] ) );
    ParseController.ParseOrderDataSet.Next;
    Sleep( 100 );
    if Self.Terminated then Break;
  end;
  ParseController.ParseOrderDataSet.EmptyDataSet;
  if ( Self.Terminated ) then CleanupList;
end;

{ ---------------------------------------------------------------------------- }

procedure TParseThread.DeleteList;
var
  aIndex: Integer;
begin
  for aIndex := FList.Count - 1 downto 0 do
  begin
    if not Assigned( FList[aIndex] ) then Continue;
    if ( PActObject( FList[aIndex]  ).ActFileType in [xfWrapper..xfAsRun] ) then
    begin
      if ( PActObject( FList[aIndex] ).ActStatus = 'C' ) or
         ( PActObject( FList[aIndex] ).ActStatus = 'E' ) then
      begin
        Dispose( PActObject( FList[aIndex] ) );
        FList.Delete( aIndex );
      end;
    end else
    begin
      Dispose( PActObject( FList[aIndex] ) );
      FList.Delete( aIndex );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TParseThread.CleanupList;
begin
  while FList.Count > 0 do
  begin
    if Assigned( FList[0] ) then
    begin
      Dispose( PActObject( FList[0] ) );
      FList[0] := nil;
    end;
    FList.Delete( 0 );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TParseThread.ParseXmlFile: Boolean;
var
  aIndex: Integer;
  aObj: PActObject;
  aMsg, aFileName: String;
  aParser: TXmlFile;
  aParserClass: TXmlFileClass;
  aStartTime: TDateTime;
begin
  Result := True; 
  for aIndex := 0 to FList.Count - 1 do
  begin
    aObj := PActObject( FList[aIndex] );
    if ( Result ) then
    begin
      aFileName := IncludeTrailingPathDelimiter( aObj.ActFilePath ) + aObj.ActFileName;
      case aObj.ActFileType of
        xfWrapper: aParserClass := TWrapperXmlFile;
        xfPpv: aParserClass := TDexPpvXmlFile;
        xfSub: aParserClass := TDexSubXmlFile;
        xfAsRun: aParserClass := TAsRunXmlFile;
      else
        aParserClass := nil;
      end;
      if not Assigned( aParserClass ) then Continue;
      aParser := aParserClass.Create;
      try
        aParser.XmlFileName := aFileName;
        FMsgSubject.Normal( Format( '�i�ѪR�֤ߡj��� %s ��.....', [aObj.ActFileName] ) );
        Sleep( 300 );
        aObj.ActStatus := 'P';
        FParseSubject.ActionObject := aObj;
        Sleep( 300 );
        aStartTime := Now;
        Result := aParser.Parse( aMsg );
        if Result then
          FMsgSubject.Normal( Format( '�i�ѪR�֤ߡj��� %s ����, �x�s��Ƥ�....', [aObj.ActFileName] ) )
        else
          FMsgSubject.Error( Format( '�i�ѪR�֤ߡj��� %s ����, ��]: %s�C ', [aMsg] ) );
        Sleep( 300 );  
        if Result then Result := aParser.Save( aMsg );
        if Result then
          FMsgSubject.Normal( '�i�ѪR�֤ߡj�x�s��Ƨ����C' )
        else
          FMsgSubject.Error( Format( '�i�ѪR�֤ߡj�x�s��ƥ���, ��]: %s�C ', [aMsg] ) );
        Sleep( 300 );
        aObj.ActCost := SecondsBetween( Now, aStartTime );
        { �� 0 ���令 1 �� }
        if ( aObj.ActCost <= 0 ) then aObj.ActCost := 1; 
      finally
        aParser.Free;
      end;
    end;  
    if ( Result ) then
    begin
      aObj.ActStatus := 'C';
      aObj.ActErrMsg := EmptyStr;
      case aObj.ActFileType of
        xfWrapper: FParam.WrapperLastProcFile := aObj.ActFileName;
        xfPpv, xfSub: FParam.DexLastProcFile := aObj.ActFileName;
        xfAsRun: FParam.AsRunLastProcFile := aObj.ActFileName;
      end;
    end else
    begin
      aObj.ActStatus := 'E';
      aObj.ActErrMsg := aMsg;
    end;
    FParseSubject.ActionObject := aObj;
    Sleep( 300 );
    if ( Self.Terminated ) then Break;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TParseThread.ChangeDetected: Boolean;
var
  aMerge: TMergeChange;
  aCount: Integer;
  aMsg: String;
begin
  aMerge := TMergeChange.Create;
  try
    aMerge.DataConnection := ParseController.DataConnection;
    aMerge.DataReader := ParseController.DataReader;
    aMerge.DataWriter := ParseController.DataWriter;
    aMerge.DataDelete := ParseController.DataDelete;
    aMerge.DataInsert := ParseController.DataInsert;
    aMerge.DataUpdate := ParseController.DataUpdate;
    Result := aMerge.DetectChangeCount( aMsg, aCount );
    if Result then
    begin
      FMsgSubject.Normal( Format( '�i�ѪR�֤ߡj�o�{ %d �����ʸ`�ت��ơC',
        [aCount]  ) );
      Sleep( 300 );
    end;
  finally
    aMerge.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TParseThread.MergeChange: Boolean;
var
  aMerge: TMergeChange;
  aMsg: String;
  aObj: PActObject;
  aStartTime: TDateTime;
begin
  FMsgSubject.Normal( '�i�ѪR�֤ߡj�}�l���ʸ�ƦX�֡C' );
  Sleep( 300 );
  aMerge := TMergeChange.Create;
  try
    New( aObj );
    aObj.KeyId := EmptyStr;
    aObj.ActDate := FormatDateTime( 'yyyy/mm/dd', Date );
    aObj.ActTime := FormatDateTime( 'hh:nn:ss', Now );
    aObj.ActFileName := '�X�ָ`�ت��ʸ��';
    aObj.ActFileSize := EmptyStr;
    aObj.ActFilePath := EmptyStr;
    aObj.ActSource := 'Merge';
    aObj.ActFileType := xfMerge;
    aObj.ActStatus := 'P';
    aObj.ActCost := 0;
    aObj.ActProgress := 0;
    aObj.ActErrMsg := EmptyStr;
    FList.Add( aObj );
    FParseSubject.ActionObject := aObj;
    Sleep( 300 );
    FParseSubject.ActionObject := aObj;
    Sleep( 300 );
    aMerge.DataConnection := ParseController.DataConnection;
    aMerge.DataReader := ParseController.DataReader;
    aMerge.DataWriter := ParseController.DataWriter;
    aMerge.DataDelete := ParseController.DataDelete;
    aMerge.DataInsert := ParseController.DataInsert;
    aMerge.DataUpdate := ParseController.DataUpdate;
    aStartTime := Now;
    Result := aMerge.Merge( aMsg );
    if Result then
      aObj.ActStatus := 'C'
    else begin
      aObj.ActStatus := 'E';
      aObj.ActErrMsg := aMsg;
    end;  
    if Result then
      FMsgSubject.OK( '�i�ѪR�֤ߡj���ʸ�ƦX�֧����C' )
    else
      FMsgSubject.Error( Format( '�i�ѪR�֤ߡj���ʸ�ƦX��, ��]: %s�C ', [aMsg] ) );
    Sleep( 300 ); 
    aObj.ActCost := SecondsBetween( Now, aStartTime );  
    FParseSubject.ActionObject := aObj;
    Sleep( 300 );
  finally
    aMerge.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

end.
