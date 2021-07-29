unit cbDataThread;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, IniFiles, cbClass;

type
  TDataThread = class(TThread)
  private
    { Private declarations }
    FNotify: TNotify;
    FComps: String;
    FSRA: TFreq;
    FCA: TFreq;
    procedure Notify;
    procedure SetCA(const Value: TFreq);
    procedure SetSRA(const Value: TFreq);
    procedure SaveConfig(const aKind: Integer);
    function CheckCanCA: Boolean;
    function CheckCanSRA: Boolean;
    function LocalFileDetect: Integer;
    procedure DoProcessDownloadFile;
  protected
    procedure Execute; override;
  public
    constructor Create(aComps: String);
    destructor Destotry;
    property CA: TFreq read FCA write SetCA;
    property SRA: TFreq read FSRA write SetSRA;
    property Terminated;
    procedure Synchronize(Method: TThreadMethod);
  end;


implementation

uses cbMain, cbDataControler, DateUtils, Encryption_TLB;

{ Important: Methods and properties of objects in visual components can only be
  used in a method called using Synchronize, for example,

      Synchronize(UpdateCaption);

  and UpdateCaption could look like,

    procedure TDataThread.UpdateCaption;
    begin
      Form1.Caption := 'Updated in a thread';
    end; }

{ TDataThread }

{ ---------------------------------------------------------------------------- }

constructor TDataThread.Create(aComps: String);
begin
  inherited Create( True );
  FComps := aComps;
  Self.FreeOnTerminate := False;
end;

{ ---------------------------------------------------------------------------- }

destructor TDataThread.Destotry;
begin
  inherited;
end;

{ ---------------------------------------------------------------------------- }

procedure TDataThread.Execute;
var
  aCount: Integer;
begin
  FNotify.MsgType := MB_ICONINFORMATION;
  FNotify.MsgText := '��Ʈw������}�l����C';
  Synchronize( Notify );
  DataControler := TDataControler.Create( FComps );
  Sleep( 3000 );
  while not Self.Terminated do
  begin
    if CheckCanCA then
    begin
      try
        FNotify.MsgType := MB_ICONINFORMATION;
        FNotify.MsgText := '�}�l�B�z�iCM�˾����u��ơj�C';
        Synchronize( Notify );
        {}
        DataControler.DoDataExport( ekCA );
        {}
        FNotify.MsgType := MB_ICONINFORMATION;
        FNotify.MsgText := '�iCM�˾����u��ơj��ƳB�z�����C';
        Synchronize( Notify );
      except
        on E: Exception do
        begin
          FNotify.MsgType := MB_ICONERROR;
          FNotify.MsgText := Format( '�B�z�iCM�˾����u��ơj����, ��]:%s�C', [E.Message] );
          Synchronize( Notify );
        end;
      end;
      FCA.LastExc := Now;
      FCA.NextExc := ( FCA.LastExc + ( ( 1 / 24 / 60 ) * FCA.CFR ) );
      {}
      SaveConfig( 2 );
      {}
      FNotify.MsgType := MB_ICONINFORMATION;
      FNotify.MsgText := Format( '�U���B�z�iCM�˾����u��ơj�ɶ���: %s�C',
        [FormatDateTime( 'yyyy-mm-dd hh:nn', FCA.NextExc )] );
      Synchronize( Notify );
    end;
    if Self.Terminated then Break;
    Sleep( 1000 );
    if CheckCanSRA then
    begin
      try
        FNotify.MsgType := MB_ICONINFORMATION;
        FNotify.MsgText := '�}�l�B�z�iCM���v��ơj�C';
        Synchronize( Notify );
        {}
        DataControler.DoDataExport( ekSA );
        {}
        FNotify.MsgType := MB_ICONINFORMATION;
        FNotify.MsgText := '�iCM���v��ơj��ƳB�z�����C';
        Synchronize( Notify );
      except
        on E: Exception do
        begin
          FNotify.MsgType := MB_ICONERROR;
          FNotify.MsgText := Format( '�B�z�iCM���v��ơj����, ��]:%s�C', [E.Message] );
          Synchronize( Notify );
        end;
      end;
      if Self.Terminated then Break;
      Sleep( 1000 );
      try
        FNotify.MsgType := MB_ICONINFORMATION;
        FNotify.MsgText := '�}�l�B�z�iCM���v��ơj�C';
        Synchronize( Notify );
        {}
        DataControler.DoDataExport( ekRA );
        {}
        FNotify.MsgType := MB_ICONINFORMATION;
        FNotify.MsgText := '�iCM���v��ơj��ƳB�z�����C';
        Synchronize( Notify );
      except
        on E: Exception do
        begin
          FNotify.MsgType := MB_ICONERROR;
          FNotify.MsgText := Format( '�B�z�iCM���v��ơj����, ��]:%s�C', [E.Message] );
          Synchronize( Notify );
        end;
      end;
      FSRA.LastExc := Now;
      FSRA.NextExc := ( FSRA.LastExc + ( ( 1 / 24 / 60 ) * FSRA.CFR ) );
      {}
      SaveConfig( 3 );
      {}
      FNotify.MsgType := MB_ICONINFORMATION;
      FNotify.MsgText := Format( '�U���B�z�iCM��/���v��ơj�ɶ���: %s�C',
        [FormatDateTime( 'yyyy-mm-dd hh:nn', FSRA.NextExc )] );
      Synchronize( Notify );
    end;
    Sleep( 3000 );
    aCount := LocalFileDetect;
    if ( aCount > 0 ) then
    begin
      FNotify.MsgType := MB_ICONINFORMATION;
      FNotify.MsgText := Format( '���%d�ӤU���^���ɡC', [aCount] );
      Synchronize( Notify );
      try
        DoProcessDownloadFile;
      except
        on E: Exception do
        begin
          FNotify.MsgType := MB_ICONINFORMATION;
          FNotify.MsgText := Format( '�B�z�U���^���ɮɵo�Ϳ��~, ��]:%s�C', [E.Message] );
          Synchronize( Notify )          ;
        end;
      end;
    end;
  end;
  SaveConfig( 2 );
  SaveConfig( 3 );
  {}
  FNotify.MsgType := MB_ICONINFORMATION;
  FNotify.MsgText := '��Ʈw�����浲���C';
  Synchronize( Notify );
end;

{ ---------------------------------------------------------------------------- }

procedure TDataThread.Notify;
begin
  Main.AddMsg(  FNotify );
end;

{ ---------------------------------------------------------------------------- }

procedure TDataThread.SetCA(const Value: TFreq);
begin
  FCA.CFR := Value.CFR;
  FCA.LastExc := Value.LastExc;
end;

{ ---------------------------------------------------------------------------- }

procedure TDataThread.SetSRA(const Value: TFreq);
begin
  FSRA.CFR := Value.CFR;
  FSRA.LastExc := Value.LastExc;
end;

{ ---------------------------------------------------------------------------- }

procedure TDataThread.SaveConfig(const aKind: Integer);
var
  aIndex: Integer;
  aIni: TIniFile;
  aTmp: TStringList;
  aObj: _Password;
  aKey: WideString;
begin
  TDataControler.CreateTemporyIniFile(
    IncludeTrailingPathDelimiter( ExtractFilePath(
      ParamStr( 0 ) ) ) + 'CONFIG.INI',
    IncludeTrailingPathDelimiter( ExtractFilePath(
      ParamStr( 0 ) ) )+ 'TMPCONFIG.INI' );
  aIni := TIniFile.Create( IncludeTrailingPathDelimiter( ExtractFilePath(
      ParamStr( 0 ) ) )+ 'TMPCONFIG.INI' );
  try
    {}
    if ( aKind in [0,2]) then
    begin
      aIni.WriteInteger( 'CAFRQUENCE', 'FR', FCA.CFR  );
      aIni.WriteString( 'CAFRQUENCE', 'LASTEXC', FormatDateTime( 'yyyy/mm/dd hh:nn:ss', FCA.LastExc ) );
    end;
    {}
    if ( aKind in [0,3] ) then
    begin
      aIni.WriteInteger( 'SRAFRQUENCE', 'FR', FSRA.CFR  );
      aIni.WriteString( 'SRAFRQUENCE', 'LASTEXC', FormatDateTime( 'yyyy/mm/dd hh:nn:ss', FSRA.LastExc ) );
    end;
    {}
    aIni.UpdateFile;
  finally
    aIni.Free;
  end;
  aTmp := TStringList.Create;
  try
    aTmp.LoadFromFile( IncludeTrailingPathDelimiter( ExtractFilePath(
      ParamStr( 0 ) ) )+ 'TMPCONFIG.INI' );
    aObj := CoPassword.Create;
    try
      aKey := 'CS';
      for aIndex := 0 to aTmp.Count - 1 do
        aTmp[aIndex] := aObj.Encrypt( aTmp[aIndex], aKey );
      aTmp.SaveToFile( IncludeTrailingPathDelimiter( ExtractFilePath(
        ParamStr( 0 ) ) )+ 'CONFIG.INI' );
    finally
      aObj := nil;
    end;
  finally
    aTmp.Free;
  end;
  if FileExists( IncludeTrailingPathDelimiter( ExtractFilePath(
      ParamStr( 0 ) ) )+ 'TMPCONFIG.INI' ) then
    DeleteFile( IncludeTrailingPathDelimiter( ExtractFilePath(
      ParamStr( 0 ) ) )+ 'TMPCONFIG.INI' );
end;

{ ---------------------------------------------------------------------------- }

function TDataThread.CheckCanCA: Boolean;
begin
  Result := ( Now >= ( FCA.NextExc ) );
end;

{ ---------------------------------------------------------------------------- }

function TDataThread.CheckCanSRA: Boolean;
begin
  Result := ( Now >= ( FSRA.NextExc ) );
end;

{ ---------------------------------------------------------------------------- }

function TDataThread.LocalFileDetect: Integer;
var
  aSearch: TSearchRec;
  aSearchMask, aPrefix: String;
begin
  Result := 0;
  aSearchMask := IncludeTrailingPathDelimiter( LocalDownloadDir ) + '*.txt';
  if FindFirst( aSearchMask, faAnyFile, aSearch ) = 0 then
  begin
    try
      repeat
         if ( aSearch.Attr and faDirectory ) <> 0 then
         begin
           if ( aSearch.Name = '.' ) or ( aSearch.Name = '..' ) then Continue;
         end else
         if ( aSearch.Attr and faArchive ) <> 0 then
         begin
           aPrefix := UpperCase( Copy( aSearch.Name, 1, 3 ) );
           if ( aPrefix = 'RP-' ) then
             Inc( Result );
         end;
      until FindNext( aSearch ) <> 0;
    finally
      FindClose( aSearch );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TDataThread.DoProcessDownloadFile;
var
  aSearch: TSearchRec;
  aSearchMask, aSource, aDest: String;
  aProcessResult: Boolean;
begin
  aSearchMask := IncludeTrailingPathDelimiter( LocalDownloadDir ) + '*.txt';
  if FindFirst( aSearchMask, faAnyFile, aSearch ) = 0 then
  begin
    try
      repeat
         if ( aSearch.Attr and faDirectory ) <> 0 then
         begin
           if ( aSearch.Name = '.' ) or ( aSearch.Name = '..' ) then Continue;
         end else
         if ( aSearch.Attr and faArchive ) <> 0 then
         begin
           FNotify.MsgType := MB_ICONINFORMATION;
           FNotify.MsgText := Format( '�^���ɮ�:%s�B�z���C', [aSearch.Name] );
           Synchronize( Notify );
           Sleep( 300 );
           aProcessResult := DataControler.DoDataImport(
             IncludeTrailingPathDelimiter( LocalDownloadDir ) + aSearch.Name );
           if aProcessResult then
           begin
             aSource := IncludeTrailingPathDelimiter( LocalDownloadDir ) + aSearch.Name;
             aDest := IncludeTrailingPathDelimiter( LocalBackupDir ) + aSearch.Name;
             MoveFile( PChar( aSource ), PChar( aDest ) );
             FNotify.MsgType := MB_ICONINFORMATION;
             FNotify.MsgText := Format( '�^���ɮ�:%s�w�B�z�����C', [aSearch.Name] );
             Synchronize( Notify );
           end else
           begin
             FNotify.MsgType := MB_ICONERROR;
             FNotify.MsgText := '�B�z�^���ɮɵo�Ϳ��~, ���^���ɱN�|����[Error]��Ƨ��U�C';
             Synchronize( Notify );
             aSource := IncludeTrailingPathDelimiter( LocalDownloadDir ) + aSearch.Name;
             aDest := IncludeTrailingPathDelimiter( LocalErrorDir ) + aSearch.Name;
             MoveFile( PChar( aSource ), PChar( aDest ) );
           end;
           Sleep( 300 );
         end;
      until FindNext( aSearch ) <> 0;
    finally
      FindClose( aSearch );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TDataThread.Synchronize(Method: TThreadMethod);
begin
  inherited Synchronize( Method );
end;

{ ---------------------------------------------------------------------------- }

end.
