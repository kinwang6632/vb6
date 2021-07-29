unit cbLiceMgr;

interface

uses
  Windows, SysUtils, Classes, Forms, Controls, LbCipher, LbClass,
  cbResStr, cbUtilis, DateUtils;


type
  TLicenceInfo = record
    LicenceKey: String;
    RegisterKey: String;
    InstallDay: String;
    ExpireDay: String;
    PiroUseDay: String;
  end;

type
  TLicenceManager = class(TDataModule)
    LbDES: TLbDES;
    procedure DataModuleCreate(Sender: TObject);
    procedure DataModuleDestroy(Sender: TObject);
  private
    FCanInputLic: Boolean;
    FLicFileName: String;
    FMACList: TStringList;
    FMACList2: TStringList;
    FLicInfo: TLicenceInfo;
    procedure ChangeKey(const aMode: Integer);
    function EncryptMAC: String;
    procedure DecryptRegisterKey;
  public
    function LoadLicenceFile: Boolean;
    function ValidateLicence: Boolean;
    function SaveLicenceFile: Boolean;
    function IsMACAvailable: Boolean;
    function ActiveUserInputForm: Boolean;
    procedure Update;
    property CanInputLicence: Boolean read FCanInputLic write FCanInputLic;
    property LicenceInfo: TLicenceInfo read FLicInfo;
  end;
  

implementation

{$R *.dfm}

uses CbLicInp, cbIphlpapi;

{ ---------------------------------------------------------------------------- }

procedure TLicenceManager.DataModuleCreate(Sender: TObject);
begin
  FCanInputLic := True;
  FLicFileName := IncludeTrailingPathDelimiter( ExtractFilePath(
    Application.ExeName ) ) + 'Key.lic';
  ZeroMemory( @FLicInfo, SizeOf( FLicInfo ) );
  FMACList := TStringList.Create;
  FMACList2 := TStringList.Create;
  GetMACAddress( FMACList );
  ChangeKey( 1 );
  FLicInfo.LicenceKey := EncryptMAC;
end;

{ ---------------------------------------------------------------------------- }

procedure TLicenceManager.DataModuleDestroy(Sender: TObject);
begin
  FMACList.Free;
  FMACList2.Free;
end;

{ ---------------------------------------------------------------------------- }

procedure TLicenceManager.ChangeKey(const aMode: Integer);
var
  aIndex: Integer;
  aPhrase: String;
begin
  if ( aMode = 1 ) then
  begin
    for aIndex := Ord( 'a' ) to Ord( 'z' ) do
      aPhrase := ( aPhrase + Chr( aIndex ) );
  end else
  begin
    for aIndex := Ord( 'z' ) downto Ord( 'a' ) do
      aPhrase := ( aPhrase + Chr( aIndex ) );
  end;
  LbDES.GenerateKey( aPhrase );
end;

{ ---------------------------------------------------------------------------- }

function TLicenceManager.EncryptMAC: String;
var
  aIndex: Integer;
  aText: String;
begin
  Result := EmptyStr;
  for aIndex := 0 to FMACList.Count - 1 do
    aText := ( aText + FMACList[aIndex] + ',' );
  Result := LbDES.EncryptString( aText );
end;

{ ---------------------------------------------------------------------------- }

procedure TLicenceManager.DecryptRegisterKey;
var
  aPos, aRef: Integer;
  aText, aValue: String;
begin
  ChangeKey( 2 );
  try
    aText := LbDES.DecryptString( FLicInfo.RegisterKey );
  except
    Exit;
  end;  
  aRef := 1;
  while ( aText <> EmptyStr ) do
  begin
    aPos := AnsiPos( ',', aText );
    if ( aPos > 0 ) then
    begin
      aValue := Copy( aText, 1, aPos - 1 );
      Delete( aText, 1, aPos );
    end else
    begin
      aValue := aText;
      aText := EmptyStr;
    end;  
    if AnsiPos( '/', aValue ) <= 0 then
      FMACList2.Add( aValue )
    else begin
      case aRef of
        1: FLicInfo.InstallDay := aValue;
        2: FLicInfo.ExpireDay := aValue;
        3: FLicInfo.PiroUseDay := aValue;
      end;
      Inc( aRef );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TLicenceManager.LoadLicenceFile: Boolean;
var
  aIndex: Integer;
  aFile: TStringList;
  aText: String;
begin
  Result := False;
  if not FileExists( FLicFileName ) then Exit;
  aFile := TStringList.Create;
  try
    aFile.LoadFromFile( FLicFileName );
    for aIndex := 0 to aFile.Count - 1 do
      aText := ( aText + aFile[aIndex] );
    FLicInfo.RegisterKey := aText;
    DecryptRegisterKey;
  finally
    aFile.Free;
  end;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

function TLicenceManager.ActiveUserInputForm: Boolean;
begin
  Result := False;
  fmLicInp := TfmLicInp.Create( nil );
  try
     fmLicInp.Caption := SInputFormTitle;
     fmLicInp.lbLicenceNote.Caption :=  SLicenceNote;
     fmLicInp.lbLicenceKey.Caption := SLicenceKey;
     fmLicInp.lbRegisterKey.Caption := SRegisterKey;
     fmLicInp.LicenceKey.Text := FLicInfo.LicenceKey;
     fmLicInp.RegsiterKey.Text := FLicInfo.RegisterKey;
     if ( fmLicInp.ShowModal = mrOk ) then
     begin
       FLicInfo.RegisterKey := fmLicInp.RegsiterKey.Text;
       SaveLicenceFile;
       LoadLicenceFile;
       Result := True;
     end;
  finally
    fmLicInp.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TLicenceManager.ValidateLicence: Boolean;

  { -------------------------------------------- }

  function CompareMAC: Boolean;
  var
    aIndex, aPos: Integer;
  begin
    Result := False;
    for aIndex := 0 to FMACList2.Count - 1 do
    begin
      aPos := FMACList.IndexOf( FMACList2[aIndex] );
      Result := ( aPos >= 0 );
      if Result then Break;
    end;
  end;

  { -------------------------------------------- }

var
  aDateTime: TDateTime;
begin
  Result := False;
  if not FileExists( FLicFileName ) then Exit;
  Result := ( TryStrToDate( FLicInfo.InstallDay, aDateTime ) and
    TryStrToDate( FLicInfo.ExpireDay, aDateTime ) and
    TryStrToDate( FLicInfo.PiroUseDay, aDateTime ) );
  if not Result then Exit;
  Result := CompareMAC;
  if not Result then Exit;
  Result := ( StrToDate( FLicInfo.ExpireDay ) > Date );
  if not Result then Exit;
  Result := ( StrToDate( FLicInfo.InstallDay ) <= Date );
  if not Result then Exit;
  Result :=
    ( StrToDate( FLicInfo.PiroUseDay ) <= Date ) and
    ( StrToDate( FLicInfo.PiroUseDay ) <= StrToDate( FLicInfo.ExpireDay  ) ) and
    ( StrToDate( FLicInfo.PiroUseDay ) >= StrToDate( FLicInfo.InstallDay ) );
end;

{ ---------------------------------------------------------------------------- }

function TLicenceManager.SaveLicenceFile: Boolean;
var
  aFile: TStringList;
begin
  aFile := TStringList.Create;
  try
    aFile.Add( FLicInfo.RegisterKey );
    aFile.SaveToFile( FLicFileName );
  finally
    aFile.Free;
  end;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

function TLicenceManager.IsMACAvailable: Boolean;
begin
  Result := ( FMACList.Count > 0 );
end;

{ ---------------------------------------------------------------------------- }

procedure TLicenceManager.Update;
var
  aIndex: Integer;
  aText: String;
begin
  ChangeKey( 2 );
  if ( StrToDate( FLicInfo.PiroUseDay ) > Date ) then Exit;
  FLicInfo.PiroUseDay := FormatDateTime( 'yyyy/mm/dd', Date );
  aText := EmptyStr;
  for aIndex := 0 to FMACList2.Count - 1 do
    aText := ( aText + FMACList2[aIndex] + ',' );
  aText := LbDES.EncryptString( aText + FLicInfo.InstallDay + ',' +
    FLicInfo.ExpireDay + ',' + FLicInfo.PiroUseDay );
  FLicInfo.RegisterKey := aText;
  SaveLicenceFile;
end;

{ ---------------------------------------------------------------------------- }

end.
