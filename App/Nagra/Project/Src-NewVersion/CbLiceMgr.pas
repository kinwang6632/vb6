unit cbLiceMgr;

interface

uses
  Windows, SysUtils, Classes, Forms, Registry, LbCipher, LbClass,
  cbResStr, cbUtils;


type
  PExpireInfo = ^TExpireInfo;

  TExpireInfo = record
    ExpireDay: TDateTime;
    InstallDay: TDateTime;
    PiroUseDay: TDateTime;
    ImmediateExpire: string;
    case Integer of
      0: ( Kdx1,  Kdx2,  Kdx3,  Kdx4: Smallint );
  end;


type
  TLicenceManager = class(TDataModule)
    LbDES: TLbDES;
    procedure DataModuleCreate(Sender: TObject);
    procedure DataModuleDestroy(Sender: TObject);
  private
    FReg: TRegistry;
    FExpire: PExpireInfo;
    FInput: Boolean;
    FKeyName: string;
    function GetRootKey: HKEY;
    procedure SetRootKey(const Value: HKEY);
    procedure SetKeyName(const Value: string);
    procedure WriteToKey(const AKeyName: string);
    procedure ReadFromKey(const AKeyName: string);
    procedure InitExpireInfo;
    function ActiveUserInputForm(const AMsg: string): string;
    procedure LicenceToExpireInto(const ALic: string);
  public
    property RootKey: HKEY read GetRootKey write SetRootKey;
    property KeyName: string read FKeyName write SetKeyName;
    property CanInputLicence: Boolean read FInput write FInput default True;
    function IsLicenceVaildate: Boolean;
    procedure SaveLicence;
  end;

implementation

{$R *.dfm}

uses CbLicInp;

const
  KdxNames: array [0..3] of string = ( 'kdx1', 'kdx2', 'kdx3', 'kdx4' );

{ ---------------------------------------------------------------------------- }

procedure TLicenceManager.DataModuleCreate(Sender: TObject);
begin
  FReg := TRegistry.Create;
  FReg.RootKey := HKEY_LOCAL_MACHINE;
  FExpire := New( PExpireInfo );
  InitExpireInfo;
  FKeyName := ExtractFileName( Application.ExeName );
end;

{ ---------------------------------------------------------------------------- }

procedure TLicenceManager.DataModuleDestroy(Sender: TObject);
begin
  Dispose( FExpire );
  FReg.CloseKey;
  FReg.Free;
end;

{ ---------------------------------------------------------------------------- }

function TLicenceManager.GetRootKey: HKEY;
begin
  Result := FReg.RootKey;
end;

{ ---------------------------------------------------------------------------- }

procedure TLicenceManager.SetRootKey(const Value: HKEY);
begin
  if Value <> FReg.RootKey then
  begin
    FReg.CloseKey;
    FReg.RootKey := Value;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TLicenceManager.SetKeyName(const Value: string);
begin
  if Value <> FKeyName then
  begin
    FKeyName := Value;
    if FKeyName <> '' then
    begin
      FReg.OpenKey( FKeyName, True );
      WriteToKey( FKeyName );
    end else InitExpireInfo;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TLicenceManager.WriteToKey(const AKeyName: string);
begin
  if not FReg.ValueExists( KdxNames[FExpire.Kdx1] ) then
    FReg.WriteDateTime( KdxNames[FExpire.Kdx1],
    StrToDate( LbDES.EncryptString( DateToStr( FExpire.ExpireDay ) ) ) );
  if not FReg.ValueExists( KdxNames[FExpire.Kdx2] ) then
    FReg.WriteDateTime( KdxNames[FExpire.Kdx1],
    StrToDate( LbDES.EncryptString( DateToStr( FExpire.InstallDay ) ) ) );
  if FReg.ValueExists( KdxNames[FExpire.Kdx3] ) then
    FReg.WriteDateTime( KdxNames[FExpire.Kdx1],
    StrToDate( LbDES.EncryptString( DateToStr( FExpire.PiroUseDay ) ) ) );
  if FReg.ValueExists( KdxNames[FExpire.Kdx4] ) then
    FReg.WriteString( KdxNames[FExpire.Kdx1],
    LbDES.EncryptString( FExpire.ImmediateExpire ) );
end;

{ ---------------------------------------------------------------------------- }

procedure TLicenceManager.ReadFromKey(const AKeyName: string);
begin
  try
    FExpire.ExpireDay :=
      StrToDate( LbDES.DecryptString( FReg.ReadString(
      KdxNames[FExpire.Kdx1] ) ) );
    FExpire.InstallDay :=
      StrToDate( LbDES.DecryptString( FReg.ReadString(
      KdxNames[FExpire.Kdx2] ) ) );
    FExpire.PiroUseDay :=
      StrToDate( LbDES.DecryptString( FReg.ReadString(
      KdxNames[FExpire.Kdx3] ) ) );
    FExpire.ImmediateExpire := LbDES.DecryptString( FReg.ReadString(
      KdxNames[FExpire.Kdx4] ) );
  except
    { ... }
  end;    
end;

{ ---------------------------------------------------------------------------- }

procedure TLicenceManager.InitExpireInfo;
begin
  FExpire.ExpireDay := 0;
  FExpire.InstallDay := Date;
  FExpire.PiroUseDay := Date;
  FExpire.ImmediateExpire := 'Y';
  FExpire.Kdx1 := Ord( FExpire.Kdx1 );
  FExpire.Kdx2 := Ord( FExpire.Kdx2 );
  FExpire.Kdx3 := Ord( FExpire.Kdx3 );
  FExpire.Kdx4 := Ord( FExpire.Kdx4 );
end;

{ ---------------------------------------------------------------------------- }

function TLicenceManager.ActiveUserInputForm(const AMsg: string): string;
begin
  frmLicInp := TfrmLicInp.Create( nil );
  try
     frmLicInp.Msg.Caption := AMsg;
     frmLicInp.ShowModal;
     Result := frmLicInp.InputLic.Text;
  finally
    frmLicInp.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TLicenceManager.LicenceToExpireInto(const ALic: string);
const
  SEPARATOR = #8;
var
  ADecry, AValue: String;
  ACount: Integer;
begin
  ADecry := LbDES.DecryptString( ALic );
  ACount := 0;
  while ( Pos( SEPARATOR, ADecry ) > 0 ) or ( Length( ADecry ) > 0 ) do
  begin
    Inc( ACount );
    if Pos( SEPARATOR, ADecry ) > 0 then
    begin
      AValue := Copy( ADecry, 1, Pos( SEPARATOR, ADecry ) - 1 ) ;
      Delete( ADecry, 1, Pos( SEPARATOR, ADecry ) );
    end
    else AValue := ADecry;
    case ACount of
      1: FExpire.ExpireDay := StrToDate( AValue );
      2: FExpire.InstallDay := StrToDate( AValue );
      3: FExpire.PiroUseDay := StrToDate( AValue );
      4: FExpire.ImmediateExpire := AValue;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TLicenceManager.IsLicenceVaildate: Boolean;
label
  ReCheck;
var
  AMsg, AInput: String;

  { ------------------------------------------------ }

  function CheckLicenceStoreFormat: Boolean;
  begin
    Result := True;
    try
      ReadFromKey( FKeyName );
    except
      begin
        AMsg := SLicenceInvalid;
        Result := False;
      end;  
    end;
  end;

  { ------------------------------------------------ }

  function CheckLicence: Boolean;
  begin
    if ( FExpire.InstallDay = Date ) then
      AMsg := Format( SLicenceNotRegister, [SVendor] )
    else if ( FExpire.ImmediateExpire = 'Y' ) then
      AMsg := Format( SLicenceImmediateExpire, [SVendor] )
    else if ( FExpire.ExpireDay < Date ) then
      AMsg := Format( SLicenceExpire, [SVendor] )
    else if ( FExpire.PiroUseDay > Date ) then
      AMsg := Format( SLicenceExpire, [SVendor] );
    Result := ( AMsg = '' );  
  end;

  { ------------------------------------------------ }

begin
ReCheck:
  AMsg := '';
  Result := CheckLicenceStoreFormat;
  if not Result then
  begin
    if FInput then
    begin
      AInput := ActiveUserInputForm( AMsg );
      LicenceToExpireInto( AInput );
      SaveLicence;
      goto ReCheck;
    end
    else WarningMsg( AMsg );
    Exit;
  end;
  Result := CheckLicence;
  if not Result then
  begin
    if FInput then
    begin
      AInput := ActiveUserInputForm( AMsg );
      LicenceToExpireInto( AInput );
      SaveLicence;
      goto ReCheck;
    end
    else WarningMsg( AMsg );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TLicenceManager.SaveLicence;
begin
  WriteToKey( FKeyName );
end;

{ ---------------------------------------------------------------------------- }

end.
