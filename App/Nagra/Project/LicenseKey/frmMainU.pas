unit frmMainU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, ExtCtrls;

type
  TfrmMain = class(TForm)
    StaticText2: TStaticText;
    edtDate: TEdit;
    StaticText3: TStaticText;
    RadioGroup1: TRadioGroup;
    RadioGroup2: TRadioGroup;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    StaticText4: TStaticText;
    edtLicenseKey: TEdit;
    Bevel1: TBevel;
    Image1: TImage;
    procedure BitBtn2Click(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;

implementation

uses ConstObjU;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

function GetComputerName: String;
var
  ALen: Cardinal;
begin
  SetLength( Result, MAX_COMPUTERNAME_LENGTH );
  ALen := Length( Result );
  Windows.GetComputerName( @Result[1], ALen );
end;

{ ---------------------------------------------------------------------------- }

function Encrypt(const ASource, AKey: String): String;
var
  AIndex, AKeyPos, AKeyLength : Integer;
begin
  Result := '';
  AKeyPos := 1;  
  AKeyLength := Length( AKey );
  for AIndex := 1 to Length( ASource ) do
  begin
    Result := Result + Chr( Ord( ASource[AIndex] ) + Ord( AKey[AKeyPos] ) );
    Inc( AKeyPos );
    if AKeyPos = ( AKeyLength ) then AKeyPos := 1;
  end;
end;

{ ---------------------------------------------------------------------------- }

function Decrypt(const ASource, AKey: String): String;
var
  AIndex, AKeyPos, AKeyLength : Integer;
begin
  Result := '';
  AKeyPos := 1;
  AKeyLength := Length( AKey );
  for AIndex := 1 to Length( ASource ) do
  begin
    Result := Result + Chr( Ord( ASource[AIndex] ) - Ord( AKey[AKeyPos] ) );
    Inc( AKeyPos );
    if AKeyPos = AKeyLength then AKeyPos:=1;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.BitBtn2Click(Sender: TObject);
begin
  Self.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.BitBtn1Click(Sender: TObject);
var
  ARadio1, ARadio2, ASourceStr : String;
begin
  case RadioGroup1.ItemIndex of
    0: ARadio1 := SYS4_Y;
    1: ARadio1 := SYS4_N;
  end;
  case RadioGroup2.ItemIndex of
    0: ARadio2 := SYS5_Y;
    1: ARadio2 := SYS5_N;
  end;

  { ¥]§t Pc Name ¨Ó¥[±K
  ASourceStr := edtDate.Text + ENCRYPTION_SEP + ARadio1 + ENCRYPTION_SEP +
    ARadio2 + ENCRYPTION_SEP + GetComputerName;
  }
  
  ASourceStr := edtDate.Text + ENCRYPTION_SEP + ARadio1 + ENCRYPTION_SEP +
    ARadio2;
  edtLicenseKey.Text := Encrypt( ASourceStr, ENCRYPTION_KEY );
end;

{ ---------------------------------------------------------------------------- }

end.

