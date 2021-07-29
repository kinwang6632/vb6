unit cbKeyGenMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Menus,
{ TurboPower LockBox }
  LbCipher, LbClass,
{Developer Express }
  cxLookAndFeelPainters, dxSkinsCore,
  dxSkinsDefaultPainters, cxControls, cxContainer, cxEdit, cxTextEdit,
  dxGDIPlusClasses, cxButtons, cxMaskEdit;

type
  TForm1 = class(TForm)
    btnOK: TcxButton;
    Bevel1: TBevel;
    Image1: TImage;
    lbLicenceKey: TLabel;
    LicenceKey: TcxTextEdit;
    lbRegisterKey: TLabel;
    RegsiterKey: TcxTextEdit;
    lbLicenceNote: TLabel;
    lbMacs: TLabel;
    MacList: TcxTextEdit;
    lbInstDate: TLabel;
    InstDate: TcxMaskEdit;
    lbExpireDate: TLabel;
    ExpireDate: TcxMaskEdit;
    procedure FormCreate(Sender: TObject);
    procedure LicenceKeyPropertiesChange(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnOKClick(Sender: TObject);
  private
    { Private declarations }
    FLbDES: TLbDES;
    procedure ChangeKey(const aMode: Integer);
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

uses cbResStr, cbUtilis;

{$R *.dfm}

function ExtractValue(var aValue: String; aSeparator: String = ','): String;
var
  aPos: Integer;
begin
  aPos := AnsiPos( aSeparator, aValue );
  if aPos = 0 then
  begin
    Result := aValue;
    aValue := EmptyStr;
  end else
  begin
    Result := Copy( aValue, 1, aPos - 1 );
    Delete( aValue, 1, aPos );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.FormCreate(Sender: TObject);
begin
  FLbDES := TLbDES.Create( nil );
  LicenceKey.Text := EmptyStr;
  RegsiterKey.Text := EmptyStr;
  {}
  lbLicenceNote.Caption := SLicenceNote2;
  lbLicenceKey.Caption := SLicenceKey;
  lbRegisterKey.Caption := SRegisterKey;
  lbMacs.Caption := SMacList;
  lbInstDate.Caption := SInstallDate;
  lbExpireDate.Caption := SExpireDate;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.FormDestroy(Sender: TObject);
begin
  FLbDES.Free;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.LicenceKeyPropertiesChange(Sender: TObject);
var
  aText: String;
begin
  if ( LicenceKey.Text = EmptyStr ) then
  begin
    RegsiterKey.Text := EmptyStr;
    Exit;
  end;
  ChangeKey( 1 );
  try
    aText := FLbDES.DecryptString( LicenceKey.Text );
    if IsDelimiter( ',', aText, Length( aText ) ) then
      Delete( aText, Length( aText ), 1 );
    MacList.Text := aText;
  except
    RegsiterKey.Text := EmptyStr;
    MacList.Text := EmptyStr;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.btnOKClick(Sender: TObject);
var
  aMacs, aKey: String;
begin
  RegsiterKey.Text := EmptyStr;
  if ( InstDate.Text = EmptyStr ) or ( MacList.Text = EmptyStr ) then Exit;
  ChangeKey( 2 );
  aKey := EmptyStr;
  aMacs := MacList.Text;
  if not IsDelimiter( ',', aMacs, Length( aMacs ) ) then aMacs := ( aMacs + ',' );
  aKey := FLbDES.EncryptString( aMacs + InstDate.Text + ',' +
    ExpireDate.Text + ',' + InstDate.Text );
  RegsiterKey.Text := aKey;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.ChangeKey(const aMode: Integer);
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
  FLbDES.GenerateKey( aPhrase );
end;

{ ---------------------------------------------------------------------------- }

end.
