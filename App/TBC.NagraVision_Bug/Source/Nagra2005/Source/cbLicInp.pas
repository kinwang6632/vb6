unit cbLicInp;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Menus, cxTextEdit, cxControls, cxContainer, cxEdit,
  cxLookAndFeelPainters, dxSkinsCore, dxSkinsDefaultPainters, cxButtons,
  cxGraphics, cxLookAndFeels;

type
  TfmLicInp = class(TForm)
    LicenceKey: TcxTextEdit;
    Bevel1: TBevel;
    btnOK: TcxButton;
    btnCancel: TcxButton;
    Image1: TImage;
    RegsiterKey: TcxTextEdit;
    lbLicenceNote: TLabel;
    lbLicenceKey: TLabel;
    lbRegisterKey: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure KeyPropertiesChange(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fmLicInp: TfmLicInp;

implementation

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TfmLicInp.FormCreate(Sender: TObject);
begin
  LicenceKey.Text := '';
  btnOK.Enabled := False;
  btnCancel.Enabled := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmLicInp.KeyPropertiesChange(Sender: TObject);
begin
  btnOK.Enabled := ( RegsiterKey.Text <> '' );
  btnOK.Default := btnOK.Enabled; 
end;

{ ---------------------------------------------------------------------------- }

end.
