unit CbLicInp;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls,
  cxTextEdit, cxControls, cxContainer, cxEdit, cxLabel,
  cxLookAndFeelPainters, cxButtons;

type
  TfrmLicInp = class(TForm)
    Msg: TcxLabel;
    InputLic: TcxTextEdit;
    Bevel1: TBevel;
    btnOK: TcxButton;
    btnCancel: TcxButton;
    Image1: TImage;
    procedure FormCreate(Sender: TObject);
    procedure KeyPropertiesChange(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmLicInp: TfrmLicInp;

implementation

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TfrmLicInp.FormCreate(Sender: TObject);
begin
  InputLic.Text := '';
  btnOK.Enabled := False;
  btnCancel.Enabled := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmLicInp.KeyPropertiesChange(Sender: TObject);
begin
  btnOK.Enabled := ( InputLic.Text <> '' );
  btnOK.Default := btnOK.Enabled; 
end;

{ ---------------------------------------------------------------------------- }

end.
