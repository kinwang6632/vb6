unit cbAbout;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, Menus, cxButtons, cxLookAndFeelPainters,
  cxGraphics, cxLookAndFeels;

type
  TfmAbout = class(TForm)
    Panel1: TPanel;
    Bevel1: TBevel;
    btnOk: TcxButton;
    Image4: TImage;
    Label1: TLabel;
    lblVendor: TLabel;
    Label3: TLabel;
    lblAppTitle: TLabel;
    lblVersion: TLabel;
    Label7: TLabel;
    Image1: TImage;
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fmAbout: TfmAbout;

implementation

{$R *.dfm}

uses cbResStr;

{ ---------------------------------------------------------------------------- }

procedure TfmAbout.FormCreate(Sender: TObject);
begin
  lblVendor.Caption := SVendor;
  lblAppTitle.Caption := SAppTitle;
  lblVersion.Caption := SVersion;
  btnOk.Caption := SButtonOK;
  Self.Caption := SMenuItem3;
end;

{ ---------------------------------------------------------------------------- }

end.
