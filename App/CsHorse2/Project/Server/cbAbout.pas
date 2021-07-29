unit cbAbout;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, jpeg, Menus,
{ App: }
  cbLanguage,
{ Developer Express: }
  cxButtons, cxLookAndFeelPainters;


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

{ ---------------------------------------------------------------------------- }

procedure TfmAbout.FormCreate(Sender: TObject);
begin
  lblVendor.Caption := LanguageManager.Get( 'SVendor' );
  lblAppTitle.Caption := LanguageManager.Get( 'SAppTitle' );
  btnOk.Caption := LanguageManager.Get( 'SConfirm' );
  Self.Caption := LanguageManager.Get( 'fmAbout' );
end;

{ ---------------------------------------------------------------------------- }

end.
