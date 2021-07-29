unit frmInvA05_3U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, Buttons;

type
  TfrmInvA05_3 = class(TForm)
    Panel1: TPanel;
    btnPrintNew: TBitBtn;
    btnExit: TBitBtn;
    GroupBox1: TGroupBox;
    ckbInvTitle: TCheckBox;
    ckbMailAddr: TCheckBox;
    ckbMemo1: TCheckBox;
    Panel2: TPanel;
    Label1: TLabel;
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmInvA05_3: TfrmInvA05_3;

implementation

uses frmMainU;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA05_3.FormShow(Sender: TObject);
begin
  Self.Caption := frmMain.GetFormTitleString( 'A05_3', '¦C¦L¿ï¶µ' );
end;

{ ---------------------------------------------------------------------------- }

end.
