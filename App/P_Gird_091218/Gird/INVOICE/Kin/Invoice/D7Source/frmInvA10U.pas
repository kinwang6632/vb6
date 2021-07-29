unit frmInvA10U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Mask, Buttons, ExtCtrls;

type
  TfrmInvA10 = class(TForm)
    Panel1: TPanel;
    Label1: TLabel;
    Panel2: TPanel;
    btnOK: TBitBtn;
    btnExit: TBitBtn;
    Label3: TLabel;
    medQueryYear: TMaskEdit;
    procedure btnExitClick(Sender: TObject);
    procedure btnOKClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmInvA10: TfrmInvA10;

implementation

uses dtmMainU, dtmMainJU, frmMainU, frmInvA10_1U;

{$R *.dfm}

procedure TfrmInvA10.btnExitClick(Sender: TObject);
begin
    Close;
end;



procedure TfrmInvA10.btnOKClick(Sender: TObject);
var
    sL_QueryYear,sL_CompID : String;
begin
    sL_CompID := dtmMain.getCompID;
    sL_QueryYear := Trim(medQueryYear.Text);

    frmInvA10_1 := TfrmInvA10_1.Create(Application);
    frmInvA10_1.sG_CompID := dtmMain.getCompID;
    frmInvA10_1.sG_QueryYear := sL_QueryYear;
    frmInvA10_1.ShowModal;
    frmInvA10_1.Free;

end;

procedure TfrmInvA10.FormShow(Sender: TObject);
begin
   Self.Caption := frmMain.GetFormTitleString( 'A0A', '發票鎖帳作業' );
end;

end.
