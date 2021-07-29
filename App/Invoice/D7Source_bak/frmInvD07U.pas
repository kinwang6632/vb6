unit frmInvD07U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, Mask, ExtCtrls, fraYMU, cxButtons, Menus,
  cxLookAndFeelPainters;

type
  TfrmInvD07 = class(TForm)
    Panel1: TPanel;
    Label5: TLabel;
    Panel2: TPanel;
    fraYM: TfraYM;
    medYear: TMaskEdit;
    rdbYearMonth: TRadioButton;
    rdbYear: TRadioButton;
    btnOk: TcxButton;
    btnClose: TcxButton;
    procedure FormShow(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure btnOKClick(Sender: TObject);
    procedure rdbYearMonthClick(Sender: TObject);
    procedure rdbYearClick(Sender: TObject);
  private
    FLOGINOK: BOOLEAN;
    procedure SetLOGINOK(const Value: BOOLEAN);
    { Private declarations }
    function isDataOK : Boolean;
  public
    { Public declarations }
    PROPERTY LOGINOK:BOOLEAN read FLOGINOK write SetLOGINOK;
  end;

var
  frmInvD07: TfrmInvD07;

implementation

uses frmMainU, dtmMainU, dtmMainJU, frmInvD07_1U, Ustru;

{$R *.dfm}

{ TfrmInvD07 }

procedure TfrmInvD07.SetLOGINOK(const Value: BOOLEAN);
begin
  FLOGINOK := Value;
end;

procedure TfrmInvD07.FormShow(Sender: TObject);
begin
  self.Caption := frmMain.GetFormTitleString( 'D07', '發票字軌維護' );
end;

procedure TfrmInvD07.btnExitClick(Sender: TObject);
begin
    Close;
end;

function TfrmInvD07.isDataOK: Boolean;
var
    sL_Year,sL_YM : String;
begin
    Result := true;

    sL_YM := Trim(fraYM.getYM);
    if rdbYearMonth.Checked then
    begin
      if sL_YM = EMPTY_YM_STR  then
      begin
        MessageDlg('請輸入發票年月',mtError,[mbOk],0);
        fraYM.mseYM.SetFocus;
        Result := false;
        Exit;
      end;
    end;

    sL_Year := Trim(medYear.Text);
    if rdbYear.Checked then
    begin
      if sL_Year = ''  then
      begin
        MessageDlg('請輸入年度條件',mtError,[mbOk],0);
        medYear.SetFocus;
        Result := false;
        Exit;
      end
      else if Length(sL_Year) <> 4  then
      begin
        MessageDlg('年度條件格式錯誤!(範例:2004)',mtError,[mbOk],0);
        medYear.SetFocus;
        Result := false;
        Exit;
      end;
    end;


end;

procedure TfrmInvD07.btnOKClick(Sender: TObject);
var
    sL_CompID,sL_Year,sL_YM : String;
begin
    if isDataOK then
    begin
      sL_CompID := dtmMain.getCompID;
      sL_Year := Trim(medYear.Text);
      sL_YM := TUstr.replaceStr(Trim(fraYM.getYM),'/','');
      if sL_YM <> '' then
      begin
        sL_YM := dtmMainJ.changePrefixYM(sL_YM);
        fraYM.setYM(Copy(sL_YM,1,4) + '/' + Copy(sL_YM,5,2));
      end;

      dtmMainJ.getInv099Data(sL_CompID,sL_Year,sL_YM);

      frmInvD07_1 := TfrmInvD07_1.Create(Application);
      frmInvD07_1.sG_CompID := sL_CompID;
      frmInvD07_1.sG_QueryYear := sL_Year;
      frmInvD07_1.sG_QueryYM := sL_YM;
      frmInvD07_1.ShowModal;
      frmInvD07_1.Free;
    end;
end;

procedure TfrmInvD07.rdbYearMonthClick(Sender: TObject);
begin
    medYear.Clear;
    fraYM.mseYM.SetFocus;
end;

procedure TfrmInvD07.rdbYearClick(Sender: TObject);
begin
    fraYM.mseYM.Clear;
    medYear.SetFocus;
end;

end.
