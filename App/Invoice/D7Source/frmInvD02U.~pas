unit frmInvD02U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Buttons, Mask;

type
  TfrmInvD02 = class(TForm)
    Panel1: TPanel;
    Label5: TLabel;
    Panel2: TPanel;
    rdbCustID1: TRadioButton;
    rdbCname: TRadioButton;
    medCustID1: TMaskEdit;
    medCname: TMaskEdit;
    btnOK: TBitBtn;
    btnExit: TBitBtn;
    Rdbsno: TRadioButton;
    medSno: TMaskEdit;
    medCustID2: TMaskEdit;
    Label1: TLabel;
    RdbTel: TRadioButton;
    medTel: TMaskEdit;
    rdbNo: TRadioButton;
    procedure FormShow(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure btnOKClick(Sender: TObject);
    procedure rdbCustID1Click(Sender: TObject);
    procedure medCustID1Exit(Sender: TObject);
    procedure rdbCnameClick(Sender: TObject);
    procedure RdbTelClick(Sender: TObject);
    procedure RdbsnoClick(Sender: TObject);
  private
    { Private declarations }
    function isDataOK : Boolean;
  public
    { Public declarations }
  end;

var
  frmInvD02: TfrmInvD02;

implementation

uses frmMainU, frmInvD02_1U, dtmMainJU, dtmMainU, cbUtilis;

{$R *.dfm}

procedure TfrmInvD02.FormShow(Sender: TObject);
begin
   self.Caption := frmMain.GetFormTitleString( 'D02', '客戶基本資料查詢' );
   medCustID1.SetFocus;
end;

procedure TfrmInvD02.btnExitClick(Sender: TObject);
begin
    Close;
end;

function TfrmInvD02.isDataOK: Boolean;
var
    sL_Var : String;
begin
   Result := true;

   if rdbCustID1.Checked then
   begin
      sL_Var := Trim(medCustID1.Text);
      if sl_Var = '' then
      begin
        MessageDlg('請輸入客戶編號條件',mtError,[mbOk],0);
        medCustID1.SetFocus;
        Result := false;
        Exit;
      end;

      sL_Var := Trim(medCustID2.Text);
      if sl_Var = '' then
      begin
        MessageDlg('請輸入客戶編號範圍條件',mtError,[mbOk],0);
        medCustID1.SetFocus;
        Result := false;
        Exit;
      end;
   end;

   if rdbCname.Checked then
   begin
      sL_Var := Trim(medCname.Text);
      if sl_Var = '' then
      begin
        MessageDlg('請輸入客戶姓名條件',mtError,[mbOk],0);
        medCname.SetFocus;
        Result := false;
        Exit;
      end;
   end;

   if Rdbsno.Checked then
   begin
      sL_Var := Trim(medSno.Text);
      if sl_Var = '' then
      begin
        MessageDlg('請輸入統一編號條件',mtError,[mbOk],0);
        medSno.SetFocus;
        Result := false;
        Exit;
      end;
   end;

   if rdbTel.Checked then
   begin
     sL_Var := Trim(medTel.Text);
     if sL_Var = '' then
     begin
       MessageDlg('請輸入電話條件',mtError,[mbOk],0);
       medTel.SetFocus;
       Result := false;
       Exit;
     end;
   end;
end;

procedure TfrmInvD02.btnOKClick(Sender: TObject);
var
    sL_CustID1,
    sL_CustID2,
    sL_CustName,
    sL_Sno,
    sL_Tel,
    sL_CompID : String;
    bL_Select : Integer;
begin
    if isDataOK then
    begin
      sL_CompID := dtmMain.getCompID;
      
      bL_Select := -1;

      if rdbCustID1.Checked then
      begin
        sL_CustID1 := Trim(medCustID1.Text);
        sL_CustID2 := Trim(medCustID2.Text);
        bL_Select := 1;
      end;

      if rdbCname.Checked then
      begin
        sL_CustName := Trim(medCname.Text);
        bL_Select := 2;
      end;

      if Rdbsno.Checked then
      begin
        sL_Sno := Trim(medSno.Text);
        bL_Select := 3;
      end;

      if rdbTel.Checked then
      begin
        sL_Tel := Trim(medTel.Text);
        bL_Select := 4;
      end;

      if rdbNo.Checked then
      begin
        bL_Select := 5;
      end;


      if not dtmMainJ.getInv002Data(bL_Select,sL_CustID1,sL_CustID2,sL_CustName,
         sL_Sno,sL_Tel,sL_CompID ) and ( bL_Select <> 5 ) then
      begin
        WarningMsg( '查詢無符合資料。' );
        Exit;
      end;
      frmInvD02_1 := TfrmInvD02_1.Create(Application);
      frmInvD02_1.sG_CompID := dtmMain.getCompID;
      frmInvD02_1.bG_Select := bL_Select;
      frmInvD02_1.sG_CustID1 := sL_CustID1;
      frmInvD02_1.sG_CustID2 := sL_CustID2;
      frmInvD02_1.sG_CustName := sL_CustName;
      frmInvD02_1.sG_Sno := sL_Sno;
      frmInvD02_1.sg_Tel := sL_Tel;
      frmInvD02_1.ShowModal;
      frmInvD02_1.Free;
    end;
end;

procedure TfrmInvD02.rdbCustID1Click(Sender: TObject);
begin
  medCustID1.Clear;
  medCustID2.Clear;
  medCname.Clear;
  medSno.Clear;
  medTel.Clear;
  medCustID1.SetFocus;
end;

procedure TfrmInvD02.medCustID1Exit(Sender: TObject);
begin
   if trim(medCustID2.Text) = '' then
   begin
      medCustID2.Text := medCustID1.Text;
   end;
end;

procedure TfrmInvD02.rdbCnameClick(Sender: TObject);
begin
  medCustID1.Clear;
  medCustID2.Clear;
  medCname.Clear;
  medSno.Clear;
  medTel.Clear;
  medCname.SetFocus;
end;

procedure TfrmInvD02.RdbTelClick(Sender: TObject);
begin
  medCustID1.Clear;
  medCustID2.Clear;
  medCname.Clear;
  medSno.Clear;
  medTel.Clear;
  medTel.SetFocus;
end;

procedure TfrmInvD02.RdbsnoClick(Sender: TObject);
begin
  medCustID1.Clear;
  medCustID2.Clear;
  medCname.Clear;
  medSno.Clear;
  medTel.Clear;
  medSno.SetFocus;
end;

end.
