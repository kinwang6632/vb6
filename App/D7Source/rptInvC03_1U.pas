unit rptInvC03_1U;

interface

uses Windows, SysUtils, Messages, Classes, Graphics, Controls,
  StdCtrls, ExtCtrls, Forms, QuickRpt, QRCtrls;

type
  TrptInvC03_1 = class(TQuickRep)
    DetailBand1: TQRBand;
    PageHeaderBand1: TQRBand;
    PageFooterBand1: TQRBand;
    lblBillID: TQRLabel;
    lblCustID: TQRLabel;
    lblCitemName: TQRLabel;
    lblChargeDate: TQRLabel;
    lblTotalAmount: TQRLabel;
    lblNotes: TQRLabel;
    QRLabel1: TQRLabel;
    QRLabel2: TQRLabel;
    QRLabel4: TQRLabel;
    QRLabel5: TQRLabel;
    QRLabel6: TQRLabel;
    QRLabel7: TQRLabel;
    QRShape1: TQRShape;
    QRShape2: TQRShape;
    QRLabel3: TQRLabel;
    QRLabel8: TQRLabel;
    lblCustName: TQRLabel;
    QRLabel9: TQRLabel;
    lblLogTime: TQRLabel;
    QRLabel10: TQRLabel;
    lblCounts: TQRLabel;
    QRLabel11: TQRLabel;
    QRLabel12: TQRLabel;
    QRLabel13: TQRLabel;
    QRLabel14: TQRLabel;
    qrlPage2: TQRSysData;
    qrlPage1: TQRLabel;
    qrlPrintTime1: TQRLabel;
    qrlPrintTime: TQRLabel;
    qrlOperator1: TQRLabel;
    qrlOperator: TQRLabel;
    //procedure QuickRepPreview(Sender: TObject);
    procedure lblBillIDPrint(sender: TObject; var Value: String);
    procedure lblCustIDPrint(sender: TObject; var Value: String);
    procedure lblCustNamePrint(sender: TObject; var Value: String);
    procedure lblCitemNamePrint(sender: TObject; var Value: String);
    procedure lblChargeDatePrint(sender: TObject; var Value: String);
    procedure lblTotalAmountPrint(sender: TObject; var Value: String);
    procedure lblNotesPrint(sender: TObject; var Value: String);
    procedure lblLogTimePrint(sender: TObject; var Value: String);
    procedure QuickRepBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
    procedure DetailBand1BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
    procedure lblCountsPrint(sender: TObject; var Value: String);
    procedure QRLabel12Print(sender: TObject; var Value: String);
    procedure QRLabel13Print(sender: TObject; var Value: String);
    procedure qrlOperatorPrint(sender: TObject; var Value: String);
    procedure qrlPrintTimePrint(sender: TObject; var Value: String);
  private
    nG_Counts : Integer;
    fG_AllTotalAmount : Double;
    sG_PrintTime : String;
  public
    sG_ConditionString,sG_UserID : String;
  end;

var
  rptInvC03_1: TrptInvC03_1;

implementation

uses frmRptPreviewU, dtmMainJU, Ustru;

{$R *.DFM}
{
procedure TrptInvC03_1.QuickRepPreview(Sender: TObject);
var
    frmPreView : TfrmRptPreview;
begin

    frmPreView := TfrmRptPreview.Create(nil);

    frmPreView.G_ActiveDataSet := self.DataSet;
    //frmPreView.bG_LockData := bG_LockData;
    frmPreView.Report := Self;
    frmPreView.QRPreview.QRPrinter := Self.QRPrinter;
    frmPreView.Show;

end;
}

procedure TrptInvC03_1.lblBillIDPrint(sender: TObject;
  var Value: String);

begin
    Value := self.DataSet.FieldByName('BILLID').AsString;
end;

procedure TrptInvC03_1.lblCustIDPrint(sender: TObject;
  var Value: String);
begin
    Value := self.DataSet.FieldByName('CUSTID').AsString;
end;

procedure TrptInvC03_1.lblCustNamePrint(sender: TObject;
  var Value: String);
begin
    Value := self.DataSet.FieldByName('CUSTNAME').AsString;
end;

procedure TrptInvC03_1.lblCitemNamePrint(sender: TObject;
  var Value: String);
begin
    Value := self.DataSet.FieldByName('Description').AsString;
end;

procedure TrptInvC03_1.lblChargeDatePrint(sender: TObject;
  var Value: String);
begin

    Value := self.DataSet.FieldByName('ChargeDate').AsString;
end;

procedure TrptInvC03_1.lblTotalAmountPrint(sender: TObject;
  var Value: String);
var
    fL_TotalAmount : Double;
begin
    fL_TotalAmount := self.DataSet.FieldByName('TotalAmount').AsFloat;
    Value := TUstr.ChangeMinusTag(fL_TotalAmount);
    //Value := TUstr.CommaNumber(self.DataSet.FieldByName('TotalAmount').AsString);

    fG_AllTotalAmount := fG_AllTotalAmount + fL_TotalAmount;
end;

procedure TrptInvC03_1.lblNotesPrint(sender: TObject;
  var Value: String);
begin
    Value := self.DataSet.FieldByName('Notes').AsString;
end;

procedure TrptInvC03_1.lblLogTimePrint(sender: TObject;
  var Value: String);
begin
    Value := self.DataSet.FieldByName('LogTime').AsString;
end;

procedure TrptInvC03_1.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
    nG_Counts := 0;
    fG_AllTotalAmount := 0;

    sG_PrintTime := DateTimeToStr(now);
end;

procedure TrptInvC03_1.DetailBand1BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
begin
    nG_Counts := nG_Counts + 1;
end;

procedure TrptInvC03_1.lblCountsPrint(sender: TObject;
  var Value: String);
begin
    Value := TUstr.CommaNumber(IntToStr(nG_Counts));
end;

procedure TrptInvC03_1.QRLabel12Print(sender: TObject;
  var Value: String);
begin
    Value := TUstr.ChangeMinusTag(fG_AllTotalAmount);
end;

procedure TrptInvC03_1.QRLabel13Print(sender: TObject;
  var Value: String);
begin
    Value := sG_ConditionString;
end;

procedure TrptInvC03_1.qrlOperatorPrint(sender: TObject;
  var Value: String);
begin
    Value := sG_UserID;
end;

procedure TrptInvC03_1.qrlPrintTimePrint(sender: TObject;
  var Value: String);
begin
    Value := sG_PrintTime;
end;

end.
