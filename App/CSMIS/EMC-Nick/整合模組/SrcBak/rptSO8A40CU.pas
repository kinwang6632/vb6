unit rptSO8A40CU;

interface

uses Windows, SysUtils, Messages, Classes, Graphics, Controls,
  StdCtrls, ExtCtrls, Forms, QuickRpt, QRCtrls;

type
  TrptSO8A40C = class(TQuickRep)
    DetailBand1: TQRBand;
    PageFooterBand1: TQRBand;
    PageHeaderBand1: TQRBand;
    lblCompName: TQRLabel;
    lblBelongYM: TQRLabel;
    lblRealDate: TQRLabel;
    QRLabel5: TQRLabel;
    QRShape1: TQRShape;
    QRGroup1: TQRGroup;
    QRBand1: TQRBand;
    QRLabel2: TQRLabel;
    QRLabel3: TQRLabel;
    QRGroup2: TQRGroup;
    QRBand2: TQRBand;
    QRLabel1: TQRLabel;
    QRLabel4: TQRLabel;
    QRLabel6: TQRLabel;
    QRLabel7: TQRLabel;
    QRLabel8: TQRLabel;
    QRLabel10: TQRLabel;
    QRLabel11: TQRLabel;
    QRShape2: TQRShape;
    QRShape3: TQRShape;
    procedure lblCompNamePrint(sender: TObject; var Value: String);
    procedure lblRealDatePrint(sender: TObject; var Value: String);
    procedure lblBelongYMPrint(sender: TObject; var Value: String);
    procedure QRLabel2Print(sender: TObject; var Value: String);
    procedure QRLabel1Print(sender: TObject; var Value: String);
    procedure QRLabel6Print(sender: TObject; var Value: String);
    procedure QRLabel8Print(sender: TObject; var Value: String);
    procedure QRGroup1AfterPrint(Sender: TQRCustomBand;
      BandPrinted: Boolean);
    procedure QuickRepBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
    procedure QRLabel10Print(sender: TObject; var Value: String);
    procedure QRLabel11Print(sender: TObject; var Value: String);
    procedure QuickRepPreview(Sender: TObject);
  private
    nL_GroupProviderCounts,nL_TotalProviderCounts : Integer;
    fL_GroupProviderBenefit,fL_TotalProviderBenefit : Double;
  public
    sG_TransData,sG_CompCode,sG_CompName,sG_StartDate : String;
    sG_EndDate,sG_BelongYM,sG_ShowDetail : String;

  end;

var
  rptSO8A40C: TrptSO8A40C;

implementation

uses UdateTimeu, Ustru, frmRptPreviewU, dtmMain2U;

{$R *.DFM}

procedure TrptSO8A40C.QRGroup1AfterPrint(Sender: TQRCustomBand;
  BandPrinted: Boolean);
begin
    //群組顯示後歸零
    nL_GroupProviderCounts := 0;
    fL_GroupProviderBenefit := 0;
end;

procedure TrptSO8A40C.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
    //列印前將
    nL_TotalProviderCounts := 0;
    fL_TotalProviderBenefit := 0;
end;

procedure TrptSO8A40C.lblCompNamePrint(sender: TObject; var Value: String);
begin
    Value := sG_CompName;
end;

procedure TrptSO8A40C.lblRealDatePrint(sender: TObject; var Value: String);
begin
    Value := '實收日期: ' + TUdateTime.GetFormatDateStr(sG_StartDate,'#',false,true)
             + ' ~ ' + TUdateTime.GetFormatDateStr(sG_EndDate,'#',false,true);
end;

procedure TrptSO8A40C.lblBelongYMPrint(sender: TObject; var Value: String);
var
    sL_YM : String;
begin
    sL_YM := Copy(sG_BelongYM,1,4) + '/' + Copy(sG_BelongYM,5,6);
    Value := '歸屬年月: ' + TUdateTime.GetFormatDateStr(sL_YM,'#',false,true);
end;

procedure TrptSO8A40C.QRLabel2Print(sender: TObject; var Value: String);
var
    sL_Provider : String;
begin
    sL_Provider := dtmMain2.cdsSO114B.FieldByName('PROVIDERID').AsString;

    Value := dtmMain2.getProviderName(sL_Provider);
end;

procedure TrptSO8A40C.QRLabel1Print(sender: TObject; var Value: String);
var
    sL_SQL,sL_ProductID : String;
begin
    sL_ProductID := dtmMain2.cdsSO114B.FieldByName('PRODUCTID').AsString;
    sL_SQL := 'SELECT * FROM So112 WHERE PRODUCT_ID=''' + sL_ProductID + '''';

    with dtmMain2.cdsCom do
    begin
      Close;
      CommandText := sL_SQL;
      Open;
    end;

    Value := dtmMain2.cdsCom.FieldByName('PRODUCT_NAME').AsString;

end;

procedure TrptSO8A40C.QRLabel6Print(sender: TObject; var Value: String);
var
    sL_ProductID,sL_ProviderID : String;
    fL_ProviderBenefit : Double;
begin
    sL_ProductID := dtmMain2.cdsSO114B.FieldByName('PRODUCTID').AsString;
    sL_ProviderID := dtmMain2.cdsSO114B.FieldByName('PROVIDERID').AsString;
    fL_ProviderBenefit := dtmMain2.getProviderBenefit(sG_CompCode,sG_BelongYM,sL_ProductID,sL_ProviderID);

    fL_GroupProviderBenefit := fL_GroupProviderBenefit + fL_ProviderBenefit;
    fL_TotalProviderBenefit := fL_TotalProviderBenefit + fL_ProviderBenefit;

    Value := TUstr.CommaNumber(FloatToStr(fL_ProviderBenefit));
end;

procedure TrptSO8A40C.QRLabel8Print(sender: TObject; var Value: String);
var
    sL_ProductID,sL_ProviderID : String;
    nL_ProviderCounts : Integer;
begin
    sL_ProductID := dtmMain2.cdsSO114B.FieldByName('PRODUCTID').AsString;
    sL_ProviderID := dtmMain2.cdsSO114B.FieldByName('PROVIDERID').AsString;
    nL_ProviderCounts := dtmMain2.getProviderCounts(sG_CompCode,sG_BelongYM,sL_ProductID,sL_ProviderID);

    nL_GroupProviderCounts := nL_GroupProviderCounts + nL_ProviderCounts;
    nL_TotalProviderCounts := nL_TotalProviderCounts + nL_ProviderCounts;
    Value := TUstr.CommaNumber(IntToStr(nL_ProviderCounts));
end;



procedure TrptSO8A40C.QRLabel10Print(sender: TObject; var Value: String);
begin
    Value := '小計:  ' + TUstr.CommaNumber(FloatToStr(fL_GroupProviderBenefit));
end;

procedure TrptSO8A40C.QRLabel11Print(sender: TObject; var Value: String);
begin
    Value := '總計:  ' + TUstr.CommaNumber(FloatToStr(fL_TotalProviderBenefit));
end;

procedure TrptSO8A40C.QuickRepPreview(Sender: TObject);
var
    frmPreView : TfrmRptPreview;
begin
  frmPreView := TfrmRptPreview.Create(nil);

  frmPreView.G_ActiveDataSet := self.DataSet;

  frmPreView.sG_ShowDetail := self.sG_ShowDetail;
  frmPreView.sG_TransData := self.sG_TransData;
  frmPreView.sG_CompCode := self.sG_CompCode;
  frmPreView.sG_BelongYM := self.sG_BelongYM;
  frmPreView.Report := Self;
  frmPreView.QRPreview.QRPrinter := Self.QRPrinter;
  frmPreView.Show;

end;

end.
