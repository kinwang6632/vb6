unit rptSO8A40BU;

interface

uses Windows, SysUtils, Messages, Classes, Graphics, Controls,
  StdCtrls, ExtCtrls, Forms, QuickRpt, QRCtrls;

type
  TrptSO8A40B = class(TQuickRep)
    DetailBand1: TQRBand;
    PageFooterBand1: TQRBand;
    PageHeaderBand1: TQRBand;
    lblCompName: TQRLabel;
    lblBelongYM: TQRLabel;
    lblRealDate: TQRLabel;
    QRLabel1: TQRLabel;
    QRLabel2: TQRLabel;
    QRLabel3: TQRLabel;
    QRLabel4: TQRLabel;
    QRLabel5: TQRLabel;
    QRShape1: TQRShape;
    QRLabel6: TQRLabel;
    QRLabel7: TQRLabel;
    procedure lblCompNamePrint(sender: TObject; var Value: String);
    procedure lblRealDatePrint(sender: TObject; var Value: String);
    procedure lblBelongYMPrint(sender: TObject; var Value: String);
    procedure QRLabel2Print(sender: TObject; var Value: String);
    procedure QRLabel4Print(sender: TObject; var Value: String);
    procedure QuickRepPreview(Sender: TObject);
    procedure QRLabel7Print(sender: TObject; var Value: String);
  private
    sG_RptRealDate,sG_RptBelongYM : String;
  public
    sG_CompName,sG_CompCode,sG_StartDate,sG_EndDate,sG_BelongYM,sG_ShowDetail : String;
    sG_TransData,sG_ShowType : String;  

  end;

var
  rptSO8A40B: TrptSO8A40B;

implementation

uses frmRptPreviewU, UdateTimeu, dtmMain2U;

{$R *.DFM}

procedure TrptSO8A40B.lblCompNamePrint(sender: TObject; var Value: String);
begin
    Value := sG_CompName;
end;

procedure TrptSO8A40B.lblRealDatePrint(sender: TObject; var Value: String);
begin
    Value := sG_RptRealDate;
end;

procedure TrptSO8A40B.lblBelongYMPrint(sender: TObject; var Value: String);
var
    sL_YM : String;
begin
    Value := sG_RptBelongYM;
end;

procedure TrptSO8A40B.QRLabel2Print(sender: TObject; var Value: String);
var
    sL_CustID : String;
begin
    sL_CustID := dtmMain2.cdsSO116.FieldByName('CUSTID').AsString;
    Value := dtmMain2.getCustName(sL_CustID);
end;



procedure TrptSO8A40B.QRLabel4Print(sender: TObject; var Value: String);
var
    sL_ChannelString : String;
begin
   sL_ChannelString := dtmMain2.cdsSO116.FieldByName('CHANNEL_STRING').AsString;
   Value := sL_ChannelString;
end;

procedure TrptSO8A40B.QuickRepPreview(Sender: TObject);
var
    frmPreView : TfrmRptPreview;
    sL_YM : String;
begin
    sG_RptRealDate := '實收日期: ' + TUdateTime.GetFormatDateStr(sG_StartDate,'#',false,true)
             + ' ~ ' + TUdateTime.GetFormatDateStr(sG_EndDate,'#',false,true);

    sL_YM := Copy(sG_BelongYM,1,4) + '/' + Copy(sG_BelongYM,5,6);
    sG_RptBelongYM := '歸屬年月: ' + TUdateTime.GetFormatDateStr(sL_YM,'#',false,true);

  frmPreView := TfrmRptPreview.Create(nil);

  frmPreView.G_ActiveDataSet := self.DataSet;

  frmPreView.sG_ShowDetail := self.sG_ShowDetail;
  frmPreView.sG_TransData := self.sG_TransData;
  frmPreView.sG_CompCode := self.sG_CompCode;
  frmPreView.sG_BelongYM := self.sG_BelongYM;

  //ToExcel 會用到
  frmPreView.sG_RptName := '客戶收視明細表';
  frmPreView.sG_RptCompName := sG_CompName;
  frmPreView.sG_RptRealDate := sG_RptRealDate;
  frmPreView.sG_RptBelongYM := sG_RptBelongYM;
  frmPreView.sG_ShowType := sG_ShowType;


  frmPreView.Report := Self;
  frmPreView.QRPreview.QRPrinter := Self.QRPrinter;
  frmPreView.Show;
end;

procedure TrptSO8A40B.QRLabel7Print(sender: TObject; var Value: String);
begin
  Value := dtmMain2.cdsSO116.FieldByName('CUSTID').AsString;
end;



end.
