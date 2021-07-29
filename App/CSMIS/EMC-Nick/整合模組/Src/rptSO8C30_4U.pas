unit rptSO8C30_4U;

interface

uses Windows, SysUtils, Messages, Classes, Graphics, Controls,
  StdCtrls, ExtCtrls, Forms, QuickRpt, QRCtrls;

type
  TrptSO8C30_4 = class(TQuickRep)
    QRBand1: TQRBand;
    QRLabel1: TQRLabel;
    QRLabel2: TQRLabel;
    QRLabel3: TQRLabel;
    QRLabel4: TQRLabel;
    QRLabel5: TQRLabel;
    QRLabel6: TQRLabel;
    QRLabel7: TQRLabel;
    QRLabel9: TQRLabel;
    QRLabel10: TQRLabel;
    QRLabel11: TQRLabel;
    QRLabel13: TQRLabel;
    QRLabel14: TQRLabel;
    QRLabel15: TQRLabel;
    QRLabel16: TQRLabel;
    QRLabel17: TQRLabel;
    QRLabel18: TQRLabel;
    QRShape1: TQRShape;
    QRShape2: TQRShape;
    QRShape3: TQRShape;
    QRShape4: TQRShape;
    QRShape5: TQRShape;
    QRShape6: TQRShape;
    QRBand2: TQRBand;
    QRLabel21: TQRLabel;
    QRLabel22: TQRLabel;
    QRLabel23: TQRLabel;
    QRLabel8: TQRLabel;
    QRLabel12: TQRLabel;
    QRLabel24: TQRLabel;
    QRLabel25: TQRLabel;
    QRLabel26: TQRLabel;
    QRShape7: TQRShape;
    QRShape8: TQRShape;
    QRDBText2: TQRDBText;
    QRShape10: TQRShape;
    QRShape11: TQRShape;
    QRLabel20: TQRLabel;
    QRShape12: TQRShape;
    QRShape13: TQRShape;
    QRDBText1: TQRDBText;
    QRShape14: TQRShape;
    QRShape15: TQRShape;
    QRShape16: TQRShape;
    QRShape17: TQRShape;
    QRDBText3: TQRDBText;
    QRDBText4: TQRDBText;
    procedure QRLabel10Print(sender: TObject; var Value: String);
    procedure QRLabel11Print(sender: TObject; var Value: String);
    procedure QRLabel13Print(sender: TObject; var Value: String);
    procedure QRLabel15Print(sender: TObject; var Value: String);
    procedure QRLabel14Print(sender: TObject; var Value: String);
    procedure QRLabel22Print(sender: TObject; var Value: String);
    procedure QRLabel21Print(sender: TObject; var Value: String);
    procedure QRLabel23Print(sender: TObject; var Value: String);
    procedure QRLabel20Print(sender: TObject; var Value: String);
    procedure QuickRepPreview(Sender: TObject);
  private

  public
    sG_Operator : String;
    sG_GetPaperEmpsName, sG_ReportPaperNumStr, sG_ReportGetPaperDateStr, sG_CompName : String;

  end;

var
  rptSO8C30_4: TrptSO8C30_4;

implementation

uses dtmMain3U, UdateTimeu, frmRptPreviewU;

{$R *.DFM}

procedure TrptSO8C30_4.QRLabel10Print(sender: TObject; var Value: String);
begin
    Value := sG_CompName;
end;

procedure TrptSO8C30_4.QRLabel11Print(sender: TObject; var Value: String);
begin
    Value := sG_ReportGetPaperDateStr;
end;

procedure TrptSO8C30_4.QRLabel13Print(sender: TObject; var Value: String);
begin
    Value := sG_GetPaperEmpsName;
end;

procedure TrptSO8C30_4.QRLabel15Print(sender: TObject; var Value: String);
begin
    Value := TUdateTime.CDateStr(date,9);
end;

procedure TrptSO8C30_4.QRLabel14Print(sender: TObject; var Value: String);
begin
    Value := sG_Operator;
end;

procedure TrptSO8C30_4.QRLabel22Print(sender: TObject; var Value: String);
begin
    Value := self.DataSet.fieldByName('EMPNAME').AsString;
end;

procedure TrptSO8C30_4.QRLabel21Print(sender: TObject; var Value: String);
var
    nL_Year,nL_Month, nL_Day : word;
begin
    DecodeDate(self.DataSet.FieldByName('GETPAPERDATE').AsDateTime, nL_Year, nL_Month, nL_Day);
    Value := Format('%3d/%.2d/%.2d', [nL_Year-1911, nL_Month, nL_Day]);
end;

procedure TrptSO8C30_4.QRLabel23Print(sender: TObject; var Value: String);
begin
    Value := self.DataSet.fieldByName('TOTALPAPERCOUNT').AsString;
end;

procedure TrptSO8C30_4.QRLabel20Print(sender: TObject; var Value: String);
begin
    Value := self.DataSet.fieldByName('SPAPERNUM').AsString + '~' + self.DataSet.fieldByName('EPAPERNUM').AsString ;
end;

procedure TrptSO8C30_4.QuickRepPreview(Sender: TObject);
var
  frmPreView : TfrmRptPreview;
begin
  frmPreView := TfrmRptPreview.Create(nil);
  frmPreView.Report := Self;
  frmPreView.QRPreview.QRPrinter := Self.QRPrinter;
  frmPreView.Show;
end;

end.
