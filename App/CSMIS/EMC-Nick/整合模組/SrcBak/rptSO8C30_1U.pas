unit rptSO8C30_1U;

interface

uses Windows, SysUtils, Messages, Classes, Graphics, Controls,
  StdCtrls, ExtCtrls, Forms, QuickRpt, QRCtrls;

type
  TrptSO8C30_1 = class(TQuickRep)
    QRBand1: TQRBand;
    QRLabel1: TQRLabel;
    QRLabel2: TQRLabel;
    QRLabel3: TQRLabel;
    QRLabel4: TQRLabel;
    QRLabel5: TQRLabel;
    QRLabel6: TQRLabel;
    QRLabel7: TQRLabel;
    QRLabel8: TQRLabel;
    QRLabel9: TQRLabel;
    QRLabel10: TQRLabel;
    QRLabel11: TQRLabel;
    QRLabel12: TQRLabel;
    QRLabel13: TQRLabel;
    QRLabel14: TQRLabel;
    QRLabel15: TQRLabel;
    QRBand2: TQRBand;
    QRLabel17: TQRLabel;
    QRLabel18: TQRLabel;
    QRLabel19: TQRLabel;
    QRShape1: TQRShape;
    QRShape2: TQRShape;
    QRShape3: TQRShape;
    QRShape4: TQRShape;
    QRShape5: TQRShape;
    QRShape6: TQRShape;
    QRShape7: TQRShape;
    QRShape8: TQRShape;
    QRLabel20: TQRLabel;
    QRLabel21: TQRLabel;
    QRLabel22: TQRLabel;
    QRLabel23: TQRLabel;
    QRDBText1: TQRDBText;
    QRBand3: TQRBand;
    QRShape9: TQRShape;
    QRLabel16: TQRLabel;
    procedure QRLabel10Print(sender: TObject; var Value: String);
    procedure QRLabel11Print(sender: TObject; var Value: String);
    procedure QRLabel12Print(sender: TObject; var Value: String);
    procedure QRLabel13Print(sender: TObject; var Value: String);
    procedure QRLabel14Print(sender: TObject; var Value: String);
    procedure QRLabel15Print(sender: TObject; var Value: String);
    procedure QRLabel20Print(sender: TObject; var Value: String);
    procedure QRLabel21Print(sender: TObject; var Value: String);
    procedure QRLabel22Print(sender: TObject; var Value: String);
    procedure QRLabel23Print(sender: TObject; var Value: String);
    procedure QuickRepPreview(Sender: TObject);
  private

  public
    sG_Operator : String;
    sG_GetPaperEmpsName, sG_ReportPaperNumStr, sG_ReportGetPaperDateStr, sG_CompName : String;
  end;

var
  rptSO8C30_1: TrptSO8C30_1;

implementation

uses dtmMain3U, UdateTimeu, frmRptPreviewU;

{$R *.DFM}

procedure TrptSO8C30_1.QRLabel10Print(sender: TObject; var Value: String);
begin
    Value := sG_CompName;
end;

procedure TrptSO8C30_1.QRLabel11Print(sender: TObject; var Value: String);
begin
    Value := sG_ReportGetPaperDateStr;
end;

procedure TrptSO8C30_1.QRLabel12Print(sender: TObject; var Value: String);
begin
    Value := sG_ReportPaperNumStr;
end;

procedure TrptSO8C30_1.QRLabel13Print(sender: TObject; var Value: String);
begin
    Value := sG_GetPaperEmpsName;
end;

procedure TrptSO8C30_1.QRLabel14Print(sender: TObject; var Value: String);
begin
    Value := sG_Operator;
end;

procedure TrptSO8C30_1.QRLabel15Print(sender: TObject; var Value: String);
begin
    Value := TUdateTime.CDateStr(date,9);
end;

procedure TrptSO8C30_1.QRLabel20Print(sender: TObject; var Value: String);
begin
    Value := self.DataSet.fieldByName('EMPNO').AsString; 
end;

procedure TrptSO8C30_1.QRLabel21Print(sender: TObject; var Value: String);
var
    nL_Year,nL_Month, nL_Day : word;
begin
    DecodeDate(self.DataSet.FieldByName('GETPAPERDATE').AsDateTime, nL_Year, nL_Month, nL_Day);
    Value := Format('%3d/%.2d/%.2d', [nL_Year-1911, nL_Month, nL_Day]);
end;

procedure TrptSO8C30_1.QRLabel22Print(sender: TObject; var Value: String);
begin
    Value := self.DataSet.fieldByName('EMPNAME').AsString;
end;

procedure TrptSO8C30_1.QRLabel23Print(sender: TObject; var Value: String);
begin
    Value := self.DataSet.fieldByName('MIN_PAPER_NUM').AsString + ' ¦Ü ' + self.DataSet.fieldByName('MAX_PAPER_NUM').AsString;

end;

procedure TrptSO8C30_1.QuickRepPreview(Sender: TObject);
var
    frmPreView : TfrmRptPreview;

begin
  //inherited;
{
  if (sG_FormatName<>'')  then
  begin
    sL_FullPathRptControlIniFileName := TEMP_REPORT_PATH + sG_FormatName+'.ini';
    TUother.restoreControlInfo(self,   sL_FullPathRptControlIniFileName);
  end;
}

  frmPreView := TfrmRptPreview.Create(nil);

  //if not Assigned(G_ExcelDataSet) then
//    frmPreView.G_ActiveDataSet := self.DataSet;
  //else
  //  frmPreView.G_ActiveDataSet := G_ExcelDataSet;

  frmPreView.Report := Self;
  frmPreView.QRPreview.QRPrinter := Self.QRPrinter;
  frmPreView.Show;
end;

end.
