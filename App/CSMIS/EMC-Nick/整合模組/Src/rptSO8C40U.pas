unit rptSO8C40U;

interface

uses Windows, SysUtils, Messages, Classes, Graphics, Controls,
  StdCtrls, ExtCtrls, Forms, QuickRpt, QRCtrls;

type
  TrptSO8C40 = class(TQuickRep)
    QRBand1: TQRBand;
    QRLabel1: TQRLabel;
    QRLabel2: TQRLabel;
    QRLabel3: TQRLabel;
    QRLabel4: TQRLabel;
    QRLabel5: TQRLabel;
    QRLabel6: TQRLabel;
    QRLabel9: TQRLabel;
    QRLabel10: TQRLabel;
    QRLabel11: TQRLabel;
    QRLabel14: TQRLabel;
    QRLabel15: TQRLabel;
    QRBand2: TQRBand;
    QRDBText1: TQRDBText;
    QRDBText2: TQRDBText;
    QRDBText4: TQRDBText;
    QRDBText5: TQRDBText;
    QRDBText7: TQRDBText;
    QRGroup1: TQRGroup;
    QRLabel16: TQRLabel;
    QRDBText3: TQRDBText;
    QRLabel17: TQRLabel;
    QRShape10: TQRShape;
    QRShape11: TQRShape;
    QRLabel18: TQRLabel;
    QRLabel7: TQRLabel;
    QRLabel8: TQRLabel;
    QRShape1: TQRShape;
    QRShape2: TQRShape;
    QRLabel12: TQRLabel;
    QRShape20: TQRShape;
    QRShape21: TQRShape;
    QRShape3: TQRShape;
    QRShape4: TQRShape;
    QRShape5: TQRShape;
    QRShape6: TQRShape;
    QRShape7: TQRShape;
    QRShape8: TQRShape;
    QRGroup2: TQRGroup;
    QRBand3: TQRBand;
    QRShape9: TQRShape;
    procedure QRLabel10Print(sender: TObject; var Value: String);
    procedure QRLabel11Print(sender: TObject; var Value: String);
    procedure QRLabel12Print(sender: TObject; var Value: String);
    procedure QRLabel13Print(sender: TObject; var Value: String);
    procedure QRLabel14Print(sender: TObject; var Value: String);
    procedure QRLabel15Print(sender: TObject; var Value: String);
    procedure QRLabel20Print(sender: TObject; var Value: String);
    procedure QuickRepPreview(Sender: TObject);
    procedure QRDBText1Print(sender: TObject; var Value: String);
  private

  public
    sG_Operator : String;
    sG_GetPaperEmpsName, sG_ReportPaperNumStr, sG_ReportGetPaperDateStr, sG_CompName : String;
  end;

var
  rptSO8C40: TrptSO8C40;

implementation

uses dtmMain3U, UdateTimeu, frmRptPreviewU;

{$R *.DFM}

procedure TrptSO8C40.QRLabel10Print(sender: TObject; var Value: String);
begin
    Value := sG_CompName;
end;

procedure TrptSO8C40.QRLabel11Print(sender: TObject; var Value: String);
begin
    Value := sG_ReportGetPaperDateStr;
end;

procedure TrptSO8C40.QRLabel12Print(sender: TObject; var Value: String);
begin
    Value := sG_ReportPaperNumStr;
end;

procedure TrptSO8C40.QRLabel13Print(sender: TObject; var Value: String);
begin
    Value := sG_GetPaperEmpsName;
end;

procedure TrptSO8C40.QRLabel14Print(sender: TObject; var Value: String);
begin
    Value := sG_Operator;
end;

procedure TrptSO8C40.QRLabel15Print(sender: TObject; var Value: String);
begin
    Value := TUdateTime.CDateStr(date,9);
end;

procedure TrptSO8C40.QRLabel20Print(sender: TObject; var Value: String);
begin
    Value := self.DataSet.fieldByName('EMPNO').AsString; 
end;

procedure TrptSO8C40.QuickRepPreview(Sender: TObject);
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

procedure TrptSO8C40.QRDBText1Print(sender: TObject; var Value: String);
begin
    if UpperCase(Value)='CUSTID' then
      Value := '?????s??'
    else if UpperCase(Value)='CUSTNAME' then
      Value := '?????W??'
    else if UpperCase(Value)='CUSTTEL' then
      Value := '?????q??'
    else if UpperCase(Value)='BILLNO' then
      Value := '???O????'
    else if UpperCase(Value)='REALDATE' then
      Value := '????????';
end;

end.
