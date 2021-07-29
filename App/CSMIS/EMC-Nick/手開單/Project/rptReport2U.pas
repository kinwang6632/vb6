unit rptReport2U;

interface

uses Windows, SysUtils, Messages, Classes, Graphics, Controls,
  StdCtrls, ExtCtrls, Forms, QuickRpt, QRCtrls;

type
  TrptReport2 = class(TQuickRep)
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
    QRGroup1: TQRGroup;
    QRDBText1: TQRDBText;
    QRDBText2: TQRDBText;
    QRDBText3: TQRDBText;
    QRBand2: TQRBand;
    QRDBText4: TQRDBText;
    QRBand4: TQRBand;
    QRShape9: TQRShape;
    QRLabel18: TQRLabel;
    QRLabel19: TQRLabel;
    QRLabel16: TQRLabel;
    QRLabel17: TQRLabel;
    QRShape1: TQRShape;
    QRShape2: TQRShape;
    QRShape3: TQRShape;
    QRShape4: TQRShape;
    procedure QRLabel10Print(sender: TObject; var Value: String);
    procedure QRLabel11Print(sender: TObject; var Value: String);
    procedure QRLabel12Print(sender: TObject; var Value: String);
    procedure QRLabel13Print(sender: TObject; var Value: String);
    procedure QRLabel14Print(sender: TObject; var Value: String);
    procedure QRLabel15Print(sender: TObject; var Value: String);
    procedure QRLabel21Print(sender: TObject; var Value: String);
    procedure QRLabel19Print(sender: TObject; var Value: String);
    procedure QuickRepPreview(Sender: TObject);
    procedure QRDBText1Print(sender: TObject; var Value: String);
    procedure QRDBText3Print(sender: TObject; var Value: String);
  private

  public
    sG_Operator, sG_ActiveEmpNo, sG_ActiveGetPaperDate : String;
    sG_GetPaperEmpsName, sG_ReportPaperNumStr, sG_ReportGetPaperDateStr, sG_CompName : String;
  end;

var
  rptReport2: TrptReport2;

implementation

uses dtmMainU, UdateTimeu, Ustru, frmRptPreviewU;

{$R *.DFM}

procedure TrptReport2.QRLabel10Print(sender: TObject; var Value: String);
begin
    Value := sG_CompName;
end;

procedure TrptReport2.QRLabel11Print(sender: TObject; var Value: String);
begin
    Value := sG_ReportGetPaperDateStr;
end;

procedure TrptReport2.QRLabel12Print(sender: TObject; var Value: String);
begin
    Value := sG_ReportPaperNumStr;
end;

procedure TrptReport2.QRLabel13Print(sender: TObject; var Value: String);
begin
    Value := sG_GetPaperEmpsName;
end;

procedure TrptReport2.QRLabel14Print(sender: TObject; var Value: String);
begin
    Value := sG_Operator;
end;

procedure TrptReport2.QRLabel15Print(sender: TObject; var Value: String);
begin
    Value := TUdateTime.CDateStr(date,9);
end;

procedure TrptReport2.QRLabel21Print(sender: TObject; var Value: String);
var
    nL_Year,nL_Month, nL_Day : word;
begin
    DecodeDate(self.DataSet.FieldByName('GETPAPERDATE').AsDateTime, nL_Year, nL_Month, nL_Day);
    Value := Format('%3d-%2d-%2d', [nL_Year-1911, nL_Month, nL_Day]);
end;

procedure TrptReport2.QRLabel19Print(sender: TObject; var Value: String);
var
    dL_GetPaperDate : TDate;
    sL_EmpNo, sL_GetPaperDate : String;
begin
{
    sL_EmpNo := self.DataSet.fieldByName('EMPNO').AsString;
    sL_GetPaperDate := self.DataSet.fieldByName('GetPaperDate').AsString;
    dL_GetPaperDate := TUdateTime.CDate2Date(sL_GetPaperDate);

    sL_GetPaperDate := TUdateTime.GetPureDateStr(dL_GetPaperDate);
    sL_GetPaperDate := TUstr.replaceStr(sL_GetPaperDate,'/','');
}    
    {
    sL_GetPaperDate := self.DataSet.fieldByName('GetPaperDate').AsString;

    dL_TmpDate := TUdateTime.CDate2Date(Label1.Caption);
    Label2.Caption := DateToStr(dL_TmpDate);
    }

//    dtmMain.getReport2PaperCounts(sL_EmpNo, dL_GetPaperDate);

    dL_GetPaperDate := TUdateTime.CDate2Date(sG_ActiveGetPaperDate);

    sL_GetPaperDate := TUdateTime.GetPureDateStr(dL_GetPaperDate);
    sL_GetPaperDate := TUstr.replaceStr(sL_GetPaperDate,'/','');
    Value := IntToStr(dtmMain.getReport2PaperCounts(sG_ActiveEmpNo, sL_GetPaperDate));

end;

procedure TrptReport2.QuickRepPreview(Sender: TObject);
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

procedure TrptReport2.QRDBText1Print(sender: TObject; var Value: String);
begin
    sG_ActiveEmpNo := Value;
end;

procedure TrptReport2.QRDBText3Print(sender: TObject; var Value: String);
begin
    sG_ActiveGetPaperDate := Value;
end;

end.

