unit rptSO8B20B_3U;

interface

uses Windows, SysUtils, Messages, Classes, Graphics, Controls,
  StdCtrls, ExtCtrls, Forms, QuickRpt, QRCtrls,Math,Dialogs;

type
  TrptSO8B20B_3 = class(TQuickRep)
    DetailBand1: TQRBand;
    PageFooterBand1: TQRBand;
    PageHeaderBand1: TQRBand;
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
    QRLabel16: TQRLabel;
    QRLabel17: TQRLabel;
    QRLabel18: TQRLabel;
    QRLabel19: TQRLabel;
    QRLabel20: TQRLabel;
    QRShape1: TQRShape;
    SummaryBand1: TQRBand;
    QRLabel21: TQRLabel;
    QRLabel22: TQRLabel;
    QRLabel23: TQRLabel;
    QRLabel24: TQRLabel;
    QRLabel25: TQRLabel;
    QRLabel26: TQRLabel;
    QRLabel27: TQRLabel;
    procedure QuickRepPreview(Sender: TObject);
    procedure QRLabel1Print(sender: TObject; var Value: String);
    procedure QRLabel2Print(sender: TObject; var Value: String);
    procedure QRLabel12Print(sender: TObject; var Value: String);
    procedure QRLabel13Print(sender: TObject; var Value: String);
    procedure QRLabel14Print(sender: TObject; var Value: String);
    procedure QRLabel15Print(sender: TObject; var Value: String);
    procedure QRLabel16Print(sender: TObject; var Value: String);
    procedure QRLabel17Print(sender: TObject; var Value: String);
    procedure QRLabel18Print(sender: TObject; var Value: String);
    procedure QRLabel19Print(sender: TObject; var Value: String);
    procedure QRLabel20Print(sender: TObject; var Value: String);
    procedure QRLabel36Print(sender: TObject; var Value: String);
    procedure QRLabel21Print(sender: TObject; var Value: String);
    procedure QRLabel22Print(sender: TObject; var Value: String);
    procedure QRLabel25Print(sender: TObject; var Value: String);
    procedure QRLabel26Print(sender: TObject; var Value: String);
    procedure QRLabel27Print(sender: TObject; var Value: String);
  private
    sG_UpdTime : String;
  public
    sG_PayType,sG_Compute, sG_BelongYM : String;
    sG_CodeNoSQL,sG_ChargeNameSQL : String;
    bG_IncludeWorkerComm : Boolean;

  end;

var
  rptSO8B20B_3: TrptSO8B20B_3;

implementation

uses frmRptPreviewU, dtmMain1U, Ustru, frmMainMenuU, UCommonU;

{$R *.DFM}

procedure TrptSO8B20B_3.QuickRepPreview(Sender: TObject);
var
    frmPreView : TfrmRptPreview;
    bL_HavaData : Boolean;
begin
    if not dtmMain1.cdsStatisticExcel.Active then
      dtmMain1.cdsStatisticExcel.CreateDataSet;
    dtmMain1.cdsStatisticExcel.EmptyDataSet;

    //查出符合條件的資料
    bL_HavaData := dtmMain1.getStatisticData(REPORT_MODE,sG_PayType,frmMainMenu.sG_CompCode,sG_BelongYM,sG_CodeNoSQL,bG_IncludeWorkerComm);

    //回復原狀態
    TUCommonFun.setDefaultCursor;

    if bL_HavaData then
    begin
      sG_UpdTime := dtmMain1.cdsSo134.FieldByName('UPDTIME').AsString;

      frmPreView := TfrmRptPreview.Create(nil);

      frmPreView.G_ActiveDataSet := self.DataSet;
      //frmPreView.bG_LockData := bG_LockData;
      frmPreView.Report := Self;
      frmPreView.QRPreview.QRPrinter := Self.QRPrinter;
      frmPreView.Show;
    end
    else
    begin
      MessageDlg('查無資料',mtInformation, [mbOK],0);
      exit;
    end;

end;

procedure TrptSO8B20B_3.QRLabel1Print(sender: TObject; var Value: String);
begin
    Value := '佣獎金統計表';
end;

procedure TrptSO8B20B_3.QRLabel2Print(sender: TObject; var Value: String);
begin
//    Value :='歸屬年月: '+IntToStr(StrToInt(Copy(sG_BelongYM,1,3)))+'年'+Copy(sG_BelongYM,4,2)+'月';
Value :='歸屬年月: '+IntToStr(StrToInt(Copy(sG_BelongYM,1,3))+1911-1911)+'年'+Copy(sG_BelongYM,4,2)+'月';
end;

procedure TrptSO8B20B_3.QRLabel12Print(sender: TObject; var Value: String);
var
    sL_EmpID : String;
begin
    sL_EmpID := self.DataSet.FieldByName('EmpID').AsString;
    value := sL_EmpID;

    if sL_EmpID = '總計' then
    begin
      QRLabel12.Font.Color := clred;
      QRLabel13.Font.Color := clred;
      QRLabel14.Font.Color := clred;
      QRLabel15.Font.Color := clred;
      QRLabel16.Font.Color := clred;
      QRLabel17.Font.Color := clred;
      QRLabel18.Font.Color := clred;
      QRLabel19.Font.Color := clred;
      QRLabel20.Font.Color := clred;
      QRLabel25.Font.Color := clred;
      QRLabel26.Font.Color := clred;
    end;
end;

procedure TrptSO8B20B_3.QRLabel13Print(sender: TObject; var Value: String);
begin
    value := self.DataSet.FieldByName('EmpName').AsString;
end;

procedure TrptSO8B20B_3.QRLabel14Print(sender: TObject; var Value: String);
begin
    value := self.DataSet.FieldByName('WorkerCounts').AsString;
end;

procedure TrptSO8B20B_3.QRLabel15Print(sender: TObject; var Value: String);
var
    fL_WorkerComm : Double;
begin
    fL_WorkerComm := self.DataSet.FieldByName('WorkerComm').AsFloat;
    value := TUstr.CommaNumber(FloatToStr(RoundTo(fL_WorkerComm,-2)));
end;

procedure TrptSO8B20B_3.QRLabel16Print(sender: TObject; var Value: String);
begin
    value := self.DataSet.FieldByName('ClctCounts').AsString;
end;

procedure TrptSO8B20B_3.QRLabel17Print(sender: TObject; var Value: String);
var
    fL_ClctComm : Double;
begin
    fL_ClctComm := self.DataSet.FieldByName('ClctComm').AsFloat;
    value := TUstr.CommaNumber(FloatToStr(RoundTo(fL_ClctComm,-2)));
end;

procedure TrptSO8B20B_3.QRLabel18Print(sender: TObject; var Value: String);
begin
    value := self.DataSet.FieldByName('BoxIntAccCounts').AsString;
end;

procedure TrptSO8B20B_3.QRLabel19Print(sender: TObject; var Value: String);
var
    fL_BoxIntAccComm : Double;
begin
    fL_BoxIntAccComm := self.DataSet.FieldByName('BoxIntAccComm').AsFloat;
    value := TUstr.CommaNumber(FloatToStr(RoundTo(fL_BoxIntAccComm,-2)));
end;

procedure TrptSO8B20B_3.QRLabel20Print(sender: TObject; var Value: String);
var
    fL_TotalComm : Double;
begin
    fL_TotalComm := self.DataSet.FieldByName('TotalComm').AsFloat;
    value := TUstr.CommaNumber(FloatToStr(RoundTo(fL_TotalComm,-2)));
end;

procedure TrptSO8B20B_3.QRLabel36Print(sender: TObject; var Value: String);
begin
    Value := '佣獎金計算時間: ' + sG_UpdTime;
end;

procedure TrptSO8B20B_3.QRLabel21Print(sender: TObject; var Value: String);
begin
    if bG_IncludeWorkerComm then
      Value := '公司別: ' + frmMainMenu.sG_CompName + ' (包含BOX工程人員獎金)'
    else
      Value := '公司別: ' + frmMainMenu.sG_CompName + ' (不包含BOX工程人員獎金)';
end;

procedure TrptSO8B20B_3.QRLabel22Print(sender: TObject; var Value: String);
var
    sL_PayTypeName : String;
begin
    if sG_PayType = BATCH_PAY then
      sL_PayTypeName := '分期給付'
    else if sG_PayType = ONCE_PAY then
      sL_PayTypeName := '一次給付';

    Value := '佣金計算原則: ' + sL_PayTypeName;

end;

procedure TrptSO8B20B_3.QRLabel25Print(sender: TObject; var Value: String);
begin
    value := self.DataSet.FieldByName('ChIntAccCounts').AsString;
end;

procedure TrptSO8B20B_3.QRLabel26Print(sender: TObject; var Value: String);
var
    fL_ChIntAccComm : Double;
begin
    fL_ChIntAccComm := self.DataSet.FieldByName('ChIntAccComm').AsFloat;
    value := TUstr.CommaNumber(FloatToStr(RoundTo(fL_ChIntAccComm,-2)));
end;

procedure TrptSO8B20B_3.QRLabel27Print(sender: TObject; var Value: String);
begin
  value := '收費項目: ' + sG_ChargeNameSQL;
end;

end.
