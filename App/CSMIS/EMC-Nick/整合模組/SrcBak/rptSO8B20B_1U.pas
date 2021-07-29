unit rptSO8B20B_1U;

interface

uses Windows, SysUtils, Messages, Classes, Graphics, Controls,
  StdCtrls, ExtCtrls, Forms, QuickRpt, QRCtrls ,Math ,Dialogs;

type
  TrptSO8B20B_1 = class(TQuickRep)
    DetailBand1: TQRBand;
    PageFooterBand1: TQRBand;
    PageHeaderBand1: TQRBand;
    QRLabel1: TQRLabel;
    QRLabel2: TQRLabel;
    QRLabel4: TQRLabel;
    QRLabel6: TQRLabel;
    QRLabel7: TQRLabel;
    QRLabel8: TQRLabel;
    QRLabel13: TQRLabel;
    QRLabel14: TQRLabel;
    QRLabel18: TQRLabel;
    QRLabel19: TQRLabel;
    QRLabel20: TQRLabel;
    QRLabel21: TQRLabel;
    QRLabel23: TQRLabel;
    QRLabel24: TQRLabel;
    QRLabel26: TQRLabel;
    SummaryBand1: TQRBand;
    QRLabel22: TQRLabel;
    QRLabel16: TQRLabel;
    QRLabel17: TQRLabel;
    QRShape3: TQRShape;
    QRShape1: TQRShape;
    QRLabel27: TQRLabel;
    QRLabel28: TQRLabel;
    QRLabel29: TQRLabel;
    QRLabel30: TQRLabel;
    QRLabel31: TQRLabel;
    QRLabel32: TQRLabel;
    QRLabel33: TQRLabel;
    QRLabel34: TQRLabel;
    QRLabel35: TQRLabel;
    QRLabel36: TQRLabel;
    QRLabel37: TQRLabel;
    QRLabel38: TQRLabel;
    QRLabel39: TQRLabel;
    QRGroup1: TQRGroup;
    QRBand1: TQRBand;
    QRLabel9: TQRLabel;
    QRLabel10: TQRLabel;
    QRLabel11: TQRLabel;
    QRLabel12: TQRLabel;
    QRLabel25: TQRLabel;
    QRLabel40: TQRLabel;
    QRLabel41: TQRLabel;
    QRLabel42: TQRLabel;
    QRLabel43: TQRLabel;
    QRLabel44: TQRLabel;
    QRShape2: TQRShape;
    QRLabel3: TQRLabel;
    QRLabel5: TQRLabel;
    QRLabel15: TQRLabel;
    procedure QRLabel2Print(sender: TObject; var Value: String);
    procedure QRLabel13Print(sender: TObject; var Value: String);
    procedure QRLabel14Print(sender: TObject; var Value: String);
    procedure QRLabel16Print(sender: TObject; var Value: String);
    procedure QuickRepPreview(Sender: TObject);
    procedure QRLabel17Print(sender: TObject; var Value: String);
    procedure QRLabel21Print(sender: TObject; var Value: String);
    procedure QRLabel20Print(sender: TObject; var Value: String);
    procedure QRLabel22Print(sender: TObject; var Value: String);
    procedure QRLabel23Print(sender: TObject; var Value: String);
    procedure QRLabel26Print(sender: TObject; var Value: String);
    procedure DetailBand1BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
    procedure QuickRepBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
    procedure QRLabel29Print(sender: TObject; var Value: String);
    procedure QRLabel32Print(sender: TObject; var Value: String);
    procedure QRLabel33Print(sender: TObject; var Value: String);
    procedure QRLabel35Print(sender: TObject; var Value: String);
    procedure QRLabel36Print(sender: TObject; var Value: String);
    procedure QRLabel37Print(sender: TObject; var Value: String);
    procedure QRLabel39Print(sender: TObject; var Value: String);
    procedure QRLabel10Print(sender: TObject; var Value: String);
    procedure QRLabel12Print(sender: TObject; var Value: String);
    procedure QRLabel25Print(sender: TObject; var Value: String);
    procedure QRLabel41Print(sender: TObject; var Value: String);
    procedure QRGroup1BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
    procedure QRLabel42Print(sender: TObject; var Value: String);
    procedure QRLabel43Print(sender: TObject; var Value: String);
    procedure QRLabel44Print(sender: TObject; var Value: String);
    procedure QRLabel3Print(sender: TObject; var Value: String);
    procedure QRLabel1Print(sender: TObject; var Value: String);
    procedure QRLabel5Print(sender: TObject; var Value: String);
    procedure QRLabel15Print(sender: TObject; var Value: String);
  private
    fG_TotIntAccCommAmt,fG_TotClctnCommAmt: Double;
    fG_IntAccCom,fG_ClctnComm: Double;
    fG_TotWorClInAccCom:Double;
    sG_UpdTime : String;
    fG_GrpTotClctnCommAmt,fG_GrpTotIntAccCommAmt,fG_GrpTotWorClInAccCom : Double;
  public
    bG_LockData : boolean;
    sG_EmpNoSQL,sG_SalesCodeSQL,sG_OtherCompCodeSQL,sG_PayType : String;
    sG_Compute, sG_BelongYM ,sG_OrderList ,sG_OrderNameList : String;
    sG_CodeNoSQL,sG_ChargeNameSQL : String;

  end;

var
  rptSO8B20B_1: TrptSO8B20B_1;

implementation

uses dtmMain1U, frmRptPreviewU, Ustru, frmMainMenuU, UCommonU;

{$R *.DFM}

procedure TrptSO8B20B_1.QRLabel2Print(sender: TObject; var Value: String);

var sL_ComputeYM : String;
begin
//    Value :='歸屬年月: '+IntToStr(StrToInt(Copy(sG_BelongYM,1,3)))+'年'+Copy(sG_BelongYM,4,2)+'月';
   Value :='歸屬年月: '+IntToStr(StrToInt(Copy(sG_BelongYM,1,3))+1911-1911)+'年'+Copy(sG_BelongYM,4,2)+'月';
end;


procedure TrptSO8B20B_1.QRLabel13Print(sender: TObject; var Value: String);
begin
    value:=self.DataSet.FieldByName('IntAccName').AsString
end;

procedure TrptSO8B20B_1.QRLabel14Print(sender: TObject; var Value: String);
begin
    fG_IntAccCom := (self.DataSet.FieldByName('IntAccComm').AsFloat);
    value := TUstr.CommaNumber(FloatToStr(RoundTo(fG_IntAccCom,-2)));

    fG_GrpTotIntAccCommAmt := fG_GrpTotIntAccCommAmt + fG_IntAccCom;
    fG_TotIntAccCommAmt := fG_TotIntAccCommAmt + fG_IntAccCom;
end;

procedure TrptSO8B20B_1.QRLabel16Print(sender: TObject; var Value: String);
begin
    Value :=TUstr.CommaNumber(FloatToStr(RoundTo(fG_TotIntAccCommAmt,-2)));
end;

procedure TrptSO8B20B_1.QuickRepPreview(Sender: TObject);
var
    frmPreView : TfrmRptPreview;
    L_StrList : TStringList;
    sL_GroupBy : String;
    bL_HavaData : Boolean;
begin
    //指定OrderBY的第一個欄位為報表Group By 欄位
    L_StrList := TStringList.Create;
    L_StrList := TUstr.ParseStrings(sG_OrderList,',');
    sL_GroupBy := L_StrList.Strings[0];
    QRGroup1.Expression := sL_GroupBy;

    //查出符合條件的資料
    bL_HavaData := dtmMain1.getRptChannelDetailData(sG_PayType,sG_BelongYM,sG_Compute,sG_EmpNoSQL,sG_SalesCodeSQL,sG_OtherCompCodeSQL,sG_OrderList,sG_CodeNoSQL);

    //回復原狀態
    TUCommonFun.setDefaultCursor;

    if bL_HavaData then
    begin
      sG_UpdTime := dtmMain1.cdsSo134.FieldByName('UPDTIME').AsString;

      frmPreView := TfrmRptPreview.Create(nil);

      frmPreView.G_ActiveDataSet := self.DataSet;
      frmPreView.bG_LockData := bG_LockData;
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

procedure TrptSO8B20B_1.QRLabel17Print(sender: TObject; var Value: String);
begin
    value:=TUstr.CommaNumber(FloatToStr(RoundTo(fG_TotWorClInAccCom,-2)));
end;

procedure TrptSO8B20B_1.QRLabel21Print(sender: TObject; var Value: String);
begin
    fG_ClctnComm:=self.DataSet.FieldByName('CLCTENCOMM').AsFloat;
    Value:=TUstr.CommaNumber(FloatToStr(RoundTo(fG_ClctnComm,-2)));

    fG_GrpTotClctnCommAmt := fG_GrpTotClctnCommAmt + fG_ClctnComm;
    fG_TotClctnCommAmt := fG_TotClctnCommAmt + fG_ClctnComm;
end;

procedure TrptSO8B20B_1.QRLabel20Print(sender: TObject; var Value: String);
begin
    value:=self.DataSet.FieldByName('CLCTENNAME').AsString;
end;

procedure TrptSO8B20B_1.QRLabel22Print(sender: TObject; var Value: String);
begin
    value:=FloatToStr(RoundTo(fG_TotClctnCommAmt,-2));
end;

procedure TrptSO8B20B_1.QRLabel23Print(sender: TObject; var Value: String);
var fL_WorkClcIntAccCom:double;
begin
    fL_WorkClcIntAccCom := (fG_ClctnComm + fG_IntAccCom);
    value:=TUstr.CommaNumber(FloatToStr(RoundTo(fL_WorkClcIntAccCom,-2)));

    fG_GrpTotWorClInAccCom := fG_GrpTotWorClInAccCom + fL_WorkClcIntAccCom;
    fG_TotWorClInAccCom := fG_TotWorClInAccCom + fL_WorkClcIntAccCom;
end;



procedure TrptSO8B20B_1.QRLabel26Print(sender: TObject; var Value: String);
begin
    value:=self.DataSet.FieldByName('BillNo').AsString;
end;

procedure TrptSO8B20B_1.DetailBand1BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
begin
//    if  self.DataSet.FieldByName('ComputeYM').AsString <> self.DataSet.FieldByName('BelongYM').AsString then
//      PrintBand := false;
    if  sG_BelongYM <> self.DataSet.FieldByName('BelongYM').AsString then
      PrintBand := false;
end;

procedure TrptSO8B20B_1.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
    //列印前先歸零
    fG_TotClctnCommAmt := 0 ;
    fG_TotIntAccCommAmt := 0 ;
    fG_TotWorClInAccCom := 0 ;
end;

procedure TrptSO8B20B_1.QRLabel29Print(sender: TObject; var Value: String);
begin
    value := self.DataSet.FieldByName('CitemName').AsString;
end;

procedure TrptSO8B20B_1.QRLabel32Print(sender: TObject; var Value: String);
var
    sL_StartDate : String;
begin
  if self.DataSet.FieldByName('RealStartDate').AsString <> '' then
    //sL_StartDate := DateToStr(self.DataSet.FieldByName('RealStartDate').AsDateTime);
    sL_StartDate := FormatDateTime( 'eee/mm/dd', self.DataSet.FieldByName('RealStartDate').AsDateTime );
   Value := sL_StartDate;
end;

procedure TrptSO8B20B_1.QRLabel33Print(sender: TObject; var Value: String);
var
    sL_StopDate : String;
begin
    if self.DataSet.FieldByName('RealStopDate').AsString <> '' then
      //sL_StopDate := DateToStr(self.DataSet.FieldByName('RealStopDate').AsDateTime);
    sL_StopDate := FormatDateTime( 'eee/mm/dd', self.DataSet.FieldByName('RealStopDate').AsDateTime);
    value := sL_StopDate;
end;

procedure TrptSO8B20B_1.QRLabel35Print(sender: TObject; var Value: String);
begin
    value := self.DataSet.FieldByName('MediaName').AsString;
end;

procedure TrptSO8B20B_1.QRLabel36Print(sender: TObject; var Value: String);
begin
    Value := '佣獎金計算時間: ' + sG_UpdTime;
end;

procedure TrptSO8B20B_1.QRLabel37Print(sender: TObject; var Value: String);
begin
    Value := '排序方式: ' + sG_OrderNameList;
end;

procedure TrptSO8B20B_1.QRLabel39Print(sender: TObject; var Value: String);
var
    fL_RealAmt : Double;
begin
    fL_RealAmt := self.DataSet.FieldByName('RealAmt').AsFloat;

    if fL_RealAmt = 0 then
      Value := ''
    else
      value:= FloatToStr(fL_RealAmt);
end;

procedure TrptSO8B20B_1.QRLabel10Print(sender: TObject; var Value: String);
begin
    if self.DataSet.FieldByName('RealDate').AsString <> '' then
      value := FormatDateTime( 'eee/mm/dd', self.DataSet.FieldByName('RealDate').AsDateTime);
end;

procedure TrptSO8B20B_1.QRLabel12Print(sender: TObject; var Value: String);
begin
    if self.DataSet.FieldByName('ShouldDate').AsString <> '' then
      value := FormatDateTime( 'eee/mm/dd', self.DataSet.FieldByName('ShouldDate').AsDateTime);
end;

procedure TrptSO8B20B_1.QRLabel25Print(sender: TObject; var Value: String);
begin
    value := Self.DataSet.FieldByName('CustID').AsString;
end;

procedure TrptSO8B20B_1.QRLabel41Print(sender: TObject; var Value: String);
var
    sL_FirstFlag : String;
begin
    sL_FirstFlag := self.DataSet.FieldByName('FirstFlag').AsString;
    if sL_FirstFlag = '1' then
      value := '首次'
    else if sL_FirstFlag = '' then
      value := '續收';

end;

procedure TrptSO8B20B_1.QRGroup1BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
begin
    fG_GrpTotClctnCommAmt := 0;
    fG_GrpTotIntAccCommAmt := 0;
    fG_GrpTotWorClInAccCom := 0;
end;

procedure TrptSO8B20B_1.QRLabel42Print(sender: TObject; var Value: String);
begin
    value:=FloatToStr(RoundTo(fG_GrpTotClctnCommAmt,-2));
end;

procedure TrptSO8B20B_1.QRLabel43Print(sender: TObject; var Value: String);
begin
    Value :=TUstr.CommaNumber(FloatToStr(RoundTo(fG_GrpTotIntAccCommAmt,-2)));
end;

procedure TrptSO8B20B_1.QRLabel44Print(sender: TObject; var Value: String);
begin
    value:=TUstr.CommaNumber(FloatToStr(RoundTo(fG_GrpTotWorClInAccCom,-2)));
end;

procedure TrptSO8B20B_1.QRLabel3Print(sender: TObject; var Value: String);
begin
  Value := '公司別: ' + frmMainMenu.sG_CompName;
end;

procedure TrptSO8B20B_1.QRLabel1Print(sender: TObject; var Value: String);
begin
    Value := '付費頻道佣金明細表';
end;

procedure TrptSO8B20B_1.QRLabel5Print(sender: TObject; var Value: String);
var
    sL_PayTypeName : String;
begin
    if sG_PayType = BATCH_PAY then
      sL_PayTypeName := '分期給付'
    else if sG_PayType = ONCE_PAY then
      sL_PayTypeName := '一次給付';

    Value := '佣金計算原則: ' + sL_PayTypeName;
end;

procedure TrptSO8B20B_1.QRLabel15Print(sender: TObject; var Value: String);
begin
  value := '收費項目: ' + sG_ChargeNameSQL;
end;

end.
