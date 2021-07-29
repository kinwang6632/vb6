unit rptSO8B20B_2U;

interface

uses Windows, SysUtils, Messages, Classes, Graphics, Controls,
  StdCtrls, ExtCtrls, Forms, QuickRpt, QRCtrls,Dialogs,Math;

type
  TrptSO8B20B_2 = class(TQuickRep)
    DetailBand1: TQRBand;
    PageFooterBand1: TQRBand;
    PageHeaderBand1: TQRBand;
    QRGroup1: TQRGroup;
    QRBand1: TQRBand;
    QRLabel1: TQRLabel;
    QRLabel2: TQRLabel;
    QRLabel4: TQRLabel;
    QRLabel18: TQRLabel;
    QRLabel24: TQRLabel;
    QRLabel28: TQRLabel;
    QRLabel30: TQRLabel;
    QRLabel31: TQRLabel;
    QRLabel37: TQRLabel;
    QRLabel38: TQRLabel;
    QRLabel9: TQRLabel;
    QRLabel11: TQRLabel;
    QRLabel40: TQRLabel;
    QRLabel3: TQRLabel;
    QRLabel6: TQRLabel;
    QRLabel7: TQRLabel;
    QRLabel8: TQRLabel;
    QRLabel19: TQRLabel;
    QRLabel34: TQRLabel;
    QRLabel36: TQRLabel;
    QRShape1: TQRShape;
    QRLabel26: TQRLabel;
    QRLabel29: TQRLabel;
    QRLabel32: TQRLabel;
    QRLabel33: TQRLabel;
    QRLabel10: TQRLabel;
    QRLabel12: TQRLabel;
    QRLabel25: TQRLabel;
    QRLabel41: TQRLabel;
    QRLabel13: TQRLabel;
    QRLabel14: TQRLabel;
    QRLabel20: TQRLabel;
    QRLabel21: TQRLabel;
    QRLabel23: TQRLabel;
    QRLabel35: TQRLabel;
    QRLabel42: TQRLabel;
    QRLabel43: TQRLabel;
    QRLabel44: TQRLabel;
    QRShape2: TQRShape;
    SummaryBand1: TQRBand;
    QRLabel22: TQRLabel;
    QRLabel16: TQRLabel;
    QRLabel17: TQRLabel;
    QRLabel27: TQRLabel;
    QRShape3: TQRShape;
    QRLabel39: TQRLabel;
    procedure QuickRepPreview(Sender: TObject);
    procedure QRLabel1Print(sender: TObject; var Value: String);
    procedure QRLabel2Print(sender: TObject; var Value: String);
    procedure QRLabel3Print(sender: TObject; var Value: String);
    procedure QRLabel37Print(sender: TObject; var Value: String);
    procedure QRLabel36Print(sender: TObject; var Value: String);
    procedure QRLabel26Print(sender: TObject; var Value: String);
    procedure QRLabel29Print(sender: TObject; var Value: String);
    procedure QRLabel32Print(sender: TObject; var Value: String);
    procedure QRLabel33Print(sender: TObject; var Value: String);
    procedure QRLabel12Print(sender: TObject; var Value: String);
    procedure QRLabel10Print(sender: TObject; var Value: String);
    procedure QRLabel25Print(sender: TObject; var Value: String);
    procedure QRLabel41Print(sender: TObject; var Value: String);
    procedure QRLabel39Print(sender: TObject; var Value: String);
    procedure QRLabel20Print(sender: TObject; var Value: String);
    procedure QRLabel21Print(sender: TObject; var Value: String);
    procedure QRLabel13Print(sender: TObject; var Value: String);
    procedure QRLabel14Print(sender: TObject; var Value: String);
    procedure QRLabel35Print(sender: TObject; var Value: String);
    procedure QRLabel23Print(sender: TObject; var Value: String);
    procedure QuickRepBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
    procedure QRGroup1BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
    procedure QRLabel42Print(sender: TObject; var Value: String);
    procedure QRLabel43Print(sender: TObject; var Value: String);
    procedure QRLabel44Print(sender: TObject; var Value: String);
    procedure QRLabel22Print(sender: TObject; var Value: String);
    procedure QRLabel16Print(sender: TObject; var Value: String);
    procedure QRLabel17Print(sender: TObject; var Value: String);
    procedure QRLabel38Print(sender: TObject; var Value: String);
    procedure DetailBand1BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
    procedure QRBand1BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
  private
    sG_UpdTime : String;
    fG_WorkerComm,fG_GrpTotWorkerCommAmt,fG_TotWorkerCommAmt : Double;
    fG_IntAccCom,fG_GrpTotIntAccCommAmt,fG_TotIntAccCommAmt : Double;
    fL_WorkClcIntAccCom,fG_GrpTotWorClInAccCom,fG_TotWorClInAccCom : Double;
  public
    sG_EmpNoSQL,sG_SalesCodeSQL,sG_OtherCompCodeSQL,sG_PayType : String;
    sG_Compute, sG_BelongYM ,sG_OrderList ,sG_OrderNameList : String;
    bG_IncludeWorkerComm : Boolean;  

  end;

var
  rptSO8B20B_2: TrptSO8B20B_2;

implementation

uses dtmMain1U, frmRptPreviewU, Ustru, frmMainMenuU, UCommonU;

{$R *.DFM}

procedure TrptSO8B20B_2.QuickRepPreview(Sender: TObject);
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
    bL_HavaData := dtmMain1.getRptBoxDetailData(sG_BelongYM,sG_Compute,sG_EmpNoSQL,sG_SalesCodeSQL,sG_OtherCompCodeSQL,sG_OrderList);

    //回復原狀態
    TUCommonFun.setDefaultCursor;
    
    if bL_HavaData then
    begin
      sG_UpdTime := dtmMain1.cdsSo122.FieldByName('UPDTIME').AsString;

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

procedure TrptSO8B20B_2.QRLabel1Print(sender: TObject; var Value: String);
begin
    Value := 'BOX 獎金明細表';
end;

procedure TrptSO8B20B_2.QRLabel2Print(sender: TObject; var Value: String);
begin
//    Value :='歸屬年月: '+IntToStr(StrToInt(Copy(sG_BelongYM,1,3)))+'年'+Copy(sG_BelongYM,4,2)+'月';
Value :='歸屬年月: '+IntToStr(StrToInt(Copy(sG_BelongYM,1,3))+1911-1911)+'年'+Copy(sG_BelongYM,4,2)+'月';
end;

procedure TrptSO8B20B_2.QRLabel3Print(sender: TObject; var Value: String);
begin
    if bG_IncludeWorkerComm then
      Value := '公司別: ' + frmMainMenu.sG_CompName + ' (包含BOX工程人員獎金)'
    else
      Value := '公司別: ' + frmMainMenu.sG_CompName + ' (不包含BOX工程人員獎金)';
end;

procedure TrptSO8B20B_2.QRLabel37Print(sender: TObject; var Value: String);
begin
    Value := '排序方式: ' + sG_OrderNameList;
end;

procedure TrptSO8B20B_2.QRLabel36Print(sender: TObject; var Value: String);
begin
    Value := '佣獎金計算時間: ' + sG_UpdTime;
end;

procedure TrptSO8B20B_2.QRLabel26Print(sender: TObject; var Value: String);
begin
    value := self.DataSet.FieldByName('StbNo').AsString;
end;

procedure TrptSO8B20B_2.QRLabel29Print(sender: TObject; var Value: String);
begin
    value := self.DataSet.FieldByName('Sno').AsString;
end;

procedure TrptSO8B20B_2.QRLabel32Print(sender: TObject; var Value: String);
begin
    if self.DataSet.FieldByName('RealStartDate').AsString <> '' then
      value := FormatDateTime( 'eee/mm/dd', self.DataSet.FieldByName('RealStartDate').AsDateTime)
    else
      value := '';
end;

procedure TrptSO8B20B_2.QRLabel33Print(sender: TObject; var Value: String);
begin
    if self.DataSet.FieldByName('RealStopDate').AsString <> '' then
      value := FormatDateTime( 'eee/mm/dd', self.DataSet.FieldByName('RealStopDate').AsDateTime)
    else
      value := '';
end;

procedure TrptSO8B20B_2.QRLabel12Print(sender: TObject; var Value: String);
begin
    if self.DataSet.FieldByName('ShouldDate').AsString <> '' then
      value := FormatDateTime( 'eee/mm/dd', self.DataSet.FieldByName('ShouldDate').AsDateTime)
    else
      value := '';
end;

procedure TrptSO8B20B_2.QRLabel10Print(sender: TObject; var Value: String);
begin
    if self.DataSet.FieldByName('RealDate').AsString <> '' then
      value := FormatDateTime( 'eee/mm/dd', self.DataSet.FieldByName('RealDate').AsDateTime)
    else
      value := '';
end;

procedure TrptSO8B20B_2.QRLabel25Print(sender: TObject; var Value: String);
begin
    value := self.DataSet.FieldByName('CustID').AsString;
end;

procedure TrptSO8B20B_2.QRLabel41Print(sender: TObject; var Value: String);
var
    nL_BuyOrRent : Integer;
begin
    nL_BuyOrRent := self.DataSet.FieldByName('BuyOrRent').AsInteger;
    if nL_BuyOrRent = 1 then
      value := '買'
    else if nL_BuyOrRent = 2 then
      value := '租';
end;

procedure TrptSO8B20B_2.QRLabel39Print(sender: TObject; var Value: String);
var
    fL_RealAmt : Double;
begin
    //先不用列
    Value := '';
{
    fL_RealAmt := self.DataSet.FieldByName('RealAmt').AsFloat;

    if fL_RealAmt = 0 then
      Value := ''
    else
      value := FloatToStr(fL_RealAmt);
}
end;

procedure TrptSO8B20B_2.QRLabel20Print(sender: TObject; var Value: String);
begin
    //有勾選包含BOX工程人員獎金
    if bG_IncludeWorkerComm then
      value := self.DataSet.FieldByName('WORKEREN1NAME').AsString
    else
      value := '';
end;

procedure TrptSO8B20B_2.QRLabel21Print(sender: TObject; var Value: String);
begin
    //有勾選包含BOX工程人員獎金
    if bG_IncludeWorkerComm then
      fG_WorkerComm := self.DataSet.FieldByName('WORKEREN1COMM').AsFloat
    else
      fG_WorkerComm := 0;

    Value:=TUstr.CommaNumber(FloatToStr(RoundTo(fG_WorkerComm,-2)));

    fG_GrpTotWorkerCommAmt := fG_GrpTotWorkerCommAmt + fG_WorkerComm;
    fG_TotWorkerCommAmt := fG_TotWorkerCommAmt + fG_WorkerComm;


end;

procedure TrptSO8B20B_2.QRLabel13Print(sender: TObject; var Value: String);
begin
    value:=self.DataSet.FieldByName('IntAccName').AsString
end;

procedure TrptSO8B20B_2.QRLabel14Print(sender: TObject; var Value: String);
begin
    fG_IntAccCom := (self.DataSet.FieldByName('IntAccComm').AsFloat);
    value := TUstr.CommaNumber(FloatToStr(RoundTo(fG_IntAccCom,-2)));

    fG_GrpTotIntAccCommAmt := fG_GrpTotIntAccCommAmt + fG_IntAccCom;
    fG_TotIntAccCommAmt := fG_TotIntAccCommAmt + fG_IntAccCom;
end;

procedure TrptSO8B20B_2.QRLabel35Print(sender: TObject; var Value: String);
begin
    value := self.DataSet.FieldByName('MediaName').AsString;
end;

procedure TrptSO8B20B_2.QRLabel23Print(sender: TObject; var Value: String);
begin
    fL_WorkClcIntAccCom := (fG_WorkerComm + fG_IntAccCom);
    value:=TUstr.CommaNumber(FloatToStr(RoundTo(fL_WorkClcIntAccCom,-2)));

    fG_GrpTotWorClInAccCom := fG_GrpTotWorClInAccCom + fL_WorkClcIntAccCom;
    fG_TotWorClInAccCom := fG_TotWorClInAccCom + fL_WorkClcIntAccCom;
end;

procedure TrptSO8B20B_2.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
    //列印前先歸零
    fG_TotWorkerCommAmt := 0 ;
    fG_TotIntAccCommAmt := 0 ;
    fG_TotWorClInAccCom := 0 ;
end;

procedure TrptSO8B20B_2.QRGroup1BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
begin
    fG_GrpTotWorkerCommAmt := 0;
    fG_GrpTotIntAccCommAmt := 0;
    fG_GrpTotWorClInAccCom := 0;
end;

procedure TrptSO8B20B_2.QRLabel42Print(sender: TObject; var Value: String);
begin
    //有勾選包含BOX工程人員獎金
    if bG_IncludeWorkerComm then
      value:=FloatToStr(RoundTo(fG_GrpTotWorkerCommAmt,-2))
    else
      value := '0';
end;

procedure TrptSO8B20B_2.QRLabel43Print(sender: TObject; var Value: String);
begin
    Value :=TUstr.CommaNumber(FloatToStr(RoundTo(fG_GrpTotIntAccCommAmt,-2)));
end;

procedure TrptSO8B20B_2.QRLabel44Print(sender: TObject; var Value: String);
begin
    value:=TUstr.CommaNumber(FloatToStr(RoundTo(fG_GrpTotWorClInAccCom,-2)));
end;

procedure TrptSO8B20B_2.QRLabel22Print(sender: TObject; var Value: String);
begin
    //有勾選包含BOX工程人員獎金
    if bG_IncludeWorkerComm then
      value:=FloatToStr(RoundTo(fG_TotWorkerCommAmt,-2))
    else
      value := '0';
end;

procedure TrptSO8B20B_2.QRLabel16Print(sender: TObject; var Value: String);
begin
    Value :=TUstr.CommaNumber(FloatToStr(RoundTo(fG_TotIntAccCommAmt,-2)));
end;

procedure TrptSO8B20B_2.QRLabel17Print(sender: TObject; var Value: String);
begin
    value:=TUstr.CommaNumber(FloatToStr(RoundTo(fG_TotWorClInAccCom,-2)));
end;

procedure TrptSO8B20B_2.QRLabel38Print(sender: TObject; var Value: String);
begin
    //先不用列
    Value := '';
end;

procedure TrptSO8B20B_2.DetailBand1BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
begin
    //有勾選包含BOX工程人員獎金
    if bG_IncludeWorkerComm then
      fG_WorkerComm := self.DataSet.FieldByName('WORKEREN1COMM').AsFloat
    else
      fG_WorkerComm := 0;

    //若工程人員及介紹人皆為0,則該筆資料不顯示出來
    fG_IntAccCom := (self.DataSet.FieldByName('IntAccComm').AsFloat);
    if (fG_IntAccCom=0) and (fG_WorkerComm=0) then
      PrintBand := false
    else
      PrintBand := true;
end;

procedure TrptSO8B20B_2.QRBand1BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
begin
    //若工程人員及介紹人皆為0,則該筆資料不顯示出來
    if (fG_GrpTotWorkerCommAmt=0) and (fG_GrpTotIntAccCommAmt=0) then
      PrintBand := false
    else
      PrintBand := true;
end;

end.
