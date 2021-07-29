unit rptInvC02_1U;

interface

uses Windows, SysUtils, Messages, Classes, Graphics, Controls,
  StdCtrls, ExtCtrls, Forms, QuickRpt, QRCtrls;

type
  TrptInvC02_1 = class(TQuickRep)
    DetailBand1: TQRBand;
    PageHeaderBand1: TQRBand;
    PageFooterBand1: TQRBand;
    QRLabel10: TQRLabel;
    qrlRptName: TQRLabel;
    qrlPage1: TQRLabel;
    qrlPage2: TQRSysData;
    qrlOperator1: TQRLabel;
    qrlOperator: TQRLabel;
    qrlPrintTime: TQRLabel;
    qrlPrintTime1: TQRLabel;
    QRBand1: TQRBand;
    lQRLabel2: TQRLabel;
    lblCondition1: TQRLabel;
    lblCondition2: TQRLabel;
    QRBand2: TQRBand;
    QRLabel1: TQRLabel;
    QRLabel2: TQRLabel;
    QRLabel3: TQRLabel;
    QRLabel4: TQRLabel;
    QRLabel5: TQRLabel;
    QRLabel6: TQRLabel;
    QRShape1: TQRShape;
    QRShape2: TQRShape;
    QRLabel7: TQRLabel;
    QRLabel8: TQRLabel;
    QRLabel9: TQRLabel;
    QRLabel11: TQRLabel;
    QRLabel12: TQRLabel;
    QRLabel13: TQRLabel;
    procedure qrlOperatorPrint(sender: TObject; var Value: String);
    procedure qrlPrintTimePrint(sender: TObject; var Value: String);
    procedure QuickRepBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
    procedure lblCondition1Print(sender: TObject; var Value: String);
    procedure lblCondition2Print(sender: TObject; var Value: String);
    procedure QRLabel7Print(sender: TObject; var Value: String);
    procedure QRLabel8Print(sender: TObject; var Value: String);
    procedure QRLabel9Print(sender: TObject; var Value: String);
    procedure QRLabel11Print(sender: TObject; var Value: String);
    procedure QRLabel12Print(sender: TObject; var Value: String);
    procedure QRLabel13Print(sender: TObject; var Value: String);
  private
    sG_PrintTime : String;
  public
    sG_CompID,sG_CompName,sG_SInvID,sG_EInvID,sG_SInvDate,sG_EInvDate : String;
    sG_ObsoleteReason,sG_UserID : String;

  end;

var
  rptInvC02_1: TrptInvC02_1;

implementation

uses dtmMainJU, Ustru;

{$R *.DFM}

procedure TrptInvC02_1.qrlOperatorPrint(sender: TObject;
  var Value: String);
begin
    Value := sG_UserID;
end;

procedure TrptInvC02_1.qrlPrintTimePrint(sender: TObject;
  var Value: String);
begin
    Value := sG_PrintTime;
end;

procedure TrptInvC02_1.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
    sG_PrintTime := DateTimeToStr(now);
end;

procedure TrptInvC02_1.lblCondition1Print(sender: TObject;
  var Value: String);
var
    sL_ConditionStr : String;
begin
    if sG_SInvID = '' then
      sL_ConditionStr := '公司簡稱 : ' + sG_CompName + ' , 發票號碼 : 全部'
    else
      sL_ConditionStr := '公司簡稱 : ' + sG_CompName + ' , 發票號碼 : '  + sG_SInvID + ' ~ ' + sG_EInvID;

    Value := sL_ConditionStr;
end;

procedure TrptInvC02_1.lblCondition2Print(sender: TObject;
  var Value: String);
var
    sL_ConditionStr : String;
begin
    if sG_SInvDate = '' then
      sL_ConditionStr := '作廢原因 : ' + sG_ObsoleteReason + ' , 發票日期: 全部'
    else
      sL_ConditionStr := '作廢原因 : ' + sG_ObsoleteReason + ' , 發票日期: ' + sG_SInvDate + ' ~ ' + sG_EInvDate;

    Value := sL_ConditionStr;

end;

procedure TrptInvC02_1.QRLabel7Print(sender: TObject; var Value: String);
begin

    Value := self.DataSet.FieldByName('InvID').AsString;
end;

procedure TrptInvC02_1.QRLabel8Print(sender: TObject; var Value: String);
begin
    Value := self.DataSet.FieldByName('InvDate').AsString;
end;

procedure TrptInvC02_1.QRLabel9Print(sender: TObject; var Value: String);
begin
    Value := TUstr.CommaNumber(self.DataSet.FieldByName('InvAmount').AsString);
end;

procedure TrptInvC02_1.QRLabel11Print(sender: TObject; var Value: String);
begin
    Value := self.DataSet.FieldByName('ObsoleteReason').AsString;
end;


procedure TrptInvC02_1.QRLabel12Print(sender: TObject; var Value: String);
begin
    Value := self.DataSet.FieldByName('CustID').AsString;
end;

procedure TrptInvC02_1.QRLabel13Print(sender: TObject; var Value: String);
begin
    Value := self.DataSet.FieldByName('HowToCreate').AsString;
end;

end.
