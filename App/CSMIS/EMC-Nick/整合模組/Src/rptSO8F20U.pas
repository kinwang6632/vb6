unit rptSO8F20U;

interface

uses Windows, SysUtils, Messages, Classes, Graphics, Controls,
  StdCtrls, ExtCtrls, Forms, QuickRpt, QRCtrls;

type
  TrptSO8F20 = class(TQuickRep)
    DetailBand1: TQRBand;
    PageFooterBand1: TQRBand;
    PageHeaderBand1: TQRBand;
    QRGroup1: TQRGroup;
    QRBand1: TQRBand;
    QRGroup2: TQRGroup;
    QRBand2: TQRBand;
    QRGroup3: TQRGroup;
    QRBand3: TQRBand;
    QRLabel1: TQRLabel;
    QRLabel4: TQRLabel;
    QRLabel5: TQRLabel;
    QRShape1: TQRShape;
    QRLabel8: TQRLabel;
    QRLabel2: TQRLabel;
    QRLabel3: TQRLabel;
    QRLabel6: TQRLabel;
    QRLabel7: TQRLabel;
    procedure QuickRepPreview(Sender: TObject);
    procedure QRLabel1Print(sender: TObject; var Value: String);
    procedure QRLabel2Print(sender: TObject; var Value: String);
    procedure QRLabel3Print(sender: TObject; var Value: String);
    procedure QRLabel4Print(sender: TObject; var Value: String);
    procedure QRLabel5Print(sender: TObject; var Value: String);
  private

  public
    sG_User : String;

  end;

var
  rptSO8F20: TrptSO8F20;

implementation

uses dtmMain4U, frmRptPreviewU, frmSO8F10_2U;

{$R *.DFM}

procedure TrptSO8F20.QuickRepPreview(Sender: TObject);
var
    frmPreView : TfrmRptPreview;
begin

    frmPreView := TfrmRptPreview.Create(nil);

    frmPreView.G_ActiveDataSet := self.DataSet;
    frmPreView.Report := Self;
    frmPreView.QRPreview.QRPrinter := Self.QRPrinter;
    frmPreView.Show;
end;

procedure TrptSO8F20.QRLabel1Print(sender: TObject; var Value: String);
var
    sL_CompCode,sL_CompName : String;
begin
    sL_CompCode := self.DataSet.FieldByName('CompCode').AsString;

    if dtmMain4.cdsCompInfo.Locate('CompCode', sL_CompCode,[]) then
      sL_CompName := dtmMain4.cdsCompInfo.FieldByName('CompName').AsString;

    Value := '公司: [ ' + sL_CompName + ' ]';

end;

procedure TrptSO8F20.QRLabel2Print(sender: TObject; var Value: String);
var
    sL_MasterCodeDesc : String;
begin
    sL_MasterCodeDesc := self.DataSet.FieldByName('Description').AsString;
    Value := '類別資料檔:[ ' + sL_MasterCodeDesc + ' ]';

end;

procedure TrptSO8F20.QRLabel3Print(sender: TObject; var Value: String);
var
    sL_TableName,sL_TableDesc : String;
begin
    sL_TableName := self.DataSet.FieldByName('TableName').AsString;

    if dtmMain4.adoCodeCD067A.Locate('TableName', sL_TableName,[]) then
      sL_TableDesc := dtmMain4.adoCodeCD067A.FieldByName('TableDescription').AsString;

    Value := '分類代碼檔:[ ' + sL_TableName + ' ] - [ ' + sL_TableDesc + ' ]';
end;

procedure TrptSO8F20.QRLabel4Print(sender: TObject; var Value: String);
var
    sL_MasterCodeNo : String;
begin
    sL_MasterCodeNo := self.DataSet.FieldByName('CodeNo').AsString;
    Value := sL_MasterCodeNo;
end;

procedure TrptSO8F20.QRLabel5Print(sender: TObject; var Value: String);
var
    sL_DetailCodeNo : String;
begin
    sL_DetailCodeNo := self.DataSet.FieldByName('DetailCodeNo').AsString;
    Value := sL_DetailCodeNo;
end;

end.
