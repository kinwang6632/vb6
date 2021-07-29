unit rptBasicu;

interface

uses Windows, SysUtils, Messages, Classes, Graphics, Controls,
  StdCtrls, ExtCtrls, Forms, Quickrpt, QRCtrls, DBTables, Db;

type
  TrptBasic = class(TQuickRep)
    QRBand1: TQRBand;
    QRLabel2: TQRLabel;
    QRSysData1: TQRSysData;
    QRBand3: TQRBand;
    QRSysData2: TQRSysData;
    QRLabel3: TQRLabel;
    QRBand2: TQRBand;
    QRLabel1: TQRLabel;
    procedure rptBasicPreview(Sender: TObject);
  private
    FSourceDataSet: TDataSet;
    FTitle: string;
    procedure SetSourceDataSet(const Value: TDataSet);
    procedure SetTitle(const Value: string);
    function CreateQRControl(QRControlClass:TControlClass;
      aParent:TControl) : TControl ;
    procedure SetQRControlPos(QRLabs: array of TQRLabel;
      QRTxts: array of TQRDBText) ;
  public
    property SourceDataSet : TDataSet read FSourceDataSet
      write SetSourceDataSet;
    property Title : string  read FTitle write SetTitle;
  end;

var
  rptBasic: TrptBasic;

implementation

uses frmRptPreviewU;

{$R *.DFM}

{ TrptBasic }

procedure TrptBasic.SetTitle(const Value: string);
begin
  FTitle := Value;
  QRLabel1.Caption := FTitle ;
end;

procedure TrptBasic.SetSourceDataSet(const Value: TDataSet);
var
  ii, wid, pos: Integer ;
  aQRLine: TQRShape;
  aQRLab : TQRLabel ;
  aQRTxt : TQRDBText ;
begin
  FSourceDataSet := Value;
  Self.DataSet := FSourceDataSet ;
  pos := 10;
  for ii := 0 to FSourceDataSet.FieldCount-1 do
  begin
    if (not FSourceDataSet.Fields[ii].Visible) or
      (UpperCase(FSourceDataSet.Fields[ii].FieldName) = 'CUPDEN') or
      (UpperCase(FSourceDataSet.Fields[ii].FieldName) = 'DUPD') then Continue;
    {create QRLabels based on Field's DisplayLabels}
    aQRLine := CreateQRControl(TQRShape, QRBand1) as TQRShape ;
    aQRLine.Shape := qrsHorLine;
    aQRLab := CreateQRControl(TQRLabel, QRBand1) as TQRLabel ;
    aQRLab.Caption := FSourceDataSet.Fields[ii].DisplayLabel ;
    wid := aQRLab.Width;
    {create QRDBTexts based on DataSet's Fields}
    aQRTxt := CreateQRControl(TQRDBText, QRBand2) as TQRDBText ;
    aQRTxt.DataSet := FSourceDataSet ;
    aQRTxt.DataField := FSourceDataSet.Fields[ii].FieldName ;
    if wid < FSourceDataSet.Fields[ii].DisplayWidth * TextWidth(aQRTxt.Font, '1') then
      wid := FSourceDataSet.Fields[ii].DisplayWidth * TextWidth(aQRTxt.Font, '1');
    aQRLab.Left := pos;
    aQRLab.Top := QRBand1.Height - aQRLab.Height - 6;
    aQRTxt.Left := pos;
    aQRTxt.Top := 3;
    aQRLine.Left := pos;
    aQRLine.Top := QRBand1.Height - 3;
    aQRLine.Width := wid;
    aQRLine.Height := 3;
    pos := pos + wid + 3;
  end ;
  {set QRDBTexts and QRLabels' position}
  //SetQRControlPos(aQRLabs, aQRTxts);
end;

function TrptBasic.CreateQRControl(QRControlClass:TControlClass;
  aParent:TControl) : TControl ;
var
  aCtrl : TControl ;
begin
  aCtrl := QRControlClass.Create(nil) ;
  TQRPrintable(aCtrl).ParentReport := Self ;
  TWinControl(aParent).InsertControl(aCtrl);
  aCtrl.Visible := True ;
  Result := aCtrl ;
end ;

procedure TrptBasic.SetQRControlPos(QRLabs: array of TQRLabel;
  QRTxts: array of TQRDBText) ;
var
  ii, tWid, Gap : Integer ;
begin
  {calculate average gap between columns}
  tWid := 0 ;
  for ii := 0 to High(QRTxts) do
    Inc(tWid, QRTxts[ii].Width) ;
  Gap := (QRTxts[0].Parent.Width-tWid) div (High(QRTxts)+2);
  {assign each control's position}
  tWid := Gap ;
  for ii := 0 to High(QRTxts) do
  begin
    QRTxts[ii].Left := tWid ;
    QRTxts[ii].Top := (QRTxts[ii].Parent.Height-QRTxts[ii].Height) div 2 ;
    QRLabs[ii].Left := tWid ;
    QRLabs[ii].Top := QRLabs[ii].Parent.Height-QRLabs[ii].Height-3 ;
    Inc(tWid, QRTxts[ii].Width+Gap) ;
  end ;
end;

procedure TrptBasic.rptBasicPreview(Sender: TObject);
var
  frmPreView: TfrmRptPreview;
begin
  frmPreView := TfrmRptPreview.Create(nil);
  frmPreView.G_ActiveDataSet := self.DataSet;  
  frmPreView.Report := Self;
  frmPreView.QRPreview.QRPrinter := Self.QRPrinter;
  frmPreView.Show;

end;

end.
