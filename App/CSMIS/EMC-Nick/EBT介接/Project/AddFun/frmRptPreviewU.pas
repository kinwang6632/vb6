unit frmRptPreviewU;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QuickRpt, QRPrntr, StdCtrls, ExtCtrls, Buttons, ComCtrls, DBClient,
  scExcelExport, DB;

type
  TfrmRptPreView = class(TForm)
    Panel1: TPanel;
    QRPreview: TQRPreview;
    stbInfo: TStatusBar;
    Panel2: TPanel;
    btnPrint: TBitBtn;
    btnCancel: TBitBtn;
    BitBtn1: TBitBtn;
    PrintDialog1: TPrintDialog;
    btnFirst: TBitBtn;
    btnPrv: TBitBtn;
    btnLast: TBitBtn;
    btnNext: TBitBtn;
    trbZoom: TTrackBar;
    scExcelExport1: TscExcelExport;
    btnToExcel: TBitBtn;
    procedure QRPreviewPageAvailable(Sender: TObject; PageNum: Integer);
    procedure btnFirstClick(Sender: TObject);
    procedure btnPrevClick(Sender: TObject);
    procedure btnNextClick(Sender: TObject);
    procedure btnLastClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure btnPrvClick(Sender: TObject);
    procedure BitBtn4Click(Sender: TObject);
    procedure BitBtn5Click(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure trbZoomChange(Sender: TObject);
    procedure btnToExcelClick(Sender: TObject);
  private
    { Private declarations }
    FReport: TQuickRep;
    procedure SetReport(const Value: TQuickRep);
  public
    { Public declarations }
    G_CDS : TClientDataSet;
    sG_MID : String;
    G_ActiveDataSet : TDataSet;
    property Report: TQuickRep read FReport write SetReport;
  end;

implementation

uses dtmMainU;

{$R *.DFM}

procedure TfrmRptPreView.QRPreviewPageAvailable(Sender: TObject;
  PageNum: Integer);
begin
  stbInfo.Panels[0].Text := IntToStr(PageNum) + '/' +
    IntToStr(QRPreview.QRPrinter.PageCount);


end;

procedure TfrmRptPreView.btnFirstClick(Sender: TObject);
begin
  QRPreview.PageNumber := 1;
  QRPreviewPageAvailable(QRPreview, QRPreview.PageNumber);
end;

procedure TfrmRptPreView.btnPrevClick(Sender: TObject);
begin
  if QRPreview.PageNumber = 1 then Exit;
  QRPreview.PageNumber := QRPreview.PageNumber - 1;
  QRPreviewPageAvailable(QRPreview, QRPreview.PageNumber);
end;

procedure TfrmRptPreView.btnNextClick(Sender: TObject);
begin
  if QRPreview.PageNumber = QRPreview.QRPrinter.PageCount then Exit;
  QRPreview.PageNumber := QRPreview.PageNumber + 1;
  QRPreviewPageAvailable(QRPreview, QRPreview.PageNumber);
end;

procedure TfrmRptPreView.btnLastClick(Sender: TObject);
begin
  QRPreview.PageNumber := QRPreview.QRPrinter.PageCount;
  QRPreviewPageAvailable(QRPreview, QRPreview.PageNumber);
end;

procedure TfrmRptPreView.SetReport(const Value: TQuickRep);
begin
  FReport := Value;
end;

procedure TfrmRptPreView.btnCancelClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmRptPreView.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  //QRPreview.QRPrinter.Free;
  //QRPreview.QRPrinter := nil;
  Action := caFree;
end;

procedure TfrmRptPreView.FormShow(Sender: TObject);

begin
    QRPreview.Zoom := 100;
    trbZoom.Position := 100;

    {
    QRPreview.ZoomToWidth;
    trbZoom.Position := QRPreview.Width;
    }
end;

procedure TfrmRptPreView.BitBtn1Click(Sender: TObject);
begin
  if not PrintDialog1.Execute then Exit;
    {
    QRPreview.QRPrinter.Copies := PrintDialog1.Copies;
    QRPreview.QRPrinter.Duplex := PrintDialog1.Collate;
    FReport.PrinterSettings.ApplySettings(QRPreview.QRPrinter);
    }
end;

procedure TfrmRptPreView.btnPrintClick(Sender: TObject);
var
    ii, nL_Start, nL_Stop : Integer;

begin
    with PrintDialog1 do
    begin
      if PrintRange = prAllPages then
      begin
        nL_Start := MinPage - 1;
        nL_Stop := MaxPage - 1;
      end
      {
      else if PrintRange = prSelection then
      begin
        nL_Start := PageControl1.ActivePage.PageIndex;
        nL_Stop := Start;
      end
      }
      else  { PrintRange = prPageNums }
      begin
        nL_Start := FromPage ;
        nL_Stop := ToPage ;
      end;
    end;

  {
  QRPreview.QRPrinter.FirstPage := nL_Start;
  QRPreview.QRPrinter.LastPage:= nL_Stop;
  QRPreview.QRPrinter.Copies := PrintDialog1.Copies;
  QRPreview.QRPrinter.Duplex := PrintDialog1.Collate;
  FReport.PrinterSettings.ApplySettings(QRPreview.QRPrinter);
  QRPreview.QRPrinter.Print;
  }



    FReport.QRPrinter := QRPreview.QRPrinter;
    FReport.PrinterSettings.Copies := PrintDialog1.Copies;
    FReport.PrinterSettings.Duplex := PrintDialog1.Collate;
    FReport.PrinterSettings.FirstPage := nL_Start;
    FReport.PrinterSettings.LastPage := nL_Stop;
    FReport.Print;

end;

procedure TfrmRptPreView.BitBtn2Click(Sender: TObject);
begin
  QRPreview.PageNumber := 1;
  QRPreviewPageAvailable(QRPreview, QRPreview.PageNumber);
end;

procedure TfrmRptPreView.btnPrvClick(Sender: TObject);
begin
  if QRPreview.PageNumber = 1 then Exit;
  QRPreview.PageNumber := QRPreview.PageNumber - 1;
  QRPreviewPageAvailable(QRPreview, QRPreview.PageNumber);
end;

procedure TfrmRptPreView.BitBtn4Click(Sender: TObject);
begin
  QRPreview.PageNumber := QRPreview.QRPrinter.PageCount;
  QRPreviewPageAvailable(QRPreview, QRPreview.PageNumber);
end;

procedure TfrmRptPreView.BitBtn5Click(Sender: TObject);
begin
  if QRPreview.PageNumber = QRPreview.QRPrinter.PageCount then Exit;
  QRPreview.PageNumber := QRPreview.PageNumber + 1;
  QRPreviewPageAvailable(QRPreview, QRPreview.PageNumber);
end;

procedure TfrmRptPreView.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
    sL_SQL1, sL_SQL2, sL_SQL3, sL_SQL4,  sL_SQL5: String;

begin
    if (Key=VK_NEXT) or (Key=VK_NUMPAD3)then
      BitBtn5Click(Sender)
    else if (Key=VK_PRIOR) or (Key=VK_NUMPAD9)then
      btnPrvClick(Sender)
    else if (Key=VK_END)then
      BitBtn4Click(Sender)
    else if (Key=VK_HOME)then
      BitBtn2Click(Sender)
    else if Key=VK_ESCAPE then
      Close;


end;

procedure TfrmRptPreView.trbZoomChange(Sender: TObject);
begin
  QRPreview.Zoom := trbZoom.Position;
  {
  QRPreview.ZoomToWidth;
  QRPreview.Zoom := 100;
  QRPreview.ZoomToFit;
  }
end;

procedure TfrmRptPreView.btnToExcelClick(Sender: TObject);
begin
    scExcelExport1.Dataset := G_ActiveDataSet;
    scExcelExport1.ExportDataset(true);    
    scExcelExport1.ExcelVisible := true;
end;


end.
