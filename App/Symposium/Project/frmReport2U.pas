unit frmReport2U;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, ComCtrls, ExtCtrls, Quickrpt, QRCtrls, qrextra, QRExport,
  qrprntr;
const
    AGENT_PERFORMANCE_HEADER = 'Agent-Per-';
type
  TfrmReport2 = class(TForm)
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    Label1: TLabel;
    dtpRptSDate: TDateTimePicker;
    rgpComp: TRadioGroup;
    Label2: TLabel;
    Label3: TLabel;
    dtpRptEDate6: TDateTimePicker;
    BitBtn3: TBitBtn;
    rgpDataSrc: TRadioGroup;
    procedure BitBtn1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure BitBtn3Click(Sender: TObject);
  private
    { Private declarations }
    procedure transReportFormat(I_QuickRep : TQuickRep; sI_ReportHeader:String; dI_SDate, dI_EDate:TDate);
  public
    { Public declarations }
  end;

var
  frmReport2: TfrmReport2;

implementation

uses dtmMainU, rptReport1U, rptReport2U, rptReport3U, frmMainU;

{$R *.DFM}

procedure TfrmReport2.BitBtn1Click(Sender: TObject);
begin
    Close;
end;

procedure TfrmReport2.FormShow(Sender: TObject);
begin
    dtpRptSDate.date := date;
    dtpRptEDate6.date := date;
end;

procedure TfrmReport2.BitBtn2Click(Sender: TObject);
var
    nL_CompNdx, nL_DataSrc : Integer;
    bL_HasRecord : boolean;
begin
    nL_CompNdx := rgpComp.ItemIndex;
    nL_DataSrc := rgpDataSrc.ItemIndex ;    
    bL_HasRecord := dtmMain.activeAgentPerformanceStat(nL_DataSrc, nL_CompNdx, dtpRptSDate.date, dtpRptEDate6.date);
    if (bL_HasRecord) then
    begin

      rptReport2 := TrptReport2.Create(Application);
      try
        rptReport2.sG_RptSDate := DateToStr(dtpRptSDate.date);
        rptReport2.sG_RptEDate := DateToStr(dtpRptEDate6.date);
        rptReport2.Exported := True;
        try
          frmMain.transReportFormat(rptReport2, AGENT_PERFORMANCE_HEADER,dtpRptSDate.date, dtpRptEDate6.date);
        finally
          rptReport2.Exported := False;
        end;
        rptReport2.Preview;
      finally
        rptReport2.free;
      end;
    end
    else
      MessageDlg('沒有資料可供列印!', mtInformation,[mbOK], 0);
end;

procedure TfrmReport2.BitBtn3Click(Sender: TObject);
var
    nL_CompNdx : Integer;
    bL_HasRecord : boolean;
begin
    nL_CompNdx := rgpComp.ItemIndex;
    bL_HasRecord := dtmMain.activeAgentBySkillsetPerformance(nL_CompNdx, dtpRptSDate.date, dtpRptEDate6.date);
    if (bL_HasRecord) then
    begin

      rptReport3 := TrptReport3.Create(Application);
//      rptReport3.bG_ShowInvalidDetail := chbShowError.Checked;
      rptReport3.sG_RptSDate := DateToStr(dtpRptSDate.date);
      rptReport3.sG_RptEDate := DateToStr(dtpRptEDate6.date);

//      rptReport3.nG_CompNdx := nL_CompNdx;
      rptReport3.Preview;
      rptReport3.free;

    end
    else
      MessageDlg('沒有資料可供列印!', mtInformation,[mbOK], 0);
end;

procedure TfrmReport2.transReportFormat(I_QuickRep: TQuickRep;
  sI_ReportHeader: String; dI_SDate, dI_EDate: TDate);
var
  AsciiExportFilter : TQRAsciiExportFilter;
  wL_Year, wL_Month, wL_Day : Word;
  sL_SDate, sL_EDate : String;
  sL_RptPath , sL_RptFileName : String;

begin

   decodeDate(dI_SDate,wL_Year, wL_Month, wL_Day);
   sL_SDate := Format('%.4d%.2d%.2d', [ wL_Year, wL_Month, wL_Day ]);

   decodeDate(dI_EDate,wL_Year, wL_Month, wL_Day);
   sL_EDate := Format('%.4d%.2d%.2d', [ wL_Year, wL_Month, wL_Day ]);
    
//   sL_RptPath := 'c:\';
   sL_RptPath := '';
   sL_RptFileName := sL_RptPath + sI_ReportHeader + sL_SDate + '~' + sL_EDate + '.txt';
   AsciiExportFilter := TQRAsciiExportFilter.Create(sL_RptFileName);

   try
      I_QuickRep.ExportToFilter(AsciiExportFilter);
   finally
      AsciiExportFilter.Free;
   end;
    
{
Other filters:
HTML: TQRHTMLDocumentFilter
ASCII: TQRAsciiExportFilter
CSV: TQRCommaSeparatedFilter

In Professional Version:
RTF: TQRRTFExportFilter
WMF: TQRWMFExportFilter
Excel: TQRXLSFilter
}

end;

end.


