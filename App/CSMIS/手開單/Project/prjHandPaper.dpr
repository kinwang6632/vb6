program prjHandPaper;

uses
  Forms,
  frmSo18C0U in 'frmSo18C0U.pas' {frmSo18C0},
  dtmMainU in 'dtmMainU.pas' {dtmMain: TDataModule},
  frmSo18C1U in 'frmSo18C1U.pas' {frmSo18C1},
  fraChineseYMDU in 'Common\fraChineseYMDU.pas' {fraChineseYMD: TFrame},
  Ustru in 'Common\Ustru.pas',
  UdateTimeu in 'Common\UdateTimeu.pas',
  frmDbMultiSelectu in 'Common\frmDbMultiSelectu.pas' {frmDbMultiSelect},
  UdbU in 'Common\UdbU.pas',
  frmDbSelectu in 'Common\frmDbSelectu.pas' {frmDbSelect},
  frmSo18C2U in 'frmSo18C2U.pas' {frmSo18C2},
  frmSo18C3U in 'frmSo18C3U.pas' {frmSo18C3},
  rptReport2U in 'rptReport2U.pas' {rptReport2: TQuickRep},
  rptReport4U in 'rptReport4U.pas' {rptReport4: TQuickRep},
  rptReport1U in 'rptReport1U.pas' {rptReport1: TQuickRep},
  frmRptPreviewU in 'Common\frmRptPreviewU.pas' {frmRptPreView},
  frmSo18C4U in 'frmSo18C4U.pas' {frmSo18C4},
  frmSo18C5U in 'frmSo18C5U.pas' {frmSo18C5},
  rptReport3U in 'rptReport3U.pas' {rptReport3: TQuickRep},
  frmSo18C6U in 'frmSo18C6U.pas' {frmSo18C6},
  scExcelExport in '..\..\..\CommoUnit\scExcelExport.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TfrmSo18C0, frmSo18C0);
  Application.CreateForm(TdtmMain, dtmMain);
  Application.Run;
end.
