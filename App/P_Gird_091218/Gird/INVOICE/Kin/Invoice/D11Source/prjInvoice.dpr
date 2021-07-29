program prjInvoice;

uses
  Forms,
  dtmMainU in 'dtmMainU.pas' {dtmMain: TDataModule},
  dtmMainJU in 'dtmMainJU.pas' {dtmMainJ: TDataModule},
  dtmMainHU in 'dtmMainHU.pas' {dtmMainH: TDataModule},
  dtmSOU in 'dtmSOU.pas' {dtmSO: TDataModule},
  dtmReportModule in 'dtmReportModule.pas' {dtmReport: TDataModule},
  frmLoginU in 'frmLoginU.pas' {frmLogin},
  frmMainU in 'frmMainU.pas' {frmMain},
  Encryption_TLB in '..\Common\Encryption_TLB.pas',
  fraYMDU in '..\Common\fraYMDU.pas' {fraYMD: TFrame},
  fraYMU in '..\Common\fraYMU.pas' {fraYM: TFrame},
  XLSFile in '..\Common\XLSFile.pas',
  xmlU in '..\Common\xmlU.pas',
  Ustru in '..\Common\Ustru.pas',
  UdateTimeu in '..\Common\UdateTimeu.pas',
  UotherU in '..\Common\UotherU.pas',
  frmInvA01U in 'frmInvA01U.pas' {frmInvA01},
  frmInvA01_1U in 'frmInvA01_1U.pas' {frmInvA01_1},
  frmInvA02_1U in 'frmInvA02_1U.pas' {frmInvA02_1},
  frmInvA02U in 'frmInvA02U.pas' {frmInvA02},
  frmInvA02_4U in 'frmInvA02_4U.pas' {frmInvA02_4},
  frmInvA02_5U in 'frmInvA02_5U.pas' {frmInvA02_5},
  frmInvA02_6U in 'frmInvA02_6U.pas' {frmInvA02_6},
  frmInvA04U in 'frmInvA04U.pas' {frmInvA04},
  frmInvA05U in 'frmInvA05U.pas' {frmInvA05},
  frmInvA05_1U in 'frmInvA05_1U.pas' {frmInvA05_1},
  frmInvA05_2U in 'frmInvA05_2U.pas' {frmInvA05_2},
  frmInvA05_3U in 'frmInvA05_3U.pas' {frmInvA05_3},
  frmInvA06U in 'frmInvA06U.pas' {frmInvA06},
  frmInvA06_1U in 'frmInvA06_1U.pas' {frmInvA06_1},
  frmInvA07_1U in 'frmInvA07_1U.pas' {frmInvA07_1},
  frmInvA07_2U in 'frmInvA07_2U.pas' {frmInvA07_2},
  frmInvA08U in 'frmInvA08U.pas' {frmInvA08},
  frmInvA09U in 'frmInvA09U.pas' {frmInvA09},
  frmInvA09_1U in 'frmInvA09_1U.pas' {frmInvA09_1},
  frmInvA10U in 'frmInvA10U.pas' {frmInvA10},
  frmInvA10_1U in 'frmInvA10_1U.pas' {frmInvA10_1},
  frmInvA11U in 'frmInvA11U.pas' {frmInvA11},
  frmInvB01U in 'frmInvB01U.pas' {frmInvB01},
  frmInvB02U in 'frmInvB02U.pas' {frmInvB02},
  frmInvB03U in 'frmInvB03U.pas' {frmInvB03},
  frmInvB04U in 'frmInvB04U.pas' {frmInvB04},
  frmInvB05U in 'frmInvB05U.pas' {frmInvB05},
  frmInvC01U in 'frmInvC01U.pas' {frmInvC01},
  frmInvC02U in 'frmInvC02U.pas' {frmInvC02},
  frmInvC03U in 'frmInvC03U.pas' {frmInvC03},
  frmInvC04U in 'frmInvC04U.pas' {frmInvC04},
  frmInvC05U in 'frmInvC05U.pas' {frmInvC05},
  frmInvD01U in 'frmInvD01U.pas' {frmInvD01},
  frmInvD02U in 'frmInvD02U.pas' {frmInvD02},
  frmInvD02_1U in 'frmInvD02_1U.pas' {frmInvD02_1},
  frmInvD02_2U in 'frmInvD02_2U.pas' {frmInvD02_2},
  frmInvD03U in 'frmInvD03U.pas' {frmInvD03},
  frmInvD04U in 'frmInvD04U.pas' {frmInvD04},
  frmInvD05U in 'frmInvD05U.pas' {frmInvD05},
  frmInvD06U in 'frmInvD06U.pas' {frmInvD06},
  frmInvD07U in 'frmInvD07U.pas' {frmInvD07},
  frmInvD07_1U in 'frmInvD07_1U.pas' {frmInvD07_1},
  frmInvD08U in 'frmInvD08U.pas' {frmInvD08},
  frmInvE01U in 'frmInvE01U.pas' {frmInvE01},
  frmInvE02U in 'frmInvE02U.pas' {frmInvE02},
  frmInvE03U in 'frmInvE03U.pas' {frmInvE03},
  frmInvE04U in 'frmInvE04U.pas' {frmInvE04},
  rptInvC02_1U in 'rptInvC02_1U.pas' {rptInvC02_1: TQuickRep},
  rptInvC03_1U in 'rptInvC03_1U.pas' {rptInvC03_1: TQuickRep},
  rptInvA07U in 'rptInvA07U.pas' {rptInvA07: TQuickRep},
  frmSearchCustU in 'frmSearchCustU.pas' {frmSearchCust},
  frmRptPreviewU in 'frmRptPreviewU.pas' {frmRptPreView},
  frmMultiSelectU in 'frmMultiSelectU.pas' {frmMultiSelect},
  cbAppClass in 'cbAppClass.pas',
  cbDataLookup in 'cbDataLookup.pas' {DataLookup: TFrame};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TdtmMain, dtmMain);
  Application.CreateForm(TfrmMain, frmMain);
  Application.CreateForm(TdtmMainJ, dtmMainJ);
  Application.CreateForm(TdtmMainH, dtmMainH);
  Application.CreateForm(TdtmReport, dtmReport);
  Application.CreateForm(TdtmSO, dtmSO);
  Application.Run;
end.
