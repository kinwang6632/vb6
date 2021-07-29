program prjCommAP;

uses
  Forms,
  UCommonU in '..\Common\UCommonU.pas',
  UObjectu in '..\Common\UObjectu.pas',
  Ustru in '..\Common\Ustru.pas',
  frmSO8B10U in 'frmSO8B10U.pas' {frmSO8B10},
  frmDbSelectu in '..\Common\frmDbSelectu.pas' {frmDbSelect},
  UdateTimeu in '..\Common\UdateTimeu.pas',
  rptSO8B20AU in 'rptSO8B20AU.pas' {rptSO8B20A: TQuickRep},
  frmRptPreviewU in '..\Common\frmRptPreviewU.pas' {frmRptPreView},
  frmSO8B20U in 'frmSO8B20U.pas' {frmSO8B20},
  fraChineseYMU in '..\Common\fraChineseYMU.pas' {fraChineseYM: TFrame},
  frmSO8B30U in 'frmSO8B30U.pas' {frmSO8B30},
  frmSO8B20_1U in 'frmSO8B20_1U.pas' {frmSO8B20_1},
  frmSO8B40U in 'frmSO8B40U.pas' {frmSO8B40},
  frmDbMultiSelectu in '..\Common\frmDbMultiSelectu.pas' {frmDbMultiSelect},
  UdbU in '..\Common\UdbU.pas',
  frmDualListDlgU in 'frmDualListDlgU.pas' {frmDualListDlg},
  XLSFile in '..\Common\XLSFile.pas',
  frmSO8B50U in 'frmSO8B50U.pas' {frmSO8B50},
  rptSO8B20B_3U in 'rptSO8B20B_3U.pas' {rptSO8B20B_3: TQuickRep},
  rptSO8B20B_2U in 'rptSO8B20B_2U.pas' {rptSO8B20B_2: TQuickRep},
  rptSO8B20B_4U in 'rptSO8B20B_4U.pas' {rptSO8B20B_4: TQuickRep},
  frmMainMenuU in 'frmMainMenuU.pas' {frmMainMenu},
  dtmMain1U in 'dtmMain1U.pas' {dtmMain1: TDataModule},
  fraChineseYMDU in '..\Common\fraChineseYMDU.pas' {fraChineseYMD: TFrame},
  frmSO8A3011U in 'frmSO8A3011U.pas' {frmSO8A3011},
  dtmMain2U in 'dtmMain2U.pas' {dtmMain2: TDataModule},
  frmSO8A60U in 'frmSO8A60U.pas' {frmSO8A60},
  frmSO8A10U in 'frmSO8A10U.pas' {frmSO8A10},
  frmSO8A40U in 'frmSO8A40U.pas' {frmSO8A40},
  frmSO8A20U in 'frmSO8A20U.pas' {frmSO8A20},
  rptSO8A40AU in 'rptSO8A40AU.pas' {rptSO8A40A: TQuickRep},
  rptSO8A40BU in 'rptSO8A40BU.pas' {rptSO8A40B: TQuickRep},
  rptSO8A40CU in 'rptSO8A40CU.pas' {rptSO8A40C: TQuickRep},
  frmSO8A50U in 'frmSO8A50U.pas' {frmSO8A50},
  frmSO8A301U in 'frmSO8A301U.pas' {frmSO8A301},
  frmSO8A30U in 'frmSO8A30U.pas' {frmSO8A30},
  frmLoginU in 'frmLoginU.pas' {frmLogin},
  dtmMain4U in 'dtmMain4U.pas' {dtmMain4: TDataModule},
  frmSO8F10_1U in 'frmSO8F10_1U.pas' {frmSO8F10_1},
  frmSO8F10_2U in 'frmSO8F10_2U.pas' {frmSO8F10_2},
  rptSO8F20U in 'rptSO8F20U.pas' {rptSO8F20: TQuickRep},
  frmSO8F20U in 'frmSO8F20U.pas' {frmSO8F20},
  dtmMain3U in 'dtmMain3U.pas' {dtmMain3: TDataModule},
  frmSO8C10U in 'frmSO8C10U.pas' {frmSO8C10},
  frmSO8C20U in 'frmSO8C20U.pas' {frmSO8C20},
  frmSO8C40_1U in 'frmSO8C40_1U.pas' {frmSO8C40_1},
  rptSO8C30_1U in 'rptSO8C30_1U.pas' {rptSO8C30_1: TQuickRep},
  rptSO8C30_2U in 'rptSO8C30_2U.pas' {rptSO8C30_2: TQuickRep},
  rptSO8C30_3U in 'rptSO8C30_3U.pas' {rptSO8C30_3: TQuickRep},
  frmSO8C30U in 'frmSO8C30U.pas' {frmSO8C30},
  rptSO8C40U in 'rptSO8C40U.pas' {rptSO8C40: TQuickRep},
  frmSO8C40_2U in 'frmSO8C40_2U.pas' {frmSO8C40_2},
  rptSO8B20B_1U in 'rptSO8B20B_1U.pas' {rptSO8B20B_1: TQuickRep},
  frmSO8F10_3U in 'frmSO8F10_3U.pas' {frmSO8F10_3},
  frmSO8F10_4U in 'frmSO8F10_4U.pas' {frmSO8F10_4},
  DevPower_TLB in '..\Common\DevPower_TLB.pas',
  frmSo8C60U in 'frmSo8C60U.pas' {frmSo8C60},
  rptSO8C30_4U in 'rptSO8C30_4U.pas' {rptSO8C30_4: TQuickRep},
  xmlU in '..\Common\xmlU.pas',
  scExcelExport in '..\Common\scExcelExport.pas',
  rptSO8C10U in 'rptSO8C10U.pas' {rptSO8C10: TQuickRep},
  Emc_Separate_TLB in '..\Server\Emc_Separate_TLB.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TfrmMainMenu, frmMainMenu);
  Application.Run;
end.
