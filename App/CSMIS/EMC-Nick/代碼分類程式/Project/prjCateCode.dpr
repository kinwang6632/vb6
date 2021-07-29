program prjCateCode;

uses
  Forms,
  frmMainU in 'frmMainU.pas' {frmMain},
  frmLoginU in 'frmLoginU.pas' {frmLogin},
  dtmMainU in 'dtmMainU.pas' {dtmMain: TDataModule},
  frmCateCode1U in 'frmCateCode1U.pas' {frmCateCode1},
  UdateTimeu in 'Common\UdateTimeu.pas',
  Ustru in 'Common\Ustru.pas',
  frmCateCode2U in 'frmCateCode2U.pas' {frmCateCode2},
  frmReportMainU in 'frmReportMainU.pas' {frmReportMain},
  UCommonU in 'Common\UCommonU.pas',
  UObjectu in 'Common\UObjectu.pas',
  rptShowDataU in 'rptShowDataU.pas' {rptShowData: TQuickRep},
  frmRptPreviewU in 'Common\frmRptPreviewU.pas' {frmRptPreView};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TdtmMain, dtmMain);
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
