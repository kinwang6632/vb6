program prjAddFun;

uses
  Forms,
  frmMainU in 'frmMainU.pas' {frmMain},
  frmGroupDataU in 'frmGroupDataU.pas' {frmGroupData},
  dtmMainU in 'dtmMainU.pas' {dtmMain: TDataModule},
  frmLoginU in 'frmLoginU.pas' {frmLogin},
  Ustru in '..\Gateway\Common\Ustru.pas',
  frmUserDataU in 'frmUserDataU.pas' {frmUserData},
  frmSTabMaintainu in 'frmSTabMaintainU.pas' {frmSTabMaintain},
  rptBasicu in 'rptBasicu.pas' {rptBasic: TQuickRep},
  frmRptPreviewU in 'frmRptPreviewU.pas' {frmRptPreView},
  dtmSTabMaintainu in 'dtmSTabMaintainU.pas' {dtmSTabMaintain: TDataModule},
  UObjectu in 'UObjectu.pas',
  frmCompDataU in 'frmCompDataU.pas' {frmCompData},
  dtmCompDataU in 'dtmCompDataU.pas' {dtmCompData: TDataModule},
  frmDeptDataU in 'frmDeptDataU.pas' {frmDeptData},
  dtmDeptDataU in 'dtmDeptDataU.pas' {dtmDeptData: TDataModule};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TdtmMain, dtmMain);
  Application.CreateForm(TfrmMain, frmMain);

  Application.Run;
end.
