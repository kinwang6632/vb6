program prjSimpleCA;

uses
  Forms,
  frmMainU in 'frmMainU.pas' {frmMain},
  Ustru in '..\Common\Ustru.pas',
  UdbU in '..\Common\UdbU.pas',
  frmReport1U in 'frmReport1U.pas' {frmReport1},
  rptReport1U in 'rptReport1U.pas' {rptReport1: TQuickRep},
  frmDbMultiSelectu in '..\Common\frmDbMultiSelectu.pas' {frmDbMultiSelect},
  frmParamU in 'frmParamU.pas' {frmParam},
  UdateTimeu in '..\Common\UdateTimeu.pas',
  frmCmdResultU in 'frmCmdResultU.pas' {frmCmdResult},
  ConstU in '..\Common\ConstU.pas',
  frmSyncU in 'frmSyncU.pas' {frmSync},
  dtmMainU in 'dtmMainU.pas' {dtmMain: TDataModule},
  xmlU in '..\Common\xmlU.pas',
  Encryption_TLB in 'C:\Program Files\Borland\Delphi6\Imports\Encryption_TLB.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TfrmMain, frmMain);
  Application.CreateForm(TfrmParam, frmParam);
  Application.CreateForm(TdtmMain, dtmMain);
  Application.Run;
end.
