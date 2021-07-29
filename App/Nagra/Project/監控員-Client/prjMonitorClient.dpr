program prjMonitorClient;

uses
  Forms,
  frmMainU in 'frmMainU.pas' {frmMain},
  Uotheru in '..\Common\UotherU.pas',
  UmsgCodeu in '..\Common\UmsgCodeu.pas',
  Ustru in '..\Common\Ustru.pas',
  prjMonitor_TLB in '..\ºÊ±±­û\prjMonitor_TLB.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
