program prjMonitor;

uses
  Forms,
  frmMainU in 'frmMainU.pas' {frmMain},
  ConstObjU in '..\Common\ConstObjU.pas',
  Ustru in '..\Common\Ustru.pas',
  Uotheru in '..\Common\UotherU.pas',
  UmsgCodeu in '..\Common\UmsgCodeu.pas',
  prjMonitor_TLB in 'prjMonitor_TLB.pas',
  CoMonitorImpU in 'CoMonitorImpU.pas' {Monitor: CoClass},
  UdateTimeu in '..\Common\UdateTimeu.pas';

{$R *.TLB}

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
