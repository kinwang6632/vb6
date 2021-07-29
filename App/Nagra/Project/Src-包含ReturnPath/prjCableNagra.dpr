program prjCableNagra;

uses
  Forms,
  frmMainU in 'frmMainU.pas' {frmMain},
  frmSysParamU in 'frmSysParamU.pas' {frmSysParam},
  SendCommandU in 'SendCommandU.pas',
  ConstObjU in '..\Common\ConstObjU.pas',
  frmLoginU in 'frmLoginU.pas' {frmLogin},
  Ustru in '..\Common\Ustru.pas',
  UdateTimeu in '..\Common\UdateTimeu.pas',
  ReceivingCommandResponseThreadU in 'ReceivingCommandResponseThreadU.pas',
  HandleResponseU in 'HandleResponseU.pas',
  FetchComDataThreadU in 'FetchComDataThreadU.pas',
  dtmMainU in 'dtmMainU.pas' {dtmMain: TDataModule},
  HandleUIU in 'HandleUIU.pas',
  prjCableNagra_TLB in 'prjCableNagra_TLB.pas',
  ReturnPathSvrU in 'ReturnPathSvrU.pas',
  ProcessReturnPathDataU in 'ProcessReturnPathDataU.pas',
  prjMonitor_TLB in 'D:\App\Nagra\Project\ºÊ±±­û\prjMonitor_TLB.pas';

{$R *.TLB}

{$R *.RES}

begin
  Application.Initialize;
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
