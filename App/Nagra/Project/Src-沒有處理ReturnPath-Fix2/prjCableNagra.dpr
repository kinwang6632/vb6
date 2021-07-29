program prjCableNagra;

uses
  Forms,
  frmMainU in 'frmMainU.pas' {frmMain},
  frmSysParamU in 'frmSysParamU.pas' {frmSysParam},
  SendCommandU in 'SendCommandU.pas',
  ConstObjU in '..\Common\ConstObjU.pas',
  frmLoginU in 'frmLoginU.pas' {frmLogin},
  UdateTimeu in '..\Common\UdateTimeu.pas',
  ReceivingCommandResponseThreadU in 'ReceivingCommandResponseThreadU.pas',
  HandleResponseU in 'HandleResponseU.pas',
  FetchComDataThreadU in 'FetchComDataThreadU.pas',
  dtmMainU in 'dtmMainU.pas' {dtmMain: TDataModule},
  HandleUIU in 'HandleUIU.pas',
  prjCableNagra_TLB in 'prjCableNagra_TLB.pas',
  Ustru in '..\Common\Ustru.pas',
  prjMonitor_TLB in '..\ºÊ±±­û\prjMonitor_TLB.pas',
  Uotheru in '..\Common\UotherU.pas',
  UmsgCodeu in '..\Common\UmsgCodeu.pas',
  ProcessReturnPathDataU in 'ProcessReturnPathDataU.pas',
  ReturnPathSvrU in 'ReturnPathSvrU.pas',
  cbQueue in 'cbQueue.pas',
  CbUtils in 'CbUtils.pas';

{$R *.TLB}

{$R *.RES}

begin
  Application.Initialize;
  Application.CreateForm(TdtmMain, dtmMain);
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
