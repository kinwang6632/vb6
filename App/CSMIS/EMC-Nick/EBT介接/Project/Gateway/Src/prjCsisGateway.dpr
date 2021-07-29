program prjCsisGateway;

uses
  Forms,
  frmMainU in 'frmMainU.pas' {frmMain},
  frmLoginU in 'frmLoginU.pas' {frmLogin},
  frmParamU in 'frmParamU.pas' {frmParam},
  frmUserU in 'frmUserU.pas' {frmUser},
  ConstObjU in '..\Common\ConstObjU.pas',
  dtmMainU in 'dtmMainU.pas' {dtmMain: TDataModule},
  Ustru in '..\Common\Ustru.pas',
  UdateTimeu in '..\Common\UdateTimeu.pas',
  Encryption_TLB in '..\Common\Encryption_TLB.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TdtmMain, dtmMain);
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
