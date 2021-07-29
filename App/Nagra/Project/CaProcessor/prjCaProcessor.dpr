program prjCaProcessor;

uses
  Forms,
  frmMainU in 'frmMainU.pas' {frmMain},
  frmSysParamU in 'frmSysParamU.pas' {frmSysParam},
  frmLoginU in 'frmLoginU.pas' {frmLogin},
  TListernerU in 'TListernerU.pas',
  ConstU in '..\Common\ConstU.pas',
  Ustru in '..\Common\Ustru.pas',
  UdateTimeu in '..\Common\UdateTimeu.pas',
  dllNdsU in '..\Common\dllNdsU.pas',
  frmAboutU in 'frmAboutU.pas' {frmAbout},
  frmDispatcherInfoU in 'frmDispatcherInfoU.pas' {frmDispatcherInfo},
  dtmMainU in 'dtmMainU.pas' {dtmMain: TDataModule};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TdtmMain, dtmMain);
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
