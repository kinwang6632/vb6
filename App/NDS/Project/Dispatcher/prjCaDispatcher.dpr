program prjCaDispatcher;

uses
  Forms,
  frmLoginU in 'frmLoginU.pas' {frmLogin},
  dtmMainU in 'dtmMainU.pas' {dtmMain: TDataModule},
  frmSysParamU in 'frmSysParamU.pas' {frmSysParam},
  frmMainU in 'frmMainU.pas' {frmMain},
  frmSendCommandU in 'frmSendCommandU.pas' {frmSendCommand},
  TListernerU in 'TListernerU.pas',
  Ustru in '..\Common\Ustru.pas',
  ConstU in '..\Common\ConstU.pas',
  frmTestDataU in 'frmTestDataU.pas' {frmTestData},
  frmAboutU in 'frmAboutU.pas' {frmAbout},
  frmDbInfoU in 'frmDbInfoU.pas' {frmDbInfo};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TdtmMain, dtmMain);
  Application.CreateForm(TfrmLogin, frmLogin);
  Application.CreateForm(TfrmAbout, frmAbout);
  Application.Run;
end.
