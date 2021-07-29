program Ngw;

uses
  ExceptionLog,
  Forms,
  cbMain in 'cbMain.pas' {fmMain},
  cbLiceMgr in 'cbLiceMgr.pas' {LicenceManager: TDataModule},
  cbLicInp in 'cbLicInp.pas' {fmLicInp},
  cbStyleModule in 'cbStyleModule.pas' {StyleModule: TDataModule},
  cbConfigModule in 'cbConfigModule.pas' {ConfigModule: TDataModule},
  cbConfig in 'cbConfig.pas' {fmConfig},
  cbAbout in 'cbAbout.pas' {fmAbout},
  cbResStr in 'cbResStr.pas',
  cbClass in 'cbClass.pas',
  cbSoDbThread in 'cbSoDbThread.pas',
  cbControlThread in 'cbControlThread.pas',
  cbNagraDevice in 'cbNagraDevice.pas',
  cbNagraCmd in 'cbNagraCmd.pas',
  cbFeedbackDbThread in 'cbFeedbackDbThread.pas',
  cbFeedbackThread in 'cbFeedbackThread.pas',
  cbControl in 'cbControl.pas',
  cbGuiCleanup in 'cbGuiCleanup.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TfmMain, fmMain);
  Application.Run;
end.
