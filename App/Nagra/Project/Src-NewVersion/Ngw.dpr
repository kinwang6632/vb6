program Ngw;

uses
  Forms,
  cbMain in 'cbMain.pas' {fmMain},
  cbLiceMgr in 'cbLiceMgr.pas' {LicenceManager: TDataModule},
  cbResStr in 'cbResStr.pas',
  cbUtils in 'cbUtils.pas',
  cbLicInp in 'cbLicInp.pas' {frmLicInp},
  cbStyleModule in 'cbStyleModule.pas' {StyleModule: TDataModule},
  cbConfigModule in 'cbConfigModule.pas' {ConfigModule: TDataModule},
  cbClass in 'cbClass.pas',
  cbSoDbThread in 'cbSoDbThread.pas',
  cbControlThread in 'cbControlThread.pas',
  cbNagraDevice in 'cbNagraDevice.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TfmMain, fmMain);
  Application.Run;
end.
