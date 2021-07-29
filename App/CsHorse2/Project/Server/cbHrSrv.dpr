program cbHrSrv;

{#ROGEN:..\CsHorse2Library.rodl}

uses
  uROCOMInit,
  Forms,
  cbDesignPattern in '..\cbDesignPattern.pas',
  cbLanguage in '..\cbLanguage.pas',
  cbSo in '..\cbSo.pas',
  cbThread in '..\cbThread.pas',
  cbStyleModule in 'cbStyleModule.pas' {StyleModule: TDataModule},
  cbSrvClass in 'cbSrvClass.pas',
  cbMain in 'cbMain.pas' {fmMain},
  cbConfigModule in 'cbConfigModule.pas' {ConfigModule: TDataModule},
  cbMainUIModule in 'cbMainUIModule.pas' {MainUIModule: TDataModule},
  cbOraModule in 'cbOraModule.pas' {OraModule: TDataModule},
  cbSyncDbDataThread in 'cbSyncDbDataThread.pas',
  cbSessionMonitorThread in 'cbSessionMonitorThread.pas',
  cbUpdateUIThread in 'cbUpdateUIThread.pas',
  CsHorse2Library_Intf in '..\CsHorse2Library_Intf.pas',
  CsHorse2Library_Invk in '..\CsHorse2Library_Invk.pas',
  cbROServerModule in 'cbROServerModule.pas' {ROServerModule: TDataModule},
  LoginService_Impl in '..\LoginService_Impl.pas' {LoginService: TRORemoteDataModule},
  AnnService_Impl in '..\AnnService_Impl.pas' {AnnService: TRORemoteDataModule},
  CallbackService_Impl in '..\CallbackService_Impl.pas' {CallbackService: TRORemoteDataModule},
  cbHrHelper in '..\cbHrHelper.pas',
  cbConfig in 'cbConfig.pas' {fmConfig},
  cbGroupSet in 'cbGroupSet.pas' {fmGroupSet},
  cbAbout in 'cbAbout.pas' {fmAbout},
  cbSyncObjs in '..\cbSyncObjs.pas';

{$R *.res}
{$R ..\RODLFile.res}

begin
  Application.Initialize;
  Application.CreateForm(TStyleModule, StyleModule);
  Application.CreateForm(TfmMain, fmMain);
  Application.Run;
end.

