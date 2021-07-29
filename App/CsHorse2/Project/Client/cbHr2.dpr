program cbHr2;

uses
  Forms,
  cbMain in 'cbMain.pas' {fmHrMain},
  CsHorse2Library_Intf in '..\CsHorse2Library_Intf.pas',
  cbROClientModule in 'cbROClientModule.pas' {ROClientModule: TDataModule},
  cbUtilis in '..\..\..\CommUnit\cbUtilis.pas',
  WtsApi32 in '..\..\..\CommUnit\WtsApi32.pas',
  cbClientThread in 'cbClientThread.pas',
  cbThread in '..\cbThread.pas',
  cbDesignPattern in '..\cbDesignPattern.pas',
  cbSo in '..\cbSo.pas',
  cbAnnThread in 'cbAnnThread.pas',
  cbHrHelper in '..\cbHrHelper.pas',
  cbCallbackEventThread in 'cbCallbackEventThread.pas',
  cbDataClientModule in 'cbDataClientModule.pas' {DataClientModule: TDataModule},
  cbStyleModule in 'cbStyleModule.pas' {StyleModule: TDataModule},
  cbClientClass in 'cbClientClass.pas',
  cbSendMsg in 'cbSendMsg.pas' {fmSendMsg},
  cbReadAnn in 'cbReadAnn.pas' {fmReadAnn},
  cbControls in '..\cbControls.pas',
  cbSyncObjs in '..\cbSyncObjs.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.Title := '客服小幫手';
  Application.CreateForm(TStyleModule, StyleModule);
  Application.CreateForm(TfmHrMain, fmHrMain);
  Application.Run;
end.
