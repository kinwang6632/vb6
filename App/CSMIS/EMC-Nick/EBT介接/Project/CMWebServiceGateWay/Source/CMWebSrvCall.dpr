program CMWebSrvCall;

uses
  Forms,
  cbMain in 'cbMain.pas' {Main},
  cbDataController in 'cbDataController.pas' {DataController: TDataModule},
  cbSoInfo in 'cbSoInfo.pas',
  Encryption_TLB in 'Encryption_TLB.pas',
  cbCMCP in 'cbCMCP.pas',
  cbUtils in 'cbUtils.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TMain, Main);
  Application.Run;
end.
