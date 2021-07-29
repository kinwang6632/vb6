program DGW;

uses
  Forms,
  cbStyleModule in 'cbStyleModule.pas' {StyleModule: TDataModule},
  cbMain in 'cbMain.pas' {fmMain},
  cbConfigModule in 'cbConfigModule.pas' {ConfigModule: TDataModule},
  cbAppClass in 'cbAppClass.pas',
  cbSoDbThread in 'cbSoDbThread.pas',
  cbSendThread in 'cbSendThread.pas',
  cbConfig in 'cbConfig.pas' {fmConfig};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TStyleModule, StyleModule);
  Application.CreateForm(TfmMain, fmMain);
  Application.Run;
end.
