program CarcySrv;

uses
  Forms,
  cbMain in 'cbMain.pas' {fmMain},
  cbStyleModule in 'cbStyleModule.pas' {StyleModule: TDataModule},
  cbConfigModule in 'cbConfigModule.pas' {ConfigModule: TDataModule},
  cbConfig in 'cbConfig.pas' {fmConfig},
  cbUtilis in '..\..\..\CommoUnit\cbUtilis.pas',
  cbClass in '..\..\..\CommoUnit\cbClass.pas',
  cbAppClass in 'cbAppClass.pas',
  cbSoDataModule in 'cbSoDataModule.pas' {SoDataModule: TDataModule},
  cbSoDbThread in 'cbSoDbThread.pas',
  OracleCI in '..\..\..\CommoUnit\OracleCI.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TStyleModule, StyleModule);
  Application.CreateForm(TfmMain, fmMain);
  Application.Run;
end.
