program PgSch;

uses
  Forms,
  cbMain in 'cbMain.pas' {fmMain},
  cbStyleModule in 'cbStyleModule.pas' {StyleModule: TDataModule},
  cbConfigModule in 'cbConfigModule.pas' {ConfigModule: TDataModule},
  cbClass in '..\..\..\CommoUnit\cbClass.pas',
  cbUtilis in '..\..\..\CommoUnit\cbUtilis.pas',
  OracleCI in '..\..\..\CommoUnit\OracleCI.pas',
  cbAppClass in 'cbAppClass.pas',
  cbParseThread in 'cbParseThread.pas',
  cbWatchThread in 'cbWatchThread.pas',
  cbOption in 'cbOption.pas' {fmOption},
  cbEmail in 'cbEmail.pas' {fmEmail},
  cbParseController in 'cbParseController.pas' {ParseController: TDataModule},
  cbLogModule in 'cbLogModule.pas' {LogModule: TDataModule};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TStyleModule, StyleModule);
  Application.CreateForm(TfmMain, fmMain);
  Application.Run;
end.
