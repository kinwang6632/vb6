program CmBatchProc;

uses
  Forms,
  unMain in 'unMain.pas' {fmMain},
  OracleCI in '..\..\CommoUnit\OracleCI.pas',
  cbUtilis in '..\..\CommoUnit\cbUtilis.pas',
  cbSQL in 'cbSQL.pas' {fmSQL};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TfmMain, fmMain);
  Application.CreateForm(TfmSQL, fmSQL);
  Application.Run;
end.
