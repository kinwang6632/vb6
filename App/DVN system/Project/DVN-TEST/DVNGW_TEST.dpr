program DVNGW_TEST;

uses
  Forms,
  cbMain in 'cbMain.pas' {fmMain},
  cbUtilis in '..\..\..\CommoUnit\cbUtilis.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TfmMain, fmMain);
  Application.Run;
end.
