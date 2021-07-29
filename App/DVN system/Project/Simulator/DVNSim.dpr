program DVNSim;

uses
  Forms,
  cbMain in 'cbMain.pas' {fmMain},
  cbAppClass in '..\DVN2006\cbAppClass.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TfmMain, fmMain);
  Application.Run;
end.
