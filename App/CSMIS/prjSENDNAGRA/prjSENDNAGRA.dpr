program prjSENDNAGRA;

uses
  Forms,
  cbMain in 'cbMain.pas' {fmMain};

{$R *.res}

begin
  Application.Initialize;
  //Application.MainFormOnTaskbar := True;
  Application.CreateForm(TfmMain, fmMain);
  Application.Run;
end.
