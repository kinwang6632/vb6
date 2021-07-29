program TestWebPool;

uses
  Forms,
  cbMain in 'cbMain.pas' {fmMain},
  WebPool_TLB in 'WebPool_TLB.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TfmMain, fmMain);
  Application.Run;
end.
