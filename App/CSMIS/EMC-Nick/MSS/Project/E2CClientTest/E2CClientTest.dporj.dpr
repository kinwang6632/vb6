program E2CClientTest.dporj;

uses
  Forms,
  cbTestMain in 'cbTestMain.pas' {TestMain};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TTestMain, TestMain);
  Application.Run;
end.
