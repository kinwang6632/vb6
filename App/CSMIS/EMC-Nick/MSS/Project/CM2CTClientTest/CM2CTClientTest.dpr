program CM2CTClientTest;

uses
  Forms,
  cbTestMain in 'cbTestMain.pas' {TestMain};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TTestMain, TestMain);
  Application.Run;
end.
