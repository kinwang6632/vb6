program NgwProfileTool;

uses
  ExceptionLog,
  Forms,
  cbNgwProfileTool in 'cbNgwProfileTool.pas' {fmNgwProfileTool};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TfmNgwProfileTool, fmNgwProfileTool);
  Application.Run;
end.
