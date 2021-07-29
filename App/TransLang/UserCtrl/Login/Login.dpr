program Login;

uses
  Forms,
  frmCheckUserU in 'frmCheckUserU.pas' {frmCheckUser},
  utlUserInoU in '..\Common\utlUserInoU.pas',
  utlUCU in '..\Common\utlUCU.pas';

{$R *.RES}

begin
  Application.Initialize;
  Application.CreateForm(TfrmCheckUser, frmCheckUser);
  Application.Run;
end.
