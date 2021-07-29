program prjGenLicenseKey;

uses
  Forms,
  frmMainU in 'frmMainU.pas' {frmMain},
  ConstObjU in '..\Common\ConstObjU.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.Title := 'LicenseKey Generator';
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
