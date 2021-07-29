program Project1;

uses
  Forms,
  Unit1 in 'Unit1.pas' {Form1},
  CbLiceMgr in 'CbLiceMgr.pas' {LicenceManager: TDataModule},
  CbResStr in 'CbResStr.pas',
  CbUtils in 'CbUtils.pas',
  CbLicInp in 'CbLicInp.pas' {frmLicInp};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
