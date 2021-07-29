program projMax;

uses
  Forms,
  frmMainU in 'frmMainU.pas' {frmMain},
  frmLoadU in 'frmLoadU.pas' {frmLoad},
  frmOptionU in 'frmOptionU.pas' {frmOption},
  dtmBasicU in '..\..\Common\dtmBasicU.pas' {dtmBasic: TDataModule},
  dtmMainU in 'dtmMainU.pas' {dtmMain: TDataModule};

{$R *.RES}

begin
  Application.Initialize;
  Application.CreateForm(TfrmMain, frmMain);
  Application.CreateForm(TdtmBasic, dtmBasic);
  Application.CreateForm(TdtmMain, dtmMain);
  Application.Run;
end.
