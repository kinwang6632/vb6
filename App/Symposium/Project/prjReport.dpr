program prjReport;

uses
  Forms,
  frmMainU in 'frmMainU.pas' {frmMain},
  dtmMainU in 'dtmMainU.pas' {dtmMain: TDataModule},
  frmReport1U in 'frmReport1U.pas' {frmReport1},
  rptReport1U in 'rptReport1U.pas' {rptReport1: TQuickRep},
  frmReport2U in 'frmReport2U.pas' {frmReport2},
  rptReport2U in 'rptReport2U.pas' {rptReport2: TQuickRep},
  Ustru in 'Ustru.pas';

{$R *.RES}

begin
  Application.Initialize;
  Application.CreateForm(TfrmMain, frmMain);
  Application.CreateForm(TdtmMain, dtmMain);
  Application.Run;
end.
