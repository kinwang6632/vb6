program prjSimpleCA;

uses
  Forms,
  frmMainU in 'frmMainU.pas' {frmMain},
  frmDbMultiSelectu in '..\Common\frmDbMultiSelectu.pas' {frmDbMultiSelect},
  Ustru in '..\Common\Ustru.pas',
  UdbU in '..\Common\UdbU.pas',
  frmReport1U in 'frmReport1U.pas' {frmReport1},
  rptReport1U in 'rptReport1U.pas' {rptReport1: TQuickRep},
  dtmMainU in 'dtmMainU.pas' {dtmMain: TDataModule},
  ConstObjU in '..\Common\ConstObjU.pas',
  UdateTimeu in '..\Common\UdateTimeu.pas',
  frmSysParamU in 'frmSysParamU.pas' {frmSysParam};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TfrmMain, frmMain);
  Application.CreateForm(TdtmMain, dtmMain);
  Application.Run;
end.
