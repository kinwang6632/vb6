program prjSimpleCA;

uses
  Forms,
  frmMainU in 'frmMainU.pas' {frmMain},
  Ustru in '..\..\..\CommoUnit\Ustru.pas',
  UdbU in '..\..\..\CommoUnit\UdbU.pas',
  frmReport1U in 'frmReport1U.pas' {frmReport1},
  rptReport1U in 'rptReport1U.pas' {rptReport1: TQuickRep},
  dtmMainU in 'dtmMainU.pas' {dtmMain: TDataModule},
  ConstObjU in '..\..\..\CommoUnit\ConstObjU.pas',
  UdateTimeu in '..\..\..\CommoUnit\UdateTimeu.pas',
  frmSysParamU in 'frmSysParamU.pas' {frmSysParam},
  frmDbMultiSelectu in '..\..\..\CommoUnit\frmDbMultiSelectu.pas',
  cbMain in 'cbMain.pas' {fmMain},
  cbUtilis in '..\..\..\CommoUnit\cbUtilis.pas',
  Encryption_TLB in '..\..\..\CommoUnit\Encryption_TLB.pas',
  cbChannel in 'cbChannel.pas' {fmChannel},
  cbUser in 'cbUser.pas' {fmUser};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TfmMain, fmMain);
  Application.Run;
end.
