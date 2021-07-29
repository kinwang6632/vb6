program pUCSet;

uses
  Forms,
  frmMainU in 'frmMainU.pas' {frmMain},
  frmUserGrpU in 'frmUserGrpU.pas' {frmUserGrp},
  frmDataFileAttrU in 'frmDataFileAttrU.pas' {frmDataFileAttr},
  frmUserGrpAttrU in 'frmUserGrpAttrU.pas' {frmUserGrpAttr},
  frmDataFileU in 'frmDataFileU.pas' {frmDataFile},
  frmAppU in 'frmAppU.pas' {frmApp},
  frmAppAttrU in 'frmAppAttrU.pas' {frmAppAttr},
  frmUserAttrU in 'frmUserAttrU.pas' {frmUserAttr},
  frmFuncAttrU in 'frmFuncAttrU.pas' {frmFuncAttr},
  frmSelFuncU in 'frmSelFuncU.pas' {frmSelFunc},
  frmSelUserU in 'frmSelUserU.pas' {frmSelUser},
  utlUCU in '..\Common\utlUCU.pas';

{$R *.RES}

begin
  Application.Initialize;
  Application.CreateForm(TfrmMain, frmMain);
  Application.CreateForm(TfrmUserGrp, frmUserGrp);
  Application.CreateForm(TfrmApp, frmApp);
  Application.Run;
end.
