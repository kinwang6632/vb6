program prjSO62M0A;

uses
  Forms,
  frmCD078B0U in 'frmCD078B0U.pas' {frmCD078B0},
  frmCD078B1U in 'frmCD078B1U.pas' {frmCD078B1},
  frmCD078B2U in 'frmCD078B2U.pas' {frmCD078B2},
  frmCD078B3U in 'frmCD078B3U.pas' {frmCD078B3},
  frmCD078B3_1U in 'frmCD078B3_1U.pas' {frmCD078B3_1},
  frmCD078B3_2U in 'frmCD078B3_2U.pas' {frmCD078B3_2},
  frmCD078B3_3U in 'frmCD078B3_3U.pas' {frmCD078B3_3},
  frmCD078B3_4U in 'frmCD078B3_4U.pas' {frmCD078B3_4},
  frmCD078B3_5U in 'frmCD078B3_5U.pas' {frmCD078B3_5},
  frmCD078B4U in 'frmCD078B4U.pas' {frmCD078B4},
  frmCD078B4_1U in 'frmCD078B4_1U.pas' {frmCD078B4_1},
  frmCD078B5U in 'frmCD078B5U.pas' {frmCD078B5},
  frmCD078B8U in 'frmCD078B8U.pas' {frmCD078B8},
  frmCD078B9U in 'frmCD078B9U.pas' {frmCD078B9},
  cbDBController in 'cbDBController.pas' {DBController: TDataModule},
  cbStyleController in 'cbStyleController.pas' {StyleController: TDataModule},
  cbDataLookup in 'cbDataLookup.pas' {DataLookup: TFrame},
  frmMultiSelectU in 'frmMultiSelectU.pas' {frmMultiSelect},
  frmCD078G0U in 'frmCD078G0U.pas' {frmCD078GU},
  frmCD078G1U in 'frmCD078G1U.pas' {frmCD078G1},
  frmCD078B3_7U in 'frmCD078B3_7U.pas' {frmCD078B3_7},
  frmCD078IU in 'frmCD078IU.pas' {frmCD078I},
  frmCD078J1U in 'frmCD078J1U.pas' {frmCD078JU},
  frmCD078BAU in 'frmCD078BAU.pas' {frmCD078BA},
  frmCD078K0U in 'frmCD078K0U.pas' {frmCD078K},
  frmCD078L0U in 'frmCD078L0U.pas' {frmCD078L};

{$R *.res}

begin
  Application.Initialize;
  Application.Title := '優惠組合管理 [SO62M0A]';
  Application.CreateForm(TfrmCD078B0, frmCD078B0);
  Application.CreateForm(TfrmCD078B3_7, frmCD078B3_7);
  Application.CreateForm(TfrmCD078BA, frmCD078BA);
  //  Application.CreateForm(TfrmCD078JU, frmCD078JU);
  //  Application.CreateForm(TForm1, Form1);
  //  Application.CreateForm(TForm1, Form1);
//    Application.CreateForm(TfrmCD078G1, frmCD078G1);
  //  Application.CreateForm(TfrmCD078GU, frmCD078GU);
  //  Application.CreateForm(TfrmCD078B3_6, frmCD078B3_6);
  //  Application.CreateForm(TfrmCD078B6, frmCD078B6);
  Application.Run;
end.
