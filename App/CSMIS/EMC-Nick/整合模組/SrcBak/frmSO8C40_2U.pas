unit frmSO8C40_2U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, StdCtrls, ExtCtrls, Buttons, DB, Grids, DBGrids;

type
  TfrmSO8C40_2 = class(TForm)
    Panel1: TPanel;
    Label1: TLabel;
    btnQuery: TBitBtn;
    btnExit: TBitBtn;
    StaticText3: TStaticText;
    dtpStartDate: TDateTimePicker;
    dtpEndDate: TDateTimePicker;
    DataSource1: TDataSource;
    procedure btnExitClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure btnQueryClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmSO8C40_2: TfrmSO8C40_2;

implementation

uses dtmMain3U, rptSO8C40U, UdateTimeu, frmMainMenuU;

{$R *.dfm}

procedure TfrmSO8C40_2.btnExitClick(Sender: TObject);
begin
    Close;
end;

procedure TfrmSO8C40_2.FormShow(Sender: TObject);
begin
    self.Caption := frmMainMenu.setCaption('SO8C50','手開單據管理-修改手開單號(已日結)-異動報表');
    dtpStartDate.Date := date;
    dtpEndDate.Date := date;    
    dtpStartDate.SetFocus;

end;

procedure TfrmSO8C40_2.btnQueryClick(Sender: TObject);
var
    nL_RecCount : Integer;
    L_Rpt4 : TrptSO8C40;
begin
    nL_RecCount := dtmMain3.getSo128Data(dtpStartDate.Date, dtpEndDate.Date);
    if nL_RecCount=0 then
    begin
      MessageDlg('查無異動資料!',mtWarning, [mbOK],0);
      dtpStartDate.SetFocus;
      Exit;
    end
    else
    begin
           L_Rpt4:= TrptSO8C40.Create(Application);
           L_Rpt4.sG_CompName := frmMainMenu.sG_CompName;

           L_Rpt4.sG_ReportGetPaperDateStr :=
             FormatDateTime( 'yyyy/mm/dd', dtpStartDate.Date ) + '~' +
             FormatDateTime( 'yyyy/mm/dd', dtpEndDate.Date );

           L_Rpt4.sG_Operator := frmMainMenu.sG_User;
           L_Rpt4.Preview;
           L_Rpt4.Free;
    end;
end;

end.
