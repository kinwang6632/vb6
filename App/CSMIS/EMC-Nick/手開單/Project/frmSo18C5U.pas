unit frmSo18C5U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, StdCtrls, ExtCtrls, Buttons, DB, Grids, DBGrids;

type
  TfrmSo18C5 = class(TForm)
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
  frmSo18C5: TfrmSo18C5;

implementation

uses frmSo18C0U, dtmMainU, rptReport4U, UdateTimeu;

{$R *.dfm}

procedure TfrmSo18C5.btnExitClick(Sender: TObject);
begin
    Close;
end;

procedure TfrmSo18C5.FormShow(Sender: TObject);
begin
    self.Caption := frmSo18C0.getCaption('SO8C50','��}��ں޲z-�ק��}�渹(�w�鵲)-���ʳ���');
    dtpStartDate.Date := date;
    dtpEndDate.Date := date;    
    dtpStartDate.SetFocus;

end;

procedure TfrmSo18C5.btnQueryClick(Sender: TObject);
var
    nL_RecCount : Integer;
    L_Rpt4 : TrptReport4;    
begin
    nL_RecCount := dtmMain.getSo128Data(dtpStartDate.Date, dtpEndDate.Date);
    if nL_RecCount=0 then
    begin
      MessageDlg('�d�L���ʸ��!',mtWarning, [mbOK],0);
      dtpStartDate.SetFocus;
      Exit;
    end
    else
    begin
           L_Rpt4:= TrptReport4.Create(Application);
           L_Rpt4.sG_CompName := frmSo18C0.sG_CompName;

           L_Rpt4.sG_ReportGetPaperDateStr := TUdateTime.CDateStr(dtpStartDate.Date,9) + '~' + TUdateTime.CDateStr(dtpEndDate.Date,9);

           L_Rpt4.sG_Operator := frmSo18C0.sG_Operator;
           L_Rpt4.Preview;
           L_Rpt4.Free;
    end;
end;

end.
