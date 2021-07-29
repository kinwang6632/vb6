unit frmReport1U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, Buttons, ComCtrls;

type
  TfrmReport1 = class(TForm)
    Panel1: TPanel;
    btnPrint: TBitBtn;
    btnCancel: TBitBtn;
    RadioGroup1: TRadioGroup;
    edtSubscruberID: TEdit;
    edtIccNo: TEdit;
    DateTimePicker1: TDateTimePicker;
    function IsDataOK(var sI_Cond:String):boolean;
    procedure btnCancelClick(Sender: TObject);
    procedure RadioGroup1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmReport1: TfrmReport1;

implementation

uses frmMainU, rptReport1U;

{$R *.dfm}

procedure TfrmReport1.btnCancelClick(Sender: TObject);
begin
    self.Close;
end;

procedure TfrmReport1.RadioGroup1Click(Sender: TObject);
begin
    if RadioGroup1.ItemIndex = 0 then
      edtSubscruberID.SetFocus
    else if RadioGroup1.ItemIndex = 1 then
      edtIccNo.SetFocus
    else
      DateTimePicker1.SetFocus;
end;

procedure TfrmReport1.FormShow(Sender: TObject);
begin
    DateTimePicker1.Date := date;
    edtSubscruberID.SetFocus;
end;

function TfrmReport1.IsDataOK(var sI_Cond:String): boolean;
var
    sL_StartDate, sL_EndDate : String;
    wL_Year, wL_Month, wL_Day : word;
    dL_StartDate, dL_EndDate : TDate;
    bL_Result : boolean;
begin
    bL_Result := true;
    if RadioGroup1.ItemIndex = 0 then
    begin
      if edtSubscruberID.Text='' then
      begin
        bL_Result := false;
        MessageDlg('請輸入 Subscriber ID.', mtWarning, [mbOK],0);
        edtSubscruberID.SetFocus;
      end
      else
        sI_Cond := ' SUBSCRIBER_ID=' + STR_SEP + Trim(edtSubscruberID.Text) + STR_SEP + ' ';
    end
    else if RadioGroup1.ItemIndex = 1 then
    begin
      if edtIccNo.Text='' then
      begin
        bL_Result := false;
        MessageDlg('請輸入 ICC NO.', mtWarning, [mbOK],0);
        edtIccNo.SetFocus;
      end
      else
        sI_Cond := ' ICC_NO=' + STR_SEP + Trim(edtIccNo.Text) + STR_SEP + ' ';
    end
    else
    begin
      DecodeDate(DateTimePicker1.date,wL_Year, wL_Month, wL_Day);
      sL_StartDate := Format('%.4d/%.2d/%.2d 00:00:01', [wL_Year, wL_Month, wL_Day]);
      sL_EndDate := Format('%.4d/%.2d/%.2d 23:59:59', [wL_Year, wL_Month, wL_Day]);
      dL_StartDate := StrToDateTime(sL_StartDate);
      dL_EndDate := StrToDateTime(sL_EndDate);
      sI_Cond := ' UPDTIME between ' + frmMain.getOracleSQLDateTimeStr(dL_StartDate) + ' and ' + frmMain.getOracleSQLDateTimeStr(dL_EndDate) + ' ';
    end;
    
    result := bL_Result;
end;

procedure TfrmReport1.btnPrintClick(Sender: TObject);
var
    nL_RecordCount : Integer;
    sL_SQL, sL_Cond : STring;
    L_Rpt : TrptReport1;
begin
    if IsDataOK(sL_Cond) then
    begin
      sL_SQL := 'select * from NDS002 where ' + sL_Cond + ' order by UPDTIME ,ACTION';
      nL_RecordCount := frmMain.activeReport1Data(sL_SQL);
      if nL_RecordCount=0 then
      begin
        MessageDlg('查無資料可供列印!', mtInformation, [mbOK],0);
        RadioGroup1Click(nil);
      end
      else
      begin
        L_Rpt := TrptReport1.Create(Application);
        L_Rpt.nG_RecCount := nL_RecordCount;
        L_Rpt.sG_Operator := frmMain.sG_UserID;
        L_Rpt.Preview;
        L_Rpt.Free;
      end;
    end;
end;

end.
