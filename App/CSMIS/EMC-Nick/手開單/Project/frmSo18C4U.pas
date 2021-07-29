unit frmSo18C4U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, ExtCtrls, Grids, DBGrids, DB;

type
  TfrmSo18C4 = class(TForm)
    Panel1: TPanel;
    btnQuery: TBitBtn;
    btnExit: TBitBtn;
    Label1: TLabel;
    GroupBox1: TGroupBox;
    dbgSo034: TDBGrid;
    edtBillNo: TEdit;
    edtSrcHandPaperNo: TEdit;
    edtNewHandPaperNo: TEdit;
    btnUpdate: TButton;
    dsrSo034: TDataSource;
    StaticText1: TStaticText;
    StaticText2: TStaticText;
    btnReport: TBitBtn;
    procedure FormShow(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure btnQueryClick(Sender: TObject);
    procedure dsrSo034DataChange(Sender: TObject; Field: TField);
    procedure btnUpdateClick(Sender: TObject);
    procedure btnReportClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmSo18C4: TfrmSo18C4;

implementation

uses frmSo18C0U, dtmMainU, frmSo18C5U;

{$R *.dfm}

procedure TfrmSo18C4.FormShow(Sender: TObject);
begin
    self.Caption := frmSo18C0.getCaption('SO8C40','手開單據管理-修改手開單號(已日結)');
    edtBillNo.SetFocus;
end;

procedure TfrmSo18C4.btnExitClick(Sender: TObject);
begin
    dsrSo034.DataSet.Close;
    Close;
end;

procedure TfrmSo18C4.btnQueryClick(Sender: TObject);
var
    sL_BillNo : STring;
    nL_RecCount : Integer;
begin
    sL_BillNo := Trim(edtBillNo.Text);
    if sL_BillNo='' then
    begin
      MessageDlg('請輸入收費單號!',mtWarning, [mbOK],0);
      edtBillNo.SetFocus;
      Exit;
    end;

    nL_RecCount := dtmMain.getSo034Data(sL_BillNo);
    if nL_RecCount=0 then
    begin
      MessageDlg('查無此收費單資料!',mtWarning, [mbOK],0);
      edtBillNo.SetFocus;
      Exit;
    end;

end;

procedure TfrmSo18C4.dsrSo034DataChange(Sender: TObject; Field: TField);
begin
    edtSrcHandPaperNo.Text := dsrSo034.DataSet.FieldByName('ManualNo').AsString;
end;

procedure TfrmSo18C4.btnUpdateClick(Sender: TObject);
var
    sL_CustID, sL_CustName, sL_Tel1, sL_Item : String;
    dL_RealDate : TDate;
    sL_ConfirmMsg : String;
    sL_BillNo, sL_OldHandPaperNo, sL_NewHandPaperNo:String;
    sL_ErrorMsg:String;
begin
    if (dsrSo034.DataSet.State=dsInactive) or  (dsrSo034.DataSet.RecordCount=0) then
    begin
      MessageDlg('請先查詢出收費單資料!',mtWarning, [mbOK],0);
      edtBillNo.SetFocus;
      Exit;
    end;

    sL_OldHandPaperNo := Trim(edtSrcHandPaperNo.Text);
    sL_NewHandPaperNo := Trim(edtNewHandPaperNo.Text);

    if sL_OldHandPaperNo=sL_NewHandPaperNo then
    begin
      MessageDlg('對應之手開單號=新手開單號!無法修改',mtWarning, [mbOK],0);
      edtNewHandPaperNo.SetFocus;
      Exit;
    end;
    sL_BillNo := Trim(dsrSo034.DataSet.FieldByName('BILLNO').AsString);
    sL_ConfirmMsg := '確定要將 ' + sL_BillNo + ' 之手開單號: ' + sL_OldHandPaperNo + ' 換成: ' + sL_NewHandPaperNo + ' ?';
    if MessageDlg(sL_ConfirmMsg, mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      sL_CustID := dsrSo034.DataSet.FieldByName('CustId').AsString;
      sL_CustName := dsrSo034.DataSet.FieldByName('CustName').AsString;
      sL_Tel1 := dsrSo034.DataSet.FieldByName('Tel1').AsString;
      sL_Item := dsrSo034.DataSet.FieldByName('Item').AsString;
      dL_RealDate := dsrSo034.DataSet.FieldByName('REALDATE').AsDateTime;


      if dtmMain.setNewHandPaperNo(sL_CustID,sL_CustName,sL_Tel1, sL_Item, dL_RealDate, sL_BillNo, sL_OldHandPaperNo, sL_NewHandPaperNo, sL_ErrorMsg) then
      begin
        MessageDlg('更新完成!',mtInformation, [mbOK],0);
        btnQueryClick(Sender);
        edtBillNo.SetFocus;
        exit;
      end
      else
      begin
        if sL_ErrorMsg='' then
          MessageDlg('更新失敗!',mtError, [mbOK],0)
        else
          MessageDlg(sL_ErrorMsg,mtError, [mbOK],0);        
        edtNewHandPaperNo.SetFocus;
        exit;
      end;
    end;
end;

procedure TfrmSo18C4.btnReportClick(Sender: TObject);
var
    L_Frm : TfrmSo18C5;
begin
    L_Frm := TfrmSo18C5.Create(Application);
    L_Frm.ShowModal;
    L_Frm.Free;

end;

end.
