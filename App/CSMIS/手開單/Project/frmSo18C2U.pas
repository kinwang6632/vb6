unit frmSo18C2U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, Buttons, DB, Grids, DBGrids;

type
  TfrmSo18C2 = class(TForm)
    Panel1: TPanel;
    StaticText1: TStaticText;
    edtPaperSNum: TEdit;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    btnQuery: TBitBtn;
    StaticText7: TStaticText;
    edtPaperENum: TEdit;
    dbgPaperData: TDBGrid;
    dsrPaperData: TDataSource;
    procedure FormShow(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure btnQueryClick(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
    function canObsolete(I_DataSet:TDataSet):boolean;
  public
    { Public declarations }
  end;

var
  frmSo18C2: TfrmSo18C2;

implementation

uses frmSo18C0U, dtmMainU, Ustru, UdateTimeu;

{$R *.dfm}

procedure TfrmSo18C2.FormShow(Sender: TObject);
begin
    self.Caption := frmSo18C0.getCaption('SO8C20','手開單據管理-單據做廢作業');
end;

procedure TfrmSo18C2.BitBtn2Click(Sender: TObject);
begin
    Close;
end;

procedure TfrmSo18C2.btnQueryClick(Sender: TObject);
var
    L_DataSet : TDataSet;
begin
    if edtPaperSNum.Text = '' then
    begin
      edtPaperSNum.SetFocus;
      MessageDlg('請輸入欲做廢之單號!', mtWarning, [mbOK],0);
      Exit;
    end;

    if edtPaperENum.Text = '' then
    begin
      edtPaperENum.SetFocus;
      MessageDlg('請輸入欲做廢之單號!', mtWarning, [mbOK],0);
      Exit;
    end;

    L_DataSet := dtmMain.getPaperData(edtPaperSNum.Text, edtPaperENum.Text);
    if L_DataSet.RecordCount>0 then
    begin
      dsrPaperData.DataSet := L_DataSet;
      L_DataSet.FieldByName('PAPERNUM').DisplayLabel := '單號';
      L_DataSet.FieldByName('EMPNAME').DisplayWidth := 10;
      L_DataSet.FieldByName('EMPNAME').DisplayLabel := '領單人員';
      L_DataSet.FieldByName('GETPAPERDATE').DisplayLabel := '領單日期';
      L_DataSet.FieldByName('BILLNO').DisplayLabel := '單據編號';
    end
    else
      MessageDlg('無可供做廢單據!', mtInformation, [mbOK],0);
  
end;

procedure TfrmSo18C2.BitBtn1Click(Sender: TObject);
var
    sL_PaperSNum, sL_PaperENum : String;
begin
    if (dsrPaperData.DataSet = nil ) or
       (dsrPaperData.DataSet.State=dsInactive) or
       (dsrPaperData.DataSet.RecordCount = 0) then
    begin
      edtPaperSNum.SetFocus;
      MessageDlg('請先查詢出欲做廢之單號!', mtInformation, [mbOK],0);
      Exit;
    end;
       
    if canObsolete(dbgPaperData.DataSource.DataSet) then
    begin
      sL_PaperSNum := TUstr.AddString(edtPaperSNum.Text,'0',false,PAPER_NUMBER_LENGTH);
      sL_PaperENum := TUstr.AddString(edtPaperENum.Text,'0',false,PAPER_NUMBER_LENGTH);
//      if MessageDlg('是否確認做廢此批單據? [單號:' + sL_PaperSNum + '~' + sL_PaperENum + ']', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
      if MessageDlg('是否確認做廢此批單據?' + '[ 共 ' + IntToStr(dsrPaperData.DataSet.recordCount) + ' 張 ]', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
      begin
        if dtmMain.disablePaperData(sL_PaperSNum, sL_PaperENum) then
          MessageDlg('作廢完成!', mtInformation, [mbOK],0);
      end;
    end
    else
    begin
      MessageDlg('已有單據編號,無法做廢!', mtInformation, [mbOK],0);
    end;

end;

procedure TfrmSo18C2.FormClose(Sender: TObject; var Action: TCloseAction);
begin
    Action := caFree;
end;

function TfrmSo18C2.canObsolete(I_DataSet: TDataSet): boolean;
var
    bL_Result : boolean;
begin
    bL_Result := true;
    with I_DataSet do
    begin
      First;
      while not Eof do
      begin
        if FieldByName('BILLNO').AsString <> '' then
        begin
          bL_Result := false;
          break;
        end;
        Next;  
      end;
    end;

    result := bL_Result;
end;

end.
