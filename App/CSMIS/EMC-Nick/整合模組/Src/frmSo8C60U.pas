unit frmSo8C60U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, fraChineseYMDU, ExtCtrls, Buttons;

type
  TfrmSo8C60 = class(TForm)
    Panel1: TPanel;
    fraChineseYMD3: TfraChineseYMD;
    StaticText5: TStaticText;
    StaticText6: TStaticText;
    btnGetPaperUser: TButton;
    StaticText7: TStaticText;
    edtName: TEdit;
    StaticText9: TStaticText;
    edtBeginNumber: TEdit;
    lblCounts: TLabel;
    StaticText1: TStaticText;
    lblEndNumber: TLabel;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    StaticText2: TStaticText;
    lblBeginNumber: TLabel;
    procedure FormShow(Sender: TObject);
    procedure btnGetPaperUserClick(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure edtBeginNumberExit(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    sG_OriginalEmpNo, sG_OriginalEmpName, sG_OriginalCount : String;
    dG_OriginalGetPaperDate : TDate;
    nG_NewBeginNum, nG_OriginalBeginNum, nG_OriginalEndNum : Integer;
    sG_OriginalSeq, sG_CompCode,  sG_EmpNo, sG_Prefix, sG_BeginNum, sG_EndNum, sG_Counts : String;
  end;

var
  frmSo8C60: TfrmSo8C60;

implementation

uses dtmMain3U, frmDbSelectu, Ustru, frmSO8C10U;

{$R *.dfm}

procedure TfrmSo8C60.FormShow(Sender: TObject);
begin
//1234    self.Caption := frmSo8C10.getCaption('SO8C60','手開單據管理-銷單領用');


    edtBeginNumber.Text := sG_BeginNum;
    lblBeginNumber.Caption := sG_BeginNum;
    nG_OriginalBeginNum := StrToIntDef(sG_BeginNum, 0 );
    nG_NewBeginNum := nG_OriginalBeginNum;

    lblEndNumber.Caption := sG_EndNum;
    nG_OriginalEndNum := StrToIntDef(sG_EndNum, 0);

    lblCounts.Caption := sG_Counts;

    fraChineseYMD3.mseYMD.SetFocus;
end;

procedure TfrmSo8C60.btnGetPaperUserClick(Sender: TObject);
begin
    dtmMain3.activeDataSet(dtmMain3.qryCm003);
    dtmMain3.qryCm003.FieldByName('EMPNO').DisplayLabel := '員工編號';
    dtmMain3.qryCm003.FieldByName('EMPNAME').DisplayLabel := '姓名';

    if SelectRecord('請點選領單人員', dtmMain3.qryCm003, 'EMPNO;EMPNAME') = mrOk then
    begin
      edtName.Text := dtmMain3.qryCm003.FieldByName('EMPNAME').AsString;
      sG_EmpNo := dtmMain3.qryCm003.FieldByName('EMPNO').AsString;
      {
      if not (dsrSo126.DataSet.State in [dsEdit, dsInsert]) then
        dsrSo126.DataSet.Edit;
      dsrSo126.DataSet.FieldByName('EMPNO').AsString := dtmMain.qryCm003.FieldByName('EMPNO').AsString;
      dsrSo126.DataSet.FieldByName('EMPNAME').AsString := dtmMain.qryCm003.FieldByName('EMPNAME').AsString;
      }
    end;

end;

procedure TfrmSo8C60.BitBtn2Click(Sender: TObject);
begin
    Close;
end;

procedure TfrmSo8C60.BitBtn1Click(Sender: TObject);
var
    sL_Result, aValue : String;
begin
    if Trim(TUstr.replaceStr(fraChineseYMD3.getYMD,'/',''))='' then
    begin
      MessageDlg('請輸入領單日期!', mtWarning, [mbOk],0);
      fraChineseYMD3.mseYMD.SetFocus;
      Exit;
    end;
    if (edtName.Text='') then
    begin
      MessageDlg('請點選領單人員!', mtWarning, [mbOk],0);
      btnGetPaperUser.SetFocus;
      Exit;
    end;
    if (edtBeginNumber.Text='') then
    begin
      MessageDlg('請輸入起始流水號!', mtWarning, [mbOk],0);
      edtBeginNumber.SetFocus;
      Exit;
    end;


    if (nG_NewBeginNum>nG_OriginalBeginNum) and (nG_NewBeginNum<=nG_OriginalEndNum) then
    begin
      if MessageDlg('確定執行此項作業?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
      begin
        aValue := TUstr.ToROCYMD( fraChineseYMD3.getYMD );
        sL_Result := dtmMain3.reusePaper(dG_OriginalGetPaperDate, sG_OriginalEmpNo, sG_OriginalEmpName, sG_BeginNum, sG_OriginalSeq, sG_CompCode,  sG_EmpNo, edtName.Text, sG_Prefix, edtBeginNumber.Text, lblEndNumber.Caption, edtName.Text,
          aValue, sG_OriginalCount, lblCounts.Caption);
        if sL_Result='' then
        begin
          MessageDlg('處理完成!', mtInformation, [mbOK], 0);
          Close;
        end
        else
          MessageDlg(sL_Result, mtError, [mbOK], 0);
      end;  
    end
    else
    begin
      Messagedlg('起始流水號應該介於 '+ IntToStr(nG_OriginalBeginNum) + ' 與 ' + IntToStr(nG_OriginalEndNum) + ' 之間', mtWarning, [mbOK],0);
      edtBeginNumber.SetFocus;
      Exit;
      
    end;
end;

procedure TfrmSo8C60.edtBeginNumberExit(Sender: TObject);
begin
    nG_NewBeginNum := StrToIntDef(edtBeginNumber.Text,-9999);
    if  nG_NewBeginNum = -9999 then
    begin
      MessageDlg('起始流水號應該為數字!', mtWarning, [mbOk],0);
      edtBeginNumber.SetFocus;
      Exit;
    end;
    lblCounts.Caption := IntToStr(nG_OriginalEndNum - nG_NewBeginNum + 1);
end;

end.
