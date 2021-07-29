unit frmSO8C20U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, Buttons, DB, Grids, DBGrids;

type
  TfrmSO8C20 = class(TForm)
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
  frmSO8C20: TfrmSO8C20;

implementation

uses dtmMain3U, Ustru, UdateTimeu, frmMainMenuU;

{$R *.dfm}

procedure TfrmSO8C20.FormShow(Sender: TObject);
begin
    self.Caption := frmMainMenu.setCaption('SO8C20','��}��ں޲z-��ڰ��o�@�~');
end;

procedure TfrmSO8C20.BitBtn2Click(Sender: TObject);
begin
    Close;
end;

procedure TfrmSO8C20.btnQueryClick(Sender: TObject);
var
    L_DataSet : TDataSet;
begin
    if edtPaperSNum.Text = '' then
    begin
      edtPaperSNum.SetFocus;
      MessageDlg('�п�J�����o���渹!', mtWarning, [mbOK],0);
      Exit;
    end;

    if edtPaperENum.Text = '' then
    begin
      edtPaperENum.SetFocus;
      MessageDlg('�п�J�����o���渹!', mtWarning, [mbOK],0);
      Exit;
    end;

    L_DataSet := dtmMain3.getPaperData(edtPaperSNum.Text, edtPaperENum.Text);
    if L_DataSet.RecordCount>0 then
    begin
      dsrPaperData.DataSet := L_DataSet;
      L_DataSet.FieldByName('PAPERNUM').DisplayLabel := '�渹';
      L_DataSet.FieldByName('EMPNAME').DisplayWidth := 10;
      L_DataSet.FieldByName('EMPNAME').DisplayLabel := '���H��';
      L_DataSet.FieldByName('GETPAPERDATE').DisplayLabel := '�����';
      L_DataSet.FieldByName('BILLNO').DisplayLabel := '��ڽs��';
    end
    else
      MessageDlg('�L�i�Ѱ��o���!', mtInformation, [mbOK],0);
  
end;

procedure TfrmSO8C20.BitBtn1Click(Sender: TObject);
var
    sL_PaperSNum, sL_PaperENum : String;
begin
    if (dsrPaperData.DataSet = nil ) or
       (dsrPaperData.DataSet.State=dsInactive) or
       (dsrPaperData.DataSet.RecordCount = 0) then
    begin
      edtPaperSNum.SetFocus;
      MessageDlg('�Х��d�ߥX�����o���渹!', mtInformation, [mbOK],0);
      Exit;
    end;
       
    if canObsolete(dbgPaperData.DataSource.DataSet) then
    begin
      sL_PaperSNum := TUstr.AddString(edtPaperSNum.Text,'0',false,PAPER_NUMBER_LENGTH);
      sL_PaperENum := TUstr.AddString(edtPaperENum.Text,'0',false,PAPER_NUMBER_LENGTH);
//      if MessageDlg('�O�_�T�{���o������? [�渹:' + sL_PaperSNum + '~' + sL_PaperENum + ']', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
      if MessageDlg('�O�_�T�{���o������?' + '[ �@ ' + IntToStr(dsrPaperData.DataSet.recordCount) + ' �i ]', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
      begin
        if dtmMain3.disablePaperData(sL_PaperSNum, sL_PaperENum) then
          MessageDlg('�@�o����!', mtInformation, [mbOK],0);
      end;
    end
    else
    begin
      MessageDlg('�w����ڽs��,�L�k���o!', mtInformation, [mbOK],0);
    end;

end;

procedure TfrmSO8C20.FormClose(Sender: TObject; var Action: TCloseAction);
begin
    Action := caFree;
end;

function TfrmSO8C20.canObsolete(I_DataSet: TDataSet): boolean;
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
