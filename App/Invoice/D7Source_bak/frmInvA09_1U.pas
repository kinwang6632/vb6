unit frmInvA09_1U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, ComCtrls, DB, Grids, DBGrids, Buttons, FileCtrl;

type
  TfrmInvA09_1 = class(TForm)
    Panel1: TPanel;
    Label6: TLabel;
    dsrCheckBusinessID: TDataSource;
    Panel2: TPanel;
    PageControl1: TPageControl;
    tbsChkNumber: TTabSheet;
    DBGrid1: TDBGrid;
    tbsChkAmount: TTabSheet;
    tbsChkMisData: TTabSheet;
    btnExit: TBitBtn;
    btnToExcel_1: TBitBtn;
    btnToExcel_3: TBitBtn;
    DBGrid2: TDBGrid;
    DBGrid3: TDBGrid;
    dsrCheckJumpInvID: TDataSource;
    DBGrid4: TDBGrid;
    dsrCheckAmount: TDataSource;
    dsrCheckMisData: TDataSource;
    btnToExcel_2: TBitBtn;
    btnToExcel_4: TBitBtn;
    btnTransferPath: TButton;
    edtExcelPath: TEdit;
    Label11: TLabel;
    procedure FormShow(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure btnTransferPathClick(Sender: TObject);
    procedure btnToExcel_1Click(Sender: TObject);
    procedure btnToExcel_2Click(Sender: TObject);
    procedure btnToExcel_3Click(Sender: TObject);
    procedure btnToExcel_4Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    sG_InvYM,sG_Prefix,sG_SInvDate1,sG_EInvDate1,sG_SInvDate2,sG_EInvDate2 : String;
    bG_IncludeObsolete : Boolean;
  end;

var
  frmInvA09_1: TfrmInvA09_1;

implementation

uses frmMainU, dtmMainJU, dtmMainU, Ustru, XLSFile;

{$R *.dfm}

procedure TfrmInvA09_1.FormShow(Sender: TObject);
begin
    self.Caption := frmMain.GetFormTitleString('A09_1','�o������ˮֵ��G');

    if dtmMainJ.cdsCheckBusinessID.RecordCount=0 then
      btnToExcel_1.Enabled := false
    else
      btnToExcel_1.Enabled := true;


    if dtmMainJ.cdsCheckJumpInvID.RecordCount=0 then
      btnToExcel_2.Enabled := false
    else
      btnToExcel_2.Enabled := true;


    if dtmMainJ.cdsCheckAmount.RecordCount=0 then
      btnToExcel_3.Enabled := false
    else
      btnToExcel_3.Enabled := true;



    if dtmMainJ.cdsCheckMisData.RecordCount=0 then
      btnToExcel_4.Enabled := false
    else
      btnToExcel_4.Enabled := true;

end;

procedure TfrmInvA09_1.btnExitClick(Sender: TObject);
begin
    Close;
end;

procedure TfrmInvA09_1.btnTransferPathClick(Sender: TObject);
var
  aPath : String;
begin
  aPath := edtExcelPath.Text;
  if SelectDirectory( 'Excel�s����|', EmptyStr, aPath ) then
    edtExcelPath.Text := aPath;
end;

procedure TfrmInvA09_1.btnToExcel_1Click(Sender: TObject);
var
    sL_FileName,sL_ExcelSavePath,sL_CurrDateTime,sL_ExcelTitle : String;
begin
    //�Τ@�s���ˬd���G
    sL_ExcelSavePath := Trim(edtExcelPath.Text);
    sL_CurrDateTime := TUstr.replaceStr(TUstr.replaceStr(TUstr.replaceStr(DateTimeToStr(now),'/',''),':',''),' ','');
    sL_ExcelTitle := '����W��:�νs�ˬd���G';
    sL_ExcelTitle := sL_ExcelTitle + '@���q�W��:' + dtmMain.getCompName;
    sL_ExcelTitle := sL_ExcelTitle + '@�o���~��:' + sG_InvYM;

    sL_FileName := sL_ExcelSavePath + '\�Τ@�s���ˬd���G_' + dtmMain.getCompName + '_' + sL_CurrDateTime + '.XLS';

    if dtmMainJ.cdsCheckBusinessID.RecordCount > 0 then
    begin
      dtmMainJ.cdsCheckBusinessID.DisableControls;
      DataSetToXLS(dtmMainJ.cdsCheckBusinessID,sL_FileName,sL_ExcelTitle,'');
      dtmMainJ.cdsCheckBusinessID.EnableControls;

      MessageDlg('��Excel����,�ɮצs���:' + sL_FileName,mtInformation,[mbOk],0);
    end;
end;

procedure TfrmInvA09_1.btnToExcel_2Click(Sender: TObject);
var
    sL_FileName,sL_ExcelSavePath,sL_CurrDateTime,sL_ExcelTitle : String;
begin
    //�����ˬd���G
    sL_ExcelSavePath := Trim(edtExcelPath.Text);
    sL_CurrDateTime := TUstr.replaceStr(TUstr.replaceStr(TUstr.replaceStr(DateTimeToStr(now),'/',''),':',''),' ','');
    sL_ExcelTitle := '����W��:�����ˬd���G';
    sL_ExcelTitle := sL_ExcelTitle + '@���q�W��:' + dtmMain.getCompName;
    sL_ExcelTitle := sL_ExcelTitle + '@�o���~��:' + sG_InvYM;
    sL_ExcelTitle := sL_ExcelTitle + '@�r�y:' + sG_Prefix;


    sL_FileName := sL_ExcelSavePath + '\�����ˬd���G_' + dtmMain.getCompName + '_' + sL_CurrDateTime + '.XLS';

    if dtmMainJ.cdsCheckJumpInvID.RecordCount > 0 then
    begin
      dtmMainJ.cdsCheckJumpInvID.DisableControls;
      DataSetToXLS(dtmMainJ.cdsCheckJumpInvID,sL_FileName,sL_ExcelTitle,'');
      dtmMainJ.cdsCheckJumpInvID.EnableControls;

      MessageDlg('��Excel����,�ɮצs���:' + sL_FileName,mtInformation,[mbOk],0);
    end;
end;

procedure TfrmInvA09_1.btnToExcel_3Click(Sender: TObject);
var
    sL_FileName,sL_ExcelSavePath,sL_CurrDateTime,sL_ExcelTitle : String;
begin
    //�D���ɪ��B�ˮֵ��G
    sL_ExcelSavePath := Trim(edtExcelPath.Text);
    sL_CurrDateTime := TUstr.replaceStr(TUstr.replaceStr(TUstr.replaceStr(DateTimeToStr(now),'/',''),':',''),' ','');
    sL_ExcelTitle := '����W��:�D���ɪ��B�ˮֵ��G';
    sL_ExcelTitle := sL_ExcelTitle + '@���q�W��:' + dtmMain.getCompName;
    sL_ExcelTitle := sL_ExcelTitle + '@�o�����:' + sG_SInvDate1 + ' ~ ' + sG_EInvDate1;
    if bG_IncludeObsolete then
      sL_ExcelTitle := sL_ExcelTitle + '@�O�_�]�t�o�o��:�O'
    else
      sL_ExcelTitle := sL_ExcelTitle + '@�O�_�]�t�o�o��:�_';

    sL_FileName := sL_ExcelSavePath + '\�D���ɪ��B�ˮֵ��G_' + dtmMain.getCompName + '_' + sL_CurrDateTime + '.XLS';

    if dtmMainJ.cdsCheckAmount.RecordCount > 0 then
    begin
      dtmMainJ.cdsCheckAmount.DisableControls;
      DataSetToXLS(dtmMainJ.cdsCheckAmount,sL_FileName,sL_ExcelTitle,'');
      dtmMainJ.cdsCheckAmount.EnableControls;

      MessageDlg('��Excel����,�ɮצs���:' + sL_FileName,mtInformation,[mbOk],0);
    end;

end;

procedure TfrmInvA09_1.btnToExcel_4Click(Sender: TObject);
var
    sL_FileName,sL_ExcelSavePath,sL_CurrDateTime,sL_ExcelTitle : String;
begin
    //�o���}�߻P�ȪA��Ƥ���ˮֵ��G
    sL_ExcelSavePath := Trim(edtExcelPath.Text);
    sL_CurrDateTime := TUstr.replaceStr(TUstr.replaceStr(TUstr.replaceStr(DateTimeToStr(now),'/',''),':',''),' ','');
    sL_ExcelTitle := '����W��:�o���}�߻P�ȪA��Ƥ���ˮֵ��G';
    sL_ExcelTitle := sL_ExcelTitle + '@���q�W��:' + dtmMain.getCompName;
    sL_ExcelTitle := sL_ExcelTitle + '@�o�����:' + sG_SInvDate2 + ' ~ ' + sG_EInvDate2;

    sL_FileName := sL_ExcelSavePath + '\�o���}�߻P�ȪA��Ƥ���ˮֵ��G_' + dtmMain.getCompName + '_' + sL_CurrDateTime + '.XLS';

    if dtmMainJ.cdsCheckMisData.RecordCount > 0 then
    begin
      dtmMainJ.cdsCheckMisData.DisableControls;
      DataSetToXLS(dtmMainJ.cdsCheckMisData,sL_FileName,sL_ExcelTitle,'');
      dtmMainJ.cdsCheckMisData.EnableControls;

      MessageDlg('��Excel����,�ɮצs���:' + sL_FileName,mtInformation,[mbOk],0);
    end;
end;

end.
