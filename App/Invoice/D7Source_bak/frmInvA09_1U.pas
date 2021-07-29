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
    self.Caption := frmMain.GetFormTitleString('A09_1','發票資料檢核結果');

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
  if SelectDirectory( 'Excel存放路徑', EmptyStr, aPath ) then
    edtExcelPath.Text := aPath;
end;

procedure TfrmInvA09_1.btnToExcel_1Click(Sender: TObject);
var
    sL_FileName,sL_ExcelSavePath,sL_CurrDateTime,sL_ExcelTitle : String;
begin
    //統一編號檢查結果
    sL_ExcelSavePath := Trim(edtExcelPath.Text);
    sL_CurrDateTime := TUstr.replaceStr(TUstr.replaceStr(TUstr.replaceStr(DateTimeToStr(now),'/',''),':',''),' ','');
    sL_ExcelTitle := '報表名稱:統編檢查結果';
    sL_ExcelTitle := sL_ExcelTitle + '@公司名稱:' + dtmMain.getCompName;
    sL_ExcelTitle := sL_ExcelTitle + '@發票年月:' + sG_InvYM;

    sL_FileName := sL_ExcelSavePath + '\統一編號檢查結果_' + dtmMain.getCompName + '_' + sL_CurrDateTime + '.XLS';

    if dtmMainJ.cdsCheckBusinessID.RecordCount > 0 then
    begin
      dtmMainJ.cdsCheckBusinessID.DisableControls;
      DataSetToXLS(dtmMainJ.cdsCheckBusinessID,sL_FileName,sL_ExcelTitle,'');
      dtmMainJ.cdsCheckBusinessID.EnableControls;

      MessageDlg('轉Excel完成,檔案存放於:' + sL_FileName,mtInformation,[mbOk],0);
    end;
end;

procedure TfrmInvA09_1.btnToExcel_2Click(Sender: TObject);
var
    sL_FileName,sL_ExcelSavePath,sL_CurrDateTime,sL_ExcelTitle : String;
begin
    //跳號檢查結果
    sL_ExcelSavePath := Trim(edtExcelPath.Text);
    sL_CurrDateTime := TUstr.replaceStr(TUstr.replaceStr(TUstr.replaceStr(DateTimeToStr(now),'/',''),':',''),' ','');
    sL_ExcelTitle := '報表名稱:跳號檢查結果';
    sL_ExcelTitle := sL_ExcelTitle + '@公司名稱:' + dtmMain.getCompName;
    sL_ExcelTitle := sL_ExcelTitle + '@發票年月:' + sG_InvYM;
    sL_ExcelTitle := sL_ExcelTitle + '@字軌:' + sG_Prefix;


    sL_FileName := sL_ExcelSavePath + '\跳號檢查結果_' + dtmMain.getCompName + '_' + sL_CurrDateTime + '.XLS';

    if dtmMainJ.cdsCheckJumpInvID.RecordCount > 0 then
    begin
      dtmMainJ.cdsCheckJumpInvID.DisableControls;
      DataSetToXLS(dtmMainJ.cdsCheckJumpInvID,sL_FileName,sL_ExcelTitle,'');
      dtmMainJ.cdsCheckJumpInvID.EnableControls;

      MessageDlg('轉Excel完成,檔案存放於:' + sL_FileName,mtInformation,[mbOk],0);
    end;
end;

procedure TfrmInvA09_1.btnToExcel_3Click(Sender: TObject);
var
    sL_FileName,sL_ExcelSavePath,sL_CurrDateTime,sL_ExcelTitle : String;
begin
    //主副檔金額檢核結果
    sL_ExcelSavePath := Trim(edtExcelPath.Text);
    sL_CurrDateTime := TUstr.replaceStr(TUstr.replaceStr(TUstr.replaceStr(DateTimeToStr(now),'/',''),':',''),' ','');
    sL_ExcelTitle := '報表名稱:主副檔金額檢核結果';
    sL_ExcelTitle := sL_ExcelTitle + '@公司名稱:' + dtmMain.getCompName;
    sL_ExcelTitle := sL_ExcelTitle + '@發票日期:' + sG_SInvDate1 + ' ~ ' + sG_EInvDate1;
    if bG_IncludeObsolete then
      sL_ExcelTitle := sL_ExcelTitle + '@是否包含廢發票:是'
    else
      sL_ExcelTitle := sL_ExcelTitle + '@是否包含廢發票:否';

    sL_FileName := sL_ExcelSavePath + '\主副檔金額檢核結果_' + dtmMain.getCompName + '_' + sL_CurrDateTime + '.XLS';

    if dtmMainJ.cdsCheckAmount.RecordCount > 0 then
    begin
      dtmMainJ.cdsCheckAmount.DisableControls;
      DataSetToXLS(dtmMainJ.cdsCheckAmount,sL_FileName,sL_ExcelTitle,'');
      dtmMainJ.cdsCheckAmount.EnableControls;

      MessageDlg('轉Excel完成,檔案存放於:' + sL_FileName,mtInformation,[mbOk],0);
    end;

end;

procedure TfrmInvA09_1.btnToExcel_4Click(Sender: TObject);
var
    sL_FileName,sL_ExcelSavePath,sL_CurrDateTime,sL_ExcelTitle : String;
begin
    //發票開立與客服資料比對檢核結果
    sL_ExcelSavePath := Trim(edtExcelPath.Text);
    sL_CurrDateTime := TUstr.replaceStr(TUstr.replaceStr(TUstr.replaceStr(DateTimeToStr(now),'/',''),':',''),' ','');
    sL_ExcelTitle := '報表名稱:發票開立與客服資料比對檢核結果';
    sL_ExcelTitle := sL_ExcelTitle + '@公司名稱:' + dtmMain.getCompName;
    sL_ExcelTitle := sL_ExcelTitle + '@發票日期:' + sG_SInvDate2 + ' ~ ' + sG_EInvDate2;

    sL_FileName := sL_ExcelSavePath + '\發票開立與客服資料比對檢核結果_' + dtmMain.getCompName + '_' + sL_CurrDateTime + '.XLS';

    if dtmMainJ.cdsCheckMisData.RecordCount > 0 then
    begin
      dtmMainJ.cdsCheckMisData.DisableControls;
      DataSetToXLS(dtmMainJ.cdsCheckMisData,sL_FileName,sL_ExcelTitle,'');
      dtmMainJ.cdsCheckMisData.EnableControls;

      MessageDlg('轉Excel完成,檔案存放於:' + sL_FileName,mtInformation,[mbOk],0);
    end;
end;

end.
