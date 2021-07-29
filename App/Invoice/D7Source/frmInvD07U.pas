unit frmInvD07U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, Mask, ExtCtrls, fraYMU, cxButtons, Menus,
  cxLookAndFeelPainters, DB, ADODB;

type
  TfrmInvD07 = class(TForm)
    Panel1: TPanel;
    Label5: TLabel;
    Panel2: TPanel;
    fraYM: TfraYM;
    medYear: TMaskEdit;
    rdbYearMonth: TRadioButton;
    rdbYear: TRadioButton;
    btnOk: TcxButton;
    btnClose: TcxButton;
    btnImport: TcxButton;
    adoQry: TADOQuery;
    dlgOpen1: TOpenDialog;
    adoCon: TADOConnection;
    procedure FormShow(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure btnOKClick(Sender: TObject);
    procedure rdbYearMonthClick(Sender: TObject);
    procedure rdbYearClick(Sender: TObject);
    procedure btnImportClick(Sender: TObject);
  private
    FLOGINOK: BOOLEAN;
    procedure SetLOGINOK(const Value: BOOLEAN);
    { Private declarations }
    function isDataOK : Boolean;
    function importCSV(csvFileName:String):Boolean;
  public
    { Public declarations }
    PROPERTY LOGINOK:BOOLEAN read FLOGINOK write SetLOGINOK;
  end;

var
  frmInvD07: TfrmInvD07;

implementation

uses frmMainU, dtmMainU, dtmMainJU, frmInvD07_1U, Ustru,cbUtilis,frmInvD07_2U ;

{$R *.dfm}

{ TfrmInvD07 }

procedure TfrmInvD07.SetLOGINOK(const Value: BOOLEAN);
begin
  FLOGINOK := Value;
end;

procedure TfrmInvD07.FormShow(Sender: TObject);
begin
  self.Caption := frmMain.GetFormTitleString( 'D07', '發票字軌維護' );
end;

procedure TfrmInvD07.btnExitClick(Sender: TObject);
begin
    Close;
end;

function TfrmInvD07.isDataOK: Boolean;
var
    sL_Year,sL_YM : String;
begin
    Result := true;

    sL_YM := Trim(fraYM.getYM);
    if rdbYearMonth.Checked then
    begin
      if sL_YM = EMPTY_YM_STR  then
      begin
        MessageDlg('請輸入發票年月',mtError,[mbOk],0);
        fraYM.mseYM.SetFocus;
        Result := false;
        Exit;
      end;
    end;

    sL_Year := Trim(medYear.Text);
    if rdbYear.Checked then
    begin
      if sL_Year = ''  then
      begin
        MessageDlg('請輸入年度條件',mtError,[mbOk],0);
        medYear.SetFocus;
        Result := false;
        Exit;
      end
      else if Length(sL_Year) <> 4  then
      begin
        MessageDlg('年度條件格式錯誤!(範例:2004)',mtError,[mbOk],0);
        medYear.SetFocus;
        Result := false;
        Exit;
      end;
    end;


end;

procedure TfrmInvD07.btnOKClick(Sender: TObject);
var
    sL_CompID,sL_Year,sL_YM : String;
begin
    if isDataOK then
    begin
      sL_CompID := dtmMain.getCompID;
      sL_Year := Trim(medYear.Text);
      sL_YM := TUstr.replaceStr(Trim(fraYM.getYM),'/','');
      if sL_YM <> '' then
      begin
        sL_YM := dtmMainJ.changePrefixYM(sL_YM);
        fraYM.setYM(Copy(sL_YM,1,4) + '/' + Copy(sL_YM,5,2));
      end;

      dtmMainJ.getInv099Data(sL_CompID,sL_Year,sL_YM);

      frmInvD07_1 := TfrmInvD07_1.Create(Application);
      frmInvD07_1.sG_CompID := sL_CompID;
      frmInvD07_1.sG_QueryYear := sL_Year;
      frmInvD07_1.sG_QueryYM := sL_YM;
      frmInvD07_1.ShowModal;
      frmInvD07_1.Free;
    end;
end;

procedure TfrmInvD07.rdbYearMonthClick(Sender: TObject);
begin
    medYear.Clear;
    fraYM.mseYM.SetFocus;
end;

procedure TfrmInvD07.rdbYearClick(Sender: TObject);
begin
    fraYM.mseYM.Clear;
    medYear.SetFocus;
end;

procedure TfrmInvD07.btnImportClick(Sender: TObject);
begin
  if dlgOpen1.Execute then
  begin
     importCSV(dlgOpen1.FileName);
     Exit;
  end;
end;

function TfrmInvD07.importCSV(csvFileName: String): Boolean;
  var connectString: String;
  test : TfrmInvD07_2;
  frmInvD07_2 : TfrmInvD07_2;
begin
  dtmMain.SetWaitingCursor;
  try
      adoCon.Connected := False;
      connectString := 'Provider=Microsoft.Jet.OLEDB.4.0;' +
                      'Data Source=' + ExtractFilePath(csvFileName) +';' +
                      'Extended Properties=Text;Persist Security Info=False';
      adoCon.ConnectionString := connectString;
      adoCon.Connected := True;
      adoQry.Connection := adoCon;
      adoQry.Close;
      adoQry.SQL.Clear;
      adoQry.SQL.Add('SELECT * FROM [' + ExtractFileName(csvFileName)+ ']');
      adoQry.Open;
      dtmMain.SetDefaultCursor;
      if adoQry.RecordCount > 0 then
      begin
        adoQry.First;
        frmInvD07_2 := TfrmInvD07_2.Create(adoQry);
        frmInvD07_2.ShowModal;
        frmInvD07_2.Free;
      end else
      begin
        MessageDlg('無任何資料可供匯入！',mtInformation,[mbOK],0);
        dtmMain.SetWaitingCursor;
      end;
      Result := true;
  except
    on E: Exception do
    begin
      dtmMain.SetDefaultCursor;
      ErrorMsg( Format( 'CSV檔案連結失敗, 訊息:%s', [E.Message] ) );
      Exit;
    end;
  end;

end;

end.
