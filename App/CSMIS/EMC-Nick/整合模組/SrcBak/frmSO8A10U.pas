unit frmSO8A10U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls, Grids, DBGrids, ExtCtrls, DBCtrls, Mask,DB,
  Buttons;

type
  TfrmSO8A10 = class(TForm)
    GroupBox1: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    Label4: TLabel;
    dedFormulaId: TDBEdit;
    DBEdit2: TDBEdit;
    DBCheckBox1: TDBCheckBox;
    Panel1: TPanel;
    dgr1: TDBGrid;
    StatusBar1: TStatusBar;
    Label3: TLabel;
    DBEdit1: TDBEdit;
    GroupBox2: TGroupBox;
    btnSave: TBitBtn;
    btnPrint: TBitBtn;
    btnInsert: TBitBtn;
    btnEdit: TBitBtn;
    btnDelete: TBitBtn;
    btnCloseCancel: TBitBtn;
    Edit1: TEdit;
    dsrSo110: TDataSource;
    Label5: TLabel;
    dsrRef_No: TDataSource;
    CobRefNo: TComboBox;
    procedure btnEditClick(Sender: TObject);
    procedure btn_Control(btn_Save,btn_Print,btn_Insert,btn_Edit,btn_Delete:Boolean);
    procedure CurrentCursorAndCdsCount;
    procedure FormCreate(Sender: TObject);
    procedure dgr1CellClick(Column: TColumn);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure dgr1Enter(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure btnInsertClick(Sender: TObject);
    procedure btnCloseCancelClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure dgr1DblClick(Sender: TObject);
    procedure CobRefNoChange(Sender: TObject);
    procedure dsrSo110DataChange(Sender: TObject; Field: TField);
  private
    { Private declarations }
    function getCobRefNoCode:String;
    function getCobRefNoName:String;
  public
    { Public declarations }
  end;

var
  frmSO8A10: TfrmSO8A10;

implementation

uses  UCommonU, Ustru, dtmMain2U, frmMainMenuU;

{$R *.dfm}

procedure TfrmSO8A10.btnEditClick(Sender: TObject);
begin
    dtmMain2.cdsSo110.Edit;
    btn_Control(True,False,False,False,False);
    Edit1.Text:='修改';
    GroupBox1.Enabled:=True;
    Panel1.Enabled:=False;
    btnCloseCancel.Caption:='取消';
    dedFormulaId.SetFocus;
end;

procedure TfrmSO8A10.btn_Control(btn_Save, btn_Print, btn_Insert,
  btn_Edit,btn_Delete: Boolean);
begin
    btnSave.Enabled:=btn_Save;
    btnPrint.Enabled:=btn_Print;
    btnInsert.Enabled:=btn_Insert;
    btnEdit.Enabled:=btn_Edit;
    btnDelete.Enabled:=btn_Delete;
end;

procedure TfrmSO8A10.CurrentCursorAndCdsCount;
begin
    StatusBar1.Panels[0].Text:='                    '+'                 '+'                  '
    +'             '+'          '+'                    '+'              '+
    IntToStr(dtmMain2.cdsSo110.RecNo)+' '+'/'+' '+IntToStr(dtmMain2.cdsSo110.RecordCount);

end;

procedure TfrmSO8A10.FormCreate(Sender: TObject);
begin
    dtmMain2.cdsSo110.Active:=True;
    dtmMain2.cdsSo110.First;
    Edit1.Text:='顯示';
    CurrentCursorAndCdsCount;
//    CobReFNo.Text:=dtmMain2.cdsSo110.fieldByName('Ref_No').AsString+'-'+dtmMain2.cdsSo110.FieldByName('REFNO_NAME').AsString;

    TUCommonFun.AddObjectToComboBox(CobRefNO.Items, dtmMain2.cdsRefNo,not INSERT_NO_DATA_ITEM,'CodeNo','Description');
end;

procedure TfrmSO8A10.dgr1CellClick(Column: TColumn);
begin
    CurrentCursorAndCdsCount;
end;

procedure TfrmSO8A10.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
    if key=113 then
    btnSaveClick(Sender);
    if key=116 then
    btnPrintClick(Sender);
    if key=117 then
    btnInsertClick(Sender);
    if key=122 then
    btnEditClick(Sender);
    if key=88 then
    btnCloseCancelClick(Sender);
end;

procedure TfrmSO8A10.dgr1Enter(Sender: TObject);
begin
    CurrentCursorAndCdsCount;
end;

procedure TfrmSO8A10.btnPrintClick(Sender: TObject);
begin
    Close;
end;

procedure TfrmSO8A10.btnInsertClick(Sender: TObject);
begin
    dtmMain2.cdsSo110.Append;
    btn_Control(True,False,False,False,False);

    Edit1.Text:='新增';
    CobRefNo.Text:='';
    dtmMain2.cdsSo110.FieldByName('STOPFLAG').AsString:='0';
    GroupBox1.Enabled:=True;
    dedFormulaId.SetFocus;
    btnCloseCancel.Caption:='取消';
    Panel1.Enabled:=False;

end;

procedure TfrmSO8A10.btnCloseCancelClick(Sender: TObject);
begin
    if (Edit1.Text='顯示') and (btnCloseCancel.Caption='(&X)結束') then
    begin
      dtmMain2.cdsSo110.Close;
      Close;
    end;
    if (Edit1.Text='新增') or (Edit1.Text='修改') then
    begin
      dtmMain2.cdsSo110.Cancel;
      btn_Control(False,True,True,True,True);
      GroupBox1.Enabled:=False;
      Panel1.Enabled:=True;
      Edit1.Text:='顯示';
      btnCloseCancel.Caption:='(&X)結束'
    end;
end;

procedure TfrmSO8A10.btnSaveClick(Sender: TObject);
var sL_SQL:String;
begin
    if dtmMain2.cdsSo110.State=dsBrowse then
        begin
        //ShowMessage('請先按新增或修改');
        Exit;
        end;
    if dedFormulaId.Text='' then
        begin
        MessageDlg('請輸入公式代碼',mtError, [mbOK],0);

        dedFormulaId.SetFocus;
        Exit;
        end;

    if CobRefNo.ItemIndex=-1 then
    begin
      MessageDlg('請點選公式參考代碼',mtError, [mbOK],0);

      CobRefNo.SetFocus;
      Exit;
    end;
//*******************************************************
     //check duplicate Provider_id
    if dtmMain2.cdsSo110.State=dsInsert then
    begin
        with dtmMain2.cdsSqlStatement do
        begin
        Close;
        sL_SQL:='SELECT count(*) CNT FROM SO110 WHERE FORMULA_ID='''+Trim(dedFormulaId.Text)+'''';
        CommandText:=sL_SQL;
        //CommandText:='SELECT count(*) CNT FROM SO110 WHERE FORMULA_ID=:FORMULA_ID';
        //PARAMS.ParamByName('FORMULA_ID').AsString:=Trim(dedFormulaId.Text);
        OPEN;
        end;
        if dtmMain2.cdsSqlStatement.FieldByName('CNT').Value>0 then
        begin
        MessageDlg('公式代碼重覆',mtError, [mbOK],0);

        dtmMain2.cdsSo110.Cancel;
        dedFormulaId.SetFocus;
        btn_Control(True,True,True,True,True);
        Exit;
        end;
    end;

    if dtmMain2.cdsSo110.State=dsEdit then
    begin
        if dtmMain2.cdsSo110.FieldByName('FORMULA_ID').NewValue<>dtmMain2.cdsSo110.FieldByName('FORMULA_ID').OldValue then
        begin
            with dtmMain2.cdsSqlStatement do
            begin
            Close;
            sL_SQL:='SELECT count(*) CNT FROM SO110 WHERE FORMULA_ID='''+Trim(dedFormulaId.Text)+'''';
            CommandText:=sL_SQL;
            //CommandText:='SELECT count(*) CNT FROM SO110 WHERE FORMULA_ID=:FORMULA_ID';
            //PARAMS.ParamByName('FORMULA_ID').AsString:=Trim(dedFormulaId.Text);
            OPEN;
            end;
            if dtmMain2.cdsSqlStatement.FieldByName('CNT').Value>0 then
            begin
            MessageDlg('公式代碼重覆',mtError, [mbOK],0);
            dtmMain2.cdsSo110.Cancel;
            dedFormulaId.SetFocus;
            btn_Control(True,True,True,True,True);
            Exit;
            end;
        end;
        if dtmMain2.cdsSo110.FieldByName('FORMULA_ID').NewValue=dtmMain2.cdsSo110.FieldByName('FORMULA_ID').OldValue then
        begin
            with dtmMain2.cdsSqlStatement do
            begin
            Close;
            sL_SQL:='SELECT count(*) CNT FROM SO110 WHERE FORMULA_ID='''+Trim(dedFormulaId.Text)+'''';
            CommandText:=sL_SQL;
            //CommandText:='SELECT count(*) CNT FROM SO110 WHERE FORMULA_ID=:FORMULA_ID';
            //PARAMS.ParamByName('FORMULA_ID').AsString:=Trim(dedFormulaId.Text);
            OPEN;
            end;
            if dtmMain2.cdsSqlStatement.FieldByName('CNT').Value>1 then
            begin
            MessageDlg('公式代碼重覆',mtError, [mbOK],0);
            dtmMain2.cdsSo110.Cancel;
            btn_Control(True,True,True,True,True);            
            dedFormulaId.SetFocus;
            Exit;
            end;
        end;
    end;

    dtmMain2.cdsSo110.FieldByName('OPERATOR').AsString:=frmMainMenu.sG_User;
    dtmMain2.cdsSo110.FieldByName('UpdTime').AsString:=IntToStr(StrToInt(FormatDateTime('YYYY',now))-1911)+ FormatDateTime('/MM/DD hh:mm:ss',now);
    try
    dtmMain2.cdsSo110.ApplyUpdates(-1);
    dtmMain2.cdsSo110.Close;
    dtmMain2.cdsSo110.Open;
    except
    MessageDlg('存檔失敗',mtError, [mbOK],0);
    Exit;
    end;
    btn_Control(False,True,True,True,True);
    Edit1.Text:='顯示';
    CurrentCursorAndCdsCount;
    GroupBox1.Enabled:=False;
    panel1.Enabled:=True;
    btnCloseCancel.Caption:='(&X)結束'

end;

procedure TfrmSO8A10.btnDeleteClick(Sender: TObject);
var sL_DelSo110:String;
begin
    if dtmMain2.cdsSo110.IsEmpty then
    begin
      MessageDlg('已無資料',mtInformation, [mbOK],0);
      Exit;
    end;
    sL_DelSo110:='Delete from So110 where FORMULA_ID='''+dtmMain2.cdsSo110.FieldByName('FORMULA_ID').AsString+'''';

    if MessageDlg('您確定要刪除此筆資料?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      with dtmMain2.cdsSqlDelete do
      begin
        Close;
        CommandText:=sL_DelSo110;
        Execute;
      end;
        dtmMain2.cdsSo110.Close;
        dtmMain2.cdsSo110.Open;
    end;

    CurrentCursorAndCdsCount;
    btnCloseCancel.Caption:='(&X)結束'

end;

procedure TfrmSO8A10.FormShow(Sender: TObject);
var
    sL_RefNo : String;
begin
    self.Caption := frmMainMenu.setCaption('SO8A10','[拆帳公式代碼]');

    sL_RefNo := dtmMain2.cdsSo110.FieldByName('REF_NO').AsString;
    TUCommonFun.setComboDefaultNdx(CobRefNO, sL_RefNo);
end;

procedure TfrmSO8A10.dgr1DblClick(Sender: TObject);
begin
    btnEditClick(Sender);
end;



procedure TfrmSO8A10.CobRefNoChange(Sender: TObject);
var nL_CobRefNo:Integer;
begin
     nL_CobRefNo:=StrToInt(copy(CobRefNo.Text,1,1));

     case nL_CobRefNo of
       1:dtmMain2.cdsSo110.FieldByName('REF_NO').AsString:='1';
       2:dtmMain2.cdsSo110.FieldByName('REF_NO').AsString:='2';
       3:dtmMain2.cdsSo110.FieldByName('REF_NO').AsString:='3';
       4:dtmMain2.cdsSo110.FieldByName('REF_NO').AsString:='4';
       5:dtmMain2.cdsSo110.FieldByName('REF_NO').AsString:='5';
     end;

end;

procedure TfrmSO8A10.dsrSo110DataChange(Sender: TObject; Field: TField);
var sL_RefNoCode:String;
begin
    if dtmMain2.cdsSo110.State=dsBrowse then
    begin
      sL_RefNoCode:=dsrSo110.DataSet.FieldByName('Ref_No').AsString;
      TUCommonFun.setComboDefaultNdx(CobRefNo, sL_RefNoCode);
    end;
end;

function TfrmSO8A10.getCobRefNoCode: String;
var
    L_RefnoCodeStrList : TStringList;
    sL_RefnoCodeAndName : String;
begin
    sL_RefnoCodeAndName := CobRefNo.Text;

    L_RefnoCodeStrList := TStringList.Create;
    L_RefnoCodeStrList := TUstr.ParseStrings(CobRefNo.Text,'-');
    getCobRefNoCode := L_RefnoCodeStrList.Strings[0];

    L_RefnoCodeStrList.Free;
end;

function TfrmSO8A10.getCobRefNoName: String;
var
    L_RefNoNameStrList : TStringList;
    sL_RefNoCodeAndName : String;
begin
    sL_RefNoCodeAndName := CobRefNo.Text;

    L_RefNoNameStrList := TStringList.Create;
    L_RefNoNameStrList := TUstr.ParseStrings(CobRefNo.Text,'-');
    getCobRefNoName := L_RefNoNameStrList.Strings[1];

    L_RefNoNameStrList.Free;
end;



end.
