unit frmSO8A20U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls, Grids, DBGrids, ExtCtrls, DBCtrls, Mask,DB,ConvUtils,
  Buttons;

type
  TfrmSO8A20 = class(TForm)
    GroupBox1: TGroupBox;
    Panel1: TPanel;
    dgr1: TDBGrid;
    StatusBar1: TStatusBar;
    Label1: TLabel;
    dedProviderId: TDBEdit;
    Label2: TLabel;
    DBEdit2: TDBEdit;
    Label3: TLabel;
    DBEdit3: TDBEdit;
    Label4: TLabel;
    DBEdit4: TDBEdit;
    Label5: TLabel;
    DBEdit5: TDBEdit;
    Label6: TLabel;
    DBEdit6: TDBEdit;
    Label7: TLabel;
    DBCheckBox1: TDBCheckBox;
    Label8: TLabel;
    GroupBox2: TGroupBox;
    btnSave: TBitBtn;
    btnPrint: TBitBtn;
    btnInsert: TBitBtn;
    btnEdit: TBitBtn;
    btnDelete: TBitBtn;
    btnCloseCancel: TBitBtn;
    Edit1: TEdit;
    dsrSo113CT: TDataSource;
    CobAttribute: TComboBox;
    DataSource1: TDataSource;
    procedure btn_Control(btn_Save,btn_Print,btn_Insert,btn_Edit,btn_Cancel,btn_Delete:Boolean);
    procedure CurrentCursorAndCdsCount;
    procedure FormCreate(Sender: TObject);
    procedure dgr1CellClick(Column: TColumn);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure dgr1Enter(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure btnInsertClick(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure btnEditClick(Sender: TObject);
    procedure btnCloseCancelClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure dgr1DblClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure dsrSo113CTDataChange(Sender: TObject; Field: TField);
    procedure CobAttributeChange(Sender: TObject);
  private
    { Private declarations }
    function getCobAttributeCode:String;
    function getCobAttributeName:String;
  public
    { Public declarations }
  end;

var
  frmSO8A20: TfrmSO8A20;

implementation

uses UCommonU, Ustru, dtmMain2U, frmMainMenuU;

{$R *.dfm}

procedure TfrmSO8A20.btn_Control(btn_Save, btn_Print, btn_Insert,
  btn_Edit,btn_Cancel,btn_Delete: Boolean);
begin
    btnSave.Enabled:=btn_Save;
    btnPrint.Enabled:=btn_Print;
    btnInsert.Enabled:=btn_Insert;
    btnEdit.Enabled:=btn_Edit;
    btnDelete.Enabled:=btn_Delete;
end;

procedure TfrmSO8A20.CurrentCursorAndCdsCount;
begin
    StatusBar1.Panels[0].Text:='                    '+'                 '+'                  '
    +'             '+'          '+'                    '+'              '+
    IntToStr(dtmMain2.cdsSo113CT.RecNo)+' '+'/'+' '+IntToStr(dtmMain2.cdsSo113CT.RecordCount);


  
end;

procedure TfrmSO8A20.FormCreate(Sender: TObject);
begin
    dtmMain2.cdsSo113CT.Active:=True;
    dtmMain2.cdsSo113CT.First;
    Edit1.Text:='顯示';
    CurrentCursorAndCdsCount;

    //sL_AttCodeNo:=dsrSo113CT.DataSet.FieldByName('Attribute').AsString;
    //TUCommonFun.setComboDefaultNdx(CobAttribute, sL_AttCodeNo);
    CobAttribute.Text:=dtmMain2.cdsSo113CT.FieldByName('ATTRIBUTE').AsString+'-'+dtmMain2.cdsSo113CT.FieldByName('ATTRIBUTE_NAME').AsString;
end;

procedure TfrmSO8A20.dgr1CellClick(Column: TColumn);
begin
    CurrentCursorAndCdsCount;
end;

procedure TfrmSO8A20.FormKeyDown(Sender: TObject; var Key: Word;
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

procedure TfrmSO8A20.dgr1Enter(Sender: TObject);
begin
CurrentCursorAndCdsCount;
end;

procedure TfrmSO8A20.btnSaveClick(Sender: TObject);
var sL_SQL:String;
begin
    if dtmMain2.cdsSo113CT.State=dsBrowse then
        begin
        //ShowMessage('請先新增或修改');
        exit;
        end;
    if dedProviderId.Text='' then
        begin
        MessageDlg('請輸入頻道商代碼',mtError, [mbOK],0);
        dedProviderId.SetFocus;
        Exit;
        end;
    if dbEdit2.Text='' then
        begin
        MessageDlg('請輸入頻道商名稱',mtError, [mbOK],0);
        dbEdit2.SetFocus;
        Exit;
        end;
//***********************************
     //check duplicate Provider_id
    if dtmMain2.cdsSo113CT.State=dsInsert then
    begin
        with dtmMain2.cdsSqlStatement do
            begin
            Close;
            sL_SQL:='SELECT count(*) CNT FROM SO113 WHERE PROVIDER_ID='''+Trim(dedProviderId.Text)+'''';
            CommandText:=sL_SQL;
            //CommandText:='SELECT count(*) CNT FROM SO113 WHERE PROVIDER_ID=:PROVIDER_ID';
            //PARAMS.ParamByName('PROVIDER_ID').AsString:=Trim(dedProviderId.Text);
            OPEN;
            end;
        if dtmMain2.cdsSqlStatement.FieldByName('CNT').Value>0 then
            begin
            MessageDlg('頻道商代碼重覆',mtError, [mbOK],0);
            dedProviderId.SetFocus;
            btn_Control(True,True,True,True,True,True);
            Exit;
            end;
    end;

    if dtmMain2.cdsSo113CT.State=dsEdit then
    begin
        if dtmMain2.cdsSo113CT.FieldByName('PROVIDER_ID').NewValue<>dtmMain2.cdsSo113CT.FieldByName('PROVIDER_ID').OldValue then
        begin
            with dtmMain2.cdsSqlStatement do
            begin
            Close;
            sL_SQL:='SELECT count(*) CNT FROM SO113 WHERE PROVIDER_ID='''+Trim(dedProviderId.Text)+'''';
            CommandText:=sL_SQL;
            //CommandText:='SELECT count(*) CNT FROM SO113 WHERE PROVIDER_ID=:PROVIDER_ID';
            //PARAMS.ParamByName('PROVIDER_ID').AsString:=Trim(dedProviderId.Text);
            OPEN;
            end;
            if dtmMain2.cdsSqlStatement.FieldByName('CNT').Value>0 then
            begin
            MessageDlg('頻道商代碼重覆',mtError, [mbOK],0);
            dedProviderId.SetFocus;
            btn_Control(True,True,True,True,True,True);
            Exit;
            end;
        end;
        if dtmMain2.cdsSo113CT.FieldByName('PROVIDER_ID').NewValue=dtmMain2.cdsSo113CT.FieldByName('PROVIDER_ID').OldValue then
        begin
            with dtmMain2.cdsSqlStatement do
            begin
            Close;
            sL_SQL:='SELECT count(*) CNT FROM SO113 WHERE PROVIDER_ID='''+Trim(dedProviderId.Text)+'''';
            CommandText:=sL_SQL;
            //CommandText:='SELECT count(*) CNT FROM SO113 WHERE PROVIDER_ID=:PROVIDER_ID';
            //PARAMS.ParamByName('PROVIDER_ID').AsString:=Trim(dedProviderId.Text);

            OPEN;
            end;
            if dtmMain2.cdsSqlStatement.FieldByName('CNT').Value>1 then
            begin
            MessageDlg('頻道商代碼重覆',mtError, [mbOK],0);
            dedProviderId.SetFocus;
            btn_Control(True,True,True,True,True,True);
            Exit;
            end;
        end;
    end;


      dtmMain2.cdsSo113CT.FieldByName('OPERATOR').AsString:=frmMainMenu.sG_User;
      dtmMain2.cdsSo113CT.FieldByName('UpdTime').AsString:=IntToStr(StrToInt(FormatDateTime('YYYY',now))-1911)+ FormatDateTime('/MM/DD hh:mm:ss',now);
    if dtmMain2.cdsSo113CT.FieldByName('ATTRIBUTE').AsString='' then
    begin
       dtmMain2.cdsSo113CT.FieldByName('ATTRIBUTE').AsString:='2';
       TUCommonFun.setComboDefaultNdx(CobAttribute, '2');
    end;

    try
     // dtmMain2.cdsSo113CT.FieldByName('ATTRIBUTE').AsString:=getCobAttributeCode;
      //dtmMain2.cdsSo113CT.FieldByName('ATTRIBUTE_NAME').AsString:=getCobAttributeName;

       dtmMain2.saveToDB(dtmMain2.cdsSo113CT);
       dtmMain2.cdsSo113CT.Close;
       dtmMain2.cdsSo113CT.Open;

    except
        MessageDlg('存檔失敗',mtError, [mbOK],0);
        Exit;
    end;


    btn_Control(False,True,True,True,True,True);
    Edit1.Text:='顯示';
    CurrentCursorAndCdsCount;
    GroupBox1.Enabled:=False;
    panel1.Enabled:=True;
    btnCloseCancel.Caption:='(&X)結束';
    dtmMain2.cdsSo113CT.Close;
    dtmMain2.cdsSo113CT.Open;
    btnSave.Enabled:=False;
end;

procedure TfrmSO8A20.btnInsertClick(Sender: TObject);
begin
    dtmMain2.cdsSo113CT.Append;
    btn_Control(True,False,False,False,True,False);
    Edit1.Text:='新增';
    CobAttribute.Text:='';
    dtmMain2.cdsSo113CT.FieldByName('STOPFLAG').AsString:='0';
    GroupBox1.Enabled:=True;
    dedProviderId.SetFocus;
    btnCloseCancel.Caption:='取消';
    Panel1.Enabled:=False;

end;
procedure TfrmSO8A20.btnPrintClick(Sender: TObject);
begin
Close;
end;

procedure TfrmSO8A20.btnEditClick(Sender: TObject);
begin
    if dtmMain2.cdsSo113CT.IsEmpty then
    begin
    MessageDlg('已無資料',mtError, [mbOK],0);
    Exit;
    end;
    dtmMain2.cdsSo113CT.Edit;
    btn_Control(True,False,False,False,True,False);
    Edit1.Text:='修改';
    GroupBox1.Enabled:=True;
    Panel1.Enabled:=False;
    btnCloseCancel.Caption:='取消';
    dedProviderId.SetFocus;
end;

procedure TfrmSO8A20.btnCloseCancelClick(Sender: TObject);
begin
if (Edit1.Text='顯示') and (btnCloseCancel.Caption='(&X)結束') then
   begin
   dtmMain2.cdsSo113CT.Close;
   Close;
   end;
if (Edit1.Text='新增') or (Edit1.Text='修改') then
   begin
    dtmMain2.cdsSo113CT.Cancel;
    btn_Control(False,True,True,True,True,True);
    GroupBox1.Enabled:=False;
    Panel1.Enabled:=True;
    Edit1.Text:='顯示';
    btnCloseCancel.Caption:='(&X)結束'
   end;
end;

procedure TfrmSO8A20.btnDeleteClick(Sender: TObject);
var sL_DelSo113:String;
begin
if dtmMain2.cdsSo113CT.IsEmpty then
begin
 MessageDlg('已無資料',mtError, [mbOK],0);
 Exit;
end;
    sL_DelSo113:='Delete from So113 where PROVIDER_ID='''+dtmMain2.cdsSo113CT.fieldByName('PROVIDER_ID').AsString+'''';
    if MessageDlg('您確定要刪除此筆資料?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        begin
          with dtmMain2.cdsSqlDelete do
          begin
          Close;
          CommandText:=sL_DelSo113;
          EXECute;
          end;
         dtmMain2.cdsSo113CT.Close;
         dtmMain2.cdsSo113CT.Open;
        end;
    CurrentCursorAndCdsCount;
    btnCloseCancel.Caption:='(&X)結束'
end;

procedure TfrmSO8A20.dgr1DblClick(Sender: TObject);
begin
    btnEditClick(Sender);
    Panel1.Enabled:=False;
end;

procedure TfrmSO8A20.FormShow(Sender: TObject);
begin
    self.Caption := frmMainMenu.setCaption('frmSO8A20','[頻道商代碼管理]');
    TUCommonFun.AddObjectToComboBox(CobAttribute.Items, dtmMain2.cdsAttribute,not INSERT_NO_DATA_ITEM,'CodeNo','Description');
end;

procedure TfrmSO8A20.dsrSo113CTDataChange(Sender: TObject; Field: TField);
var sL_AttCodeNo:String;
begin
    if dtmMain2.cdsSo113CT.State=dsBrowse then
    begin
      sL_AttCodeNo:=dsrSo113CT.DataSet.FieldByName('Attribute').AsString;
      TUCommonFun.setComboDefaultNdx(CobAttribute, sL_AttCodeNo);
    end;
end;


function TfrmSO8A20.getCobAttributeCode: String;
var
    L_AttributeCodeStrList : TStringList;
    sL_AttributeCodeAndName : String;
begin
    //取得所屬公司
    sL_AttributeCodeAndName := CobAttribute.Text;

    L_AttributeCodeStrList := TStringList.Create;
    L_AttributeCodeStrList := TUstr.ParseStrings(CobAttribute.Text,'-');
    getCobAttributeCode := L_AttributeCodeStrList.Strings[0];

    L_AttributeCodeStrList.Free;

end;

function TfrmSO8A20.getCobAttributeName: String;
var
    L_AttributeNameStrList : TStringList;
    sL_AttributeCodeAndName : String;
begin
    //取得所屬公司
    sL_AttributeCodeAndName := CobAttribute.Text;

    L_AttributeNameStrList := TStringList.Create;
    L_AttributeNameStrList := TUstr.ParseStrings(CobAttribute.Text,'-');
    getCobAttributeName := L_AttributeNameStrList.Strings[1];

    L_AttributeNameStrList.Free;

end;

procedure TfrmSO8A20.CobAttributeChange(Sender: TObject);
begin
     if (copy(CobAttribute.Text,1,1)='') or (copy(CobAttribute.Text,1,1)='0') then
        dtmMain2.cdsSo113CT.FieldByName('ATTRIBUTE').AsInteger:=2
     else if Copy(CobAttribute.Text,1,1)='1' then
        dtmMain2.cdsSo113CT.FieldByName('ATTRIBUTE').AsInteger:=1
     else if Copy(CobAttribute.Text,1,1)='2' then
        dtmMain2.cdsSo113CT.FieldByName('ATTRIBUTE').AsInteger:=2
     else if Copy(CobAttribute.Text,1,1)='3' then
        dtmMain2.cdsSo113CT.FieldByName('ATTRIBUTE').AsInteger:=3;
end;

end.                                                      //
