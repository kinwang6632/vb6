unit frmGroupDataU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, ComCtrls, Grids, DBGrids, DB, DBCtrls, StdCtrls, Mask,
  Buttons;

type
  TfrmGroupData = class(TForm)
    StatusBar1: TStatusBar;
    pnlFun: TPanel;
    pnlSingleData: TPanel;
    pnlMultiData: TPanel;
    dbgGroupData: TDBGrid;
    dsrGroupData: TDataSource;
    dbtCodeNo: TDBEdit;
    StaticText1: TStaticText;
    StaticText3: TStaticText;
    DBRadioGroup1: TDBRadioGroup;
    dbtDesc: TDBEdit;
    btnAppend: TBitBtn;
    btnEdit: TBitBtn;
    btnDelete: TBitBtn;
    btnCancel: TBitBtn;
    btnSave: TBitBtn;
    btnExit: TBitBtn;
    rgbSo: TGroupBox;
    chb9: TCheckBox;
    chb10: TCheckBox;
    chb12: TCheckBox;
    chb11: TCheckBox;
    chb8: TCheckBox;
    chb13: TCheckBox;
    chb14: TCheckBox;
    chb16: TCheckBox;
    chb7: TCheckBox;
    chb6: TCheckBox;
    chb5: TCheckBox;
    chb3: TCheckBox;
    chb1: TCheckBox;
    chb2: TCheckBox;
    procedure FormShow(Sender: TObject);
    procedure btnAppendClick(Sender: TObject);
    procedure btnEditClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
  private
    { Private declarations }
    procedure ChangeBtnStatus;    
  public
    { Public declarations }
    procedure setSoUi;
  end;

var
  frmGroupData: TfrmGroupData;

implementation

uses dtmMainU, frmMainU, Ustru;

{$R *.dfm}

procedure TfrmGroupData.ChangeBtnStatus;
begin
     with dsrGroupData.DataSet do
     begin
       if state in [dsInactive] then
         Exit     
       else if state in [dsEdit, dsInsert] then
       begin
          btnCancel.Enabled := TRUE;
          btnSave.Enabled := TRUE;
          btnDelete.Enabled := FALSE;
          btnEdit.Enabled := FALSE;
          btnExit.Enabled := FALSE;
          btnAppend.Enabled := FALSE;
          pnlSingleData.Enabled := TRUE;
          pnlMultiData.Enabled := FALSE;
          if  state =dsEdit then
            dbtDesc.SetFocus
          else
            dbtCodeNo.SetFocus;
       end
       else
       begin
          if RecordCount>0 then
          begin
            btnEdit.Enabled := TRUE;
            btnDelete.Enabled := TRUE;
          end
          else
          begin
            btnEdit.Enabled := FALSE;
            btnDelete.Enabled := FALSE;
          end;
          btnCancel.Enabled := FALSE;
          btnSave.Enabled := FALSE;
          btnAppend.Enabled := TRUE;
          btnExit.Enabled := TRUE;          
          pnlSingleData.Enabled := FALSE;
          pnlMultiData.Enabled := TRUE;
       end;
     end;

end;

procedure TfrmGroupData.FormShow(Sender: TObject);
begin
    self.Caption := frmMain.getCaption('群組資料維護');
    dtmMain.getAllGroupData;
    ChangeBtnStatus;    
end;

procedure TfrmGroupData.btnAppendClick(Sender: TObject);
begin
    dsrGroupData.dataset.Append;
    ChangeBtnStatus;
    dbtCodeNo.ReadOnly := false;
end;

procedure TfrmGroupData.btnEditClick(Sender: TObject);
begin
    dsrGroupData.dataset.Edit;
    ChangeBtnStatus;
    dbtCodeNo.ReadOnly := true;
end;

procedure TfrmGroupData.btnDeleteClick(Sender: TObject);
begin
    if MessageDlg('是否確認刪除?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      dsrGroupData.dataset.Delete;
      ChangeBtnStatus;

    end;

end;

procedure TfrmGroupData.btnCancelClick(Sender: TObject);
begin
      dsrGroupData.dataset.Cancel;
      ChangeBtnStatus;

end;

procedure TfrmGroupData.btnSaveClick(Sender: TObject);
var
    ii : Integer;
    sL_AllSoName, sL_AllSoCode : String;
begin
    sL_AllSoCode := '';
    sL_AllSoName := '';

    if dsrGroupData.dataset.State in [dsEdit, dsInsert] then
    begin
      for ii:=0 to rgbSo.ControlCount-1 do
      begin
        if rgbSo.Controls[ii] is TCheckBox then
        begin
          if TCheckBox(rgbSo.Controls[ii]).Checked then
          begin
            if sL_AllSoCode='' then
            begin
              sL_AllSoName := TCheckBox(rgbSo.Controls[ii]).Caption;
              sL_AllSoCode := IntToStr(TCheckBox(rgbSo.Controls[ii]).Tag);
            end
            else
            begin
              sL_AllSoName := sL_AllSoName + SO_SEP + TCheckBox(rgbSo.Controls[ii]).Caption;
              sL_AllSoCode := sL_AllSoCode + SO_SEP + IntToStr(TCheckBox(rgbSo.Controls[ii]).Tag);
            end;
          end;
        end;
      end;
      dsrGroupData.dataset.FieldByName('SONAME').AsString := sL_AllSoName;
      dsrGroupData.dataset.FieldByName('SOCODE').AsString := sL_AllSoCode;
      dsrGroupData.dataset.FieldByName('UpdTime').AsDateTime := now;
      dsrGroupData.dataset.FieldByName('UpdEn').AsString := dtmMain.getUserName;      

    end;
    dsrGroupData.dataset.Post;
    ChangeBtnStatus;

end;

procedure TfrmGroupData.btnExitClick(Sender: TObject);
begin
    Close;
end;

procedure TfrmGroupData.setSoUi;
var
    L_Control : TControl;
    L_SoCodeStrList : TStringList;
    sL_AllSoCode, sL_ComponentName : String;
    ii : Integer;
begin
    if not Assigned(dsrGroupData)  then Exit;
    if dsrGroupData = nil then Exit;
    if dsrGroupData.DataSet = nil then Exit;
    if dsrGroupData.DataSet.State = dsInactive then Exit;

    sL_AllSoCode := dsrGroupData.DataSet.FieldByName('SOCODE').AsString;


    L_SoCodeStrList := TUStr.ParseStrings(sL_AllSoCode,SO_SEP);

    for ii:=0 to rgbSo.ControlCount-1 do
    begin
      if rgbSo.Controls[ii] is TCheckBox then
      begin
        TCheckBox(rgbSo.Controls[ii]).Checked := false;
      end;
    end;

    if L_SoCodeStrList<> nil then
    begin
      for ii:=0 to L_SoCodeStrList.Count -1 do
      begin
        sL_ComponentName := 'chb' + L_SoCodeStrList.Strings[ii];
        L_Control := rgbSo.FindChildControl(sL_ComponentName);
  //      rgbSo.FindComponent()
        if L_Control<>nil then
        begin
          TCheckBox(L_Control).Checked := true;
        end;


      end;
    end;
end;

end.
