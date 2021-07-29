unit frmUserDataU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, Grids, DBGrids, StdCtrls, ExtCtrls, DBCtrls, Mask, Buttons,
  ComCtrls;

type
  TfrmUserData = class(TForm)
    StatusBar1: TStatusBar;
    pnlFun: TPanel;
    btnAppend: TBitBtn;
    btnEdit: TBitBtn;
    btnDelete: TBitBtn;
    btnCancel: TBitBtn;
    btnSave: TBitBtn;
    pnlSingleData: TPanel;
    dbtUserID: TDBEdit;
    StaticText1: TStaticText;
    StaticText3: TStaticText;
    DBRadioGroup1: TDBRadioGroup;
    dbtDesc: TDBEdit;
    pnlMultiData: TPanel;
    dbgGroupData: TDBGrid;
    dsrUserData: TDataSource;
    StaticText2: TStaticText;
    StaticText4: TStaticText;
    StaticText6: TStaticText;
    StaticText7: TStaticText;
    StaticText9: TStaticText;
    StaticText10: TStaticText;
    StaticText11: TStaticText;
    DBEdit1: TDBEdit;
    DBEdit2: TDBEdit;
    DBEdit3: TDBEdit;
    DBEdit4: TDBEdit;
    DBLookupComboBox1: TDBLookupComboBox;
    DBLookupComboBox2: TDBLookupComboBox;
    DBLookupComboBox3: TDBLookupComboBox;
    DBRadioGroup2: TDBRadioGroup;
    DBRadioGroup3: TDBRadioGroup;
    dsrCompCode: TDataSource;
    dsrDeptCode: TDataSource;
    dsrGroupCode: TDataSource;
    pnlQuery: TPanel;
    BitBtn1: TBitBtn;
    btnExit: TBitBtn;
    dsrAllGroupCode: TDataSource;
    dsrAllCompCode: TDataSource;
    dsrAllDeptCode: TDataSource;
    rgpQryCond: TGroupBox;
    RadioButton0: TRadioButton;
    RadioButton1: TRadioButton;
    RadioButton2: TRadioButton;
    RadioButton3: TRadioButton;
    RadioButton4: TRadioButton;
    RadioButton5: TRadioButton;
    RadioButton6: TRadioButton;
    RadioButton7: TRadioButton;
    RadioButton8: TRadioButton;
    dblCompCode: TDBLookupComboBox;
    dblDeptCode: TDBLookupComboBox;
    dblGroupCode: TDBLookupComboBox;
    chbIsSysOp: TComboBox;
    chbIsSupervisor: TComboBox;
    chbStopFlag: TComboBox;
    edtUserID: TEdit;
    edtUserName: TEdit;
    procedure FormShow(Sender: TObject);
    procedure btnAppendClick(Sender: TObject);
    procedure btnEditClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure dbgGroupDataTitleClick(Column: TColumn);
  private
    { Private declarations }
    procedure ChangeBtnStatus;    
  public
    { Public declarations }
  end;

var
  frmUserData: TfrmUserData;

implementation

uses frmMainU, dtmMainU;

{$R *.dfm}

procedure TfrmUserData.FormShow(Sender: TObject);
begin
    self.Caption := frmMain.getCaption('使用者資料維護');

    if dsrCompCode.DataSet.State = dsInactive then
      dsrCompCode.DataSet.Open;

    if dsrDeptCode.DataSet.State = dsInactive then
      dsrDeptCode.DataSet.Open;

    if dsrGroupCode.DataSet.State = dsInactive then
      dsrGroupCode.DataSet.Open;

    if dsrAllCompCode.DataSet.State = dsInactive then
      dsrAllCompCode.DataSet.Open;

    if dsrAllDeptCode.DataSet.State = dsInactive then
      dsrAllDeptCode.DataSet.Open;

    if dsrAllGroupCode.DataSet.State = dsInactive then
      dsrAllGroupCode.DataSet.Open;


    ChangeBtnStatus;
end;

procedure TfrmUserData.btnAppendClick(Sender: TObject);
begin
    if dsrUserData.State = dsInactive then
      dsrUserData.DataSet.Open;
    dsrUserData.dataset.Append;
    ChangeBtnStatus;
    dbtUserID.ReadOnly := false;

end;

procedure TfrmUserData.ChangeBtnStatus;
begin
     with dsrUserData.DataSet do
     begin
       if state in [dsInactive] then
       begin
         pnlQuery.Enabled := true;
         pnlFun.Enabled := true;
         pnlSingleData.Enabled := false;
         pnlMultiData.Enabled := false;
         Exit;
       end
       else if state in [dsEdit, dsInsert] then
       begin
          btnCancel.Enabled := TRUE;
          btnSave.Enabled := TRUE;
          btnDelete.Enabled := FALSE;
          btnEdit.Enabled := FALSE;
//          btnExit.Enabled := FALSE;
          btnAppend.Enabled := FALSE;
          pnlSingleData.Enabled := TRUE;
          pnlMultiData.Enabled := FALSE;
          pnlQuery.Enabled := false;
          if  state =dsEdit then
            dbtDesc.SetFocus
          else
            dbtUserID.SetFocus;
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
//          btnExit.Enabled := TRUE;
          pnlSingleData.Enabled := FALSE;
          pnlMultiData.Enabled := TRUE;
          pnlQuery.Enabled := true;
       end;
     end;

end;

procedure TfrmUserData.btnEditClick(Sender: TObject);
begin
    if dsrUserData.State = dsInactive then
      MessageDlg('沒有資料可以修改!', mtWarning, [mbOK],0)
    else
    begin
      dsrUserData.dataset.Edit;
      ChangeBtnStatus;
      dbtUserID.ReadOnly := true;
    end;

end;

procedure TfrmUserData.btnDeleteClick(Sender: TObject);
begin
    if dsrUserData.State = dsInactive then
      MessageDlg('沒有資料可以刪除!', mtWarning, [mbOK],0)
    else
    begin

      if MessageDlg('是否確認刪除?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
      begin
        dsrUserData.dataset.Delete;
        ChangeBtnStatus;

      end;
    end;
end;

procedure TfrmUserData.btnCancelClick(Sender: TObject);
begin
    if dsrUserData.State = dsInactive then
      MessageDlg('沒有資料可以取消!', mtWarning, [mbOK],0)
    else
    begin

      dsrUserData.dataset.Cancel;
      ChangeBtnStatus;
    end;
end;

procedure TfrmUserData.btnSaveClick(Sender: TObject);
begin
    if dsrUserData.State = dsInactive then
      MessageDlg('沒有資料可以儲存!', mtWarning, [mbOK],0)
    else
    begin

      if dsrUserData.DataSet.State = dsInsert then
      begin
        dsrUserData.dataset.FieldByName('CreateDate').AsDateTime := now;
        dsrUserData.dataset.FieldByName('Creater').AsString := dtmMain.getUserName;

      end;

      dsrUserData.dataset.FieldByName('UpdTime').AsDateTime := now;
      dsrUserData.dataset.FieldByName('UpdEn').AsString := dtmMain.getUserName;

      dsrUserData.dataset.Post;
      ChangeBtnStatus;
    end;
end;

procedure TfrmUserData.btnExitClick(Sender: TObject);
begin
    Close;
end;

procedure TfrmUserData.BitBtn1Click(Sender: TObject);
    function getSelectedIndex:Integer;
    var
        nL_Result : Integer;
    begin
      if RadioButton0.Checked  then
        nL_Result := 0
      else if RadioButton1.Checked  then
        nL_Result := 1
      else if RadioButton2.Checked  then
        nL_Result := 2
      else if RadioButton3.Checked  then
        nL_Result := 3
      else if RadioButton4.Checked  then
        nL_Result := 4
      else if RadioButton5.Checked  then
        nL_Result := 5
      else if RadioButton6.Checked  then
        nL_Result := 6
      else if RadioButton7.Checked  then
        nL_Result := 7
      else if RadioButton8.Checked  then
        nL_Result := 8;

      result :=  nL_Result;
    end;

var
    nL_RecCount : Integer;
    sL_QueryValue : String;
    nL_SelectedNdx : Integer;
begin
    nL_SelectedNdx := getSelectedIndex;
    case nL_SelectedNdx of
     6://是否系統管理者
      begin
        if chbIsSysOp.ItemIndex=0 then
          sL_QueryValue := '1'
        else
          sL_QueryValue := '0';        
      end;
     7://是否群組管理者
      begin
        if chbIsSupervisor.ItemIndex=0 then
          sL_QueryValue := '1'
        else
          sL_QueryValue := '0';
      end;
     8://是否停用
      begin
        if chbStopFlag.ItemIndex=0 then
          sL_QueryValue := '1'
        else
          sL_QueryValue := '0';
      end;

     5://全部資料
      sL_QueryValue := '';
     0://公司別
      begin
       if dblCompCode.Text ='' then
       begin
         dblCompCode.SetFocus;
         MessageDlg('請先點選公司別',mtWarning, [mbOK],0);
         Exit;
       end
       else
         sL_QueryValue := dblCompCode.ListSource.DataSet.fieldByName('CODENO').AsString;
      end;

     1://部門
      begin
       if dblDeptCode.Text ='' then
       begin
         dblDeptCode.SetFocus;
         MessageDlg('請先點選部門',mtWarning, [mbOK],0);
         Exit;
       end
       else
         sL_QueryValue := dblDeptCode.ListSource.DataSet.fieldByName('CODENO').AsString;
      end;

     2://群組
      begin
       if dblGroupCode.Text ='' then
       begin
         dblGroupCode.SetFocus;
         MessageDlg('請先點選群組',mtWarning, [mbOK],0);
         Exit;
       end
       else
         sL_QueryValue := dblGroupCode.ListSource.DataSet.fieldByName('CODENO').AsString;
      end;

     3://帳號
      begin
       if edtUserID.Text ='' then
       begin
         edtUserID.SetFocus;
         MessageDlg('請先輸入帳號',mtWarning, [mbOK],0);
         Exit;
       end
       else
         sL_QueryValue := Trim(edtUserID.Text);
      end;

     4://姓名
      begin
       if edtUserName.Text ='' then
       begin
         edtUserName.SetFocus;
         MessageDlg('請先輸入姓名',mtWarning, [mbOK],0);
         Exit;
       end
       else
         sL_QueryValue := Trim(edtUserName.Text);
      end;

    end;
    nL_RecCount := dtmMain.getUserData(nL_SelectedNdx, sL_QueryValue);

    if nL_RecCount=0 then
    begin
      MessageDlg('查無資料,請重新查詢!', mtInformation, [mbOK],0);
    end;
    ChangeBtnStatus;
end;

procedure TfrmUserData.dbgGroupDataTitleClick(Column: TColumn);
var
    sL_OrderByStr : String;
begin
    case Column.Index of
      0://USERID
        sL_OrderByStr := ' order by USERID';
      1://USERNAME
        sL_OrderByStr := ' order by USERNAME';
      2://USERPASSWD
        sL_OrderByStr := ' order by USERPASSWD';
      3://USERGROUP
        sL_OrderByStr := ' order by USERGROUP';
      4://ISSUPERVISOR
        sL_OrderByStr := ' order by ISSUPERVISOR';
      5://ISSYSOP
        sL_OrderByStr := ' order by ISSYSOP';
      6://COMPCODE
        sL_OrderByStr := ' order by COMPCODE';
      7://DEPTCODE
        sL_OrderByStr := ' order by DEPTCODE';
      8://TEL
        sL_OrderByStr := ' order by TEL';
      9://TELEXT
        sL_OrderByStr := ' order by TELEXT';
      10://EMAIL
        sL_OrderByStr := ' order by EMAIL';
      11://STOPFLAG
        sL_OrderByStr := ' order by STOPFLAG';
      12://CREATER
        sL_OrderByStr := ' order by CREATER';
      13://CREATEDATE
        sL_OrderByStr := ' order by CREATEDATE';
      14://UPDEN
        sL_OrderByStr := ' order by UPDEN';
      15://UPDTIME
        sL_OrderByStr := ' order by UPDTIME';
    end;

    dtmMain.getUserData(sL_OrderByStr);
end;

end.
