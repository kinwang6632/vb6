unit frmCateCode1U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, Buttons, Grids, DBGrids, DB, 
  Mask, DBCtrls, DBClient;

type
  TfrmCateCode1 = class(TForm)
    pnlCommand: TPanel;
    pnlMultiData: TPanel;
    Splitter1: TSplitter;
    pnlSingleData: TPanel;
    dsrCD067A: TDataSource;
    DBGrid1: TDBGrid;
    StaticText5: TStaticText;
    StaticText7: TStaticText;
    DBEdit2: TDBEdit;
    btnAppend: TBitBtn;
    btnEdit: TBitBtn;
    btnDelete: TBitBtn;
    btnCancelExit: TBitBtn;
    btnSave: TBitBtn;
    DBEdit5: TDBEdit;
    btnEditData: TBitBtn;
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnAppendClick(Sender: TObject);
    procedure btnEditClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure btnCancelExitClick(Sender: TObject);
    procedure DBGrid1TitleClick(Column: TColumn);
    procedure DBGrid1DblClick(Sender: TObject);
    procedure btnEditDataClick(Sender: TObject);

  private
    { Private declarations }
    sG_EmpSQL : String;
    function IsDataOK:boolean;
    procedure ChangeBtnState;    
  public
    { Public declarations }
  end;

var
  frmCateCode1: TfrmCateCode1;

implementation

uses dtmMainU, frmMainU, UdateTimeu, frmCateCode2U;

{$R *.dfm}

procedure TfrmCateCode1.FormShow(Sender: TObject);
var
    bL_IsEnable : Boolean;
begin
    dtmMain.activeDataSet(dtmMain.adoCD067A);
    dtmMain.activeDataSet(dtmMain.adoCodeCD067A);
    ChangeBtnState;

    self.Caption := dtmMain.setCaption('SO8F10','[二階代碼分類資料維護]');


    //down, 設定權限...
    if (dtmMain.sG_IsSupervisor='Y') or (dtmMain.sG_NeedLogin='Y') then
    begin
      btnAppend.Enabled := true;
      btnEdit.Enabled := true;
      btnDelete.Enabled := true;
    end
    else
    begin
      dtmMain.activePrivDataset(dtmMain.sG_CompCode,dtmMain.sG_UserID);
      bL_IsEnable := dtmMain.checkPriv('SO8F10');
      btnAppend.Enabled := bL_IsEnable;
      btnEdit.Enabled := bL_IsEnable;
      btnDelete.Enabled := bL_IsEnable;
    end;
    //up, 設定權限...
end;

procedure TfrmCateCode1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
    Action := caFree;
end;

procedure TfrmCateCode1.ChangeBtnState;
begin
     with dsrCD067A.DataSet do
     begin
       if state in [dsInactive] then
         Exit
       else if state in [dsEdit, dsInsert] then
       begin
          btnCancelExit.Caption := '取消';
          btnCancelExit.Enabled := TRUE;
          btnSave.Enabled := TRUE;
          btnDelete.Enabled := FALSE;
          btnEditData.Enabled := FALSE;
          btnEdit.Enabled := FALSE;
          btnAppend.Enabled := FALSE;
          pnlSingleData.Enabled := TRUE;
          pnlMultiData.Enabled := FALSE;

       end
       else
       begin
          if RecordCount>0 then
          begin
            btnEditData.Enabled := true;
            btnEdit.Enabled := TRUE;
            btnDelete.Enabled := TRUE;
          end
          else
          begin
            btnEditData.Enabled := FALSE;          
            btnEdit.Enabled := FALSE;
            btnDelete.Enabled := FALSE;
          end;
          btnCancelExit.Caption := '結束';
          btnCancelExit.Enabled := TRUE;
          btnSave.Enabled := FALSE;
          btnAppend.Enabled := TRUE;
          pnlSingleData.Enabled := FALSE;
          pnlMultiData.Enabled := TRUE;
       end;
     end;
end;

procedure TfrmCateCode1.btnAppendClick(Sender: TObject);
begin
    if dsrCD067A.DataSet.State = dsInactive then
      dsrCD067A.DataSet.Active := True;
    dsrCD067A.DataSet.Append;
    ChangeBtnState;
    dsrCD067A.DataSet.FieldByName('TABLENAME').ReadOnly := false;
    dsrCD067A.DataSet.FieldByName('TABLENAME').FocusControl;
end;

procedure TfrmCateCode1.btnEditClick(Sender: TObject);
var
    nL_Seq : Integer;
begin
    if dsrCD067A.DataSet.State = dsInactive then Exit;
    if dsrCD067A.DataSet.RecordCount =0 then Exit;
    dsrCD067A.DataSet.Edit;
    ChangeBtnState;
    dsrCD067A.DataSet.FieldByName('TABLENAME').ReadOnly := true;
    dsrCD067A.DataSet.FieldByName('TABLEDESCRIPTION').FocusControl;
end;

procedure TfrmCateCode1.btnDeleteClick(Sender: TObject);
var
    nL_Seq,nL_MasterDataCounts,nL_DetailDataCounts : Integer;
    sL_TableName : String;
begin
    sL_TableName := dtmMain.adoCD067A.FieldByName('TableName').AsString;
    nL_MasterDataCounts := dtmMain.checkCodeDataCounts('Cd067B',sL_TableName);
    nL_DetailDataCounts := dtmMain.checkCodeDataCounts('Cd067C',sL_TableName);

    if (nL_MasterDataCounts = 0) and (nL_DetailDataCounts = 0) then
    begin
      if dsrCD067A.DataSet.State = dsInactive then Exit;
      if dsrCD067A.DataSet.RecordCount =0 then Exit;
        if MessageDlg('是否確認刪除?',mtConfirmation, [mbYes, mbNo], 0) = mrNo then
          Exit;
        dsrCD067A.DataSet.Delete;
        ChangeBtnState;
    end
    else
    begin
      MessageDlg('尚有代碼對應資料,不能刪除!', mtWarning, [mbOK],0);
    end;
//      MessageDlg('此組單據號碼,已經有對應之回單資料,所以無法刪除!', mtWarning, [mbOK],0);
end;

procedure TfrmCateCode1.btnSaveClick(Sender: TObject);
var
    sL_UpdTime : String;
begin
    try
      if dsrCD067A.DataSet.State = dsInactive then Exit;
      if IsDataOK then
      begin
        with dsrCD067A.DataSet do
        begin
          FieldByName('OPERATOR').AsString := dtmMain.sG_USerID;

          sL_UpdTime := TUdateTime.CDateStr(date, 9) + ' ' + timeToStr(time);
          if Copy(sL_UpdTime,1,1)='0' then
            sL_UpdTime := Copy(sL_UpdTime,2,Length(sL_UpdTime));

          FieldByName('UPDTIME').AsString := sL_UpdTime;
          Post;
        end;
        ChangeBtnState;
      end;
    except
      on E: Exception do
      begin
        if dsrCD067A.DataSet.State = dsInsert then
          dsrCD067A.DataSet.FieldByName('TABLENAME').FocusControl
        else if dsrCD067A.DataSet.State = dsEdit then
          dsrCD067A.DataSet.FieldByName('TABLEDESCRIPTION').FocusControl;
        MessageDlg(E.Message, mtWarning, [mbok],0);
      end;
    end;
end;

procedure TfrmCateCode1.btnCancelExitClick(Sender: TObject);
begin
    if dsrCD067A.DataSet.State in [dsEdit, dsInsert ] then
    begin
      dsrCD067A.DataSet.Cancel;
      ChangeBtnState;
    end
    else if dsrCD067A.DataSet.State in [dsInactive, dsBrowse] then
      Close;


end;

function TfrmCateCode1.IsDataOK: boolean;
var
    bL_Result : boolean;
begin
    bL_Result := true;
    if dsrCD067A.DataSet.FieldByName('TABLENAME').AsString='' then
    begin
      MessageDlg('請輸入 Table Name!', mtWarning, [mbOK],0);
      dsrCD067A.DataSet.FieldByName('TABLENAME').FocusControl;
      bL_Result := false;
    end
    else if dsrCD067A.DataSet.FieldByName('TABLEDESCRIPTION').AsString='' then
    begin
      MessageDlg('請輸入 Table Description!', mtWarning, [mbOK],0);
      dsrCD067A.DataSet.FieldByName('TABLEDESCRIPTION').FocusControl;
      bL_Result := false;
    end;
    
    result := bL_Result;
end;

procedure TfrmCateCode1.DBGrid1TitleClick(Column: TColumn);
var
   ii : Integer;
   bFlag : Boolean;
begin

     bFlag := False;
     for ii:=0 to dsrCD067A.DataSet.FieldCount-1 do
     begin
       if (dsrCD067A.DataSet.Fields[ii].FieldName=Column.FieldName) and
          (dsrCD067A.DataSet.Fields[ii].FieldKind=fkData) then
       begin
         bFlag := True;
         break;
       end;
     end;

     if not bFlag then
       Exit;//該FieldName不存在於FDataSet中
     if dsrCD067A.DataSet IS TClientDataSet then
     begin
       if TClientDataSet(dsrCD067A.DataSet).IndexName = 'TmpIndex' then
         TClientDataSet(dsrCD067A.DataSet).DeleteIndex('TmpIndex');
       TClientDataSet(dsrCD067A.DataSet).AddIndex('TmpIndex', Column.FieldName, [ixCaseInsensitive],'','',0);
       TClientDataSet(dsrCD067A.DataSet).IndexName := 'TmpIndex';
     end;
end;

procedure TfrmCateCode1.DBGrid1DblClick(Sender: TObject);
begin
//    btnEditClick(nil);
    btnEditDataClick(nil);
end;

procedure TfrmCateCode1.btnEditDataClick(Sender: TObject);
var
    sL_TableName : String;
begin
    sL_TableName := dsrCD067A.DataSet.fieldByName('TABLENAME').AsString;

    frmCateCode2 := TfrmCateCode2.Create(Application);
    dtmMain.adoCD067C.Close;
    dtmMain.activeCd067B(sL_TableName);
    frmCateCode2.stxTableName.Caption := '目前 active 的 table name : ' +  sL_TableName;
    frmCateCode2.sG_ActiveTableName := sL_TableName;
    frmCateCode2.ShowModal;
    frmCateCode2.Free;

end;

end.

