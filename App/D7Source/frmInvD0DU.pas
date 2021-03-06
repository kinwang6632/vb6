unit frmInvD0DU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, DBGrids, StdCtrls, Mask, Buttons, ExtCtrls,DB;

type
  TfrmInvD0D = class(TForm)
    pnl1: TPanel;
    lbl1: TLabel;
    pnl2: TPanel;
    btnAppend: TBitBtn;
    btnEdit: TBitBtn;
    btnDelete: TBitBtn;
    btnCancel: TBitBtn;
    btnSave: TBitBtn;
    btnExit: TBitBtn;
    pnlEdit: TPanel;
    lbl2: TLabel;
    lbl3: TLabel;
    txtItemId: TMaskEdit;
    txtDesc: TMaskEdit;
    pnlGrid: TPanel;
    dbgrd1: TDBGrid;
    txtStt_Show: TStaticText;
    dsrInv045: TDataSource;
    procedure FormShow(Sender: TObject);
    procedure btnAppendClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure dsrInv045DataChange(Sender: TObject; Field: TField);
    procedure btnEditClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
  private
    { Private declarations }
    procedure ChangeBtnStatus;
    procedure initialForm;
    function isDataOK : Boolean;
  public
    { Public declarations }
  end;

var
  frmInvD0D: TfrmInvD0D;

implementation
uses frmMainU, dtmMainJU, dtmMainU, cbUtilis;
{$R *.dfm}

procedure TfrmInvD0D.ChangeBtnStatus;
begin
   with dtmMainJ.adoInv045Code do
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
          pnlGrid.Enabled := false;
          pnlEdit.Enabled := true;
       end
       else
       begin
          if dtmMainJ.adoInv045Code.RecordCount>0 then
          begin
            btnAppend.Enabled :=TRUE;
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
          pnlGrid.Enabled := true;
          pnlEdit.Enabled := FALSE;
       end;
     end;

end;

procedure TfrmInvD0D.FormShow(Sender: TObject);
begin
  self.Caption := frmMain.GetFormTitleString( 'D0D', '?o???C?L???]???????@' );
  dtmMainJ.getAllInv045Data;
  ChangeBtnStatus;
//  setButtonCompetence;
  initialForm;
end;

procedure TfrmInvD0D.initialForm;
begin
  with dsrInv045.DataSet do
    begin
      FindField('ItemId').DisplayLabel := '?C?L???]?N?X';
      FindField('Description').DisplayLabel := '?C?L???]?W??';
      FindField('UptTime').DisplayLabel := '????????';
      FindField('UptEn').DisplayLabel := '?????H??';

      FindField('Description').DisplayWidth := 30;
      FindField('UptEn').DisplayWidth := 10;


      if RecordCount = 0 then
      begin
        btnEdit.Enabled := false;
        btnDelete.Enabled := false;
      end;
    end;
end;

procedure TfrmInvD0D.btnAppendClick(Sender: TObject);
begin
   dtmMainJ.adoInv045Code.Append;
   ChangeBtnStatus;
   txtItemId.SetFocus;
end;

procedure TfrmInvD0D.btnSaveClick(Sender: TObject);
  var
    sL_ItemId,sL_Desc,sL_SQL : String;
    bL_HavePKValue : Boolean;
begin
  if isDataOK then
  begin
    sL_ItemId := Trim(txtItemId.Text);
    sL_Desc := Trim(txtDesc.Text);
    if  dsrInv045.DataSet.State in [dsInsert] then
    begin
      bL_HavePKValue := false;
      //????PK??
      sL_SQL := 'select * from inv045 where ItemId=''' + sL_ItemId + '''';
      bL_HavePKValue := dtmMainJ.checkPK(sL_SQL);
      if bL_HavePKValue then
      begin
        WarningMsg( '?H?????@???????C' );
        txtItemId.SetFocus;
        Exit;
      end;
      sL_SQL := 'select * from inv045 where description=''' + sL_Desc + '''';
      bL_HavePKValue := dtmMainJ.checkPK(sL_SQL);
      if bL_HavePKValue then
      begin
        WarningMsg( '?C?L???]???i?????C' );
        txtDesc.SetFocus;
        Exit;
      end;
    end;

    with dsrInv045.DataSet do
    begin
      FieldByName('ItemId').AsString := sL_ItemId;
      FieldByName('Description').AsString := sL_Desc;

      FieldByName('UptTime').AsString := DateTimeToStr(now);
      FieldByName('UptEn').AsString := dtmMain.getLoginUser;

      dsrInv045.DataSet.Post;
    end;
    dtmMainJ.getAllInv045Data;
    txtItemId.Enabled := true;

    ChangeBtnStatus;

    initialForm;
  end;
end;

function TfrmInvD0D.isDataOK: Boolean;
var
    sL_ItemId,sL_Description : String;
begin
  Result := true;
  sL_ItemId := Trim(txtItemId.Text);
  sL_Description := Trim(txtDesc.Text);

  if sL_ItemId = '' then
  begin
    MessageDlg('?????J???]?N?X',mtError,[mbOk],0);
    txtItemId.SetFocus;
    Result := false;
    Exit;
  end;

  if sL_Description = '' then
  begin
    MessageDlg('?????J???]',mtError,[mbOk],0);
    txtDesc.SetFocus;
    Result := false;
    Exit;
  end;
end;

procedure TfrmInvD0D.dsrInv045DataChange(Sender: TObject; Field: TField);
begin
  with dsrInv045.DataSet do
    begin
      if RecordCount > 0 then
      begin
        txtItemId.Text := FieldByName('ItemId').AsString;
        txtDesc.Text := FieldByName('Description').AsString;
        txtStt_Show.Caption := inttostr(dsrInv045.DataSet.RecNo)+'?@/?@'+inttostr(dsrInv045.DataSet.recordcount);

      end
      else
      begin
        txtItemId.Text := '';
        txtDesc.Text := '';
      end;
    end;
end;

procedure TfrmInvD0D.btnEditClick(Sender: TObject);
begin
  dtmMainJ.adoInv045Code.Edit;
  ChangeBtnStatus;
  txtItemId.Enabled := false;
end;

procedure TfrmInvD0D.btnDeleteClick(Sender: TObject);
begin
  if ConfirmMsg( '?T?{?n?R???????????' ) then
  begin
    dtmMainJ.adoInv045Code.Delete;
    dtmMainJ.getAllInv045Data;
    ChangeBtnStatus;
  end;

  initialForm;
end;

procedure TfrmInvD0D.btnCancelClick(Sender: TObject);
begin
  if  dsrInv045.DataSet.State in [dsEdit] then
  begin
    txtItemId.Enabled := true;
    //medDesc.Enabled := true;
  end;

  dtmMainJ.adoInv045Code.Cancel;
  dtmMainJ.adoInv045Code.First;
  ChangeBtnStatus;

//  setButtonCompetence;
  initialForm;
end;

procedure TfrmInvD0D.btnExitClick(Sender: TObject);
begin
  Close;
end;

end.
