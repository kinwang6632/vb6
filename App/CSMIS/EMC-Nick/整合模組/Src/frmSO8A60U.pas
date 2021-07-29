unit frmSO8A60U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DBClient, Provider, DB, ADODB, StdCtrls, Buttons, ComCtrls,
  ExtCtrls, Grids, DBGrids, Mask, DBCtrls, Menus, xmldom, XMLIntf,
  msxmldom, XMLDoc;

const
    MAX_PERCENT = 999.99;
    MIN_PERCENT = 1;

type
  TfrmSO8A60 = class(TForm)
    pnlMultiData: TPanel;
    StatusBar1: TStatusBar;
    pnlSingleData: TPanel;
    Panel3: TPanel;
    btnAppend: TBitBtn;
    btnEdit: TBitBtn;
    btnDelete: TBitBtn;
    btnCancel: TBitBtn;
    btnExit: TBitBtn;
    dsrSo112A: TDataSource;
    dsrCd019: TDataSource;
    DBGrid1: TDBGrid;
    btnSave: TBitBtn;
    pgcPercentData: TPageControl;
    TabSheet1: TTabSheet;
    dbgPercentXmlData: TDBGrid;
    dsrPercentXmlData: TDataSource;
    PopupMenu1: TPopupMenu;
    miInsert: TMenuItem;
    miEdit: TMenuItem;
    miDelete: TMenuItem;
    miCancel: TMenuItem;
    miSave: TMenuItem;
    dsrSO113ForUSeful: TDataSource;
    pnlMasterData: TPanel;
    DBText1: TDBText;
    StaticText1: TStaticText;
    StaticText2: TStaticText;
    DBLookupComboBox1: TDBLookupComboBox;
    procedure btnExitClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure btnAppendClick(Sender: TObject);
    procedure btnEditClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure DBLookupComboBox1CloseUp(Sender: TObject);
    procedure dbgPercentXmlDataEditButtonClick(Sender: TObject);
    procedure miInsertClick(Sender: TObject);
    procedure miEditClick(Sender: TObject);
    procedure miDeleteClick(Sender: TObject);
    procedure miCancelClick(Sender: TObject);
    procedure miSaveClick(Sender: TObject);
    procedure dsrPercentXmlDataStateChange(Sender: TObject);
  private
    { Private declarations }
    procedure refreshXmlData;
    procedure ChangeState;
    procedure changePercentDataBtnState;
  public
    { Public declarations }
    
  end;

var
  frmSO8A60: TfrmSO8A60;

implementation

uses  dtmMain2U, frmMainMenuU, frmDbSelectu, xmlU;

{$R *.dfm}

procedure TfrmSO8A60.btnExitClick(Sender: TObject);
begin
    Close;
end;

procedure TfrmSO8A60.FormShow(Sender: TObject);
begin
    dtmMain2.initPercentXmlData;


    //if dtmMain2.cdsSO113ForUSeful.State = dsInactive then
    //  dtmMain2.cdsSO113ForUSeful.Open;


    if dtmMain2.cdsCd019A.State = dsInactive then
      dtmMain2.cdsCd019A.Open;

    if dtmMain2.cdsSo112A.State = dsInactive then
      dtmMain2.cdsSo112A.Open;

    ChangeState;
    changePercentDataBtnState;

    dsrSo112A.DataSet.First;
end;

procedure TfrmSO8A60.ChangeState;
begin
     DBLookupComboBox1.Enabled := true;
     with dtmMain2.cdsSo112A do
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
          pnlMasterData.Enabled := TRUE;
          pgcPercentData.Enabled := FALSE;
          pnlMultiData.Enabled := FALSE;
          if state=dsEdit then
            DBLookupComboBox1.Enabled := false;

       end
       else
       begin
          if dtmMain2.cdsSo112A.RecordCount>0 then
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
          pnlMasterData.Enabled := FALSE;
          pgcPercentData.Enabled := TRUE;

          pnlMultiData.Enabled := TRUE;
       end;
     end;
    changePercentDataBtnState;
end;

procedure TfrmSO8A60.btnAppendClick(Sender: TObject);
begin
    if dsrPercentXmlData.State in [dsEdit,dsInsert] then
      dsrPercentXmlData.DataSet.Cancel;
    if dtmMain2.cdsSo112A.State = dsInactive then
      dtmMain2.cdsSo112A.Active := True;
    dtmMain2.cdsSo112A.Append;
    ChangeState;
end;

procedure TfrmSO8A60.btnEditClick(Sender: TObject);
begin
    if dsrPercentXmlData.State in [dsEdit,dsInsert] then
      dsrPercentXmlData.DataSet.Cancel;
    if dtmMain2.cdsSo112A.State = dsInactive then Exit;
    dtmMain2.cdsSo112A.Edit;
    ChangeState;
end;

procedure TfrmSO8A60.btnDeleteClick(Sender: TObject);
begin
    if dsrPercentXmlData.State in [dsEdit,dsInsert] then
      dsrPercentXmlData.DataSet.Cancel;
    if dtmMain2.cdsSo112A.State = dsInactive then Exit;
    if MessageDlg('是否確認刪除?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
     dtmMain2.cdsSo112A.Delete;
     ChangeState;
    end;
end;

procedure TfrmSO8A60.btnCancelClick(Sender: TObject);
begin
    if dsrPercentXmlData.State in [dsEdit,dsInsert] then
      dsrPercentXmlData.DataSet.Cancel;
    if dtmMain2.cdsSo112A.State = dsInactive then Exit;
    dtmMain2.cdsSo112A.Cancel;
    ChangeState;
end;

procedure TfrmSO8A60.btnSaveClick(Sender: TObject);
begin
    if dsrPercentXmlData.State in [dsEdit,dsInsert] then
      dsrPercentXmlData.DataSet.Cancel;
    if dtmMain2.cdsSo112A.State = dsInactive then Exit;
    if dtmMain2.cdsSo112A.State in [dsEdit, dsInsert] then
    begin

      dtmMain2.cdsSo112A.FieldByName('OPERATOR').AsString := frmMainMenu.sG_User;
      dtmMain2.cdsSo112A.FieldByName('UPDTIME').AsString := DateTimeToStr(now);
      dtmMain2.cdsSo112A.Post;
      dtmMain2.saveToDB(dtmMain2.cdsSo112A);

      ChangeState;
    end;
end;

procedure TfrmSO8A60.DBLookupComboBox1CloseUp(Sender: TObject);
var
    nL_CodeNo : Integer;
    sL_Desc : String;
begin
    if dtmMain2.cdsSo112A.State in [dsEdit, dsInsert] then
    begin
      nL_CodeNo := dtmMain2.cdsCd019A.FieldByName('CODENO').AsInteger;
      sL_Desc := dtmMain2.cdsCd019A.FieldByName('DESCRIPTION').AsString;
      dtmMain2.cdsSo112A.FieldByName('CODENO').AsInteger := nL_CodeNo;
      dtmMain2.cdsSo112A.FieldByName('DESCRIPTION').AsString := sL_Desc;
    end;
end;

procedure TfrmSO8A60.dbgPercentXmlDataEditButtonClick(Sender: TObject);
var
    sL_XmlDataStr : String;
begin
    if dsrPercentXmlData.DataSet.State in [dsEdit, dsInsert] then
    begin
      dsrSO113ForUSeful.DataSet.FieldByName('PROVIDER_ID').DisplayLabel := '供應商代碼';
      dsrSO113ForUSeful.DataSet.FieldByName('PROVIDER_NAME').DisplayLabel := '供應商名稱';
      if SelectRecord('請選取供應商名稱', dsrSO113ForUSeful.DataSet, 'PROVIDER_ID;PROVIDER_NAME') = mrOk then
      begin
        dsrPercentXmlData.DataSet.FieldByName('sProviderID').AsString := dsrSO113ForUSeful.DataSet.FieldByName('PROVIDER_ID').AsString;      
        dsrPercentXmlData.DataSet.FieldByName('sProviderName').AsString := dsrSO113ForUSeful.DataSet.FieldByName('PROVIDER_NAME').AsString;

      end;
    end;
end;

procedure TfrmSO8A60.changePercentDataBtnState;
begin
    if dsrSo112A.State=dsInactive then
    begin
      dbgPercentXmlData.Columns[1].ReadOnly := true;

      Exit;
    end
    else if dsrSo112A.State in [dsEdit, dsInsert] then
    begin
      dbgPercentXmlData.Columns[1].ReadOnly := false;

      miInsert.Enabled := false;
      miEdit.Enabled := false;
      miDelete.Enabled := false;
      miCancel.Enabled := false;
      miSave.Enabled := false;
    end
    else
    begin
      if dsrPercentXmlData.DataSet.State = dsBrowse then
      begin
        dbgPercentXmlData.Columns[1].ReadOnly := true;

        miInsert.Enabled := true;
        if dsrPercentXmlData.DataSet.RecordCount>0 then
        begin
          miEdit.Enabled := true;
          miDelete.Enabled := true;
        end
        else
        begin
          miEdit.Enabled := false;
          miDelete.Enabled := false;
        end;
        miCancel.Enabled := false;
        miSave.Enabled := false;
      end
      else if dsrPercentXmlData.DataSet.State in [dsEdit, dsInsert] then
      begin
        dbgPercentXmlData.Columns[1].ReadOnly := false;

        miInsert.Enabled := false;
        miEdit.Enabled := false;
        miDelete.Enabled := false;
        miCancel.Enabled := true;
        miSave.Enabled := true;
      end;
    end;
end;

procedure TfrmSO8A60.miInsertClick(Sender: TObject);
begin
    if dsrPercentXmlData.DataSet.State = dsBrowse then
    begin
      dsrPercentXmlData.DataSet.Append;
      changePercentDataBtnState;
    end;
end;

procedure TfrmSO8A60.miEditClick(Sender: TObject);
begin
    if (dsrPercentXmlData.DataSet.State = dsBrowse) and (dsrPercentXmlData.DataSet.RecordCount>0) then
    begin
      dsrPercentXmlData.DataSet.Edit;
      changePercentDataBtnState;
    end;
end;

procedure TfrmSO8A60.miDeleteClick(Sender: TObject);
begin
    if (dsrPercentXmlData.DataSet.State = dsBrowse) and (dsrPercentXmlData.DataSet.RecordCount>0) then
    begin
      if MessageDlg('是否確認刪除?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
      begin
        dsrPercentXmlData.DataSet.Delete;
        refreshXmlData;        
        changePercentDataBtnState;
      end;
    end;
end;

procedure TfrmSO8A60.miCancelClick(Sender: TObject);
begin
    if (dsrPercentXmlData.DataSet.State in [dsEdit, dsInsert])then
    begin
      dsrPercentXmlData.DataSet.Cancel;
      changePercentDataBtnState;
    end;
end;

procedure TfrmSO8A60.miSaveClick(Sender: TObject);
begin
    if (dsrPercentXmlData.DataSet.State in [dsEdit, dsInsert])then
    begin
      dsrPercentXmlData.DataSet.Post;


      changePercentDataBtnState;
    end;
end;



procedure TfrmSO8A60.dsrPercentXmlDataStateChange(Sender: TObject);

begin
    if dtmMain2.cdsSo112A.State=dsInactive then
      exit
    else
    begin
      if dsrPercentXmlData.DataSet.State=dsBrowse then
      begin
        refreshXmlData;

  //      showmessage('browse')
      end;
      changePercentDataBtnState;      
    end;
      
end;

procedure TfrmSO8A60.refreshXmlData;
var
    sL_XmlDataStr, sL_XmlDatStr : String;
begin
    sL_XmlDataStr := dtmMain2.genXmlDataStr;
    sL_XmlDatStr := dtmMain2.genXmlDataStr1;
    dtmMain2.cdsSo112A.Edit;
    dtmMain2.cdsSo112A.FieldByName('PERCENTXMLDATA').AsString := sL_XmlDataStr;
    dtmMain2.cdsSo112A.FieldByName('AMOUNTXMLDATA').AsString := sL_XmlDatStr;
    dtmMain2.cdsSo112A.Post;
    dtmMain2.saveToDB(dtmMain2.cdsSo112A);

end;

end.


