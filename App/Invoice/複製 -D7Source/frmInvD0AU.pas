unit frmInvD0AU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Grids, DBGrids, Mask, Buttons, ExtCtrls,
  cbAppClass, DB, ADODB;

type
  TfrmInvD0A = class(TForm)
    Panel1: TPanel;
    Label5: TLabel;
    Panel4: TPanel;
    btnAppend: TBitBtn;
    btnEdit: TBitBtn;
    btnDelete: TBitBtn;
    btnCancel: TBitBtn;
    btnSave: TBitBtn;
    btnExit: TBitBtn;
    pnlEdit: TPanel;
    Label8: TLabel;
    Label9: TLabel;
    Label23: TLabel;
    Label1: TLabel;
    txtItemId: TMaskEdit;
    txtDescription: TMaskEdit;
    txtIsSelfCreated: TComboBox;
    txtRefNo: TMaskEdit;
    pnlGrid: TPanel;
    DBGrid1: TDBGrid;
    Stt_Show: TStaticText;
    DataSource1: TDataSource;
    adoInv041: TADOQuery;
    LoadTimer: TTimer;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure LoadTimerTimer(Sender: TObject);
    procedure adoInv041AfterScroll(DataSet: TDataSet);
    procedure btnAppendClick(Sender: TObject);
    procedure btnEditClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure btnExitClick(Sender: TObject);

  private
    { Private declarations }
    FDMLMode: TDMLMode;
    function DoDMLModeChange(AMode: TDMLMode): Boolean;
    function DoDMLDataSetChange(const AMode: TDMLMode): Boolean;
    function DoDMLButtonChange(const AMode: TDMLMode): Boolean;
    function DoDMLGridChange(const AMode: TDMLMode): Boolean;
    function DoDMLEditorChange(const AMode: TDMLMode): Boolean;
    function IsDataOK: Boolean;
    procedure OpenDataSet;
    procedure SetDataSetField;
    procedure ClearEditor;
  public
    { Public declarations }
  end;

var
  frmInvD0A: TfrmInvD0A;

implementation
uses frmMainU, dtmMainJU, dtmMainU, cbUtilis;
{$R *.dfm}

function TfrmInvD0A.DoDMLButtonChange(const AMode: TDMLMode): Boolean;
begin
   case AMode of
    dmBrowser, dmCancel:
      begin
        btnAppend.Enabled := True;
        btnEdit.Enabled := True;
        btnDelete.Enabled := True;
        btnCancel.Enabled := False;
        btnSave.Enabled := False;
        btnExit.Enabled := True;
        if adoInv041.IsEmpty then
        begin
          btnEdit.Enabled := False;
          btnDelete.Enabled := False;
        end;
      end;
    dmInsert:
      begin
        btnAppend.Enabled := False;
        btnEdit.Enabled := False;
        btnDelete.Enabled := False;
        btnCancel.Enabled := True;
        btnSave.Enabled := True;
        btnExit.Enabled := False;
      end;
    dmUpdate:
      begin
        btnAppend.Enabled := False;
        btnEdit.Enabled := False;
        btnDelete.Enabled := False;
        btnCancel.Enabled := True;
        btnSave.Enabled := True;
        btnExit.Enabled := False;
      end;
  else
    begin
      btnAppend.Enabled := True;
      btnEdit.Enabled := True;
      btnDelete.Enabled := True;
      btnCancel.Enabled := False;
      btnSave.Enabled := False;
      btnExit.Enabled := True;
    end;
  end;
  Result := True;
end;

function TfrmInvD0A.DoDMLDataSetChange(const AMode: TDMLMode): Boolean;
begin
  Result := False;
  case AMode of
    dmBrowser, dmCancel:
      begin
        if ( adoInv041.State in [dsEdit, dsInsert] ) then
        begin
          adoInv041.AfterScroll := nil;
          try
            adoInv041.Cancel;
          finally
            adoInv041.AfterScroll := adoInv041AfterScroll;
          end;
          adoInv041AfterScroll( adoInv041 );
        end;
      end;
    dmInsert:
      adoInv041.Append;
    dmUpdate:
      begin
        if not adoInv041.IsEmpty then
          adoInv041.Edit;
      end;
    dmDelete:
      begin
        adoInv041.Delete;
      end;
    dmSave:
      begin
        adoInv041.FieldByName( 'IDENTIFYID1' ).AsString := IDENTIFYID1;
        adoInv041.FieldByName( 'IDENTIFYID2' ).AsString := IDENTIFYID2;
        adoInv041.FieldByName( 'COMPID' ).AsString := dtmMain.getCompID;
        adoInv041.FieldByName( 'ITEMID' ).AsString := Trim( txtItemId.Text );
        adoInv041.FieldByName( 'DESCRIPTION' ).AsString := txtDescription.Text;
        adoInv041.FieldByName( 'REFNO' ).AsString := Trim( txtRefNo.Text );
        adoInv041.FieldByName( 'UPDEN' ).AsString := dtmMain.getLoginUser;
        adoInv041.FieldByName( 'UPDTIME' ).AsString := FormatDateTime( 'yyyy/mm/dd hh:nn:ss', Now );
        adoInv041.FieldByName( 'ISSELFCREATED' ).AsString := 'Y';
        if txtIsSelfCreated.ItemIndex = 1 then
          adoInv041.FieldByName( 'ISSELFCREATED' ).AsString := 'N';
        adoInv041.Post;
      end;
  end;
  Result := True;
end;

function TfrmInvD0A.DoDMLEditorChange(const AMode: TDMLMode): Boolean;
begin
  case AMode of
    dmBrowser, dmCancel:
      begin
        txtItemId.ReadOnly := True;
        txtDescription.ReadOnly := True;
        txtIsSelfCreated.Enabled := False;
        txtRefNo.ReadOnly := True;
      end;
    dmInsert:
      begin
        txtItemId.Enabled := True;
        txtDescription.Enabled := True;
        txtIsSelfCreated.Enabled := True;
        txtRefNo.Enabled := True;
        {}
        txtItemId.ReadOnly := False;
        txtDescription.ReadOnly := False;
        txtIsSelfCreated.Enabled := False;
        txtRefNo.ReadOnly := False;
        txtIsSelfCreated.ItemIndex := 0;
        if txtItemId.CanFocus then txtItemId.SetFocus;
      end;
    dmUpdate:
      begin
        if ( adoInv041.FieldByName( 'ISSELFCREATED' ).AsString = 'N' ) then
        begin
          txtItemId.Enabled := False;
          txtDescription.Enabled := False;
          txtIsSelfCreated.Enabled := False;
          txtRefNo.Enabled := False;
        end else
        begin
          txtItemId.Enabled := True;
          txtDescription.Enabled := True;
          txtIsSelfCreated.Enabled := True;
          txtRefNo.Enabled := True;
          {}
          txtItemId.ReadOnly := True;
          txtDescription.ReadOnly := False;
          txtIsSelfCreated.Enabled := False;
          txtRefNo.ReadOnly := False;
          if txtDescription.CanFocus then txtDescription.SetFocus;
        end;
      end;
  else
    pnlGrid.Enabled := True;
  end;
  Result := True;
end;

function TfrmInvD0A.DoDMLGridChange(const AMode: TDMLMode): Boolean;
begin
  case AMode of
    dmBrowser, dmCancel:
      pnlGrid.Enabled := True;
    dmInsert, dmUpdate:
      pnlGrid.Enabled := False;
  else
    pnlGrid.Enabled := True;
  end;
  Result := True;
end;

function TfrmInvD0A.DoDMLModeChange(AMode: TDMLMode): Boolean;
begin
  if DoDMLDataSetChange( AMode ) then
  begin
    if ( AMode in [dmSave, dmCancel] ) then AMode := dmBrowser;
    DoDMLButtonChange( AMode );
    DoDMLGridChange( AMode );
    DoDMLEditorChange( AMode );
    FDMLMode := AMode;
  end;
end;

procedure TfrmInvD0A.FormCreate(Sender: TObject);
begin
  self.Caption := frmMain.GetFormTitleString( 'D0A', '受贈單位代碼資料維護' );

end;

procedure TfrmInvD0A.FormShow(Sender: TObject);
begin
  LoadTimer.Enabled := True;
end;

procedure TfrmInvD0A.OpenDataSet;
begin
  adoInv041.AfterScroll := nil;
  try
    adoInv041.Close;
    adoInv041.SQL.Text := Format(
      ' SELECT * FROM INV041           ' +
      '  WHERE IDENTIFYID1 = ''%s''    ' +
      '    AND IDENTIFYID2 = ''%s''    ' +
      '    AND COMPID = ''%s''         ' +
      '  ORDER BY ITEMID               ',
      [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID] );
    adoInv041.Open;
    adoInv041.First;
  finally
    adoInv041.AfterScroll := adoInv041AfterScroll;
  end;
  adoInv041AfterScroll( adoInv041 );
end;

procedure TfrmInvD0A.SetDataSetField;
begin
  adoInv041.FieldByName( 'ITEMID' ).DisplayLabel := '受贈單位代碼';
  adoInv041.FieldByName( 'DESCRIPTION' ).DisplayLabel := '受贈單位名稱';
  adoInv041.FieldByName( 'REFNO' ).DisplayLabel := '參考號';
  adoInv041.FieldByName( 'UPDEN' ).DisplayLabel := '異動人員';
  adoInv041.FieldByName( 'UPDTIME' ).DisplayLabel := '異動時間';
  adoInv041.FieldByName( 'ISSELFCREATED' ).Visible := False;
  adoInv041.FieldByName( 'IDENTIFYID1' ).Visible := False;
  adoInv041.FieldByName( 'IDENTIFYID2' ).Visible := False;
  adoInv041.FieldByName( 'COMPID' ).Visible := False;
  adoInv041.FieldByName( 'STOPFLAG' ).Visible := False;
end;

procedure TfrmInvD0A.LoadTimerTimer(Sender: TObject);
begin
  LoadTimer.Enabled := False;
  Screen.Cursor := crSQLWait;
  try
    OpenDataSet;
    SetDataSetField;
    DoDMLModeChange( dmBrowser );
  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TfrmInvD0A.ClearEditor;
begin
  txtItemId.Text := EmptyStr;
  txtDescription.Text := EmptyStr;
  txtIsSelfCreated.ItemIndex := 0;
  txtRefNo.Text := EmptyStr;
end;

procedure TfrmInvD0A.adoInv041AfterScroll(DataSet: TDataSet);
begin
  if not adoInv041.IsEmpty then
  begin
    txtItemId.Text := adoInv041.FieldByName( 'ItemId' ).AsString;
    txtDescription.Text := adoInv041.FieldByName( 'Description' ).AsString;
    txtIsSelfCreated.ItemIndex := 1;
    if adoInv041.FieldByName( 'IsSelfCreated' ).AsString = 'Y' then
      txtIsSelfCreated.ItemIndex := 0;
    txtRefNo.Text := adoInv041.FieldByName( 'RefNo' ).AsString;
  end;
  Stt_Show.Caption := Format( '%d　/　%d', [adoInv041.RecNo, adoInv041.RecordCount] );
end;

procedure TfrmInvD0A.btnAppendClick(Sender: TObject);
begin
  DoDMLModeChange( dmInsert );
end;

procedure TfrmInvD0A.btnEditClick(Sender: TObject);
begin
  DoDMLModeChange( dmUpdate );
end;

procedure TfrmInvD0A.btnDeleteClick(Sender: TObject);
begin
  if ConfirmMsg( '確認要刪除此筆資料?' ) then
  begin
    DoDMLModeChange( dmDelete );
  end;
end;

procedure TfrmInvD0A.btnCancelClick(Sender: TObject);
begin
  DoDMLModeChange( dmCancel );
end;

procedure TfrmInvD0A.btnSaveClick(Sender: TObject);
begin
  if IsDataOK then
    DoDMLModeChange( dmSave );
end;

function TfrmInvD0A.IsDataOK: Boolean;
begin
  Result := False;
  if ( Trim( txtItemId.Text ) = EmptyStr ) then
  begin
    WarningMsg( '請輸入受贈單位代碼。' );
    if txtItemId.CanFocus then txtItemId.SetFocus;
    Exit;
  end;
  {}
  if ( Trim( txtDescription.Text ) = EmptyStr ) then
  begin
    WarningMsg( '請輸入受贈單位名稱。' );
    if txtDescription.CanFocus then txtDescription.SetFocus;
    Exit;
  end;
  {}
  Result := True;
end;

procedure TfrmInvD0A.btnExitClick(Sender: TObject);
begin
  Self.Close;
end;

end.



