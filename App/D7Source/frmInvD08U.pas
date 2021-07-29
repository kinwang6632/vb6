unit frmInvD08U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Grids, DBGrids, Mask, Buttons, ExtCtrls,
  cbAppClass, DB, ADODB;

type
  TfrmInvD08 = class(TForm)
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
    txtItemId: TMaskEdit;
    txtDescription: TMaskEdit;
    txtIsSelfCreated: TComboBox;
    pnlGrid: TPanel;
    DBGrid1: TDBGrid;
    Stt_Show: TStaticText;
    Label1: TLabel;
    txtRefNo: TMaskEdit;
    adoInv028: TADOQuery;
    DataSource1: TDataSource;
    LoadTimer: TTimer;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure adoInv028AfterScroll(DataSet: TDataSet);
    procedure LoadTimerTimer(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure btnAppendClick(Sender: TObject);
    procedure btnEditClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
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
    procedure ClearEditor;
    procedure SetDataSetField;
  public
    { Public declarations }
  end;

var
  frmInvD08: TfrmInvD08;

implementation

uses frmMainU, dtmMainJU, dtmMainU, cbUtilis;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD08.FormCreate(Sender: TObject);
begin
  self.Caption := frmMain.GetFormTitleString( 'D08', '發票用途代碼資料維護' );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD08.FormShow(Sender: TObject);
begin
  LoadTimer.Enabled := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD08.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
  if ( FDMLMode in [dmInsert, dmUpdate, dmDelete] ) then
    DoDMLModeChange( dmCancel );
  CanClose := True;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvD08.DoDMLModeChange(AMode: TDMLMode): Boolean;
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

{ ---------------------------------------------------------------------------- }

function TfrmInvD08.DoDMLDataSetChange(const AMode: TDMLMode): Boolean;
begin
  Result := False;
  case AMode of
    dmBrowser, dmCancel:
      begin
        if ( adoInv028.State in [dsEdit, dsInsert] ) then
        begin
          adoInv028.AfterScroll := nil;
          try
            adoInv028.Cancel;
          finally
            adoInv028.AfterScroll := adoInv028AfterScroll;
          end;
          adoInv028AfterScroll( adoInv028 );  
        end;   
      end;
    dmInsert:
      adoInv028.Append;
    dmUpdate:
      begin
        if not adoInv028.IsEmpty then
          adoInv028.Edit;
      end;
    dmDelete:
      begin
        adoInv028.Delete;
      end;
    dmSave:
      begin
        adoInv028.FieldByName( 'IDENTIFYID1' ).AsString := IDENTIFYID1;
        adoInv028.FieldByName( 'IDENTIFYID2' ).AsString := IDENTIFYID2;
        adoInv028.FieldByName( 'COMPID' ).AsString := dtmMain.getCompID;
        adoInv028.FieldByName( 'ITEMID' ).AsString := Trim( txtItemId.Text );
        adoInv028.FieldByName( 'DESCRIPTION' ).AsString := txtDescription.Text;
        adoInv028.FieldByName( 'REFNO' ).AsString := Trim( txtRefNo.Text );
        adoInv028.FieldByName( 'UPDEN' ).AsString := dtmMain.getLoginUser;
        adoInv028.FieldByName( 'UPDTIME' ).AsString := FormatDateTime( 'yyyy/mm/dd hh:nn:ss', Now );
        adoInv028.FieldByName( 'ISSELFCREATED' ).AsString := 'Y';
        if txtIsSelfCreated.ItemIndex = 1 then
          adoInv028.FieldByName( 'ISSELFCREATED' ).AsString := 'N';
        adoInv028.Post;  
      end;
  end;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvD08.DoDMLButtonChange(const AMode: TDMLMode): Boolean;
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
        if adoInv028.IsEmpty then
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

{ ---------------------------------------------------------------------------- }

function TfrmInvD08.DoDMLGridChange(const AMode: TDMLMode): Boolean;
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

{ ---------------------------------------------------------------------------- }

function TfrmInvD08.DoDMLEditorChange(const AMode: TDMLMode): Boolean;
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
        if ( adoInv028.FieldByName( 'ISSELFCREATED' ).AsString = 'N' ) then
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

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD08.OpenDataSet;
begin
  adoInv028.AfterScroll := nil;
  try
    adoInv028.Close;
    adoInv028.SQL.Text := Format(
      ' SELECT * FROM INV028           ' +
      '  WHERE IDENTIFYID1 = ''%s''    ' +
      '    AND IDENTIFYID2 = ''%s''    ' +
      '    AND COMPID = ''%s''         ' +
      '  ORDER BY ITEMID               ',
      [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID] );
    adoInv028.Open;
    adoInv028.First;
  finally
    adoInv028.AfterScroll := adoInv028AfterScroll;
  end;
  adoInv028AfterScroll( adoInv028 );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD08.adoInv028AfterScroll(DataSet: TDataSet);
begin
  ClearEditor;
  if not adoInv028.IsEmpty then
  begin
    txtItemId.Text := adoInv028.FieldByName( 'ItemId' ).AsString;
    txtDescription.Text := adoInv028.FieldByName( 'Description' ).AsString;
    txtIsSelfCreated.ItemIndex := 1;
    if adoInv028.FieldByName( 'IsSelfCreated' ).AsString = 'Y' then
      txtIsSelfCreated.ItemIndex := 0;
    txtRefNo.Text := adoInv028.FieldByName( 'RefNo' ).AsString;  
  end;
  Stt_Show.Caption := Format( '%d　/　%d', [adoInv028.RecNo, adoInv028.RecordCount] );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD08.LoadTimerTimer(Sender: TObject);
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

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD08.ClearEditor;
begin
  txtItemId.Text := EmptyStr;
  txtDescription.Text := EmptyStr;
  txtIsSelfCreated.ItemIndex := 0;
  txtRefNo.Text := EmptyStr;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD08.SetDataSetField;
begin
  adoInv028.FieldByName( 'ITEMID' ).DisplayLabel := '發票用途代碼';
  adoInv028.FieldByName( 'DESCRIPTION' ).DisplayLabel := '發票用途說明';
  adoInv028.FieldByName( 'REFNO' ).DisplayLabel := '參考號';
  adoInv028.FieldByName( 'UPDEN' ).DisplayLabel := '異動人員';
  adoInv028.FieldByName( 'UPDTIME' ).DisplayLabel := '異動時間';
  adoInv028.FieldByName( 'ISSELFCREATED' ).Visible := False;
  adoInv028.FieldByName( 'IDENTIFYID1' ).Visible := False;
  adoInv028.FieldByName( 'IDENTIFYID2' ).Visible := False;
  adoInv028.FieldByName( 'COMPID' ).Visible := False;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD08.btnExitClick(Sender: TObject);
begin
  Self.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD08.btnAppendClick(Sender: TObject);
begin
  DoDMLModeChange( dmInsert );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD08.btnEditClick(Sender: TObject);
begin
  DoDMLModeChange( dmUpdate );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD08.btnDeleteClick(Sender: TObject);
begin
  if ConfirmMsg( '確認要刪除此筆資料?' ) then
  begin
    DoDMLModeChange( dmDelete );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD08.btnCancelClick(Sender: TObject);
begin
  DoDMLModeChange( dmCancel );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD08.btnSaveClick(Sender: TObject);
begin
  if IsDataOK then
    DoDMLModeChange( dmSave );
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvD08.IsDataOK: Boolean;
begin
  Result := False;
  if ( Trim( txtItemId.Text ) = EmptyStr ) then
  begin
    WarningMsg( '請輸入發票用途代碼。' );
    if txtItemId.CanFocus then txtItemId.SetFocus;
    Exit;
  end;
  {}
  if ( Trim( txtDescription.Text ) = EmptyStr ) then
  begin
    WarningMsg( '請輸入發票用途品名。' );
    if txtDescription.CanFocus then txtDescription.SetFocus;
    Exit;
  end;
  {}
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

end.
