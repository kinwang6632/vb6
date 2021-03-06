unit frmInvD0CU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Mask, Buttons, Grids, DBGrids, DB, ADODB,cbAppClass;

type
  TfrmInvD0C = class(TForm)
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
    lbl5: TLabel;
    txtCodeNo: TMaskEdit;
    txtDescription: TMaskEdit;
    txtRefNo: TMaskEdit;
    txtStt_Show: TStaticText;
    tmrLoadTimer: TTimer;
    adoInv044: TADOQuery;
    ds1: TDataSource;
    pnlGrid: TPanel;
    dbgrd1: TDBGrid;
    procedure FormCreate(Sender: TObject);
    procedure tmrLoadTimerTimer(Sender: TObject);
    procedure OpenDataSet;
    procedure adoInv044AfterScroll(DataSet: TDataSet);

    procedure FormShow(Sender: TObject);
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
     function DoDMLEditorChange(const AMode: TDMLMode): Boolean;
     function DoDMLGridChange(const AMode: TDMLMode): Boolean;
     function IsDataOK: Boolean;
     procedure ClearEditor;
     procedure SetDataSetField;

  public
    { Public declarations }
  end;

var
  frmInvD0C: TfrmInvD0C;

implementation

uses frmMainU, dtmMainJU, dtmMainU, cbUtilis;

{$R *.dfm}

procedure TfrmInvD0C.FormCreate(Sender: TObject);
begin
  self.Caption := frmMain.GetFormTitleString( 'D0C', '客戶載具類別代碼維護' );
end;

procedure TfrmInvD0C.OpenDataSet;
begin
  adoInv044.AfterScroll := nil;
  try
    adoInv044.Close;
    adoInv044.SQL.Text := format(' SELECT * FROM INV044           ' +
            ' WHERE COMPID = ''%s''                               ' +
            ' ORDER BY CODENO ',[dtmMain.getCompID]);
    adoInv044.Open;
    adoInv044.First;
  finally
    adoInv044.AfterScroll := adoInv044AfterScroll;
  end;
  adoInv044AfterScroll( adoInv044 );
end;

procedure TfrmInvD0C.tmrLoadTimerTimer(Sender: TObject);
begin
  tmrLoadTimer.Enabled := False;
  Screen.Cursor := crSQLWait;
  try
    OpenDataSet;
    SetDataSetField;
    DoDMLModeChange( dmBrowser );
  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TfrmInvD0C.adoInv044AfterScroll(DataSet: TDataSet);
 var a:Integer;
begin
  ClearEditor;
  if not adoInv044.IsEmpty then
  begin
    txtCodeNo.Text := adoInv044.FieldByName( 'CODENO' ).AsString;
    txtDescription.Text := adoInv044.FieldByName( 'DESCRIPTION' ).AsString;
    txtRefNo.Text := adoInv044.FieldByName( 'REFNO' ).AsString;
  end;
  txtStt_Show.Caption := Format( '%d　/　%d', [adoInv044.RecNo, adoInv044.RecordCount] );
end;

procedure TfrmInvD0C.SetDataSetField;
begin
  adoInv044.FieldByName( 'CODENO' ).DisplayLabel := '客戶載具代碼';
  adoInv044.FieldByName( 'DESCRIPTION' ).DisplayLabel := '客戶載具名稱';
  adoInv044.FieldByName( 'REFNO' ).DisplayLabel := '參考號';
  adoInv044.FieldByName( 'UPDEN' ).DisplayLabel := '異動人員';
  adoInv044.FieldByName( 'UPDTIME' ).DisplayLabel := '異動時間';
  adoInv044.FieldByName( 'COMPID' ).Visible := False;
end;

procedure TfrmInvD0C.FormShow(Sender: TObject);
begin
  tmrLoadTimer.Enabled := True;
end;

function TfrmInvD0C.DoDMLDataSetChange(const AMode: TDMLMode): Boolean;
begin
  Result := False;
  case AMode of
    dmBrowser, dmCancel:
      begin
        if ( adoInv044.State in [dsEdit, dsInsert] ) then
        begin
          adoInv044.AfterScroll := nil;
          try
            adoInv044.Cancel;
          finally
            adoInv044.AfterScroll := adoInv044AfterScroll;
          end;
          adoInv044AfterScroll( adoInv044 );
        end;   
      end;
    dmInsert:
      adoInv044.Append;
    dmUpdate:
      begin
        if not adoInv044.IsEmpty then
          adoInv044.Edit;
      end;
    dmDelete:
      begin
        adoInv044.Delete;
      end;
    dmSave:
      begin
        adoInv044.FieldByName( 'COMPID' ).AsString := dtmMain.getCompID;
        adoInv044.FieldByName( 'CODENO' ).AsString := Trim( txtCodeNo.Text );
        adoInv044.FieldByName( 'DESCRIPTION' ).AsString := txtDescription.Text;
        adoInv044.FieldByName( 'REFNO' ).AsString := Trim( txtRefNo.Text );
        adoInv044.FieldByName( 'UPDEN' ).AsString := dtmMain.getLoginUser;
        adoInv044.FieldByName( 'UPDTIME' ).AsString := FormatDateTime( 'yyyy/mm/dd hh:nn:ss', Now );
        adoInv044.Post;  
      end;
  end;
  Result := True;
end;

function TfrmInvD0C.DoDMLButtonChange(const AMode: TDMLMode): Boolean;
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
        if adoInv044.IsEmpty then
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

function TfrmInvD0C.DoDMLEditorChange(const AMode: TDMLMode): Boolean;
begin
   case AMode of
    dmBrowser, dmCancel:
      begin
        txtCodeNo.ReadOnly := True;
        txtDescription.ReadOnly := True;
        txtRefNo.ReadOnly := True;
      end;
    dmInsert:
      begin
        txtCodeNo.Enabled := True;
        txtDescription.Enabled := True;
        txtRefNo.Enabled := True;
        {}
        txtCodeNo.ReadOnly := False;
        txtDescription.ReadOnly := False;
        txtRefNo.ReadOnly := False;
        txtCodeNo.SetFocus;
      end;
    dmUpdate:
      begin
        txtCodeNo.Enabled := True;
        txtDescription.Enabled := True;
        txtRefNo.Enabled := True;
        {}
        txtCodeNo.ReadOnly := True;
        txtDescription.ReadOnly := False;
        txtRefNo.ReadOnly := False;
        if txtDescription.CanFocus then txtDescription.SetFocus;
      end;

  else
    pnlGrid.Enabled := True;
  end;
  Result := True;
end;

function TfrmInvD0C.IsDataOK: Boolean;
begin
  Result := False;
  if ( Trim( txtCodeNo.Text ) = EmptyStr ) then
  begin
    WarningMsg( '請輸入客戶載具代碼。' );
    if txtCodeNo.CanFocus then txtCodeNo.SetFocus;
    Exit;
  end;
  {}
  if ( Trim( txtDescription.Text ) = EmptyStr ) then
  begin
    WarningMsg( '請輸入客戶載具名稱。' );
    if txtDescription.CanFocus then txtDescription.SetFocus;
    Exit;
  end;
  {}
  Result := True;
end;

procedure TfrmInvD0C.ClearEditor;
begin
  txtCodeNo.Text := EmptyStr;
  txtDescription.Text := EmptyStr;
  txtRefNo.Text := EmptyStr;
end;

function TfrmInvD0C.DoDMLModeChange(AMode: TDMLMode): Boolean;
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

function TfrmInvD0C.DoDMLGridChange(const AMode: TDMLMode): Boolean;
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

procedure TfrmInvD0C.btnAppendClick(Sender: TObject);
begin
  DoDMLModeChange( dmInsert );
end;

procedure TfrmInvD0C.btnEditClick(Sender: TObject);
begin
  DoDMLModeChange( dmUpdate );
end;

procedure TfrmInvD0C.btnDeleteClick(Sender: TObject);
begin
  if ConfirmMsg( '確認要刪除此筆資料?' ) then
  begin
    DoDMLModeChange( dmDelete );
  end;
end;

procedure TfrmInvD0C.btnCancelClick(Sender: TObject);
begin
  DoDMLModeChange( dmCancel );
end;

procedure TfrmInvD0C.btnSaveClick(Sender: TObject);
begin
  if IsDataOK then
    DoDMLModeChange( dmSave );
end;

procedure TfrmInvD0C.btnExitClick(Sender: TObject);
begin
  Self.Close;
end;

end.
