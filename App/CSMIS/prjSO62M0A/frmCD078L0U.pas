unit frmCD078L0U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs,DB, DBClient, ExtCtrls, cxStyles, dxSkinsCore, dxSkinBlack,
  dxSkinBlue, dxSkinCaramel, dxSkinCoffee, dxSkinsDefaultPainters,
  dxSkinscxPCPainter, cxCustomData, cxGraphics, cxFilter, cxData,
  cxDataStorage, cxEdit, cxDBData, cxTextEdit, cxLabel, cxCheckBox,
  cxGridLevel, cxGridCustomTableView, cxGridTableView,
  cxGridBandedTableView, cxGridDBBandedTableView, cxClasses, cxControls,
  cxGridCustomView, cxGrid, cxContainer, cxCurrencyEdit, StdCtrls, ActnList,
  cbDataLookup, cxMaskEdit, Menus, cxGridDBTableView;

type
   TDML = ( dmBrowser = 0,
           dmInsert = 1,
           dmUpdate = 2,
           dmDelete = 3,
           dmCancel = 4,
           dmSave = 5 );
  TfrmCD078L = class(TForm)
    pnl2: TPanel;
    CD078LDataSetSource: TClientDataSet;
    dsCD078L: TDataSource;
    CD078LGrid: TcxGrid;
    CD078LBandedTableView: TcxGridDBBandedTableView;
    cxGridDBBandedColumn1: TcxGridDBBandedColumn;
    cxGridDBBandedColumn2: TcxGridDBBandedColumn;
    cxGridDBBandedColumn3: TcxGridDBBandedColumn;
    cxGridDBBandedColumn4: TcxGridDBBandedColumn;
    cxGridDBBandedColumn5: TcxGridDBBandedColumn;
    cxGridDBBandedColumn6: TcxGridDBBandedColumn;
    cxGridDBBandedColumn7: TcxGridDBBandedColumn;
    cxGridDBBandedColumn8: TcxGridDBBandedColumn;
    cxGridDBBandedColumn9: TcxGridDBBandedColumn;
    cxGridDBBandedColumn10: TcxGridDBBandedColumn;
    cxGridDBBandedColumn11: TcxGridDBBandedColumn;
    cxGridDBBandedColumn12: TcxGridDBBandedColumn;
    cxGridDBBandedColumn13: TcxGridDBBandedColumn;
    cxGridDBBandedColumn14: TcxGridDBBandedColumn;
    cxGridLevel2: TcxGridLevel;
    pnl3: TPanel;
    ActionList1: TActionList;
    actSave: TAction;
    actCancel: TAction;
    actUpdate: TAction;
    actInsert: TAction;
    btnUpdate: TButton;
    btnSave: TButton;
    btnInsert: TButton;
    btnCancel: TButton;
    lbl1: TLabel;
    ServiceType: TDataLookup;
    edtMonth: TcxMaskEdit;
    lbl2: TLabel;
    lbl3: TLabel;
    edtCashAdj: TcxMaskEdit;
    lbl4: TLabel;
    edtSubTotal: TcxMaskEdit;
    lbl5: TLabel;
    edtCMAmt: TcxMaskEdit;
    lbl6: TLabel;
    edtCMTax: TcxMaskEdit;
    lbl7: TLabel;
    edtCMAdj: TcxMaskEdit;
    lbl8: TLabel;
    edtDTVAmt: TcxMaskEdit;
    lbl9: TLabel;
    edtDTVTax: TcxMaskEdit;
    lbl10: TLabel;
    edtDTVAdj: TcxMaskEdit;
    lbl11: TLabel;
    edtPRVAmt: TcxMaskEdit;
    lbl12: TLabel;
    edtPRVTax: TcxMaskEdit;
    lbl13: TLabel;
    chkStopFlag: TcxCheckBox;
    edtPRVAdj: TcxMaskEdit;
    btnDelete: TButton;
    actDelete: TAction;
    procedure FormDestroy(Sender: TObject);
    procedure ServiceTypeCodeNamePropertiesInitPopup(Sender: TObject);
    procedure ServiceTypeCodeNamePropertiesChange(Sender: TObject);
    procedure ServiceTypeCodeNoPropertiesChange(Sender: TObject);
    procedure CD078LDataSetSourceAfterScroll(DataSet: TDataSet);
    procedure FormShow(Sender: TObject);
    procedure btnUpdateClick(Sender: TObject);
    procedure actUpdateExecute(Sender: TObject);
    procedure actInsertExecute(Sender: TObject);
    procedure actSaveExecute(Sender: TObject);
    procedure actCancelExecute(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure btnInsertClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure actDeleteExecute(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);


  private
    { Private declarations }
    FMode:TDML;
    FKeyValue: String;
    procedure DMLModeChange(aMode: TDML);
    function DataSetStateChange(const aMode: TDML): Boolean;
    function DataValidate: Boolean;
    procedure ClearEditor;
    procedure DataToEditor;
    procedure EditorStateChange;
    procedure ButtonStateChange;
    function EditorToData : Boolean ;
  public
    { Public declarations }
    constructor Create( aKeyValue: String;aDataSet: TClientDataSet ); reintroduce;
  end;

var
  frmCD078L: TfrmCD078L;

implementation

uses cbDBController,cbUtilis;

{$R *.dfm}

{ TfrmCD078L }



{ TfrmCD078L }

constructor TfrmCD078L.Create(aKeyValue: String; aDataSet: TClientDataSet);
begin
  inherited Create( Application );
  //FMode := aMode;
  FKeyValue := aKeyValue;
  CD078LDataSetSource := aDataSet;
end;

procedure TfrmCD078L.FormDestroy(Sender: TObject);
begin
  CD078LDataSetSource.AfterScroll := nil;
  ServiceType.Finaliza;
end;

procedure TfrmCD078L.ServiceTypeCodeNamePropertiesInitPopup(
  Sender: TObject);
begin
  { 服務類別 }
  Screen.Cursor := crSQLWait;
  try
    ServiceType.SQL.Text := Format(
    ' SELECT CODENO, DESCRIPTION    ' +
    '  FROM %s.CD046                ',
    [DBController.LoginInfo.DbAccount] );
    ServiceType.CodeNamePropertiesInitPopup( Sender );
  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TfrmCD078L.ServiceTypeCodeNamePropertiesChange(Sender: TObject);
begin
  ServiceType.CodeNamePropertiesChange( Sender );
end;

procedure TfrmCD078L.ServiceTypeCodeNoPropertiesChange(Sender: TObject);
begin
   ServiceType.CodeNoPropertiesChange( Sender );
end;

procedure TfrmCD078L.DMLModeChange(aMode: TDML);
begin
  if not DataSetStateChange( aMode ) then Exit;
  if ( aMode = dmSave  ) and ( FMode = dmInsert ) then
  begin
    FMode := dmBrowser;
    DMLModeChange( dmBrowser );
    Exit;
  end else
  if ( aMode in [dmSave, dmCancel, dmDelete] ) then
    FMode := dmBrowser;
  ButtonStateChange;
  EditorStateChange;

end;

function TfrmCD078L.DataSetStateChange(const aMode: TDML): Boolean;
begin
  Result := True;
  try
     case aMode of
     dmInsert:
      begin
        CD078LDataSetSource.AfterScroll := nil;
        CD078LDataSetSource.Append;
        ClearEditor;
      end;
     dmUpdate:
      begin
        CD078LDataSetSource.AfterScroll := nil;
//        ClearEditor;
//        DataToEditor;
        CD078LDataSetSource.Edit;
      end;
      dmCancel:
      begin
        CD078LDataSetSource.Cancel;
        ClearEditor;
        DataToEditor;
      end;
      dmBrowser:
      begin
        ClearEditor;
        DataToEditor;
      end;
      dmDelete:
      begin
        CD078LDataSetSource.AfterScroll := nil;
        CD078LDataSetSource.Delete;
        ClearEditor;
        DataToEditor;
      end;
      dmSave:
      begin
        Screen.Cursor := crSQLWait;
        try
          Result := DataValidate;
          if Result then Result := EditorToData;
          if Result then
          begin
            try
              CD078LDataSetSource.Post;
            except
              on E: Exception do
              begin
                ErrorMsg( Format( '存檔有誤, 原因:%s。', [E.Message] ) );
                Result := False;
              end;
            end;

          end;
        finally

          Screen.Cursor := crDefault;
        end;
      end;
    end;
  finally
    CD078LDataSetSource.AfterScroll := CD078LDataSetSourceAfterScroll;
  end;


end;

procedure TfrmCD078L.ClearEditor;
begin
  ServiceType.Clear;
  edtMonth.Clear;
  edtCashAdj.Clear;
  edtSubTotal.Clear;
  chkStopFlag.Checked := False;
  edtCMAmt.Clear;
  edtCMTax.Clear;
  edtCMAdj.Clear;
  edtDTVAmt.Clear;
  edtDTVTax.Clear;
  edtDTVAdj.Clear;
  edtPRVAmt.Clear;
  edtPRVTax.Clear;
  edtPRVAdj.Clear;
end;

procedure TfrmCD078L.CD078LDataSetSourceAfterScroll(DataSet: TDataSet);
begin
  ServiceType.Value := CD078LDataSetSource.FieldByName('MasterServicetype').AsString;
  edtMonth.EditValue := CD078LDataSetSource.FieldByName('Month').AsVariant;
  edtCashAdj.EditValue := CD078LDataSetSource.FieldByName('CashAdj').AsVariant;
  if ( CD078LDataSetSource.FieldByName('StopFlag').AsString = '1' ) then
    chkStopFlag.Checked := True
  else
    chkStopFlag.Checked := False;
  edtSubTotal.EditValue := CD078LDataSetSource.FieldByName('SubTotal').AsVariant;
  edtCMAmt.EditValue := CD078LDataSetSource.FieldByName('CMAmt').AsVariant;
  edtCMTax.EditValue := CD078LDataSetSource.FieldByName('CMTax').AsVariant;
  edtCMAdj.EditValue := CD078LDataSetSource.FieldByName('CMAdj').AsVariant;

  edtDTVAmt.EditValue := CD078LDataSetSource.FieldByName('DTVAmt').AsVariant;
  edtDTVTax.EditValue := CD078LDataSetSource.FieldByName('DTVTax').AsVariant;
  edtDTVAdj.EditValue := CD078LDataSetSource.FieldByName('DTVAdj').AsVariant;

  edtPRVAmt.EditValue := CD078LDataSetSource.FieldByName('PRVAmt').AsVariant;
  edtPRVTax.EditValue := CD078LDataSetSource.FieldByName('PRVTax').AsVariant;
  edtPRVAdj.EditValue := CD078LDataSetSource.FieldByName('PRVAdj').AsVariant;
end;

procedure TfrmCD078L.DataToEditor;
begin
  CD078LDataSetSourceAfterScroll( CD078LDataSetSource );
end;

function TfrmCD078L.DataValidate: Boolean;
begin
  Result := False;
  if ServiceType.Value = '' then
  begin
    WarningMsg('請設定主服務別！');
    if ServiceType.CodeNo.CanFocus then
      ServiceType.CodeNo.SetFocus;
    Exit;
  end;
  edtMonth.PostEditValue;
  if edtMonth.EditText = '' then
  begin
    WarningMsg('請設定月數！');
    if edtMonth.CanFocus then
      edtMonth.SetFocus;
    Exit;
  end;

  Result := True;
end;

function TfrmCD078L.EditorToData: Boolean;
begin
  Result := False;
  try
    if (FMode in [dmInsert] ) then
    begin
      CD078LDataSetSource.FieldByName( 'BPCODE' ).AsString := FKeyValue;
    end;
    CD078LDataSetSource.FieldByName( 'Month' ).AsVariant := edtMonth.EditValue;
    CD078LDataSetSource.FieldByName( 'MasterServicetype' ).AsVariant := ServiceType.Value;
    CD078LDataSetSource.FieldByName( 'CashAdj' ).AsVariant := edtCashAdj.EditValue;
    CD078LDataSetSource.FieldByName( 'SubTotal' ).AsVariant := edtSubTotal.EditValue;
    CD078LDataSetSource.FieldByName( 'CMAMT' ).AsVariant := edtCMAmt.EditValue;
    CD078LDataSetSource.FieldByName( 'CMTax' ).AsVariant := edtCMTax.EditValue;
    CD078LDataSetSource.FieldByName( 'CMAdj' ).AsVariant := edtCMAdj.EditValue;
    CD078LDataSetSource.FieldByName( 'DTVAMT' ).AsVariant := edtDTVAmt.EditValue;
    CD078LDataSetSource.FieldByName( 'DTVTax' ).AsVariant := edtDTVTax.EditValue;
    CD078LDataSetSource.FieldByName( 'DTVAdj' ).AsVariant := edtDTVAdj.EditValue;
    CD078LDataSetSource.FieldByName( 'PRVAMT' ).AsVariant := edtPRVAmt.EditValue;
    CD078LDataSetSource.FieldByName( 'PRVTAX' ).AsVariant := edtPRVTax.EditValue;
    CD078LDataSetSource.FieldByName( 'PRVAdj' ).AsVariant := edtPRVAdj.EditValue;
    CD078LDataSetSource.FieldByName( 'StopFlag' ).AsInteger := 0;
    if chkStopFlag.Checked then
      CD078LDataSetSource.FieldByName( 'StopFlag' ).AsInteger := 1;


    {
    cdsCD078G1.FieldByName( 'DESCRIPTION' ).AsString := edtDescription.Text;
    cdsCD078G1.FieldByName( 'MASTERSALE' ).AsInteger := 0;
    edtDiscountAmt.PostEditValue;
    cdsCD078G1.FieldByName( 'DISCOUNTAMT' ).AsString := VarToStrDef(edtDiscountAmt.EditValue,'0');
    cdsCD078G1.FieldByName( 'SALEPOINTCODE' ).AsString := MSalePoint.Value;
    cdsCD078G1.FieldByName( 'SALEPOINTNAME' ).AsString := MSalePoint.ValueName;
    }
  except
    on E: Exception do
    begin
      ErrorMsg( Format( '存檔有誤, 原因:%s', [E.Message] ) );
      Exit;
    end;
  end;

  Result := True;
end;

procedure TfrmCD078L.EditorStateChange;
begin
  ServiceType.Enabled := False;
  edtMonth.Properties.ReadOnly := True;
  edtCashAdj.Properties.ReadOnly := True;
  edtSubTotal.Properties.ReadOnly := True;
  chkStopFlag.Properties.ReadOnly := True;

  edtCMAmt.Properties.ReadOnly := True;
  edtCMTax.Properties.ReadOnly := True;
  edtCMAdj.Properties.ReadOnly := True;

  edtDTVAmt.Properties.ReadOnly := True;
  edtDTVTax.Properties.ReadOnly := True;
  edtDTVAdj.Properties.ReadOnly := True;

  edtPRVAmt.Properties.ReadOnly := True;
  edtPRVTax.Properties.ReadOnly := True;
  edtPRVAdj.Properties.ReadOnly := True;
  CD078LGrid.Enabled := False;
  case FMode of
    dmInsert:
      begin
        ServiceType.Enabled := True;
        edtMonth.Properties.ReadOnly := False;
        edtCashAdj.Properties.ReadOnly := False;
        edtSubTotal.Properties.ReadOnly := False;
        chkStopFlag.Properties.ReadOnly := False;

        edtCMAmt.Properties.ReadOnly := False;
        edtCMTax.Properties.ReadOnly := False;
        edtCMAdj.Properties.ReadOnly := False;

        edtDTVAmt.Properties.ReadOnly := False;
        edtDTVTax.Properties.ReadOnly := False;
        edtDTVAdj.Properties.ReadOnly := False;

        edtPRVAmt.Properties.ReadOnly := False;
        edtPRVTax.Properties.ReadOnly := False;
        edtPRVAdj.Properties.ReadOnly := False;
        ServiceType.CodeNo.SetFocus;
      end;
    dmUpdate:
      begin
        edtCashAdj.Properties.ReadOnly := False;
        edtSubTotal.Properties.ReadOnly := False;
        chkStopFlag.Properties.ReadOnly := False;

        edtCMAmt.Properties.ReadOnly := False;
        edtCMTax.Properties.ReadOnly := False;
        edtCMAdj.Properties.ReadOnly := False;

        edtDTVAmt.Properties.ReadOnly := False;
        edtDTVTax.Properties.ReadOnly := False;
        edtDTVAdj.Properties.ReadOnly := False;

        edtPRVAmt.Properties.ReadOnly := False;
        edtPRVTax.Properties.ReadOnly := False;
        edtPRVAdj.Properties.ReadOnly := False;
        edtCashAdj.SetFocus;
      end;
    dmBrowser:
      begin
        CD078LGrid.Enabled := True;
        if CD078LGrid.CanFocus then
          CD078LGrid.SetFocus;
      end;
  end;
end;

procedure TfrmCD078L.ButtonStateChange;
begin
  btnUpdate.Enabled := ( FMode in [dmBrowser] ) and  ( Not CD078LDataSetSource.IsEmpty);
  btnInsert.Enabled := ( FMode in [dmBrowser] );
  btnSave.Enabled := not (FMode in [dmBrowser]);
  btnDelete.Enabled := btnUpdate.Enabled;
  actSave.Enabled := False;
  actDelete.Enabled := False;
  actInsert.Enabled := False;
  actUpdate.Enabled := False;
  if ( FMode = dmInsert ) then
  begin
    actSave.Enabled := True;
  end else
  if ( FMode = dmUpdate ) then
  begin
    actSave.Enabled := True;
  end else
  begin
    actSave.Enabled := False;
    actDelete.Enabled := True;
    actInsert.Enabled := True;
    actUpdate.Enabled := True;
    actCancel.Enabled := True;
    actCancel.Caption := '結束(&C)';
    actCancel.ShortCut := TextToShortCut( 'Alt+C' );
  end;
end;

procedure TfrmCD078L.FormShow(Sender: TObject);
begin
  Screen.Cursor := crSQLWait;
  try
    ServiceType.Initializa;
    if ( Not CD078LDataSetSource.IsEmpty ) then
      CD078LDataSetSource.First;
    dsCD078L.DataSet := CD078LDataSetSource;
    DMLModeChange( dmBrowser );
  finally
    Screen.Cursor := crDefault;
  end;
end;



procedure TfrmCD078L.btnUpdateClick(Sender: TObject);
begin
  actUpdate.Execute;
end;

procedure TfrmCD078L.actUpdateExecute(Sender: TObject);
begin
  if  not btnUpdate.Enabled then
  begin
    Exit;
  end;
  FMode := dmUpdate;
  DMLModeChange(dmUpdate);
end;

procedure TfrmCD078L.actInsertExecute(Sender: TObject);
begin
  if ( not btnInsert.Enabled ) then
  begin
    Exit;
  end;
  FMode := dmInsert;
  DMLModeChange( dmInsert );

end;

procedure TfrmCD078L.actSaveExecute(Sender: TObject);
begin
  if ( not btnSave.Enabled ) then
  begin
    Exit;
  end;
  edtMonth.PostEditValue;
  edtCashAdj.PostEditValue;
  edtSubTotal.PostEditValue;
  edtCMAmt.PostEditValue;
  edtCMTax.PostEditValue;
  edtCMAdj.PostEditValue;
  edtDTVAmt.PostEditValue;
  edtDTVTax.PostEditValue;
  edtDTVAdj.PostEditValue;
  edtPRVAmt.PostEditValue;
  edtPRVAdj.PostEditValue;
  edtPRVTax.PostEditValue;

  DMLModeChange( dmSave );
end;

procedure TfrmCD078L.actCancelExecute(Sender: TObject);

begin
  if ( FMode in [dmBrowser] ) then
  begin
    CD078LDataSetSource.AfterScroll := nil;
    Self.Close;
    Exit;
  end;

  DMLModeChange( dmCancel );

end;

procedure TfrmCD078L.btnCancelClick(Sender: TObject);
begin
  actCancel.Execute;
end;

procedure TfrmCD078L.btnInsertClick(Sender: TObject);
begin
 actInsert.Execute;
end;

procedure TfrmCD078L.btnSaveClick(Sender: TObject);
begin
  actSave.Execute;
end;

procedure TfrmCD078L.actDeleteExecute(Sender: TObject);
begin
  if ( btnDelete.Visible ) and ( btnDelete.Enabled ) then
  begin
    if MessageDlg('確認刪除此筆資料？',mtConfirmation,[mbYes,mbNo],0) = mrYes then
    begin
      DMLModeChange( dmDelete );
    end;
  end;

end;

procedure TfrmCD078L.btnDeleteClick(Sender: TObject);
begin
  actDelete.Execute;

end;

end.
