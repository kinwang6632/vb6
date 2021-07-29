unit frmCD078B3_5U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, DBClient,
  cbStyleController, frmCD078B1U, 
  dxSkinsCore, dxSkinsDefaultPainters,
  cxControls, cxContainer, cxEdit, cxTextEdit, ActnList, dxSkinscxPCPainter,
  cxGraphics, cxDataStorage, DB, cxDBData,
  cxMaskEdit, cxImageComboBox, cxCurrencyEdit, cxButtonEdit, cxGridLevel,
  cxGridCustomTableView, cxGridTableView, cxGridBandedTableView,
  cxGridDBBandedTableView, cxClasses, cxGridCustomView, cxGrid, cxDataUtils,
  cxStyles, cxFilter, cxData, cxCustomData, dxSkinBlack, dxSkinBlue,
  dxSkinCaramel, dxSkinCoffee;

const
  MAXITEM = 48;  

type
  TfrmCD078B3_5 = class(TForm)
    ActionList1: TActionList;
    actSave: TAction;
    actCancel: TAction;
    Panel1: TPanel;
    edtDML: TcxTextEdit;
    btnSave: TButton;
    Button1: TButton;
    Panel4: TPanel;
    Panel2: TPanel;
    Bevel1: TBevel;
    Label30: TLabel;
    PenalRuleGrid: TcxGrid;
    glPenalRule: TcxGridLevel;
    gvPenalRule: TcxGridDBBandedTableView;
    gvPenalRuleCol1: TcxGridDBBandedColumn;
    gvPenalRuleCol2: TcxGridDBBandedColumn;
    gvPenalRuleCol3: TcxGridDBBandedColumn;
    gvPenalRuleCol4: TcxGridDBBandedColumn;
    gvPenalRuleCol5: TcxGridDBBandedColumn;
    PenalRuleDataSet: TClientDataSet;
    PenalRuleDataSource: TDataSource;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure gvPenalRuleCol1PropertiesChange(Sender: TObject);
    procedure gvPenalRuleCol1PropertiesEditValueChanged(Sender: TObject);
    procedure gvPenalRuleEditing(Sender: TcxCustomGridTableView;
      AItem: TcxCustomGridTableItem; var AAllow: Boolean);
    procedure gvPenalRuleCustomDrawCell(Sender: TcxCustomGridTableView;
      ACanvas: TcxCanvas; AViewInfo: TcxGridTableDataCellViewInfo;
      var ADone: Boolean);
    procedure gvPenalRuleCol3PropertiesEditValueChanged(Sender: TObject);
    procedure actSaveExecute(Sender: TObject);
    procedure actCancelExecute(Sender: TObject);
  private
    { Private declarations }
    FMode: TDML;
    FSourceDataSet: TClientDataSet;
    FKeyArray: array [1..4] of String;
    FMonMap: array [1..MAXITEM] of Boolean;
    procedure PreparePenalRuleDataSet;
    procedure UnPreparePenalRuleDataSet;
    procedure DeCompoundPenalRuleData;
    procedure PenalRuleDataSetBindToArray;
    procedure MakeColumnEditMode;
    procedure MoveToLastRecord;
    function VdMonthBreak: Boolean;
    function VdMonthOverlap: Boolean;
    function VdAmt: Boolean;
    procedure SaveData;
  public
    { Public declarations }
    constructor Create(aMode: TDML; aDataSet: TClientDataSet;  aKeys: array of String); reintroduce;
  end;

var
  frmCD078B3_5: TfrmCD078B3_5;

implementation

{$R *.dfm}

uses cbUtilis;

{ ---------------------------------------------------------------------------- }

constructor TfrmCD078B3_5.Create(aMode: TDML; aDataSet: TClientDataSet;
  aKeys: array of String);
begin
  inherited Create( Application );
  FMode := aMode;
  FSourceDataSet := aDataSet;
  FKeyArray[1] := aKeys[0];  //BPCode
  FKeyArray[2] := aKeys[1];  //CItemCode
  FKeyArray[3] := aKeys[2];  //ServerTYpe
  FKeyArray[4] := aKeys[3];  //PenalCode
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_5.FormCreate(Sender: TObject);
begin
  PreparePenalRuleDataSet;
  DeCompoundPenalRuleData;
  PenalRuleDataSetBindToArray;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_5.FormShow(Sender: TObject);
begin
  MoveToLastRecord;
  MakeColumnEditMode;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_5.FormDestroy(Sender: TObject);
begin
  FSourceDataSet := nil;
  UnPreparePenalRuleDataSet;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_5.PreparePenalRuleDataSet;
var
  aIndex: Integer;
begin
  if ( PenalRuleDataSet.FieldDefs.Count <= 0 ) then
  begin
    PenalRuleDataSet.FieldDefs.Add( 'Item', ftInteger );
    PenalRuleDataSet.FieldDefs.Add( 'MonthStart', ftInteger );
    PenalRuleDataSet.FieldDefs.Add( 'MonthStop', ftInteger );
    PenalRuleDataSet.FieldDefs.Add( 'MonthAmt', ftInteger );
    PenalRuleDataSet.FieldDefs.Add( 'DecreaseAmt', ftInteger );
    //PenalRuleDataSet.FieldDefs.Add( 'IncreaseMonth', ftInteger );
    PenalRuleDataSet.CreateDataSet;
    PenalRuleDataSet.AddIndex( 'IDX1', 'Item', [ixUnique] );
    PenalRuleDataSet.IndexName := 'IDX1';
  end;
  PenalRuleDataSet.EmptyDataSet;
  PenalRuleDataSet.DisableControls;
  try
    for aIndex := 1 to MAXITEM do
    begin
      PenalRuleDataSet.Append;
      PenalRuleDataSet.FieldByName( 'Item' ).AsInteger := aIndex;
      PenalRuleDataSet.FieldByName( 'MonthAmt' ).AsInteger := 0;
      PenalRuleDataSet.FieldByName( 'DecreaseAmt' ).AsInteger := 0;
      //PenalRuleDataSet.FieldByName( 'IncreaseMonth' ).AsInteger := 0;
      PenalRuleDataSet.Post;
    end;
  finally
    PenalRuleDataSet.EnableControls;
  end;
  PenalRuleDataSet.First;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_5.UnPreparePenalRuleDataSet;
begin
  PenalRuleDataSet.EmptyDataSet;
  if ( PenalRuleDataSet.IndexName <> EmptyStr ) then
    PenalRuleDataSet.DeleteIndex( PenalRuleDataSet.IndexName );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_5.DeCompoundPenalRuleData;
begin
  PenalRuleDataSet.First;
  FSourceDataSet.First;
  while not FSourceDataSet.Eof do
  begin
    PenalRuleDataSet.Edit;
    PenalRuleDataSet.FieldByName( 'MonthStart' ).AsString := FSourceDataSet.FieldByName( 'MonthStart' ).AsString;
    PenalRuleDataSet.FieldByName( 'MonthStop' ).AsString := FSourceDataSet.FieldByName( 'MonthStop' ).AsString;
    PenalRuleDataSet.FieldByName( 'MonthAmt' ).AsString := FSourceDataSet.FieldByName( 'MonthAmt' ).AsString;
    PenalRuleDataSet.FieldByName( 'DecreaseAmt' ).AsString := FSourceDataSet.FieldByName( 'DecreaseAmt' ).AsString;
    //PenalRuleDataSet.FieldByName( 'IncreaseMonth' ).AsString := FSourceDataSet.FieldByName( 'IncreaseMonth' ).AsString;
    PenalRuleDataSet.Post;
    FSourceDataSet.Next;
    PenalRuleDataSet.Next;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_5.PenalRuleDataSetBindToArray;
begin
  FillChar( FMonMap, SizeOf( FMonMap ), Ord( False ) );
  PenalRuleDataSet.DisableControls;
  try
    PenalRuleDataSet.First;
    while not PenalRuleDataSet.Eof do
    begin
      FMonMap[PenalRuleDataSet.FieldByName( 'Item' ).AsInteger] :=
        ( PenalRuleDataSet.FieldByName( 'MonthStart' ).AsString <> EmptyStr );
      PenalRuleDataSet.Next;
    end;
    PenalRuleDataSet.First;
  finally
    PenalRuleDataSet.EnableControls;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_5.MakeColumnEditMode;
begin
  if ( PenalRuleGrid.CanFocusEx ) then PenalRuleGrid.SetFocus;
  gvPenalRule.Controller.EditingItem := gvPenalRuleCol1;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_5.MoveToLastRecord;
var
  aIndex: Integer;
begin
  for aIndex := Low( FMonMap ) to High( FMonMap ) - 1 do
  begin
    if ( not FMonMap[aIndex] ) then
    begin
      PenalRuleDataSet.First;
      PenalRuleDataSet.MoveBy( aIndex - 1 );
      Break;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B3_5.VdMonthBreak: Boolean;
var
  aLastIdx: Integer;
begin
  Result := True;
  PenalRuleDataSet.DisableControls;
  try
    aLastIdx := 0;
    PenalRuleDataSet.Last;
    while not PenalRuleDataSet.Bof do
    begin
      if ( aLastIdx = 0 ) and
         ( PenalRuleDataSet.FieldByName( 'MonthStart' ).AsString <> EmptyStr ) then
      begin
        aLastIdx := PenalRuleDataSet.FieldByName( 'Item' ).AsInteger;
      end else
      if ( aLastIdx > 0 ) and
         ( PenalRuleDataSet.FieldByName( 'MonthStart' ).AsString = EmptyStr ) then
      begin
        Result := False;
        Break;
      end;
      PenalRuleDataSet.Prior;
    end;
  finally
    PenalRuleDataSet.EnableControls;
  end;
  {}
  if ( not Result ) then
  begin
    WarningMsg( '合約月數(起) 必須連續且不可中斷。' );
    if PenalRuleGrid.CanFocusEx then PenalRuleGrid.SetFocus;
    gvPenalRule.Controller.EditingItem := gvPenalRuleCol1;
    Exit;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B3_5.VdMonthOverlap: Boolean;
var
  aPriorMonth: Integer;
  aErrMsg: String;
  aItem: TcxGridDBBandedColumn;
begin
  Result := True;
  {}
  aPriorMonth := -1;
  PenalRuleDataSet.DisableControls;
  try
    PenalRuleDataSet.First;
    while not PenalRuleDataSet.Eof do
    begin
      if ( PenalRuleDataSet.FieldByName( 'MonthStart' ).AsString <> EmptyStr ) then
      begin
        if ( PenalRuleDataSet.FieldByName( 'MonthStop' ).AsInteger <
             PenalRuleDataSet.FieldByName( 'MonthStart' ).AsInteger ) then
        begin
          Result := False;
          aErrMsg := '合約月數(迄) 不可小於合約月數(起)';
          aItem := gvPenalRuleCol2;
          Break;
        end;
        if ( aPriorMonth <= 0 ) then
          aPriorMonth := PenalRuleDataSet.FieldByName( 'MonthStop' ).AsInteger
        else begin
          if ( ( aPriorMonth + 1 ) <> PenalRuleDataSet.FieldByName( 'MonthStart' ).AsInteger ) then
          begin
            Result := False;
            aErrMsg := '合約月數(起) 必須連續, 且不可與上筆合約月數(迄)重疊';
            aItem := gvPenalRuleCol1;
            Break;
          end;
          aPriorMonth := PenalRuleDataSet.FieldByName( 'MonthStop' ).AsInteger;
        end;
      end;
      PenalRuleDataSet.Next;
    end;
  finally
    PenalRuleDataSet.EnableControls;
  end;
  if ( not Result ) then
  begin
    WarningMsg( aErrMsg );
    if PenalRuleGrid.CanFocusEx then PenalRuleGrid.SetFocus;
    gvPenalRule.Controller.EditingItem := aItem;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B3_5.VdAmt: Boolean;
var
  aItem: TcxGridDBBandedColumn;
  aErrMsg: String;
begin
  Result := True;
  PenalRuleDataSet.DisableControls;
  try
    aItem := nil;
    aErrMsg := EmptyStr;
    PenalRuleDataSet.First;
    while not PenalRuleDataSet.Eof do
    begin
      if ( PenalRuleDataSet.FieldByName( 'MonthStart' ).AsString <> EmptyStr ) then
      begin
        {}
        if ( PenalRuleDataSet.FieldByName( 'MonthStop' ).AsString = EmptyStr ) then
        begin
          aItem := gvPenalRuleCol2;
          aErrMsg := '請輸入【合約月數(迄)】。';
          Result := False;
          Break;
        end;
        {}
        if ( PenalRuleDataSet.FieldByName( 'MonthAmt' ).AsInteger < 0 ) then
        begin
          aItem := gvPenalRuleCol3;
          aErrMsg := '請輸入【總綁約金額】, 且不可為負值。';
          Result := False;
          Break;
        end;
        {}
      end;
      PenalRuleDataSet.Next;
    end;
  finally
    PenalRuleDataSet.EnableControls;
  end;
  {}
  if ( not Result ) then
  begin
    WarningMsg( aErrMsg );
    if PenalRuleGrid.CanFocusEx then PenalRuleGrid.SetFocus;
    gvPenalRule.Controller.EditingItem := aItem;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_5.gvPenalRuleCol1PropertiesChange(Sender: TObject);
var
  aInnerEdit: TcxMaskEdit;
  aIndex: Integer;
begin
  aInnerEdit := TcxMaskEdit( gvPenalRule.Controller.EditingController.Edit );
  aIndex := PenalRuleDataSet.FieldByName( 'Item' ).AsInteger;
  FMonMap[aIndex] := ( aInnerEdit.EditingText <> EmptyStr );
  if ( not FMonMap[aIndex] ) then aInnerEdit.PostEditValue;
  gvPenalRule.Invalidate( True );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_5.gvPenalRuleCol1PropertiesEditValueChanged(
  Sender: TObject);
var
  aInnerEdit: TcxMaskEdit;
  aIndex: Integer;
begin
  aInnerEdit := TcxMaskEdit( Sender );
  aInnerEdit.PostEditValue;
  aIndex := PenalRuleDataSet.FieldByName( 'Item' ).AsInteger;
  FMonMap[aIndex] := ( aInnerEdit.EditingText <> EmptyStr );
  if ( not FMonMap[aIndex] ) then
  begin
    gvPenalRule.DataController.SetEditValue( gvPenalRuleCol2.Index, Null, evsValue );
    gvPenalRule.DataController.SetEditValue( gvPenalRuleCol3.Index, 0, evsValue );
    gvPenalRule.DataController.SetEditValue( gvPenalRuleCol4.Index, 0, evsValue );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_5.gvPenalRuleCol3PropertiesEditValueChanged(
  Sender: TObject);
var
  aInnerEdit: TcxCurrencyEdit;
  aValue: Double;
begin
  aInnerEdit := TcxCurrencyEdit( Sender );
  aInnerEdit.LockChangeEvents( True, False );
  try
    aValue := 0;
    if VarToStrDef( aInnerEdit.EditValue, EmptyStr ) <> EmptyStr then
      aValue := Abs( VarAsType( aInnerEdit.EditValue, varDouble ) );
    aInnerEdit.EditValue := aValue;
    aInnerEdit.PostEditValue;
  finally
    aInnerEdit.LockChangeEvents( False, False );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_5.gvPenalRuleCustomDrawCell(Sender: TcxCustomGridTableView;
  ACanvas: TcxCanvas; AViewInfo: TcxGridTableDataCellViewInfo;
  var ADone: Boolean);
var
  aItem: TcxGridDBBandedColumn;
  aGridRecord: TcxGridDataRow;
  aDrawDisable: Boolean;
  aIndex: Integer;
begin
  aItem := TcxGridDBBandedColumn( AViewInfo.Item );
  aGridRecord := TcxGridDataRow( AViewInfo.GridRecord );
  aIndex := VarAsType( aGridRecord.Values[gvPenalRuleCol5.Index], varInteger );
  aDrawDisable := ( ( aItem <> gvPenalRuleCol1 ) and ( FMonMap[aIndex] = False ) );
  if ( aDrawDisable ) then
  begin
    ACanvas.Brush.Color := clBtnFace;
    ACanvas.Font.Color := clDefault;
  end else
  begin
    ACanvas.Brush.Color := clWindow;
    ACanvas.Font.Color := clDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_5.gvPenalRuleEditing(Sender: TcxCustomGridTableView;
  AItem: TcxCustomGridTableItem; var AAllow: Boolean);
begin
  if ( AItem <> gvPenalRuleCol1  ) then
  begin
    AAllow := FMonMap[PenalRuleDataSet.FieldByName( 'Item' ).AsInteger];
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_5.actCancelExecute(Sender: TObject);
begin
  Self.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_5.actSaveExecute(Sender: TObject);
begin
  if not VdMonthBreak then Exit;
  if not VdMonthOverlap then Exit;
  if not VdAmt then Exit;
  SaveData;
  Self.ModalResult := mrOk;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_5.SaveData;
begin
  FSourceDataSet.First;
  while not FSourceDataSet.Eof do
    FSourceDataSet.Delete;
  {}
  FSourceDataSet.First;
  PenalRuleDataSet.First;
  while not PenalRuleDataSet.Eof do
  begin
    if ( PenalRuleDataSet.FieldByName( 'MONTHSTART' ).AsString <> EmptyStr ) then
    begin
      FSourceDataSet.Append;
      FSourceDataSet.FieldByName( 'BPCODE' ).AsString := FKeyArray[1];
      FSourceDataSet.FieldByName( 'CITEMCODE' ).AsString := FKeyArray[2];
      FSourceDataSet.FieldByName( 'SERVICETYPE' ).AsString := FKeyArray[3];
      FSourceDataSet.FieldByName( 'PENALCODE' ).AsString := FKeyArray[4];
      FSourceDataSet.FieldByName( 'ITEM' ).AsInteger := PenalRuleDataSet.FieldByName( 'ITEM' ).AsInteger;
      FSourceDataSet.FieldByName( 'MONTHSTART' ).AsInteger := PenalRuleDataSet.FieldByName( 'MONTHSTART' ).AsInteger;
      FSourceDataSet.FieldByName( 'MONTHSTOP' ).AsInteger := PenalRuleDataSet.FieldByName( 'MONTHSTOP' ).AsInteger;
      FSourceDataSet.FieldByName( 'MONTHAMT' ).AsInteger := PenalRuleDataSet.FieldByName( 'MONTHAMT' ).AsInteger;
      FSourceDataSet.FieldByName( 'DECREASEAMT' ).AsInteger := PenalRuleDataSet.FieldByName( 'DECREASEAMT' ).AsInteger;
      //FSourceDataSet.FieldByName( 'INCREASEMONTH' ).AsInteger := PenalRuleDataSet.FieldByName( 'INCREASEMONTH' ).AsInteger;
      FSourceDataSet.Post;
    end;
    PenalRuleDataSet.Next;
  end;
end;

{ ---------------------------------------------------------------------------- }

end.
