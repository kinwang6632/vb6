{#6881 需求 By Kin 2014/10/16}
unit frmCD078K0U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, cxStyles, dxSkinsCore, dxSkinBlack, dxSkinBlue,
  dxSkinCaramel, dxSkinCoffee, dxSkinsDefaultPainters, dxSkinscxPCPainter,
  cxCustomData, cxGraphics, cxFilter, cxData, cxDataStorage, cxEdit, DB,
  cxDBData, cxTextEdit, cxMaskEdit, cxImageComboBox, cxCurrencyEdit,
  cxButtonEdit, cxGridLevel, cxGridCustomTableView, cxGridTableView,
  cxGridBandedTableView, cxGridDBBandedTableView, cxClasses, cxControls,
  cxGridCustomView, cxGrid, cxCheckBox, DBClient, cxMemo,cxDataUtils,
  Math, ActnList, StdCtrls;

type
  TDML = ( dmBrowser = 0,
           dmInsert = 1,
           dmUpdate = 2,
           dmDelete = 3,
           dmCancel = 4,
           dmSave = 5 );
  TfrmCD078K = class(TForm)
    pnl1: TPanel;
    pnl2: TPanel;
    RankGrid: TcxGrid;
    gvRank: TcxGridDBBandedTableView;
    gvRankColumn1: TcxGridDBBandedColumn;
    gvRankColumn2: TcxGridDBBandedColumn;
    gvRankColumn3: TcxGridDBBandedColumn;
    gvRankColumn5: TcxGridDBBandedColumn;
    gvRankColumn4: TcxGridDBBandedColumn;
    glRank: TcxGridLevel;
    cdsRankDataSet: TClientDataSet;
    dsRankDataSource: TDataSource;
    actlstact: TActionList;
    actSave: TAction;
    actCancel: TAction;
    btnSave: TButton;
    btnCancel: TButton;
    procedure FormShow(Sender: TObject);
    procedure gvRankCustomDrawCell(Sender: TcxCustomGridTableView;
      ACanvas: TcxCanvas; AViewInfo: TcxGridTableDataCellViewInfo;
      var ADone: Boolean);
    procedure gvRankColumn2PropertiesEditValueChanged(Sender: TObject);
    procedure actSaveExecute(Sender: TObject);
    procedure gvRankColumn4PropertiesValidate(Sender: TObject;
      var DisplayValue: Variant; var ErrorText: TCaption;
      var Error: Boolean);
    procedure gvRankTcxGridDBDataControllerTcxDataSummaryFooterSummaryItems0GetText(
      Sender: TcxDataSummaryItem; const AValue: Variant;
      AIsFooter: Boolean; var AText: String);
    procedure actCancelExecute(Sender: TObject);
    procedure gvRankCustomDrawFooterCell(Sender: TcxGridTableView;
      ACanvas: TcxCanvas; AViewInfo: TcxGridColumnHeaderViewInfo;
      var ADone: Boolean);
    procedure gvRankTcxGridDBDataControllerTcxDataSummaryFooterSummaryItems3GetText(
      Sender: TcxDataSummaryItem; const AValue: Variant;
      AIsFooter: Boolean; var AText: String);
    procedure gvRankColumn3PropertiesValidate(Sender: TObject;
      var DisplayValue: Variant; var ErrorText: TCaption;
      var Error: Boolean);
  private
    { Private declarations }
    FMode: TDML;
    FCD078ADataSet : TClientDataSet;
    FCD078KDataSet : TClientDataSet;
    procedure PrepareDataSet;
    procedure FillDataSet;
    procedure DMLModeChange(aMode: TDML);
    function SaveData : Boolean;
    function DataSetStateChange(const aMode: TDML): Boolean;
    function DataValidate: Boolean;


  public
    { Public declarations }
    constructor Create( aCD078ADataSet: TClientDataSet;aCD078KDataSet:TClientDataSet ); reintroduce;
  end;

var
  frmCD078K: TfrmCD078K;

implementation
  uses cbUtilis;
{$R *.dfm}

{ TfrmCD078K }

constructor TfrmCD078K.Create(aCD078ADataSet: TClientDataSet;aCD078KDataSet:TClientDataSet);
begin
  inherited Create( Application );
  FCD078ADataSet := aCD078ADataSet;
  FCD078KDataSet := aCD078KDataSet;
  PrepareDataSet;
end;

procedure TfrmCD078K.PrepareDataSet;
begin
  { Detail DataSet }
  if ( cdsRankDataSet.FieldDefs.Count <= 0 ) then
  begin
    cdsRankDataSet.FieldDefs.Add( 'BPCode', ftString, 10 );
    cdsRankDataSet.FieldDefs.Add( 'CitemCode', ftString, 5 );
    cdsRankDataSet.FieldDefs.Add( 'CitemName', ftString, 50 );
    cdsRankDataSet.FieldDefs.Add( 'MasterItem', ftInteger );
    cdsRankDataSet.FieldDefs.Add( 'Amount', ftInteger );
    cdsRankDataSet.FieldDefs.Add( 'IFRSPercent',ftFloat );
    cdsRankDataSet.FieldDefs.Add( 'IFRSAmt', ftInteger );
    cdsRankDataSet.CreateDataSet;
  end;
  cdsRankDataSet.EmptyDataSet;
end;

procedure TfrmCD078K.FormShow(Sender: TObject);
begin
  Screen.Cursor := crSQLWait;
  try
    FillDataSet;
  finally
    Screen.Cursor := crDefault;
  end;



end;

procedure TfrmCD078K.FillDataSet;
begin
  FCD078ADataSet.First;
  while not FCD078ADataSet.Eof do
  begin
    if ( FCD078ADataSet.FieldByName( 'MCitemCode' ).AsString = EmptyStr ) then
    begin
      cdsRankDataSet.Append;
      cdsRankDataSet.FieldByName( 'BPCode' ).AsString := FCD078ADataSet.FieldByName( 'BPCode' ).AsString;
      cdsRankDataSet.FieldByName( 'CitemCode' ).AsString := FCD078ADataSet.FieldByName( 'CitemCode' ).AsString;
      cdsRankDataSet.FieldByName( 'CitemName' ).AsString := FCD078ADataSet.FieldByName( 'CitemName' ).AsString;
      cdsRankDataSet.Post;
      FCD078KDataSet.Filtered := False;
      try
        FCD078KDataSet.Filter := Format( 'CITEMCODE = ''%s'' ',
          [VarToStrDef( FCD078ADataSet.FieldByName( 'CITEMCODE' ).AsString,'-1')]);
        FCD078KDataSet.Filtered := True;
        if ( not FCD078KDataSet.Eof ) then
        begin
          cdsRankDataSet.Edit;
          cdsRankDataSet.FieldByName( 'MasterItem' ).AsString :=
              FCD078KDataSet.FieldByName( 'MasterItem' ).AsString;

          cdsRankDataSet.FieldByName( 'Amount' ).AsString :=
              FCD078KDataSet.FieldByName( 'Amount' ).AsString;

          cdsRankDataSet.FieldByName( 'IFRSPercent' ).AsString :=
              FCD078KDataSet.FieldByName( 'IFRSPercent' ).AsString;

          cdsRankDataSet.FieldByName( 'IFRSAmt' ).AsString :=
              FCD078KDataSet.FieldByName( 'IFRSAmt' ).AsString;
          cdsRankDataSet.Post;
        end;
      finally
        FCD078KDataSet.Filter := EmptyStr;
        FCD078KDataSet.Filtered := False;

      end;
    end;
    FCD078ADataSet.Next;
  end;
  FCD078ADataSet.First;
  gvRank.DataController.GotoFirst;
end;

procedure TfrmCD078K.gvRankCustomDrawCell(Sender: TcxCustomGridTableView;
  ACanvas: TcxCanvas; AViewInfo: TcxGridTableDataCellViewInfo;
  var ADone: Boolean);
  var
  aItem: TcxGridDBBandedColumn;
  aGridRecord: TcxGridDataRow;
  aDrawDisable: Boolean;
begin
  aItem := TcxGridDBBandedColumn( AViewInfo.Item );
  aGridRecord := TcxGridDataRow( AViewInfo.GridRecord );
//  aIndex := VarAsType( aGridRecord.Values[gvRankColumn10.Index], varInteger );
//  aDiscountRate := StrToInt( VarToStrDef( aGridRecord.Values[gvRankColumn4.Index], '0' ) );
  {}
  aDrawDisable :=
    ( aItem = gvRankColumn1 ) or ( aItem = gvRankColumn5 ) ;
  {}
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

procedure TfrmCD078K.gvRankColumn2PropertiesEditValueChanged(
  Sender: TObject);
  var
    aCInnerCheck: TcxCheckBox;
    aCInnerEdit: TcxCurrencyEdit;
    aIndex : Integer;
    aMasterAmt,aNotMasterMat : Extended;
    aIFRSAmt : Extended;
    aBookMark : Pointer;
begin
  aMasterAmt := 0;
  aNotMasterMat := 0;
  Screen.Cursor := crSQLWait;
  Self.Enabled := False;
  gvRank.BeginUpdate;
  try
    if ( gvRank.Controller.EditingItem = gvRankColumn2 ) then
    begin
      aCInnerCheck := TcxCheckBox( Sender );

      aCInnerCheck.LockChangeEvents(True,False);
      try
        aCInnerCheck.PostEditValue;
      finally
         aCInnerCheck.LockChangeEvents(False,False);
      end;
    end;
    if ( gvRank.Controller.EditingItem = gvRankColumn3 ) or ( gvRank.Controller.EditingItem = gvRankColumn4 ) then
    begin
      aCInnerEdit := TcxCurrencyEdit ( Sender );
      aCInnerEdit.LockChangeEvents(True,False);
      try
        aCInnerEdit.ModifiedAfterEnter := True;
        aCInnerEdit.PostEditValue;
      finally
        aCInnerEdit.LockChangeEvents(False,False);
      end;
    end;
    for aIndex := 0 to gvRank.DataController.RecordCount - 1 do
    begin
      if VarToStrDef( gvRank.DataController.GetValue(aIndex,1),'0') = '1' then
      begin
        aMasterAmt := aMasterAmt + StrToInt( VarToStrDef( gvRank.DataController.GetValue( aIndex , 2), '0') );
      end else
      begin
        aNotMasterMat := aNotMasterMat + StrToInt( VarToStrDef( gvRank.DataController.GetValue( aIndex , 2), '0') );
      end;
    end;
    cdsRankDataSet.DisableControls;
    aBookMark := cdsRankDataSet.GetBookmark;
    cdsRankDataSet.First;

    try
      while not cdsRankDataSet.Eof do
      begin
        if ( cdsRankDataSet.FieldByName( 'MasterItem' ).AsString = '1' ) then
        begin
          cdsRankDataSet.Edit;
          cdsRankDataSet.FieldByName( 'IFRSAmt' ).AsVariant :=
             CbRoundTo(  aMasterAmt * ( StrToFloat( VarToStrDef( cdsRankDataSet.FieldByName( 'IFRSPercent' ).AsVariant,'0') ) / 100 ),0);
             //CbRoundTo( aMasterAmt * StrToFloat( ( VarToStrDef( cdsRankDataSet.FieldByName( 'IFRSPercent' ).AsString,'0') / 100 ) ) ,0);
          cdsRankDataSet.Post;
        end else
        begin
          cdsRankDataSet.Edit;
          cdsRankDataSet.FieldByName( 'IFRSAmt' ).AsVariant :=
             CbRoundTo(  aMasterAmt * ( StrToFloat( VarToStrDef( cdsRankDataSet.FieldByName( 'IFRSPercent' ).AsVariant,'0') ) / 100 ) +
                  cdsRankDataSet.FieldByName( 'Amount' ).AsInteger,0);
          cdsRankDataSet.Post;
        end;

        cdsRankDataSet.Next;
      end;
    finally
      cdsRankDataSet.GotoBookmark( aBookMark );
      cdsRankDataSet.FreeBookmark( aBookMark );
      cdsRankDataSet.EnableControls;
    end;
  finally
    gvRank.DataController.PostEditingData;
    gvRank.DataController.Summary.CalculateFooterSummary;
    gvRank.EndUpdate;
    Screen.Cursor := crDefault;
    Self.Enabled := True;
  end;
end;

procedure TfrmCD078K.actSaveExecute(Sender: TObject);
begin
  gvRank.DataController.PostEditingData;
  DMLModeChange( dmSave );
end;

procedure TfrmCD078K.DMLModeChange(aMode: TDML);
begin
  if not DataSetStateChange( aMode ) then Exit;
  if ( aMode in [dmSave, dmCancel] ) then
  begin
    FMode := aMode;
    Self.Close;
    Exit;
  end;

end;

function TfrmCD078K.DataSetStateChange(const aMode: TDML): Boolean;
begin

  Result := True;
  case aMode of
    dmInsert:
      begin
      end;
    dmUpdate:
      begin
      end;
    dmCancel:
      begin
       
        cdsRankDataSet.Cancel;
      end;
    dmBrowser:
      begin

        FillDataSet;

      end;
    dmSave:
      begin
        Self.Enabled := False;
        Screen.Cursor := crSQLWait;
        try
          Result := DataValidate;
          if Result then
          begin

            try
              Result := SaveData;
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
          Self.Enabled := True;
        end;


      end;
  end;
end;

function TfrmCD078K.DataValidate: Boolean;
begin
  Screen.Cursor := crSQLWait;
  try
    Result := True;
    if VarToStrDef( gvRank.DataController.Summary.FooterSummaryValues[1] , '0') <> '100' then
    begin
      WarningMsg( '拆帳比例必須為100%！');
      Result := False;
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

function TfrmCD078K.SaveData: Boolean;
begin
  FCD078KDataSet.EmptyDataSet;
  cdsRankDataSet.First;
  cdsRankDataSet.DisableControls;
  try
    while not cdsRankDataSet.Eof do
    begin
      FCD078KDataSet.Append;
      FCD078KDataSet.FieldByName( 'BPCode' ).AsString:= cdsRankDataSet.FieldByName('BPCode').AsString;
      FCD078KDataSet.FieldByName( 'CitemCode' ).AsString:= cdsRankDataSet.FieldByName('CitemCode').AsString;
      FCD078KDataSet.FieldByName( 'MasterItem' ).AsString:= cdsRankDataSet.FieldByName('MasterItem').AsString;
      FCD078KDataSet.FieldByName( 'Amount' ).AsString:= cdsRankDataSet.FieldByName('Amount').AsString;
      FCD078KDataSet.FieldByName( 'IFRSPercent' ).AsString:= cdsRankDataSet.FieldByName('IFRSPercent').AsString;
      FCD078KDataSet.FieldByName( 'IFRSAmt' ).AsString:= cdsRankDataSet.FieldByName('IFRSAmt').AsString;
      FCD078KDataSet.Post;
      cdsRankDataSet.Next;
    end;
  finally
    cdsRankDataSet.EnableControls;
  end;
  Result := True;
end;

procedure TfrmCD078K.gvRankColumn4PropertiesValidate(Sender: TObject;
  var DisplayValue: Variant; var ErrorText: TCaption; var Error: Boolean);
begin
  if not Error then Exit;
  if VarToStrDef( DisplayValue, EmptyStr ) = EmptyStr then Exit;
  if ( TcxCurrencyEdit( Sender ).Properties.DecimalPlaces = 0 ) then
  begin

  end else
  begin
     ErrorText := Format( '輸入值有誤, 值必須介於%10.3f~%10.3f',
      [TcxCurrencyEdit( Sender ).Properties.MinValue,
      TcxCurrencyEdit( Sender ).Properties.MaxValue] );
  end;
  WarningMsg( ErrorText );
  Error := False;
  ErrorText := EmptyStr;
  TcxCurrencyEdit( Sender ).SelectAll;
  Abort;
end;

procedure TfrmCD078K.gvRankTcxGridDBDataControllerTcxDataSummaryFooterSummaryItems0GetText(
  Sender: TcxDataSummaryItem; const AValue: Variant; AIsFooter: Boolean;
  var AText: String);
begin
  if ( Sender.Index = 0 ) or ( Sender.Index = 2 ) then
    AText:= FormatCurr(',0',StrToCurr(VarToStrDef( AValue ,'0' )));
end;

procedure TfrmCD078K.actCancelExecute(Sender: TObject);
begin
  Self.Close;
end;

procedure TfrmCD078K.gvRankCustomDrawFooterCell(Sender: TcxGridTableView;
  ACanvas: TcxCanvas; AViewInfo: TcxGridColumnHeaderViewInfo;
  var ADone: Boolean);
begin
  ACanvas.Font.Color := clRed;
  AViewInfo.AlignmentVert := vaCenter;
  if AViewInfo.Column.Summary.Item.Index = 1 then
    AViewInfo.AlignmentHorz := taRightJustify
  else
    AViewInfo.AlignmentHorz := taCenter;


end;

procedure TfrmCD078K.gvRankTcxGridDBDataControllerTcxDataSummaryFooterSummaryItems3GetText(
  Sender: TcxDataSummaryItem; const AValue: Variant; AIsFooter: Boolean;
  var AText: String);
begin
  AText := '合計';
end;

procedure TfrmCD078K.gvRankColumn3PropertiesValidate(Sender: TObject;
  var DisplayValue: Variant; var ErrorText: TCaption; var Error: Boolean);
begin
  if not Error then Exit;

  if TcxCurrencyEdit ( Sender ).Value < 0 then
  begin
    ErrorText := '輸入值有誤, 值必須 >0';
  end;
  WarningMsg( ErrorText );
  Error := False;
  ErrorText := EmptyStr;
  TcxCurrencyEdit( Sender ).SelectAll;
  Abort;
end;

end.
