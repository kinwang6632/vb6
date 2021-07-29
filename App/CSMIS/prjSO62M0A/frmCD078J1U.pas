unit frmCD078J1U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, cxStyles, dxSkinsCore, dxSkinBlack, dxSkinBlue,
  dxSkinCaramel, dxSkinCoffee, dxSkinsDefaultPainters, dxSkinscxPCPainter,
  cxCustomData, cxGraphics, cxFilter, cxData, cxDataStorage, cxEdit, DB,
  cxDBData, cxContainer, cxCheckBox, cxGridLevel, cxGridCustomTableView,
  cxGridTableView, cxGridDBTableView, cxClasses, cxControls,
  cxGridCustomView, cxGrid, cbDataLookup, StdCtrls,ADODB,DBClient,
  frmCD078B1U, ActnList,Menus, Buttons, cxLookAndFeelPainters, cxTextEdit,
  cxButtons,frmCD078B0U ;

type
  TfrmCD078JU = class(TForm)
    Panel1: TPanel;
    ServiceType: TDataLookup;
    Label2: TLabel;
    ProductItem: TDataLookup;
    Label1: TLabel;
    Label3: TLabel;
    btnInstCodeStr: TButton;
    cxButton1: TcxButton;
    edtInstCodeStr: TcxTextEdit;
    Panel2: TPanel;
    Bevel1: TBevel;
    edtDML: TcxTextEdit;
    btnSave: TButton;
    btnCancel: TButton;
    edtFaciItem: TcxTextEdit;
    ActionList1: TActionList;
    actSave: TAction;
    actCancel: TAction;
    CloneDataSet: TClientDataSet;

    procedure ServiceTypeCodeNamePropertiesInitPopup(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure ServiceTypeCodeNoPropertiesChange(Sender: TObject);
    procedure ServiceTypeCodeNamePropertiesChange(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure ProductItemCodeNoPropertiesChange(Sender: TObject);
    procedure ProductItemCodeNamePropertiesChange(Sender: TObject);
    procedure ProductItemCodeNamePropertiesInitPopup(Sender: TObject);
    procedure btnInstCodeStrClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure actSaveExecute(Sender: TObject);
    procedure actCancelExecute(Sender: TObject);
  private
    { Private declarations }
    FMode: TDML;
    FKeyValue: String;
    FDataFrom: String;
    FDataSet: TClientDataSet;
    FReader: TADOQuery;
    FMaxFaciItem: String;
    FQueryParam: TQueryParam;
    FInstCodeDataSet: TClientDataSet;
    FCanEnable : Boolean;
    procedure ClearEditor;
    procedure DataToEditor;
    procedure ButtonStateChange;
    procedure EditorStateChange;
    function EditorToData: Boolean;
    function VdMustInput: Boolean;
    function DataValidate(): Boolean;
    function DataSetStateChange(const aMode: TDML): Boolean;
    procedure DMLModeChange(aMode: TDML);
  public
    { Public declarations }
    constructor Create(aMode: TDML; aKeyValue, aDataFrom: String;
      aDataSet: TClientDataSet); reintroduce;
    property InstCodeDataSet: TClientDataSet read FInstCodeDataSet write FInstCodeDataSet;
  end;

var
  frmCD078JU: TfrmCD078JU;

implementation

uses cbDBController, frmMultiSelectU, cbUtilis;

{$R *.dfm}

constructor TfrmCD078JU.Create(aMode: TDML; aKeyValue, aDataFrom: String;
  aDataSet: TClientDataSet);
begin
  inherited Create( Application );
  FMode := aMode;
  FKeyValue := aKeyValue;
  FDataFrom := aDataFrom;
  FDataSet := aDataSet;

  FReader := DBController.CodeReader;
end;

procedure TfrmCD078JU.ServiceTypeCodeNamePropertiesInitPopup(
  Sender: TObject);
begin
  { 服務類別 }
  Screen.Cursor := crSQLWait;
  try
    ServiceType.SQL.Text := Format(
    ' SELECT CODENO, DESCRIPTION    ' +
    '  FROM %s.CD046                ' ,
    [DBController.LoginInfo.DbAccount] );
    ServiceType.CodeNamePropertiesInitPopup( Sender );
  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TfrmCD078JU.FormCreate(Sender: TObject);
begin
  ServiceType.Initializa;
  ProductItem.Initializa;

end;

procedure TfrmCD078JU.ServiceTypeCodeNoPropertiesChange(Sender: TObject);
begin
  ServiceType.CodeNoPropertiesChange( Sender );
  ProductItem.Clear;

end;

procedure TfrmCD078JU.ServiceTypeCodeNamePropertiesChange(Sender: TObject);
begin
  ServiceType.CodeNamePropertiesChange( Sender );
  ProductItem.Clear;

end;

procedure TfrmCD078JU.FormDestroy(Sender: TObject);
begin
  ServiceType.Finaliza;
  ProductItem.Finaliza;

end;

procedure TfrmCD078JU.ProductItemCodeNoPropertiesChange(Sender: TObject);
begin
  ProductItem.CodeNoPropertiesChange( Sender );
end;

procedure TfrmCD078JU.ProductItemCodeNamePropertiesChange(Sender: TObject);
begin
  ProductItem.CodeNamePropertiesChange( Sender );
end;

procedure TfrmCD078JU.ProductItemCodeNamePropertiesInitPopup(
  Sender: TObject);
  var aWhere : String;

begin
  Screen.Cursor := crSQLWait;

  CloneDataSet.CloneCursor( FDataSet , False);
  aWhere := '-99';
  //aBMark := FDataSet.GetBookmark;
  //#7238 some of items description cannot be appeared,so add dmCannotModify and dmBrowser to check by kin 2016/06/08
  if ( FMode in [ dmUpdate,dmCannotModify,dmBrowser ] ) then
  begin
    CloneDataSet.Filter := Format( 'PRODUCTCODE<>''%s''', [FDataSet.FieldByName( 'PRODUCTCODE' ).AsString] );
    CloneDataSet.Filtered := True;
  end;

  try
    CloneDataSet.First;
    while not CloneDataSet.Eof do
    begin
      if CloneDataSet.FieldByName( 'PRODUCTCODE' ).AsString <> EmptyStr then
      begin
        aWhere := aWhere + ',' + CloneDataSet.FieldByName( 'PRODUCTCODE' ).AsString;
      end;

      CloneDataSet.Next;
    end;

  finally
    CloneDataSet.Filter := EmptyStr;
    CloneDataSet.Filtered := False;
    //FDataSet.GotoBookmark( aBMark );
    //FDataSet.FreeBookmark( aBMark );
  end;
  aWhere := Format(' CODENO NOT IN ( %s )',[aWhere]);
  try
    ProductItem.SQL.Text := Format(
    ' SELECT CODENO, DESCRIPTION    ' +
    '  FROM %s.CD116                ' +
    '  WHERE NVL(STOPFLAG,0) = 0    ' +
    '  AND %s                       ',
    [DBController.LoginInfo.DbAccount,aWhere] );
    ProductItem.CodeNamePropertiesInitPopup( Sender );
  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TfrmCD078JU.btnInstCodeStrClick(Sender: TObject);
begin
  if ( FMode in [dmInsert, dmUpdate] ) then
  begin
    frmMultiSelect := TfrmMultiSelect.Create( Application );
    try
      frmMultiSelect.Connection := DBController.DataConnection;
      frmMultiSelect.KeyFields := 'INSTCODE';
      frmMultiSelect.KeyValues := FQueryParam.Param1;
      frmMultiSelect.DisplayFields := 'INSTCODE,裝機代碼,INSTNAME,裝機類別名稱';
      frmMultiSelect.ResultFields := 'INSTNAME';
      FInstCodeDataSet.Filtered := False;
      FInstCodeDataSet.Filter := Format( 'SERVICETYPE=''%s''', [ServiceType.Value] );
      FInstCodeDataSet.Filtered := True;
      try
        frmMultiSelect.DataSet := FInstCodeDataSet;
        if frmMultiSelect.ShowModal = mrOk then
        begin
          if ( FMode in [dmInsert, dmUpdate] ) then
          begin
            FQueryParam.Param1 := frmMultiSelect.SelectedValue;
            edtInstCodeStr.Text := frmMultiSelect.SelectedDisplay;
          end;
        end;
      finally
        FInstCodeDataSet.Filtered := False;
        FInstCodeDataSet.Filter := EmptyStr;
      end;
    finally
      frmMultiSelect.Free;
    end;
  end;
end;

procedure TfrmCD078JU.ClearEditor;
begin
  edtFaciItem.Clear;
  ServiceType.Clear;
  ProductItem.Clear;
  FQueryParam.Param1 := EmptyStr;
  edtInstCodeStr.Text := EmptyStr;
end;

procedure TfrmCD078JU.DataToEditor;
begin
  ServiceType.Value := FDataSet.FieldByName( 'SERVICETYPE' ).AsString;
  ProductItem.Value := FDataSet.FieldByName( 'PRODUCTCODE' ).AsString;
  edtFaciItem.Text := FDataSet.FieldByName( 'FACIITEM' ).AsString;
  FMaxFaciItem := FDataSet.FieldByName( 'FACIITEM' ).AsString;
  FQueryParam.Param1 :=  FDataSet.FieldByName( 'INSTCODESTR' ).AsString;
  edtInstCodeStr.Text := DBController.GetInstName( FQueryParam.Param1 );
end;

procedure TfrmCD078JU.DMLModeChange(aMode: TDML);
begin
  if not DataSetStateChange( aMode ) then Exit;
  if ( aMode = dmSave  ) and ( FMode = dmInsert ) then
  begin
    DMLModeChange( dmInsert );
    Exit;
  end else
  if ( aMode in [dmSave, dmCancel] ) then
    FMode := dmBrowser;
  ButtonStateChange;
  EditorStateChange;
  edtDML.Text := TDMLString[FMode];
  //#7238 check the outside buttons  whether can be use by kin 2016/06/8
  if ( not FDataSet.IsEmpty ) and ( FCanEnable ) then
  begin
    frmCD078B1.btnUpdate.Enabled := True;
    frmCD078B1.actUpdate.Enabled := True;
    frmCD078B1.btnDelete.Enabled := True;
    frmCD078B1.actDelete.Enabled := True;
  end;
end;

function TfrmCD078JU.DataSetStateChange(const aMode: TDML): Boolean;
begin
  Result := True;
  case aMode of
    dmInsert:
      begin
        ClearEditor;
        FMaxFaciItem := DBController.GetFaciItemSeqNo( FDataSet );
        FDataSet.Append;
      end;
    dmUpdate:
      begin
        ClearEditor;
        DataToEditor;
        FDataSet.Edit;
      end;
    dmCancel:
      begin
        FDataSet.Cancel;
      end;
    dmBrowser,dmCannotModify:
      begin
        ClearEditor;
        DataToEditor;
      end;  
    dmSave:
      begin
        Result := DataValidate;
        if Result then Result := EditorToData;
        if Result then
        begin
          try
            FDataSet.Post;
//            DoSubDetailSave;
//            FAlreadySetCItemValue := frmCD078B1.GetCD078GCItemValue;

          except
            on E: Exception do
            begin
              ErrorMsg( Format( '存檔有誤, 原因:%s。', [E.Message] ) );
              Result := False;
            end;  
          end;
        end;
      end;
  end;
end;

procedure TfrmCD078JU.ButtonStateChange;
begin
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
    actCancel.Caption := '結束(&C)';
    actCancel.ShortCut := TextToShortCut( 'Alt+C' );
  end;
end;

procedure TfrmCD078JU.EditorStateChange;
 var aIndex : Integer;
begin
  edtFaciItem.Enabled := False;
  edtFaciItem.Text := FMaxFaciItem;
  for aIndex := 0 to ControlCount - 1 do
  begin
    if ( Controls[aIndex].Name <> 'btnSave' ) and
       ( Controls[aIndex].Name <> 'Panel2' )  then
      Controls[aIndex].Enabled := not ( FMode in [dmBrowser,dmCannotModify] );
  end;

//  btnUpdate.Enabled := ( FMode in [dmBrowser] );
  {
  case FMode of
    dmInsert:
      begin
        if ( ServiceType.CodeNo.CanFocusEx ) then ServiceType.CodeNo.SetFocus;
        ProductItem.Enabled := True;
        ServiceType.Enabled := True;
      end;
    dmUpdate:
      begin
        if ( ServiceType.CodeNo.CanFocusEx ) then ServiceType.CodeNo.SetFocus;
        ProductItem.Enabled := True;

      end;
    dmBrowser:
      begin
        if ( ServiceType.CodeNo.CanFocusEx ) then ServiceType.CodeNo.SetFocus;
      end;
  end;
  }
end;

procedure TfrmCD078JU.FormShow(Sender: TObject);
begin
  Screen.Cursor := crSQLWait;
  try
//    FaciItem.DataSet := FaciDataSet;
    //#7238 add globle variable to check whether the outside buttons can be use By Kin 2016/06/08
    FCanEnable := True;
    if FMode = dmCannotModify then FCanEnable := False;
    DMLModeChange( FMode );
  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TfrmCD078JU.actSaveExecute(Sender: TObject);
begin
  DMLModeChange( dmSave );
end;

procedure TfrmCD078JU.actCancelExecute(Sender: TObject);
begin
  DMLModeChange( dmCancel );
  Self.Close;
end;

function TfrmCD078JU.EditorToData: Boolean;
begin
  Result := False;
  try
    if ( FMode in [dmInsert] ) then
    begin
      FDataSet.FieldByName( 'BPCODE' ).AsString := FKeyValue;
    end;
    FDataSet.FieldByName( 'PRODUCTCODE' ).AsString := ProductItem.Value;
    FDataSet.FieldByName( 'PRODUCTNAME' ).AsString := ProductItem.ValueName;
    FDataSet.FieldByName( 'SERVICETYPE' ).AsString := ServiceType.Value;
    FDataSet.FieldByName( 'FACIITEM' ).AsString := edtFaciItem.Text;
    FDataSet.FieldByName( 'INSTCODESTR' ).AsString := TrimChar( FQueryParam.Param1, [''''] );
   except
    on E: Exception do
    begin
      ErrorMsg( Format( '存檔有誤, 原因:%s', [E.Message] ) );
      Exit;
    end;
  end;
  Result := True;
end;

function TfrmCD078JU.VdMustInput: Boolean;
 var aExit : Boolean;
begin
  Result := False;
  aExit := False;
  if ( ServiceType.Value = EmptyStr ) then
  begin
    WarningMsg( '請輸入【服務類別】。' );
    if ServiceType.CodeNo.CanFocusEx then ServiceType.CodeNo.SetFocus;
    Exit;
  end;
  {}
  if ( ProductItem.Value = EmptyStr ) then
  begin
    WarningMsg( '請輸入【產品項目】。' );
    if ProductItem.CodeNo.CanFocusEx then ProductItem.CodeNo.SetFocus;
    Exit;
  end;
  Result := True;
  {
  if ( FQueryParam.Param1 = EmptyStr ) then
  begin
    WarningMsg( '請指定派工類別！');
    if btnInstCodeStr.CanFocus then btnInstCodeStr.Click;
    Exit;
  end;
  }


end;

function TfrmCD078JU.DataValidate: Boolean;
begin
  Result := VdMustInput;
end;

end.
