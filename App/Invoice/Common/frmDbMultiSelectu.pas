
unit frmDbMultiSelectu;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, Db, DBGrids, ExtCtrls, StdCtrls, Buttons, Variants, DBClient;

type
  TQryFieldObj = class(TObject)
    s_FieldName : String;
    s_DisplayLabel : String;
  end;

  TfrmDbMultiSelect = class(TForm)
    pnlFunc: TPanel;
    dbgData: TDBGrid;
    dsrData: TDataSource;
    pnlLSide: TPanel;
    cobQryField: TComboBox;
    edtKey: TEdit;
    btnQuery: TBitBtn;
    Label1: TLabel;
    btnOk: TBitBtn;
    btnCancel: TBitBtn;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure dbgDataKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure dbgDataDblClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btnQueryClick(Sender: TObject);
    procedure btnOkClick(Sender: TObject);
    procedure edtKeyKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure dbgDataCellClick(Column: TColumn);
    procedure dbgDataTitleClick(Column: TColumn);
  private
  private
    { Private declarations }
    FDataSet : TDataSet ;
    FDisplayFields : TStrings ;
    FDisplayCaptions : TStrings ;
    FBookmark : TBookmark ;
    FLeftPos, FTopPos: Integer;
    bG_SelectedAllData : boolean;
    G_TargetFieldNameStrList : TStringList;
    G_TargetFieldValueStrList : TStringList;
    sG_RecSep, sG_SQL : String;
    procedure setSelectedItems(I_TargetFieldNameStrList, I_SelectedItemsValueStrList:TStringList; sI_RecSep:String);
    procedure removeIndex;
    procedure setGridColor(I_DbGrid : TDBGrid);
    procedure selectAllData(bI_SelectedAll:Boolean);
    procedure getSelectedRecordValues(I_DbGrid : TDBGrid; I_TargetFieldNameStrList : TStringList; var I_TargetFieldValueStrList : TStringList; sI_RecSep:String; var sI_SQL : String);
    procedure SetDisplayFields(const Value: string);
    procedure SetDisplayCaptions(const Value: string);
    procedure SetDelete(const Value: TNotifyEvent);
    procedure SetNew(const Value: TNotifyEvent);
    procedure SetUpdate(const Value: TNotifyEvent);
    procedure SetDataSet(const Value: TDataSet);
  public
    { Public declarations }
    property DisplayFields : string write SetDisplayFields;
    property DisplayCaptions : string write SetDisplayCaptions;
    property NewFunc: TNotifyEvent write SetNew;
    property UpdateFunc: TNotifyEvent write SetUpdate;
    property DeleteFunc: TNotifyEvent write SetDelete;
    property DataSet: TDataSet read FDataSet write SetDataSet;
    procedure SetFunc(fNew, fUpdate, fDelete: TNotifyEvent);
    procedure AddFunc(sCaption: string; Callback: TNotifyEvent;
      Modal: TModalResult);
    procedure SetConfirmState(EnableOk, EnableCanCel: Boolean;
      OKCap, CancelCap: string);
  end;

function SelectMultiRecords(Caption:string; DataSet:TDataSet;
  Fields, sI_ValueSep:string; I_TargetFieldNamesStrList:TStringList;
  var I_TargetFieldValueStrList : TStringList;
  bI_SelectedAllData:boolean;
  var sI_SQL : String): TModalResult ; overload;

function SelectMultiRecords(Caption:string; DataSet:TDataSet;
  Fields, Captions, sI_ValueSep: string; I_TargetFieldNamesStrList:TStringList;
  var I_TargetFieldValueStrList : TStringList;
  bI_SelectedAllData:boolean;
  var sI_SQL : String): TModalResult ; overload;


implementation

uses Ustru, UdbU;

{$R *.DFM}


function SelectMultiRecords(Caption:string; DataSet:TDataSet;
  Fields, sI_ValueSep:string;
  I_TargetFieldNamesStrList:TStringList;
  var I_TargetFieldValueStrList : TStringList;
  bI_SelectedAllData:boolean; var sI_SQL : String): TModalResult ; overload;
var
  aFrm : TfrmDbMultiSelect ;
begin
  aFrm := TfrmDbMultiSelect.Create(Application) ;
  aFrm.sG_RecSep := sI_ValueSep;
  aFrm.G_TargetFieldNameStrList := I_TargetFieldNamesStrList;
  aFrm.G_TargetFieldValueStrList := I_TargetFieldValueStrList;
  aFrm.bG_SelectedAllData := bI_SelectedAllData;
  aFrm.Caption := Caption ;
  aFrm.DataSet := DataSet ;
  aFrm.DisplayFields := Fields ;
  Result := aFrm.ShowModal ;
  I_TargetFieldValueStrList := aFrm.G_TargetFieldValueStrList;
  sI_SQL := aFrm.sG_SQL;
  aFrm.Free;
end ;

function SelectMultiRecords(Caption:string; DataSet:TDataSet;
  Fields, Captions, sI_ValueSep: string; I_TargetFieldNamesStrList:TStringList; var I_TargetFieldValueStrList : TStringList; bI_SelectedAllData:boolean; var sI_SQL : String): TModalResult ; overload;

var
  aFrm : TfrmDbMultiSelect ;
begin
  aFrm := TfrmDbMultiSelect.Create(Application) ;
  aFrm.sG_RecSep := sI_ValueSep;
  aFrm.G_TargetFieldNameStrList := I_TargetFieldNamesStrList;
  aFrm.G_TargetFieldValueStrList := I_TargetFieldValueStrList;
  aFrm.bG_SelectedAllData := bI_SelectedAllData;
  aFrm.Caption := Caption ;
  aFrm.DataSet := DataSet ;
  aFrm.DisplayFields := Fields ;
  aFrm.DisplayCaptions := Captions ;
  Result := aFrm.ShowModal ;
  I_TargetFieldValueStrList := aFrm.G_TargetFieldValueStrList;
  sI_SQL := aFrm.sG_SQL;
  aFrm.Free;
end ;

{ TfrmSelect }

procedure TfrmDbMultiSelect.SetDisplayFields(const Value: string);
begin
  if Value = '' then Exit ;
  {parese the input string to get all display fields}
  FDisplayFields := TUstr.ParseStrings(Value, ';') ;
end;

procedure TfrmDbMultiSelect.SetDisplayCaptions(const Value: string);
begin
  if Value = '' then Exit ;
  {parese the input string to get all display fields}
  FDisplayCaptions := TUstr.ParseStrings(Value, ';') ;
end;

procedure TfrmDbMultiSelect.FormClose(Sender: TObject; var Action: TCloseAction);
begin
    removeIndex;
  if (ModalResult <> mrOk) and (FDataSet.RecordCount > 0) then
  begin
    {back to original record}
    FDataSet.GotoBookmark(FBookmark) ;
  end ;
  FDataSet.FreeBookmark(FBookmark) ;
  if Assigned(FDisplayFields) then FDisplayFields.Free;
  if Assigned(FDisplayCaptions) then FDisplayCaptions.Free;
end;

procedure TfrmDbMultiSelect.FormShow(Sender: TObject);
var

  ii, wid : Integer ;
  fld : TField ;
  clm : TColumn ;
  L_Obj : TQryFieldObj;
begin
  {record the dataset and it's bookmark}
  if not Assigned(dsrData.DataSet) then Exit;
  FDataSet := dsrData.DataSet;

  if (not (FDataSet.State = dsInactive)) and (FDataSet Is TClientDataSet) then
  begin
    if (FDataSet as TClientDataSet).IndexName = 'TmpIndex' then
      (FDataSet as TClientDataSet).DeleteIndex('TmpIndex');

  end;
    
  FBookmark := FDataSet.GetBookmark ;
  {if DisplayFields not specified, then display all fields}
  if not Assigned(FDisplayFields) or (FDisplayFields.Count=0) then Exit ;
  {else, display the specified fields}
  wid := 0 ;
  cobQryField.Items.Clear;
  for ii:=0 to FDisplayFields.Count-1 do
  begin
    fld := FDataSet.FieldByName(FDisplayFields[ii]);

    if not Assigned(fld) then Continue ;
    L_Obj := TQryFieldObj.Create;
    L_Obj.s_FieldName := fld.FieldName;
    L_Obj.s_DisplayLabel := fld.DisplayLabel;

    cobQryField.Items.AddObject(L_Obj.s_DisplayLabel, L_Obj);

    clm := dbgData.Columns.Add ;
    clm.Field := fld ;
    if Assigned(FDisplayCaptions) then
      clm.Title.Caption := FDisplayCaptions[ii];
    wid := wid + clm.Width ;
  end ;
  {resize the form to fit display}
  if ClientWidth < wid then
    if wid > Screen.Width then ClientWidth := Screen.Width
    else ClientWidth := wid ;
  dbgData.SetFocus ;
  cobQryField.ItemIndex := 0;

  //若 G_TargetFieldValueStrList 沒有值,則由 bG_SelectedAllData 來決定是否要選取所有的項目
  //若 G_TargetFieldValueStrList 有值,表示曾點選過項目,所以要將其上次點選的項目設定為 selected
  if G_TargetFieldValueStrList.Count =0 then
    selectAllData(bG_SelectedAllData)
  else
    setSelectedItems(G_TargetFieldNameStrList, G_TargetFieldValueStrList, sG_RecSep);

end;

procedure TfrmDbMultiSelect.dbgDataKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = VK_RETURN then btnOk.Click
  else if Key = VK_ESCAPE then  ModalResult := mrCancel ;
end;

procedure TfrmDbMultiSelect.dbgDataDblClick(Sender: TObject);
begin
    btnOkClick(Sender);
//  ModalResult := mrOk ;
end;

procedure TfrmDbMultiSelect.SetDelete(const Value: TNotifyEvent);
begin
end;

procedure TfrmDbMultiSelect.SetNew(const Value: TNotifyEvent);
begin
end;

procedure TfrmDbMultiSelect.SetUpdate(const Value: TNotifyEvent);
begin
end;

procedure TfrmDbMultiSelect.SetFunc(fNew, fUpdate, fDelete: TNotifyEvent);
begin
  NewFunc := fNew;
  UpdateFunc := fUpdate;
  DeleteFunc := fDelete;
end;

procedure TfrmDbMultiSelect.AddFunc(sCaption: string; Callback: TNotifyEvent;
  Modal: TModalResult);
var
  Btn: TButton;
begin
  Btn := TButton.Create(Self);
  Btn.Caption := sCaption;
  Btn.ModalResult := Modal;
  Btn.Left := FLeftPos;
  Btn.Top := FTopPos;
  Btn.Visible := True;
  Btn.OnClick := Callback;
end;

procedure TfrmDbMultiSelect.SetDataSet(const Value: TDataSet);
begin
  FDataSet := Value;
  dsrData.DataSet := FDataSet;
end;

procedure TfrmDbMultiSelect.FormCreate(Sender: TObject);
begin
  FLeftPos := 0;
end;

procedure TfrmDbMultiSelect.SetConfirmState(EnableOk, EnableCanCel: Boolean;
  OKCap, CancelCap: string);
begin
  if OkCap <> '' then btnOk.Caption := '&O' + OKCap;
  if CancelCap <> '' then btnCancel.Caption := '&C' + CancelCap;
end;

procedure TfrmDbMultiSelect.btnQueryClick(Sender: TObject);
var
    sL_QryFieldName, sL_KeyValue : Variant;
begin
    if cobQryField.Items.Count =0 then exit;
    sL_QryFieldName := (cobQryField.Items.Objects[cobQryField.ItemIndex] as TQryFieldObj).s_FieldName;
    sL_KeyValue := edtKey.Text;
    if dsrData.DataSet.Locate(sL_QryFieldName, sL_KeyValue, [loPartialKey,loCaseInsensitive]) then
      dbgData.SetFocus
    else
      edtKey.SetFocus;
end;

procedure TfrmDbMultiSelect.btnOkClick(Sender: TObject);
begin
    getSelectedRecordValues(dbgData, G_TargetFieldNameStrList, G_TargetFieldValueStrList, sG_RecSep, sG_SQL);
    if FDataSet.RecordCount > 0 then
      ModalResult := mrOk
    else
      ModalResult := mrCancel;

end;

procedure TfrmDbMultiSelect.edtKeyKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
    if Key=VK_RETURN then
      btnQueryClick(sender);
end;

procedure TfrmDbMultiSelect.getSelectedRecordValues(I_DbGrid: TDBGrid;
  I_TargetFieldNameStrList: TStringList;
  var I_TargetFieldValueStrList: TStringList; sI_RecSep: String; var sI_SQL : String);
var
    sL_NotSelectedValueStr, sL_SelectedValueStr : String;
    sL_Value, sL_CurFieldName : String;
    sL_Sep, sL_SqlColumnName, sL_SelectedValue : String;
    ii, jj, nL_TotalRecCount, nL_HalfRecCount, nL_SelectedCount : Integer;
    L_FieldType : TFieldType;
    L_BookMarkList : TBookMarkList;
begin


    I_TargetFieldValueStrList.Clear;
    with I_DbGrid.DataSource.DataSet do
    begin
      for ii:=0 to I_DbGrid.SelectedRows.Count-1 do
      begin
        GotoBookmark(pointer(I_DbGrid.SelectedRows.Items[ii]));
        for jj:=0 to I_TargetFieldNameStrList.Count -1 do
        begin
          sL_CurFieldName := G_TargetFieldNameStrList.Strings[jj];
          if jj =0 then
            sL_SelectedValue := I_DbGrid.DataSource.DataSet.FieldByName(sL_CurFieldName).AsString
          else
            sL_SelectedValue := sL_SelectedValue + sI_RecSep + I_DbGrid.DataSource.DataSet.FieldByName(sL_CurFieldName).AsString;
        end;
        I_TargetFieldValueStrList.Add(sL_SelectedValue);
      end;
    end;

    //down, 處理將傳回之 SQL
    sL_SqlColumnName := I_TargetFieldNameStrList.Strings[0];
    L_FieldType := TUdb.getDataType(I_DbGrid.DataSource.DataSet, sL_SqlColumnName);
    case L_FieldType of
      ftString :
        sL_Sep := '''';
      ftInteger, ftFloat :
        sL_Sep := '';
    end;

    L_BookMarkList := dbgData.SelectedRows;
    with dbgData.DataSource.DataSet do
    begin
      DisableControls;
      First;
      while not Eof do
      begin
        if L_BookMarkList.CurrentRowSelected then
        begin
          sL_Value := I_DbGrid.DataSource.DataSet.FieldByName(sL_SqlColumnName).AsString;

          if sL_SelectedValueStr='' then
            sL_SelectedValueStr := sL_Sep + sL_Value + sL_Sep
          else
            sL_SelectedValueStr := sL_SelectedValueStr + ',' + sL_Sep + sL_Value + sL_Sep;
        end;
        if not L_BookMarkList.CurrentRowSelected then
        begin
          if sL_NotSelectedValueStr='' then
            sL_NotSelectedValueStr := sL_Sep + I_DbGrid.DataSource.DataSet.FieldByName(sL_SqlColumnName).AsString + sL_Sep
          else
            sL_NotSelectedValueStr := sL_NotSelectedValueStr + ',' + sL_Sep + I_DbGrid.DataSource.DataSet.FieldByName(sL_SqlColumnName).AsString + sL_Sep;
        end;

        Next;
      end;
      EnableControls;
    end;
    //up, 處理將傳回之 SQL

    nL_TotalRecCount := I_DbGrid.DataSource.DataSet.RecordCount;
    nL_HalfRecCount := Round(I_DbGrid.DataSource.DataSet.RecordCount/2);
    nL_SelectedCount := I_DbGrid.SelectedRows.Count;
    if (nL_SelectedCount=0) or (nL_SelectedCount=nL_TotalRecCount) then
      sI_SQL := '';
    if nL_SelectedCount=1 then
      sI_SQL := sL_SqlColumnName + '=' + sL_SelectedValueStr;
    if (nL_SelectedCount<>0)and(nL_SelectedCount<>1) and (nL_SelectedCount<nL_HalfRecCount) then
      sI_SQL := sL_SqlColumnName + ' in ' + '(' + sL_SelectedValueStr + ')';
    if (nL_SelectedCount<>0)and(nL_SelectedCount<>1) and (nL_SelectedCount>=nL_HalfRecCount)and (nL_SelectedCount<>nL_TotalRecCount) then
      sI_SQL := sL_SqlColumnName + ' not in ' + '(' + sL_NotSelectedValueStr + ')';


end;

procedure TfrmDbMultiSelect.selectAllData(bI_SelectedAll: Boolean);
var
    bL_Selected : boolean;
    L_BookMarkList : TBookMarkList;
begin
    if bI_SelectedAll then
      bL_Selected := true
    else
      bL_Selected := false;  

    L_BookMarkList := dbgData.SelectedRows;
    with dbgData.DataSource.DataSet do
    begin
      DisableControls;
      First;
      while not Eof do
      begin
        L_BookMarkList.CurrentRowSelected := bL_Selected;
        Next;
      end;
      EnableControls;
    end;

    setGridColor(dbgData);
end;

procedure TfrmDbMultiSelect.SpeedButton1Click(Sender: TObject);
begin
    selectAllData(true);
end;

procedure TfrmDbMultiSelect.SpeedButton2Click(Sender: TObject);
begin
    selectAllData(false);
end;

procedure TfrmDbMultiSelect.setGridColor(I_DbGrid : TDBGrid);
var
    nL_TotalRecCount,nL_SelectedCount : Integer;
begin
    nL_TotalRecCount := I_DbGrid.DataSource.DataSet.RecordCount;
    nL_SelectedCount := I_DbGrid.SelectedRows.Count;
    if nL_SelectedCount=0 then
      dbgData.Color := clMoneyGreen
    else if nL_TotalRecCount=nL_SelectedCount then
      dbgData.Color := clGradientActiveCaption
    else
      dbgData.Color := clInfoBk;
end;

procedure TfrmDbMultiSelect.dbgDataCellClick(Column: TColumn);
begin
    setGridColor(dbgData);
end;

procedure TfrmDbMultiSelect.dbgDataTitleClick(Column: TColumn);
var
   ii : Integer;
   bFlag : Boolean;
   FDataSet : TDataSet;
begin
     if not (dsrData.DataSet IS TClientDataSet) then Exit;
     FDataSet := dsrData.DataSet AS TClientDataSet;
     bFlag := False;
     for ii:=0 to FDataSet.FieldCount-1 do
     begin
       if (FDataSet.Fields[ii].FieldName=Column.FieldName) and
          (FDataSet.Fields[ii].FieldKind=fkData) then
       begin
         bFlag := True;
         break;
       end;
     end;

     if not bFlag then
       Exit;//該FieldName不存在於FDataSet中
     if FDataSet IS TClientDataSet then
     begin
       if TClientDataSet(FDataSet).IndexName = 'TmpIndex' then
         TClientDataSet(FDataSet).DeleteIndex('TmpIndex');
       TClientDataSet(FDataSet).AddIndex('TmpIndex', Column.FieldName, [ixCaseInsensitive],'','',0);
       TClientDataSet(FDataSet).IndexName := 'TmpIndex';
     end;
end;

procedure TfrmDbMultiSelect.removeIndex;
begin
       if TClientDataSet(FDataSet).IndexName = 'TmpIndex' then
         TClientDataSet(FDataSet).DeleteIndex('TmpIndex');
end;

procedure TfrmDbMultiSelect.setSelectedItems(
  I_TargetFieldNameStrList,
  I_SelectedItemsValueStrList: TStringList; sI_RecSep:String);
var
    ii, jj : Integer;
    sL_SelectedFieldNames, sL_SelectedItemValue : String;
    L_TmpStrList : TStringList;
begin
    //down, 串出 locate 時會用到的 Fields Names...
    for ii:=0 to I_TargetFieldNameStrList.Count -1 do
    begin
      if ii=0 then
        sL_SelectedFieldNames := I_TargetFieldNameStrList.Strings[ii];
      {//在此不能用多個欄位做 locate, 所以 mark 起來這些程式...
      else
        sL_SelectedFieldNames := sL_SelectedFieldNames + ';' + I_TargetFieldNameStrList.Strings[ii];
      }
    end;
    //up, 串出 locate 時會用到的 Fields Names...


    for ii:=0 to I_SelectedItemsValueStrList.Count -1 do
    begin
      //down, 串出 locate 時會用到的 目標值 字串

      L_TmpStrList := TUstr.ParseStrings(I_SelectedItemsValueStrList.Strings[ii], sI_RecSep);
      for jj:=0 to L_TmpStrList.Count -1 do
      begin
        if jj=0 then
          sL_SelectedItemValue := L_TmpStrList.Strings[jj];
        {//在此不能用多個欄位做 locate, 所以 mark 起來這些程式...
        else
          sL_SelectedItemValue := sL_SelectedItemValue + ',' + L_TmpStrList.Strings[jj];
        }
      end;
      //up, 串出 locate 時會用到的 目標值 字串
      if AnsiPos(',',sL_SelectedFieldNames)=0 then
      begin
        if dbgData.DataSource.DataSet.Locate(sL_SelectedFieldNames,sL_SelectedItemValue,[loPartialKey]) then
          dbgData.SelectedRows.CurrentRowSelected := true;
      end
      else
      begin
        if dbgData.DataSource.DataSet.Locate(sL_SelectedFieldNames,VarArrayOf([sL_SelectedItemValue]),[loPartialKey]) then
        begin
          dbgData.SelectedRows.CurrentRowSelected := true;
        end;
      end;
    end;
    setGridColor(dbgData);
end;

end.
