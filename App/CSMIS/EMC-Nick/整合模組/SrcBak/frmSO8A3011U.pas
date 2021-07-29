
unit frmSO8A3011U;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, Db, DBGrids, ExtCtrls, StdCtrls, Buttons, Variants, DBClient;

type
  TQryFieldObj = class(TObject)
    s_FieldName : String;
    s_DisplayLabel : String;
  end;

  TfrmSO8A3011 = class(TForm)
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
    chbWholeRec: TCheckBox;
    BitBtn1: TBitBtn;
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
    procedure btnCancelClick(Sender: TObject);
    procedure dbgDataTitleClick(Column: TColumn);
  private
    { Private declarations }
    FDataSet : TDataSet ;
    FDisplayFields : TStrings ;
    FDisplayCaptions : TStrings ;
    FBookmark : TBookmark ;
    FLeftPos, FTopPos: Integer;
    FDefaultQryStr: String;
    procedure cancelFilter;        
    procedure SetDisplayFields(const Value: string);
    procedure SetDisplayCaptions(const Value: string);
    procedure SetDelete(const Value: TNotifyEvent);
    procedure SetNew(const Value: TNotifyEvent);
    procedure SetUpdate(const Value: TNotifyEvent);
    procedure SetDataSet(const Value: TDataSet);
    procedure SetDefaultQryStr(const Value: String);
  public
    { Public declarations }
    property DefaultQryStr : String read FDefaultQryStr write SetDefaultQryStr;
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

function SelectRecord(Caption:string; DataSet:TDataSet;
  Fields:string): TModalResult ; overload;
function SelectRecord(Caption:string; DataSet:TDataSet;
  Fields, Captions: string): TModalResult ; overload;


implementation

uses Ustru, dtmMain2U;

{$R *.DFM}

function SelectRecord(Caption:string; DataSet:TDataSet;
  Fields:string): TModalResult ;
var
  aFrm : TfrmSO8A3011 ;
begin
  aFrm := TfrmSO8A3011.Create(Application) ;
  aFrm.Caption := Caption ;
  aFrm.DataSet := DataSet ;
  aFrm.DisplayFields := Fields ;
//  aFrm.SetDefaultQryStr(sI_DefaultQryStr);  
  Result := aFrm.ShowModal ;
  aFrm.Free;
end ;

function SelectRecord(Caption:string; DataSet:TDataSet;
  Fields, Captions: string): TModalResult ;
var
  aFrm : TfrmSO8A3011 ;
begin
  aFrm := TfrmSO8A3011.Create(Application) ;
  aFrm.Caption := Caption ;
  aFrm.DataSet := DataSet ;
  aFrm.DisplayFields := Fields ;
  aFrm.DisplayCaptions := Captions ;
//  aFrm.SetDefaultQryStr(sI_DefaultQryStr);

  Result := aFrm.ShowModal ;
  aFrm.Free;
end ;

{ TfrmSelect }

procedure TfrmSO8A3011.SetDisplayFields(const Value: string);
begin
  if Value = '' then Exit ;
  {parese the input string to get all display fields}
  FDisplayFields := TUstr.ParseStrings(Value, ';') ;
end;

procedure TfrmSO8A3011.SetDisplayCaptions(const Value: string);
begin
  if Value = '' then Exit ;
  {parese the input string to get all display fields}
  FDisplayCaptions := TUstr.ParseStrings(Value, ';') ;
end;

procedure TfrmSO8A3011.FormClose(Sender: TObject; var Action: TCloseAction);
var
        L_BookMark : TBookMark;
begin
  L_BookMark := FDataSet.GetBookmark;
  cancelFilter;
  FDataSet.GotoBookmark(L_BookMark);
  if (ModalResult <> mrOk) and (FDataSet.RecordCount > 0) then
  begin
    {back to original record}
    FDataSet.GotoBookmark(FBookmark) ;
  end ;
  FDataSet.FreeBookmark(FBookmark) ;
  if Assigned(FDisplayFields) then FDisplayFields.Free;
  if Assigned(FDisplayCaptions) then FDisplayCaptions.Free;  
end;

procedure TfrmSO8A3011.FormShow(Sender: TObject);
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

  if FDefaultQryStr<>'' then
  begin
    edtKey.Text := FDefaultQryStr;
    btnQueryClick(Sender);
  end;


end;

procedure TfrmSO8A3011.dbgDataKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = VK_RETURN then btnOk.Click
  else if Key = VK_ESCAPE then  ModalResult := mrCancel ;
end;

procedure TfrmSO8A3011.dbgDataDblClick(Sender: TObject);
begin
    btnOkClick(Sender);
//  ModalResult := mrOk ;
end;

procedure TfrmSO8A3011.SetDelete(const Value: TNotifyEvent);
begin
end;

procedure TfrmSO8A3011.SetNew(const Value: TNotifyEvent);
begin
end;

procedure TfrmSO8A3011.SetUpdate(const Value: TNotifyEvent);
begin
end;

procedure TfrmSO8A3011.SetFunc(fNew, fUpdate, fDelete: TNotifyEvent);
begin
  NewFunc := fNew;
  UpdateFunc := fUpdate;
  DeleteFunc := fDelete;
end;

procedure TfrmSO8A3011.AddFunc(sCaption: string; Callback: TNotifyEvent;
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

procedure TfrmSO8A3011.SetDataSet(const Value: TDataSet);
begin
  FDataSet := Value;
  dsrData.DataSet := FDataSet;
end;

procedure TfrmSO8A3011.FormCreate(Sender: TObject);
begin
  FLeftPos := 0;
end;

procedure TfrmSO8A3011.SetConfirmState(EnableOk, EnableCanCel: Boolean;
  OKCap, CancelCap: string);
begin
  if OkCap <> '' then btnOk.Caption := '&O' + OKCap;
  if CancelCap <> '' then btnCancel.Caption := '&C' + CancelCap;
end;

procedure TfrmSO8A3011.btnQueryClick(Sender: TObject);
var
    sL_QrySep, sL_TmpFilter, sL_WildFlag : String;
    sL_QryFieldName, sL_KeyValue : Variant;
begin

    if cobQryField.Items.Count =0 then exit;
    sL_QryFieldName := (cobQryField.Items.Objects[cobQryField.ItemIndex] as TQryFieldObj).s_FieldName;
    sL_KeyValue := edtKey.Text;
    if dsrData.DataSet is TClientDataSet then
    begin
      with (dsrData.DataSet As TClientDataSet) do
      begin
        if sL_KeyValue <> '' then
        begin
          sL_QrySep := '''';
          sL_WildFlag := '%';
          Filtered := False;
          if chbWholeRec.Checked then
            sL_TmpFilter := sL_QryFieldName + ' like ' + sL_QrySep + sL_WildFlag + sL_KeyValue  + sL_WildFlag  + sL_QrySep
          else
            sL_TmpFilter := sL_QryFieldName + ' like ' + sL_QrySep + sL_KeyValue  + sL_WildFlag  + sL_QrySep;            
        end
        else
          sL_TmpFilter := '';
        Filter := sL_TmpFilter;
        Filtered := True;
        if RecordCount>0 then
          dbgData.SetFocus
        else
          edtKey.SetFocus;  

      end; 
    end
    else
    begin
      if dsrData.DataSet.Locate(sL_QryFieldName, sL_KeyValue, [loPartialKey,loCaseInsensitive]) then
        dbgData.SetFocus
      else
        edtKey.SetFocus;
    end;
end;

procedure TfrmSO8A3011.btnOkClick(Sender: TObject);
begin
  if FDataSet.RecordCount > 0 then ModalResult := mrOk
  else ModalResult := mrCancel;

end;

procedure TfrmSO8A3011.edtKeyKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
    if Key=VK_RETURN then
      btnQueryClick(sender);
end;

procedure TfrmSO8A3011.SetDefaultQryStr(const Value: String);
begin
  FDefaultQryStr := Value;
end;

procedure TfrmSO8A3011.SpeedButton1Click(Sender: TObject);
begin
    edtKey.Text := '';
    btnQueryClick(Sender);
end;

procedure TfrmSO8A3011.btnCancelClick(Sender: TObject);
begin
    cancelFilter;

end;

procedure TfrmSO8A3011.cancelFilter;
begin
    if dsrData.DataSet is TClientDataSet then
    begin
      with (dsrData.DataSet As TClientDataSet) do
      begin
          Filtered := False;
      end;
    end;

end;

procedure TfrmSO8A3011.dbgDataTitleClick(Column: TColumn);
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

end.
