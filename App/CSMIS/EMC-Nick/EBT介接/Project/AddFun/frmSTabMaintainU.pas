unit frmSTabMaintainu;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, ExtCtrls, Grids, DBGrids, ComCtrls, Spin, DB, DBTables,
  DBClient, dbctrls, mask, rptBasicu;

const
    QRYWITHCDS=168;
    QRYWITHQUERY=169;

type


  TfrmSTabMaintain = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    btnSave: TBitBtn;
    btnCancel: TBitBtn;
    edtQryString: TEdit;
    Label7: TLabel;
    btnQuery: TBitBtn;
    btnAppend: TBitBtn;
    btnEdit: TBitBtn;
    btnDelete: TBitBtn;
    btnExit: TBitBtn;
    dsrSTabMaintain: TDataSource;
    cobQueryCond: TComboBox;
    Label1: TLabel;
    Panel3: TPanel;
    pnlSingleData: TPanel;
    pnlMultiData: TPanel;
    DBGrid1: TDBGrid;
    Splitter1: TSplitter;
    chbContinueAppend: TCheckBox;
    btnPrint: TBitBtn;
    procedure btnSaveClick(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure btnQueryClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure btnAppendClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure btnEditClick(Sender: TObject);
    procedure edtQryStringKeyPress(Sender: TObject; var Key: Char);
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure DBGrid1DblClick(Sender: TObject);
    procedure DBGrid1TitleClick(Column: TColumn);
    procedure FormMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure Panel2MouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure FormDestroy(Sender: TObject);
    procedure cobQueryCondChange(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
  private
    FDefaultEditFieldName: String;
    FUnEditedFieldNames: TStringList;
    FActionTabName: String;
    FNoneDupFieldName: TStringList;
    FNoneDupFieldNameType: TStringList;
    procedure TrimStr(var sI_Str:String);    
    procedure PreAppendData;
    procedure SetDefaultEditFieldName(const Value: String);
    procedure SetUnEditedFieldNames(const Value: TStringList);
    procedure SetActionTabName(const Value: String);
    procedure SetNoneDupFieldName(const Value: TStringList);
    procedure SetNoneDupFieldNameType(const Value: TStringList);
  private
    FReport : TrptBasic ;  
    FDtm: TDataModule;
    FDataSet: TDataSet;
    FQuery : TQuery;
    sG_QryFieldName : String;
    nG_QryType : Integer;
//    ConstFieldNameStrList : TStringList;
//    ConstFieldValueStrList : TStringList;
    procedure SetReadonlyFields(bI_ReadOnly:Boolean);
    procedure SetDefaultFocus;
    procedure ApplyToDB;
    procedure ChangeState;
    procedure SetDataSet(const Value: TDataSet);
    { Private declarations }
  public
    { Public declarations }
    nG_QryData : Integer;
    bCloseFrm : Boolean;//按完成後是否結束畫面
    MainFrmEntranceFlag : Boolean;
    property Dtm : TDataModule read FDtm ;
  protected
    QueryString : String;
    property NoneDupFieldNameType : TStringList read FNoneDupFieldNameType write SetNoneDupFieldNameType;
    property NoneDupFieldName : TStringList read FNoneDupFieldName write SetNoneDupFieldName;
    property ActionTabName : String read FActionTabName write SetActionTabName;
 
    property UnEditedFieldNames : TStringList read FUnEditedFieldNames write SetUnEditedFieldNames;
    property DefaultEditFieldName:String read FDefaultEditFieldName write SetDefaultEditFieldName;
    procedure SetQueryCond(sI_Desc, sI_RealFieldName:string; nI_Type:Integer);
    procedure SetDataModule(Dtm : TDataModule);
    procedure ReSetControls;
    procedure MyMouseMove(Sender: TObject; Shift: TShiftState; X,Y: Integer);virtual;
    procedure MyEnter(Sender: TObject);
    procedure RealEnter(Sender: TObject);virtual;
    procedure RealExit(Sender: TObject);virtual;
//    procedure SetConstQryData(sFieldName, sFieldValue:String);
  end;

var
  frmSTabMaintain: TfrmSTabMaintain;

implementation

uses dtmSTabMaintainu, UObjectu;

{$R *.DFM}

procedure TfrmSTabMaintain.btnSaveClick(Sender: TObject);
var
  sL_KeyValue : String;
begin
    if FDataSet.State = dsInactive then Exit;
     if sG_QryFieldName='' then
       sG_QryFieldName := TQryCondObj(cobQueryCond.Items.Objects[cobQueryCond.ItemIndex]).s_FieldName;
     sL_KeyValue := FDataSet.FieldByName(sG_QryFieldName).AsString;
     with FDataSet do
     begin
       Post;
     end;
     if bCloseFrm then
       ApplyToDB;
     ChangeState;

     with FDataSet do
     begin
//       Close;
       if not Active then
         Open;
       Filtered := False;
     end;
     FDataSet.Locate(sG_QryFieldName,sL_KeyValue,[]);
     if bCloseFrm then
       ModalResult := mrOK;
     if chbContinueAppend.Checked then
       PreAppendData;  
end;

procedure TfrmSTabMaintain.btnExitClick(Sender: TObject);
begin
//     FDataSet.Cancel;
     ApplyToDB;
     ModalResult := mrOK;       
//     Close;
end;

procedure TfrmSTabMaintain.ChangeState;
begin
     with FDataSet do   
     begin
       if state in [dsInactive] then
         Exit     
       else if state in [dsEdit, dsInsert] then
       begin
          btnCancel.Enabled := TRUE;
          btnSave.Enabled := TRUE;
          btnDelete.Enabled := FALSE;
          btnQuery.Enabled := FALSE;
          btnEdit.Enabled := FALSE;
          btnExit.Enabled := FALSE;
          btnAppend.Enabled := FALSE;
          btnPrint.Enabled := FALSE;
          pnlSingleData.Enabled := TRUE;
          pnlMultiData.Enabled := FALSE;
       end
       else
       begin
          if FDataSet.RecordCount>0 then
          begin
            btnEdit.Enabled := TRUE;
            btnDelete.Enabled := TRUE;
            btnPrint.Enabled := True;            
          end
          else
          begin
            btnEdit.Enabled := FALSE;
            btnDelete.Enabled := FALSE;
            btnPrint.Enabled := FALSE;
          end;
          btnCancel.Enabled := FALSE;
          btnSave.Enabled := FALSE;
          btnQuery.Enabled := TRUE;
          btnAppend.Enabled := TRUE;
          btnExit.Enabled := TRUE;          
          pnlSingleData.Enabled := FALSE;
          pnlMultiData.Enabled := TRUE;
       end;
     end;
end;

procedure TfrmSTabMaintain.SetDataModule(Dtm: TDataModule);
begin
  FDtm := Dtm ;
  FDataSet :=  TdtmSTabMaintain(Dtm).MasterDataSet ;
  FQuery :=  TdtmSTabMaintain(Dtm).MasterQuery ;
  if FDataSet is TClientDataSet then
    TClientDataSet(FDataSet).FetchParams;
  dsrSTabMaintain.DataSet := FDataSet ;
end;

procedure TfrmSTabMaintain.SetDataSet(const Value: TDataSet);
begin
  FDataSet := Value;
end;


procedure TfrmSTabMaintain.btnQueryClick(Sender: TObject);
var
   sL_QrySep, sL_WildFlag : String; 
   nRec_Count : Integer;
   ii:Integer;
   sTmp_Filter, sL_CommandText : String;
   nL_ItemIndex : Integer;
begin
     QueryString := edtQryString.Text;
     TrimStr(QueryString);

     if QueryString='' then
     begin
       if MessageDlg('即將顯示全部資料!!是否確定', mtConfirmation, [mbYes, mbNo], 0) = mrNo then
         Exit;
     end;

     nRec_Count:=0;
     ApplyToDB;
     if nG_QryData = QRYWITHCDS then//在前端透過CDS去Filter Data
     begin
       nL_ItemIndex := cobQueryCond.ItemIndex;
       sG_QryFieldName := TQryCondObj(cobQueryCond.Items.Objects[nL_ItemIndex]).s_FieldName;
       nG_QryType := TQryCondObj(cobQueryCond.Items.Objects[nL_ItemIndex]).n_Type;

       with FDataSet do
       begin
         if TClientDataSet(FDataSet).IndexName='TmpIndex' then
           TClientDataSet(FDataSet).DeleteIndex('TmpIndex');
  //       if not active then
         Close;
         Open;
         Filtered := False;
         sTmp_Filter := '';
         for ii:=0 to TdtmSTabMaintain(Dtm).ConstFieldNameStrList.Count-1 do
         begin
           if ii =0 then
             sTmp_Filter := TdtmSTabMaintain(Dtm).ConstFieldNameStrList.Strings[ii] + '=''' + TdtmSTabMaintain(Dtm).ConstFieldValueStrList.Strings[ii] + ''''
           else
             sTmp_Filter := sTmp_Filter + ' AND ' + TdtmSTabMaintain(Dtm).ConstFieldNameStrList.Strings[ii] + '=''' + TdtmSTabMaintain(Dtm).ConstFieldValueStrList.Strings[ii] + '''';
         end;
         case nG_QryType of
           1://表示此 field 的型態是 number
            begin
             sL_QrySep := '';
             sL_WildFlag := '';
            end;
           2://表示此 field 的型態是 string
            begin
             sL_QrySep := '''';
             sL_WildFlag := '*';
            end;

         end;
         if QueryString<>'' then
         begin
           if sTmp_Filter<>'' then
             sTmp_Filter := sTmp_Filter + 'AND ' + sG_QryFieldName + '=' + sL_QrySep + QueryString + sL_WildFlag + sL_QrySep
           else
             sTmp_Filter := sG_QryFieldName + '=' + sL_QrySep + QueryString + sL_WildFlag + sL_QrySep;
           Filter := sTmp_Filter;
           Filtered := True;
         end
         else
         begin
           if sTmp_Filter<>'' then
           begin
             Filter := sTmp_Filter;
             Filtered := True;
           end;
         end;
         nRec_Count := RecordCount;
       end;
     end
     else//透過CDS到後端去select data
     begin
       nL_ItemIndex := cobQueryCond.ItemIndex;
       sG_QryFieldName := TQryCondObj(cobQueryCond.Items.Objects[nL_ItemIndex]).s_FieldName;
       nG_QryType := TQryCondObj(cobQueryCond.Items.Objects[nL_ItemIndex]).n_Type;

//       with TClientDataSet(FDataSet) do
       with FQuery do
       begin
         Close;
         SQL.Clear;
         SQL.Add('select * from ' + ActionTabName );
         case nG_QryType of
           1://表示此 field 的型態是 number
            begin
             if QueryString<>'' then
               SQL.Add('where ' + sG_QryFieldName  + '=' +  QueryString );
            end;
           2://表示此 field 的型態是 string
            begin
               SQL.Add('where ' + sG_QryFieldName + ' like '+  '''' + QueryString + '%''');
            end;

         end;

         Prepare;
         Open;
         nRec_Count := RecordCount;
         FDataSet.Close;
         FDataSet.Open;
       end;
     end;
     
     if nRec_Count=0 then
     begin
       MessageDlg('查無資料.', mtInformation,[mbOk], 0);
     end;
     ChangeState;

end;

procedure TfrmSTabMaintain.btnCancelClick(Sender: TObject);
begin
    if FDataSet.State = dsInactive then Exit;
     FDataSet.Cancel;
     ChangeState;
end;

procedure TfrmSTabMaintain.btnAppendClick(Sender: TObject);
begin
    if FDataSet.State = dsInactive then
      FDataSet.Active := True;
     SetReadonlyFields(False);
     PreAppendData;
end;

procedure TfrmSTabMaintain.btnDeleteClick(Sender: TObject);
begin
    if FDataSet.State = dsInactive then Exit;
     if MessageDlg('是否確認刪除?',mtConfirmation, [mbYes, mbNo], 0) = mrNo then
       Exit;
     FDataSet.Delete;
     ChangeState;
end;

procedure TfrmSTabMaintain.btnEditClick(Sender: TObject);
begin
    if FDataSet.State = dsInactive then Exit;
     SetReadonlyFields(True);
     FDataSet.Edit;
     ChangeState;
     SetDefaultFocus;
end;

procedure TfrmSTabMaintain.edtQryStringKeyPress(Sender: TObject; var Key: Char);
begin
     if Key=Char(13) then//Press Enter
       btnQueryClick(nil);
end;

procedure TfrmSTabMaintain.FormShow(Sender: TObject);
begin
     if cobQueryCond.ItemIndex<0 then
       cobQueryCond.ItemIndex := 0;
     ChangeState;
end;

procedure TfrmSTabMaintain.ApplyToDB;
begin
     with FDataSet  do
     begin
       if state in [dsEdit, dsInsert] then
       begin
         Post;
       end;
       if FDataSet IS TQuery then
       begin
         if TQuery(FDataSet).UpdatesPending then
           TQuery(FDataSet).Applyupdates;
       end
       else if FDataSet IS TClientDataSet then
       begin
         if TClientDataSet(FDataSet).ChangeCount>0 then
           TClientDataSet(FDataSet).Applyupdates(-1);
       end;
     end;
end;

{
procedure TfrmSTabMaintain.SetConstQryData(sFieldName,
  sFieldValue: String);
begin
    ConstFieldNameStrList.Add(sFieldName);
    ConstFieldValueStrList.Add(sFieldValue);
end;
}
procedure TfrmSTabMaintain.FormCreate(Sender: TObject);
var
   ii : Integer;
begin
//     nG_QryData := QRYWITHQUERY;
     nG_QryData := QRYWITHCDS;
     bCloseFrm := False;
     for ii:=0 to ComponentCount-1 do
     begin
       if Components[ii] IS TBitbtn then
       begin
         TBitbtn(Components[ii]).tag := TBitbtn(Components[ii]).font.color;
         TBitbtn(Components[ii]).OnMouseMove := MyMouseMove;
       end;
       if Components[ii] IS TDBEdit then
       begin
         TDBEdit(Components[ii]).OnEnter := MyEnter;
         TDBEdit(Components[ii]).OnExit := RealExit;         
       end;
       if Components[ii] IS TEdit then
       begin
         TEdit(Components[ii]).OnEnter := MyEnter;
         TEdit(Components[ii]).OnExit := RealExit;
       end;
       if Components[ii] IS TMaskEdit then
       begin
         TMaskEdit(Components[ii]).OnEnter := MyEnter;
         TMaskEdit(Components[ii]).OnExit := RealExit;
       end;
     end;
     UnEditedFieldNames := TStringList.Create;
     NoneDupFieldName := TStringList.Create;
     NoneDupFieldNameType := TStringList.Create;
{
    ConstFieldNameStrList := TStringList.Create;
    ConstFieldValueStrList := TStringList.Create;
}
end;

procedure TfrmSTabMaintain.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
{
    ConstFieldNameStrList.Free;
    ConstFieldValueStrList.Free;
}
end;

procedure TfrmSTabMaintain.DBGrid1DblClick(Sender: TObject);
begin
     btnEditClick(nil);
end;

procedure TfrmSTabMaintain.DBGrid1TitleClick(Column: TColumn);
var
   ii : Integer;
   bFlag : Boolean;
begin

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

procedure TfrmSTabMaintain.MyMouseMove(Sender: TObject; Shift: TShiftState;
  X, Y: Integer);
begin
     if TBitBtn(Sender).font.color = clRed then Exit;
     ReSetControls;
     TBitbtn(Sender).font.color := clRed;
     Screen.Cursor := crHandPoint;
end;

procedure TfrmSTabMaintain.ReSetControls;
var
   ii : Integer;
begin
     Screen.Cursor := crDefault;
     for ii:=0 to ComponentCount-1 do
     begin
       if Components[ii] IS TBitbtn then
         TBitbtn(Components[ii]).font.color := TBitbtn(Components[ii]).tag;
     end;
end;

procedure TfrmSTabMaintain.FormMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
     ReSetControls;
end;

procedure TfrmSTabMaintain.Panel2MouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
     ReSetControls;
end;

procedure TfrmSTabMaintain.MyEnter(Sender: TObject);
begin
    RealEnter(Sender);
end;


procedure TfrmSTabMaintain.SetQueryCond(sI_Desc, sI_RealFieldName: string; nI_Type:Integer);
var
   L_QryCondObj : TQryCondObj;
begin
   L_QryCondObj := TQryCondObj.Create;
   L_QryCondObj.s_Desc := sI_Desc;
   L_QryCondObj.s_FieldName := sI_RealFieldName;
   L_QryCondObj.n_Type := nI_Type;
   cobQueryCond.Items.AddObject(sI_Desc, L_QryCondObj);
   
end;

procedure TfrmSTabMaintain.FormDestroy(Sender: TObject);
var
   ii : Integer;
begin
     for ii:=0 to cobQueryCond.Items.Count-1 do
     begin
       cobQueryCond.Items.Objects[ii].Free;
     end;

     UnEditedFieldNames.Free;
end;

procedure TfrmSTabMaintain.SetDefaultEditFieldName(const Value: String);
begin
  FDefaultEditFieldName := Value;
end;

procedure TfrmSTabMaintain.SetDefaultFocus;
var
   ii : Integer;
begin
     if DefaultEditFieldName='' then Exit;
     for ii:=0 to ComponentCount-1 do
     begin
       if Components[ii] IS TDBEdit then
       begin
         if TDBEdit(Components[ii]).DataField = DefaultEditFieldName then
         begin
           TDBEdit(Components[ii]).SetFocus;
           break;
         end;
       end;
     end;
end;

procedure TfrmSTabMaintain.RealExit(Sender: TObject);
var
    nL_Ndx : Integer;
    sL_FieldName,  sL_Value, sL_FieldType : String;
begin
    if Sender IS TDBEdit then
    begin
      TDBEdit(Sender).Color := clWindow;
      if TDBEdit(Sender).DataSource.DataSet.state= dsInsert then
      begin
        sL_Value := TDBEdit(Sender).Text;
        if sL_Value<>'' then
        begin
          sL_FieldName := TDBEdit(Sender).DataField;

          nL_Ndx := NoneDupFieldName.IndexOf(sL_FieldName);
          if nL_Ndx<>-1 then
          begin
            sL_FieldType := NoneDupFieldNameType.Strings[nL_Ndx];
            if TdtmSTabMaintain(Dtm).CheckDup( ActionTabName, sL_FieldName, sL_Value, sL_FieldType) then
            begin
              MessageDlg('此欄位資料重覆!!',mtWarning,[mbOK],0);
              TDBEdit(Sender).setFocus;
              Exit;
            end;
          end;
        end;
      end;
    end;
    if Sender IS TEdit then
      TEdit(Sender).Color := clWindow;

end;

procedure TfrmSTabMaintain.RealEnter(Sender: TObject);
begin
    if Sender IS TDBEdit then
    begin
      TDBEdit(Sender).Color := clAqua;
    end;
    if Sender IS TEdit then
    begin
      TEdit(Sender).Color := clAqua;
    end;

end;

procedure TfrmSTabMaintain.cobQueryCondChange(Sender: TObject);
begin
     edtQryString.SetFocus;
end;

procedure TfrmSTabMaintain.PreAppendData;
begin
     FDataSet.Append;
     ChangeState;
     SetDefaultFocus;

end;

procedure TfrmSTabMaintain.TrimStr(var sI_Str: String);
begin
    sI_Str := AnsiQuotedStr(sI_Str,'''');
    Delete(sI_Str,1,1);
    Delete(sI_Str,length(sI_Str),1);
end;

procedure TfrmSTabMaintain.SetUnEditedFieldNames(const Value: TStringList);
begin
  FUnEditedFieldNames := Value;
end;

procedure TfrmSTabMaintain.SetReadonlyFields(bI_ReadOnly: Boolean);
var
    ii : Integer;
begin

     for ii:=0 to UnEditedFieldNames.Count-1 do
     begin
       dsrSTabMaintain.DataSet.FieldByName(UnEditedFieldNames.Strings[ii]).ReadOnly := bI_ReadOnly;
     end;
end;

procedure TfrmSTabMaintain.SetActionTabName(const Value: String);
begin
  FActionTabName := Value;
end;


procedure TfrmSTabMaintain.btnPrintClick(Sender: TObject);
begin
  FReport := TrptBasic.Create(Application);
  FReport.Title := Caption + '報表';
  FReport.SourceDataSet := FDataSet ;
  FReport.Preview ;
  FReport.Free ;
end;

procedure TfrmSTabMaintain.SetNoneDupFieldName(const Value: TStringList);
begin
  FNoneDupFieldName := Value;
end;

procedure TfrmSTabMaintain.SetNoneDupFieldNameType(
  const Value: TStringList);
begin
  FNoneDupFieldNameType := Value;
end;

end.
