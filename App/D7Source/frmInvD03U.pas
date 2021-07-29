unit frmInvD03U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, DB, Buttons, Grids, DBGrids, Mask,
  ADOInt, ADODB, DBClient;

type
  TfrmInvD03 = class(TForm)
    Panel1: TPanel;
    Label5: TLabel;
    pnlEdit: TPanel;
    Label8: TLabel;
    Label9: TLabel;
    Label11: TLabel;
    Label23: TLabel;
    medItemId: TMaskEdit;
    medDesc: TMaskEdit;
    cobIsSelfCreated: TComboBox;
    pnlGrid: TPanel;
    DBGrid1: TDBGrid;
    Panel4: TPanel;
    btnAppend: TBitBtn;
    btnEdit: TBitBtn;
    btnDelete: TBitBtn;
    btnCancel: TBitBtn;
    btnSave: TBitBtn;
    btnExit: TBitBtn;
    dsrInv005: TDataSource;
    cobTaxType: TComboBox;
    Label1: TLabel;
    cobSign: TComboBox;
    Stt_Show: TStaticText;
    Label2: TLabel;
    medItemIdRef: TComboBox;
    cdsInv005: TClientDataSet;
    procedure FormShow(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure btnAppendClick(Sender: TObject);
    procedure btnEditClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure cdsInv005AfterScroll(DataSet: TDataSet);
  private
    { Private declarations }
    FSign: String;
    procedure ChangeBtnStatus;
    procedure ChangeEditorSattus(const aState: TDataSetState);
    procedure setButtonCompetence;
    procedure initialForm;
    function isDataOK : Boolean;
    procedure PreparedInv005DataSet;
    procedure LoadItemIdRef;
    procedure LocateItemIdRef(const aCode: String);
    procedure GetTaxInfo(var aCode, aName: String);
    procedure GetSelfCreateText(var aSelfCreate: String);
    function Delete: Boolean;
    function Update: Boolean;
    function Insert: Boolean;
    function GetItemIdRef: String;
    function LoadData: Integer;
    function GetSign: String;
  public
    { Public declarations }
  end;

var
  frmInvD03: TfrmInvD03;

implementation

uses frmMainU, dtmMainJU, dtmMainU, cbUtilis;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD03.FormShow(Sender: TObject);
begin
  self.Caption := frmMain.GetFormTitleString( 'D03', '發票品名代碼資料維護' );
  LoadItemIdRef;
  LoadData;
  ChangeBtnStatus;
  setButtonCompetence;
  initialForm;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD03.ChangeBtnStatus;
begin
     with cdsInv005 do
     begin
       if state in [dsInactive] then
         Exit
       else if state in [dsEdit, dsInsert] then
       begin
          btnCancel.Enabled := TRUE;
          btnSave.Enabled := TRUE;
          btnDelete.Enabled := FALSE;
          btnEdit.Enabled := FALSE;
          btnExit.Enabled := FALSE;
          btnAppend.Enabled := FALSE;
          pnlGrid.Enabled := false;
          pnlEdit.Enabled := true;
       end
       else
       begin
          if cdsInv005.RecordCount>0 then
          begin
            btnAppend.Enabled :=TRUE;
            btnEdit.Enabled := TRUE;
            btnDelete.Enabled := TRUE;
          end
          else
          begin
            btnEdit.Enabled := FALSE;
            btnDelete.Enabled := FALSE;
          end;
          btnCancel.Enabled := FALSE;
          btnSave.Enabled := FALSE;
          btnAppend.Enabled := TRUE;
          btnExit.Enabled := TRUE;
          pnlGrid.Enabled := true;
          pnlEdit.Enabled := FALSE;
       end;
     end;

end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD03.ChangeEditorSattus(const aState: TDataSetState);
begin
  case aState of
    dsEdit:
      begin
        medItemId.Enabled := False;
        if ( cdsInv005.FieldByName( 'ISSELFCREATED' ).AsString = 'N' ) then
        begin
          medDesc.Enabled := False;
          cobSign.Enabled := False;
          cobTaxType.Enabled := False;
          cobIsSelfCreated.Enabled := False;
          medItemIdRef.Enabled := True;
        end else
        begin
          medDesc.Enabled := True;
          cobSign.Enabled := True;
          cobTaxType.Enabled := True;
          cobIsSelfCreated.Enabled := True;
          medItemIdRef.Enabled := True;
        end;
      end;
   else
      medItemId.Enabled := True;
      medDesc.Enabled := True;
      cobSign.Enabled := True;
      cobTaxType.Enabled := True;
      cobIsSelfCreated.Enabled := True;
      medItemIdRef.Enabled := True;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD03.initialForm;
begin
  cdsInv005.FindField('ItemId').DisplayLabel := '品名代碼';
  cdsInv005.FindField('Description').DisplayLabel := '品名';
  cdsInv005.FindField('TaxName').DisplayLabel := '稅率名稱';
  cdsInv005.FindField('UptTime').DisplayLabel := '異動時間';
  cdsInv005.FindField('UptEn').DisplayLabel := '異動人員';
  cdsInv005.FindField('ItemIdRef').DisplayLabel := '合併品名代碼';
  cdsInv005.FindField('Description2').DisplayLabel := '合併品名';
  {}
  cdsInv005.FindField('Description').DisplayWidth := 30;
  cdsInv005.FindField('Description2').DisplayWidth := 30;
  cdsInv005.FindField('UptEn').DisplayWidth := 10;
  {}
  cdsInv005.FindField('IdentifyId1').Visible := False;
  cdsInv005.FindField('IdentifyId2').Visible := False;
  cdsInv005.FindField('Sign').Visible := False;
  cdsInv005.FindField('IsSelfCreated').Visible := False;
  cdsInv005.FindField('TaxCode').Visible := False;
  cdsInv005.FindField('CompId').Visible := False;
  {}
  btnEdit.Enabled := ( cdsInv005.RecordCount > 0 );
  btnDelete.Enabled := ( cdsInv005.RecordCount > 0 );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD03.setButtonCompetence;
var
    sL_QueryCompetence,sL_InsertCompetence,sL_DeleteCompetence,sL_UpdateCompetence,sL_ExecuteCompetence : String;
    sL_VerifyCompetence : String;
begin
     //設定權限
     dtmMain.getCompetence('D03',sL_QueryCompetence,sL_InsertCompetence,sL_DeleteCompetence,
            sL_UpdateCompetence,sL_ExecuteCompetence,sL_VerifyCompetence);

     if sL_InsertCompetence = 'Y' then
       btnAppend.Enabled := true
     else
       btnAppend.Enabled := false;


     if sL_DeleteCompetence = 'Y' then
       btnDelete.Enabled := true
     else
       btnDelete.Enabled := false;


     if sL_UpdateCompetence = 'Y' then
       btnEdit.Enabled := true
     else
       btnEdit.Enabled := false;


end;

procedure TfrmInvD03.cdsInv005AfterScroll(DataSet: TDataSet);
var
    sL_Sign,sL_IsSelfCreated,sL_TaxType, aItemIdRef : String;
begin
    with dsrInv005.DataSet do
    begin
      if RecordCount > 0 then
      begin
        medItemId.Text := FieldByName('ItemId').AsString;
        medDesc.Text := FieldByName('Description').AsString;
        sL_Sign := FieldByName('Sign').AsString;
        sL_IsSelfCreated := FieldByName('IsSelfCreated').AsString;
        sL_TaxType := FieldByName('TaxCode').AsString;
        aItemIdRef := FieldByName( 'ItemIdRef' ).AsString;

        if sL_Sign = '+' then
          cobSign.ItemIndex := 0
        else
          cobSign.ItemIndex := 1;

        if sL_TaxType = '1' then
          cobTaxType.ItemIndex := 0
        else if sL_TaxType = '2' then
          cobTaxType.ItemIndex := 1
        else if sL_TaxType = '3' then
          cobTaxType.ItemIndex := 2
        else if sL_TaxType = '4' then
          cobTaxType.ItemIndex := 3
        else if sL_TaxType = '5' then
          cobTaxType.ItemIndex := 4
        else if sL_TaxType = '6' then
          cobTaxType.ItemIndex := 5;


        if sL_IsSelfCreated = 'Y' then
          cobIsSelfCreated.ItemIndex := 0
        else
          cobIsSelfCreated.ItemIndex := 1;

        LocateItemIdRef( aItemIdRef );

      end
      else
      begin
        medItemId.Text := '';
        medDesc.Text := '';
        cobSign.ItemIndex := 0;
        cobTaxType.ItemIndex := 0;
        cobIsSelfCreated.ItemIndex := 0;
        medItemIdRef.ItemIndex := 0;
      end;
      Stt_Show.Caption := inttostr(RecNo)+'　/　'+inttostr(recordcount);      
    end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD03.btnExitClick(Sender: TObject);
begin
  if not VarIsNull( cdsInv005.Data ) then
    cdsInv005.EmptyDataSet;
  cdsInv005.Data := Null;
  Self.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD03.btnAppendClick(Sender: TObject);
begin
  cdsInv005.Append;
  cobSign.ItemIndex := 0;
  cobTaxType.ItemIndex := 0;
  cobIsSelfCreated.ItemIndex := 0;
  ChangeBtnStatus;
  ChangeEditorSattus( dsInsert );
  medItemId.SetFocus;
  FSign := EmptyStr;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD03.btnEditClick(Sender: TObject);
begin
  cdsInv005.Edit;
  ChangeBtnStatus;
  ChangeEditorSattus( dsEdit );
  FSign := GetSign;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD03.btnDeleteClick(Sender: TObject);
begin
  if ConfirmMsg( '請確認要刪除此筆資料?' ) then
  begin
    Self.Delete;
    cdsInv005.Delete;
    ChangeBtnStatus;
  end;
  setButtonCompetence;
  initialForm;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD03.btnCancelClick(Sender: TObject);
begin
  if ( cdsInv005.State in [dsEdit] ) then
    medItemId.Enabled := True;
  cdsInv005.Cancel;
  cdsInv005.First;
  ChangeBtnStatus;
  ChangeEditorSattus( dsBrowse );
  setButtonCompetence;
  initialForm;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvD03.isDataOK: Boolean;
var
    sL_ItemId,sL_Description, aItemIdRef, aMsg: String;
    aSign, aTaxCode, aTaxName: String;
begin
    Result := true;
    sL_ItemId := Trim(medItemId.Text);
    sL_Description := Trim(medDesc.Text);

    if sL_ItemId = '' then
    begin
      WarningMsg( '請輸入品名代碼。' );
      medItemId.SetFocus;
      Result := false;
      Exit;
    end;

    if sL_Description = '' then
    begin
      MessageDlg('請輸入品名',mtError,[mbOk],0);
      medDesc.SetFocus;
      Result := false;
      Exit;
    end;

    if dsrInv005.DataSet.State in [dsInsert,dsEdit] then
    begin
      if (cobTaxType.ItemIndex=3) or (cobTaxType.ItemIndex=4) or (cobTaxType.ItemIndex=5) then
      begin
        WarningMsg('稅率別僅能選擇:[1-應稅],[2-零稅率],[3-免稅]。' );
        cobTaxType.SetFocus;
        Result := false;
        Exit;
      end;
      {}
      aItemIdRef := GetItemIdRef;
      if ( aItemIdRef <> EmptyStr ) and
         ( aItemIdRef = Trim( medItemId.Text ) ) then
      begin
        WarningMsg( '合併品名代碼, 不可與原品名代碼相同。' );
        if medItemIdRef.CanFocus then medItemIdRef.SetFocus;
        Result := False;
        Exit;
      end;
      {}
      if ( aItemIdRef <> EmptyStr ) then
      begin
      
          aSign := GetSign;
          GetTaxInfo( aTaxCode, aTaxName );

          dtmMain.adoComm.Close;
          dtmMain.adoComm.SQL.Text := Format(
            ' SELECT A.SIGN, A.TAXCODE FROM INV005 A ' +
            '  WHERE A.IDENTIFYID1 = ''%s''          ' +
            '    AND A.IDENTIFYID2 = ''%s''          ' +
            '    AND A.COMPID = ''%s''               ' +
            '    AND A.ITEMID = ''%s''               ',
            [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID, aItemIdRef] );
          dtmMain.adoComm.Open;
          if ( dtmMain.adoComm.FieldByName( 'TAXCODE' ).AsString <> aTaxCode ) then
          begin
            aMsg := ( aMsg + '該合併品名代碼, 稅別不同。'#13 );
          end;
          dtmMain.adoComm.Close;
          {}
          if ( aMsg <> EmptyStr ) then
          begin
            WarningMsg( aMsg );
            Result := False;
            Exit;
          end;
          
      end;
    end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD03.btnSaveClick(Sender: TObject);
var
    sL_ItemId,sL_Desc,sL_Sign,sL_TaxCode,sL_TaxName,sL_IsSelfCreated : String;
    sL_SQL : String;
    bL_HavePKValue : Boolean;
    aBookMark: TBookmark;
begin
//    aBookMark := nil;
    if isDataOK then
    begin
      sL_ItemId := Trim(medItemId.Text);
      sL_Desc := Trim(medDesc.Text);

      if cobSign.ItemIndex = 0 then
        sL_Sign := '+'
      else
        sL_Sign := '-';


      if cobTaxType.ItemIndex = 0 then
      begin
        sL_TaxCode := '1';
        sL_TaxName := '應稅';
      end
      else if cobTaxType.ItemIndex = 1 then
      begin
        sL_TaxCode := '2';
        sL_TaxName := '零稅率';
      end
      else if cobTaxType.ItemIndex = 2 then
      begin
        sL_TaxCode := '3';
        sL_TaxName := '免稅';
      end
      else if cobTaxType.ItemIndex = 3 then
      begin
        sL_TaxCode := '4';
        sL_TaxName := '印花稅';
      end
      else if cobTaxType.ItemIndex = 4 then
      begin
        sL_TaxCode := '5';
        sL_TaxName := '應稅-不拋';
      end
      else if cobTaxType.ItemIndex = 5 then
      begin
        sL_TaxCode := '6';
        sL_TaxName := '免稅-不拋';
      end;

      if cobIsSelfCreated.ItemIndex = 0 then
        sL_IsSelfCreated := 'Y'
      else
        sL_IsSelfCreated := 'N';



      if ( dsrInv005.DataSet.State in [dsInsert] ) then
      begin
        bL_HavePKValue := False;
        sL_SQL := Format(
          ' SELECT * FROM INV005               ' +
          '  WHERE IDENTIFYID1 = ''%s''        ' +
          '    AND IDENTIFYID2 = ''%s''        ' +
          '    AND COMPID = ''%s''             ' +
          '    AND ITEMID = ''%s''             ',
          [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID, sL_ItemId] );
          
        bL_HavePKValue := dtmMainJ.checkPK(sL_SQL);

        if bL_HavePKValue then
        begin
          WarningMsg( '違反唯一值條件。' );
          if medItemId.CanFocus then  medItemId.SetFocus;
          Exit;
        end;

      end;
      if (dsrInv005.DataSet.State in [dsEdit] ) then
      begin
        aBookMark := cdsInv005.GetBookmark;
      end;
      dtmMainJ.adoQryCommon.Close;
      dtmMainJ.adoQryCommon.SQL.Text := Format(
        ' SELECT * FROM INV005               ' +
        '  WHERE IDENTIFYID1 = ''%s''        ' +
        '    AND IDENTIFYID2 = ''%s''        ' +
        '    AND COMPID = ''%s''             ' +
        '    AND DESCRIPTION = ''%s''        ' +
        '    AND SIGN = ''%s''               ',
        [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID, sL_Desc, sL_Sign] );
        
      dtmMainJ.adoQryCommon.Open;

      if ( dsrInv005.DataSet.State in [dsInsert] ) then
      begin
        bL_HavePKValue := ( dtmMainJ.adoQryCommon.RecordCount > 0 );
      end else
      begin
        if ( FSign = sL_Sign ) then
          bL_HavePKValue := ( dtmMainJ.adoQryCommon.RecordCount > 1 )
        else
          bL_HavePKValue := ( dtmMainJ.adoQryCommon.RecordCount > 0 );
      end;

      dtmMainJ.adoQryCommon.Close;

      if ( bL_HavePKValue ) then
      begin
        WarningMsg( '收費項目品名/性質(正負項)不可重覆。' );
        if medDesc.CanFocus then medDesc.SetFocus;
        Exit;
      end;

      if cdsInv005.State in [dsInsert] then
        Self.Insert
      else if cdsInv005.State in [dsEdit] then
        Self.Update;

      LoadItemIdRef;
      LoadData;
      medItemId.Enabled := true;
      ChangeBtnStatus;
      ChangeEditorSattus( dsBrowse );
      setButtonCompetence;
      initialForm;

      try
        cdsInv005.GotoBookmark( aBookMark );
      finally
        cdsInv005.FreeBookmark( aBookMark );
      end;

      {
      if aBookMark <> nil then
      begin
        aBookMark
      end;
      }
    end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD03.LoadItemIdRef;
var
  aText: String;
begin
  medItemIdRef.Items.Clear;
  dtmMain.adoComm.Close;
  dtmMain.adoComm.SQL.Text := Format(
    ' SELECT A.ITEMID, A.DESCRIPTION,                            ' +
    '        DECODE( A.SIGN, ''-'', ''支出'', ''收入'' ) SIGN,   ' +
    '        DECODE( A.TAXCODE, ''1'', ''應稅'',                 ' +
    '                           ''2'', ''零稅率'',               ' +
    '                           ''3'', ''免稅'',                 ' +
    '                           ''4'', ''印花稅'',               ' +
    '                           ''5'', ''應稅-不拋'',            ' +
    '                           ''6'', ''免稅-不拋'' ) TAXNAME   ' +
    '   FROM INV005 A                                            ' +
    '  WHERE IDENTIFYID1 = ''%s''                                ' +
    '    AND IDENTIFYID2 = ''%s''                                ' +
    '    AND COMPID = ''%s''                                     ' +
    '  ORDER BY ITEMID                                           ',
    [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID] );
  dtmMain.adoComm.Open;
  dtmMain.adoComm.First;
  while not dtmMain.adoComm.Eof do
  begin
    aText := Format( '%s'#1'%s'#1'%s'#1'%s', [
      Rpad( dtmMain.adoComm.FieldByName( 'ITEMID' ).AsString, 8, #32 ),
      Rpad( dtmMain.adoComm.FieldByName( 'SIGN' ).AsString, 6, #32 ),
      Rpad( dtmMain.adoComm.FieldByName( 'TAXNAME' ).AsString, 10, #32 ),
      dtmMain.adoComm.FieldByName( 'DESCRIPTION' ).AsString] );
    medItemIdRef.Items.Add( aText );
    dtmMain.adoComm.Next;
  end;
  dtmMain.adoComm.Close;
  medItemIdRef.Items.Insert( 0, EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD03.LocateItemIdRef(const aCode: String);
var
  aText: String;
  aIndex: Integer;
  aFound: Boolean;
begin
  aFound := False;
  for aIndex := 0 to medItemIdRef.Items.Count - 1 do
  begin
    aText := Trim( medItemIdRef.Items[aIndex] );
    if ( aText = EmptyStr ) then Continue;
    aFound := ( aCode = Trim( Copy( aText, 1, Pos( #1, aText ) - 1 ) ) );
    if aFound then
    begin
      medItemIdRef.ItemIndex := aIndex;
      Break;
    end;
  end;
  if not aFound then medItemIdRef.ItemIndex := 0;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvD03.GetItemIdRef: String;
var
  aText: String;
begin
  Result := EmptyStr;
  aText := Trim( medItemIdRef.Text );
  if ( aText = EmptyStr ) then Exit;
  Result := Trim( Copy( aText, 1, Pos( #1, aText ) - 1 ) );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD03.PreparedInv005DataSet;
begin
  if ( cdsInv005.FieldDefs.Count <= 0 ) then
  begin
    cdsInv005.FieldDefs.Add( 'IDENTIFYID1', ftString, 2 );
    cdsInv005.FieldDefs.Add( 'IDENTIFYID2', ftInteger );
    cdsInv005.FieldDefs.Add( 'ITEMID', ftInteger );
    cdsInv005.FieldDefs.Add( 'DESCRIPTION', ftString, 40 );
    cdsInv005.FieldDefs.Add( 'ITEMIDREF', ftInteger );
    cdsInv005.FieldDefs.Add( 'DESCRIPTION2', ftString, 40 );
    cdsInv005.FieldDefs.Add( 'TAXNAME', ftString, 12 );
    cdsInv005.FieldDefs.Add( 'SIGN', ftString, 1 );
    cdsInv005.FieldDefs.Add( 'ISSELFCREATED', ftString, 1 );
    cdsInv005.FieldDefs.Add( 'UPTTIME', ftDateTime );
    cdsInv005.FieldDefs.Add( 'UPTEN', ftString, 20 );
    cdsInv005.FieldDefs.Add( 'TAXCODE', ftInteger );
    cdsInv005.FieldDefs.Add( 'COMPID', ftString, 6 );
  end;
  if not VarIsNull( cdsInv005.Data ) then
    cdsInv005.EmptyDataSet
  else
    cdsInv005.CreateDataSet;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvD03.LoadData: Integer;
var
  aDataSet: TADOQuery;
begin
  cdsInv005.AfterScroll := nil;
  try
    PreparedInv005DataSet;
    aDataSet := dtmMain.adoComm;
    aDataSet.Close;
    aDataSet.SQL.Text := Format(
      ' SELECT A.IDENTIFYID1, A.IDENTIFYID2,                ' +
      '        A.ITEMID, A.DESCRIPTION, A.ITEMIDREF,        ' +
      '        B.DESCRIPTION AS DESCRIPTION2, A.TAXNAME,    ' +
      '        A.SIGN, A.ISSELFCREATED, A.UPTTIME,          ' +
      '        A.UPTEN, A.TAXCODE, A.COMPID                 ' +
      '   FROM INV005 A, INV005 B                           ' +
      '   WHERE A.IDENTIFYID1 = B.IDENTIFYID1(+)            ' +
      '     AND A.IDENTIFYID2 = B.IDENTIFYID2(+)            ' +
      '     AND A.COMPID = B.COMPID(+)                      ' +
      '     AND A.ITEMIDREF = B.ITEMID(+)                   ' +
      '     AND A.IDENTIFYID1 = ''%s''                      ' +
      '     AND A.IDENTIFYID2 = ''%s''                      ' +
      '     AND A.COMPID = ''%s''                           ' +
      '  ORDER BY A.ITEMID                                  ',
      [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID] );
    aDataSet.Open;
    aDataSet.First;
    cdsInv005.DisableControls;
    try
      while not aDataSet.Eof do
      begin
        cdsInv005.Append;
        cdsInv005.FieldByName( 'IDENTIFYID1' ).AsString := aDataSet.FieldByName( 'IDENTIFYID1' ).AsString;
        cdsInv005.FieldByName( 'IDENTIFYID2' ).AsString := aDataSet.FieldByName( 'IDENTIFYID2' ).AsString;
        cdsInv005.FieldByName( 'ITEMID' ).AsString := aDataSet.FieldByName( 'ITEMID' ).AsString;
        cdsInv005.FieldByName( 'DESCRIPTION' ).AsString := aDataSet.FieldByName( 'DESCRIPTION' ).AsString;
        cdsInv005.FieldByName( 'ITEMIDREF' ).AsString := aDataSet.FieldByName( 'ITEMIDREF' ).AsString;
        cdsInv005.FieldByName( 'DESCRIPTION2' ).AsString := aDataSet.FieldByName( 'DESCRIPTION2' ).AsString;
        cdsInv005.FieldByName( 'TAXNAME' ).AsString := aDataSet.FieldByName( 'TAXNAME' ).AsString;
        cdsInv005.FieldByName( 'SIGN' ).AsString := aDataSet.FieldByName( 'SIGN' ).AsString;
        cdsInv005.FieldByName( 'ISSELFCREATED' ).AsString := aDataSet.FieldByName( 'ISSELFCREATED' ).AsString;
        if not VarIsNull( aDataSet.FieldByName( 'UPTTIME' ).Value ) then
          cdsInv005.FieldByName( 'UPTTIME' ).Value := aDataSet.FieldByName( 'UPTTIME' ).Value;
        cdsInv005.FieldByName( 'UPTEN' ).AsString := aDataSet.FieldByName( 'UPTEN' ).AsString;
        cdsInv005.FieldByName( 'TAXCODE' ).AsString := aDataSet.FieldByName( 'TAXCODE' ).AsString;
        cdsInv005.FieldByName( 'COMPID' ).AsString := aDataSet.FieldByName( 'COMPID' ).AsString;
        cdsInv005.Post;
        aDataSet.Next;
      end;
      cdsInv005.First;
      aDataSet.Close;
    finally
      cdsInv005.EnableControls;
    end;
  finally
    cdsInv005.AfterScroll := cdsInv005AfterScroll;
  end;
  cdsInv005AfterScroll( cdsInv005 );  
  Result := cdsInv005.RecordCount;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD03.GetTaxInfo(var aCode, aName: String);
begin
  if cobTaxType.ItemIndex = 0 then
  begin
    aCode := '1';
    aName := '應稅';
  end
  else if cobTaxType.ItemIndex = 1 then
  begin
    aCode := '2';
    aName := '零稅率';
  end
  else if cobTaxType.ItemIndex = 2 then
  begin
    aCode := '3';
    aName := '免稅';
  end
  else if cobTaxType.ItemIndex = 3 then
  begin
    aCode := '4';
    aName := '印花稅';
  end
  else if cobTaxType.ItemIndex = 4 then
  begin
    aCode := '5';
    aName := '應稅-不拋';
  end
  else if cobTaxType.ItemIndex = 5 then
  begin
    aCode := '6';
    aName := '免稅-不拋';
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD03.GetSelfCreateText(var aSelfCreate: String);
begin
  aSelfCreate := 'N';
  if ( cobIsSelfCreated.ItemIndex = 0 ) then aSelfCreate := 'Y';
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvD03.Delete: Boolean;
begin
  Result := True;
  dtmMain.adoComm.Close;
  dtmMain.adoComm.SQL.Text := Format(
    ' DELETE FROM INV005 WHERE IDENTIFYID1 = ''%s''    ' +
    '   AND IDENTIFYID2 = ''%s'' AND COMPID = ''%s''   ' +
    '   AND ITEMID = ''%s''                            ',
    [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID,
     cdsInv005.FieldByName( 'ITEMID' ).AsString,
     cdsInv005.FieldByName( 'DESCRIPTION' ).AsString] );
  dtmMain.adoComm.ExecSQL;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvD03.Insert: Boolean;
var
  aCreate, aTaxCode, aTaxName, aItemIdRef, aSign: String;
begin
  Result := True;
  {}
  GetSelfCreateText( aCreate );
  GetTaxInfo( aTaxCode, aTaxName );
  aItemIdRef := GetItemIdRef;
  aSign := GetSign;
  {}
  dtmMain.adoComm.Close;
  dtmMain.adoComm.SQL.Text := Format(
    ' INSERT INTO INV005 ( IDENTIFYID1, IDENTIFYID2, ITEMID,      ' +
    '    DESCRIPTION, SIGN, ISSELFCREATED, UPTTIME, UPTEN,        ' +
    '    TAXCODE, TAXNAME, ITEMIDREF, COMPID )                    ' +
    ' VALUES ( ''%s'', ''%s'',  ''%s'',                           ' +
    '    ''%s'', ''%s'',  ''%s'', SYSDATE, ''%s'',                ' +
    '    ''%s'', ''%s'', ''%s'', ''%s''  )                        ',
    [IDENTIFYID1, IDENTIFYID2, Trim( medItemId.Text ),
     medDesc.Text, aSign, aCreate, dtmMain.getLoginUser,
     aTaxCode, aTaxName, aItemIdRef, dtmMain.getCompID] );
  dtmMain.adoComm.ExecSQL;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvD03.Update: Boolean;
var
  aSql1, aSql2, aCreate, aTaxCode, aTaxName, aItemIdRef, aSign: String;
begin
  Result := True;
  { 由發票系統見建立的收費項目, 可以更新所有欄位 }
  aSql1 :=
    ' UPDATE INV005                       ' +
    '    SET ITEMID = ''%s'',             ' +
    '        DESCRIPTION = ''%s'',        ' +
    '        SIGN = ''%s'',               ' +
    '        ISSELFCREATED = ''%s'',      ' +
    '        UPTTIME = SYSDATE,           ' +
    '        UPTEN = ''%s'',              ' +
    '        TAXCODE = ''%s'',            ' +
    '        TAXNAME = ''%s'',            ' +
    '        ITEMIDREF = ''%s''           ' +
    '  WHERE IDENTIFYID1 = ''%s''         ' +
    '    AND IDENTIFYID2 = ''%s''         ' +
    '    AND COMPID = ''%s''              ' +
    '    AND ITEMID = ''%s''              ';
  { 非由發票系統建立的收費項目(由客服匯進來), 只可更新合併品名 }
  aSql2 :=
    ' UPDATE INV005                       ' +
    '    SET UPTTIME = SYSDATE,           ' +
    '        UPTEN = ''%s'',              ' +
    '        ITEMIDREF = ''%s''           ' +
    '  WHERE IDENTIFYID1 = ''%s''         ' +
    '    AND IDENTIFYID2 = ''%s''         ' +
    '    AND COMPID = ''%s''              ' +
    '    AND ITEMID = ''%s''              ';
  {}
  GetSelfCreateText( aCreate );
  GetTaxInfo( aTaxCode, aTaxName );
  aItemIdRef := GetItemIdRef;
  aSign := GetSign;
  {}
  dtmMain.adoComm.Close;
  if ( cdsInv005.FieldByName( 'ISSELFCREATED' ).AsString = 'Y' ) then
  begin
    dtmMain.adoComm.SQL.Text := Format( aSql1,
      [Trim( medItemId.Text ), medDesc.Text, aSign, aCreate, dtmMain.getLoginUser,
       aTaxCode, aTaxName, aItemIdRef, IDENTIFYID1, IDENTIFYID2,
       dtmMain.getCompID,
       cdsInv005.FieldByName( 'ITEMID' ).AsString] );
  end else
  begin
    dtmMain.adoComm.SQL.Text := Format( aSql2,
      [dtmMain.getLoginUser,  aItemIdRef, IDENTIFYID1, IDENTIFYID2,
       dtmMain.getCompID, 
       cdsInv005.FieldByName( 'ITEMID' ).AsString] );
  end;
  dtmMain.adoComm.ExecSQL;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvD03.GetSign: String;
begin
  Result := '-';
  if cobSign.ItemIndex = 0 then Result := '+';
end;

{ ---------------------------------------------------------------------------- }

end.
