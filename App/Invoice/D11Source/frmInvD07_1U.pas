unit frmInvD07_1U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, StdCtrls, Buttons, Grids, DBGrids, Mask, ExtCtrls, fraYMU,
  fraYMDU;

type
  TfrmInvD07_1 = class(TForm)
    pnlEdit: TPanel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    medStartNum: TMaskEdit;
    pnlGrid: TPanel;
    DBGrid1: TDBGrid;
    Panel4: TPanel;
    btnAppend: TBitBtn;
    btnEdit: TBitBtn;
    btnDelete: TBitBtn;
    btnCancel: TBitBtn;
    btnSave: TBitBtn;
    btnExit: TBitBtn;
    dsrInv099: TDataSource;
    fraYM: TfraYM;
    medPrefix: TMaskEdit;
    Label1: TLabel;
    medEndNum: TMaskEdit;
    medCurNum: TMaskEdit;
    fraLastInvDate: TfraYMD;
    Label7: TLabel;
    cobInvFormat: TComboBox;
    Label2: TLabel;
    cobUseful: TComboBox;
    Label3: TLabel;
    edtMemo: TEdit;
    Panel1: TPanel;
    Label5: TLabel;
    Stt_Show: TStaticText;
    btnRefresh: TBitBtn;
    procedure btnExitClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure btnAppendClick(Sender: TObject);
    procedure btnEditClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure dsrInv099DataChange(Sender: TObject; Field: TField);
    procedure btnSaveClick(Sender: TObject);
    procedure medPrefixExit(Sender: TObject);
    procedure medStartNumExit(Sender: TObject);
    procedure fraYMmseYMExit(Sender: TObject);
    procedure fraYMExit(Sender: TObject);
    procedure medEndNumExit(Sender: TObject);
    procedure btnRefreshClick(Sender: TObject);
  private
    { Private declarations }
    FIsUse: Boolean;
    procedure ChangeBtnStatus;
    procedure setButtonCompetence;
    procedure initialForm;
    function isDataOK : Boolean;
    function CheckBeforeDelete: Boolean;
  public
    { Public declarations }
    sG_CompID,sG_QueryYear,sG_QueryYM : String;
  end;

var
  frmInvD07_1: TfrmInvD07_1;

implementation

uses cbUtilis, dtmMainJU, frmMainU, dtmMainU, Ustru;

{$R *.dfm}

procedure TfrmInvD07_1.btnExitClick(Sender: TObject);
begin
    Close;
end;

procedure TfrmInvD07_1.ChangeBtnStatus;
begin
     with dsrInv099.DataSet do
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
          btnRefresh.Enabled := False;
       end
       else
       begin
          if dsrInv099.DataSet.RecordCount>0 then
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
          btnRefresh.Enabled := True;
       end;
     end;

end;

procedure TfrmInvD07_1.FormShow(Sender: TObject);
begin
  Self.Caption := frmMain.GetFormTitleString( 'D07_1', '?o???r?y???@' );
  ChangeBtnStatus;
  setButtonCompetence;
  initialForm;
  medCurNum.Enabled := false;
  fraLastInvDate.mseYMD.Enabled := false;
  FIsUse := False;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD07_1.initialForm;
begin
  dsrInv099.DataSet.FieldByName('YEARMONTH').DisplayLabel := '?????~??';
  dsrInv099.DataSet.FieldByName('PREFIX').DisplayLabel := '?r?y';
  dsrInv099.DataSet.FieldByName('INVFORMAT').DisplayLabel := '?o??????';
  dsrInv099.DataSet.FieldByName('STARTNUM').DisplayLabel := '?_?l???X';
  dsrInv099.DataSet.FieldByName('STARTNUM').DisplayWidth := 15;
  dsrInv099.DataSet.FieldByName('ENDNUM').DisplayLabel := '???????X';
  dsrInv099.DataSet.FieldByName('ENDNUM').DisplayWidth := 15;
  dsrInv099.DataSet.FieldByName('CURNUM').DisplayLabel := '?U?@?????s??';
  dsrInv099.DataSet.FieldByName('LASTINVDATE').DisplayLabel := '?o???????}??????';
  dsrInv099.DataSet.FieldByName('LASTINVDATE').DisplayWidth := 15;
  dsrInv099.DataSet.FieldByName('CURNUM').DisplayWidth := 15;
  dsrInv099.DataSet.FieldByName('USEFUL').DisplayLabel := '?O?_?i??';
  dsrInv099.DataSet.FieldByName('MEMO').DisplayLabel := '????';
  dsrInv099.DataSet.FieldByName('IDENTIFYID1').Visible := false;
  dsrInv099.DataSet.FieldByName('IDENTIFYID2').Visible := false;
  dsrInv099.DataSet.FieldByName('COMPID').Visible := false;
  dsrInv099.DataSet.FieldByName('UPTEN').DisplayWidth := 10;
  dsrInv099.DataSet.FieldByName('UPTEN').DisplayLabel := '?????H??';
  dsrInv099.DataSet.FieldByName('UPTTIME').Visible := false;
  if dsrInv099.DataSet.RecordCount = 0 then
  begin
    btnEdit.Enabled := false;
    btnDelete.Enabled := false;
  end;
  medStartNum.Enabled := True;
  cobInvFormat.Enabled := True;

end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD07_1.setButtonCompetence;
var
    sL_QueryCompetence,sL_InsertCompetence,sL_DeleteCompetence,sL_UpdateCompetence,sL_ExecuteCompetence : String;
begin
     //?]?w?v??
     dtmMain.getCompetence('D07',sL_QueryCompetence,sL_InsertCompetence,sL_DeleteCompetence,sL_UpdateCompetence,sL_ExecuteCompetence);

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

procedure TfrmInvD07_1.btnAppendClick(Sender: TObject);
begin
    dsrInv099.DataSet.Append;

    cobUseful.ItemIndex := 0;
    cobInvFormat.ItemIndex := 0;

    ChangeBtnStatus;
    fraYM.mseYM.SetFocus;;
end;

procedure TfrmInvD07_1.btnEditClick(Sender: TObject);
var
  aYearMonth: String;
begin
  aYearMonth := fraYM.mseYM.Text;
  if dtmMainJ.InvHasBeenLock( aYearMonth ) then
  begin
    WarningMsg( Format( '???i?o???????~??: %s, ?w?g???b,?L?k?????C',
      [FormatMaskText( '####/##;0;_', aYearMonth )] ) );
    Exit;
  end;
  dsrInv099.DataSet.Edit;
  ChangeBtnStatus;
  fraYM.mseYM.Enabled := false;
  FIsUse := False;
  IF dsrInv099.DataSet.FieldByName('STARTNUM').AsString <> dsrInv099.DataSet.FieldByName('CURNUM').AsString then
  begin
    medStartNum.Enabled := False;
    cobInvFormat.Enabled := False;
    FIsUse := True;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD07_1.btnDeleteClick(Sender: TObject);
begin
  if not CheckBeforeDelete then Exit;
  if ConfirmMsg( '???T?{?n?R???????????' ) then
  begin
    dsrInv099.DataSet.Delete;
    dtmMainJ.getInv099Data(sG_CompID,sG_QueryYear,sG_QueryYM);
    ChangeBtnStatus;
  end;
  setButtonCompetence;
  initialForm;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD07_1.btnCancelClick(Sender: TObject);
begin
    if  dsrInv099.DataSet.State in [dsEdit] then
    begin
      fraYM.mseYM.Enabled := true;
    end;

    dtmMainJ.adoInv099Code.Cancel;
    dtmMainJ.adoInv099Code.First;
    ChangeBtnStatus;

    setButtonCompetence;
    initialForm;

end;

procedure TfrmInvD07_1.dsrInv099DataChange(Sender: TObject; Field: TField);
var
    sL_InvFormat,sL_LastInvDate,sL_Useful,sL_YearMonth : String;
begin
    with dsrInv099.DataSet do
    begin
      if RecordCount > 0 then
      begin
        sL_YearMonth := FieldByName('YearMonth').AsString;
        sL_YearMonth := Copy(sL_YearMonth,1,4) + '/' + Copy(sL_YearMonth,5,2);
        fraYM.setYM(sL_YearMonth);

        medPrefix.Text := FieldByName('Prefix').AsString;
        medStartNum.Text := FieldByName('StartNum').AsString;
        medEndNum.Text := FieldByName('EndNum').AsString;
        medCurNum.Text := FieldByName('CurNum').AsString;
        edtMemo.Text := FieldByName('Memo').AsString;
        sL_LastInvDate := FieldByName('LastInvDate').AsString;
        fraLastInvDate.setYMD( sL_LastInvDate );

        sL_InvFormat := FieldByName('InvFormat').AsString;
        sL_Useful := FieldByName('Useful').AsString;

        if sL_InvFormat = '1' then  //1:?q?l 2:???G(Default)3:???T
          cobInvFormat.ItemIndex := 0
        else if sL_InvFormat = '2' then
          cobInvFormat.ItemIndex := 1
        else if sL_InvFormat = '3' then
          cobInvFormat.ItemIndex := 2;

        if sL_Useful = 'Y' then  //Y:?O   N:?_
          cobUseful.ItemIndex := 0
        else if sL_Useful = 'N' then
          cobUseful.ItemIndex := 1;
        Stt_Show.Caption := inttostr(RecNo)+'?@/?@'+inttostr(recordcount);          
      end
      else
      begin
        fraYM.setYM('');
        medPrefix.Text := '';
        medStartNum.Text := '';

        medEndNum.Text := '';
        medCurNum.Text := '';
        edtMemo.Text := '';
        fraLastInvDate.setYMD('');

        cobInvFormat.ItemIndex := 1;
        cobUseful.ItemIndex := 0;
      end;
    end;
end;

function TfrmInvD07_1.isDataOK: Boolean;
var
    sL_YearMonth,sL_Prefix,sL_StartNum,sL_EndNum : String;
    nL_DiffCounts : Integer;
begin
    Result := true;

    sL_YearMonth := Trim(fraYM.getYM);
    if sL_YearMonth = EMPTY_YM_STR then
    begin
      MessageDlg('?????J?o???~??',mtError,[mbOk],0);
      fraYM.mseYM.SetFocus;
      Result := false;
      Exit;
    end;


    sL_Prefix := Trim(medPrefix.Text);
    if sL_Prefix = '' then
    begin
      MessageDlg('?????J?r?y',mtError,[mbOk],0);
      medPrefix.SetFocus;
      Result := false;
      Exit;
    end;


    sL_StartNum := Trim(medStartNum.Text);
    if sL_StartNum = '' then
    begin
      MessageDlg('?????J?}?l???X',mtError,[mbOk],0);
      medStartNum.SetFocus;
      Result := false;
      Exit;
    end;


    sL_EndNum := Trim(medEndNum.Text);
    if sL_EndNum = '' then
    begin
      MessageDlg('?????J???????X',mtError,[mbOk],0);
      medEndNum.SetFocus;
      Result := false;
      Exit;
    end;


    if sL_StartNum > sL_EndNum then
    begin
      MessageDlg('???????X?????j???}?l???X',mtError,[mbOk],0);
      medEndNum.SetFocus;
      Result := false;
      Exit;
    end;


    nL_DiffCounts := StrToInt(sL_EndNum) - StrToInt(sL_StartNum) + 1;
    if (nL_DiffCounts Mod 50)<>0 then
    begin
      MessageDlg('?o???_???`?i????????50??????',mtError,[mbOk],0);
      medEndNum.SetFocus;
      Result := false;
      Exit;
    end;
end;

procedure TfrmInvD07_1.btnSaveClick(Sender: TObject);
var
    sL_YearMonth,sL_Prefix,sL_StartNum,sL_EndNum,sL_CurNum : String;
    sL_LastInvDate,sL_Memo,sL_InvFormat,sL_Useful,sL_SQL : String;
    bL_HavePKValue : Boolean;
    aCheckCurNum: String;
begin
    if isDataOK then
    begin
      sL_YearMonth := TUstr.replaceStr(Trim(fraYM.getYM),'/','');
      sL_YearMonth := dtmMainJ.changePrefixYM(sL_YearMonth);

      sL_Prefix :=  UpperCase(Trim(medPrefix.Text) );
      sL_StartNum := Lpad( Trim(medStartNum.Text), 8, '0' );
      sL_EndNum := Lpad( Trim(medEndNum.Text), 8, '0' );
      sL_CurNum := Lpad( Trim(medCurNum.Text), 8, '0' );

      { ??????, ???????_?l???X, ?N?U?@???i?????X???????_?l???X }
      if ( not FIsUse ) and ( sL_CurNum <> sL_StartNum  ) then
        sL_CurNum := sL_StartNum;

      sL_LastInvDate := fraLastInvDate.getYMD;
      if sL_LastInvDate = EMPTY_DATE_STR then
        sL_LastInvDate := '';

      sL_Memo := Trim(edtMemo.Text);

      if cobInvFormat.ItemIndex = 0  then  //1:?q?l 2:???G(Default)3:???T
        sL_InvFormat := '1'
      else if cobInvFormat.ItemIndex = 1  then
        sL_InvFormat := '2'
      else if cobInvFormat.ItemIndex = 2 then
        sL_InvFormat := '3';

      if cobUseful.ItemIndex = 0  then  //Y:?O  N:?_
        sL_Useful := 'Y'
      else if cobUseful.ItemIndex = 1  then
        sL_Useful := 'N';


      //????PK??
      sL_SQL := 'select * from inv099 where IdentifyId1=''' + IDENTIFYID1 +
                ''' and IdentifyId2=' + IDENTIFYID2 + ' and CompID=''' + sG_CompID +
                ''' and YearMonth=''' + sL_YearMonth +
                ''' and Prefix=''' + sL_Prefix +
                ''' and StartNum=' + sL_StartNum;


      if  dsrInv099.DataSet.State in [dsInsert] then
        bL_HavePKValue := dtmMainJ.checkPK(sL_SQL)
      else
        bL_HavePKValue := false;


      if bL_HavePKValue then
      begin
        MessageDlg('?H?????@??????',mtError,[mbOk],0);
        medPrefix.SetFocus;
        Exit;
      end;

      if ( dsrInv099.DataSet.State in [dsEdit] ) then
      begin
        dtmMain.adoComm.Close;
        dtmMain.adoComm.SQL.Text := Format(
         ' select curnum from inv099     ' +
         '  where identifyid1 = ''%s''   ' +
         '    and identifyid2 = ''%s''   ' +
         '    and compid = ''%s''        ' +
         '    and yearmonth = ''%s''     ' +
         '    and prefix = ''%s''        ' +
         '    and startnum = ''%s''      ',
         [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID,
          sL_YearMonth, sL_Prefix, sL_StartNum] );
        dtmMain.adoComm.Open;
        aCheckCurNum := Nvl( dtmMain.adoComm.FieldByName( 'CURNUM' ).AsString, '0' );
        dtmMain.adoComm.Close;
        if ( StrToInt( aCheckCurNum ) >= StrToInt( Nvl( sL_EndNum, '0' ) ) ) then
        begin
          WarningMsg( '???i?????o?????????s?X, ?o???????????X>?w???????X,???T?{???o?????O?_???b?}???o???C' );
          if medEndNum.CanFocus then medEndNum.SetFocus;
          Exit;
        end;
      end;
      
      
      with dsrInv099.DataSet do
      begin

        FieldByName('IdentifyId1').AsString := IDENTIFYID1;
        FieldByName('IdentifyId2').AsString := IDENTIFYID2;

        FieldByName('CompID').AsString := sG_CompID;
        FieldByName('YearMonth').AsString := sL_YearMonth;

        FieldByName('InvFormat').AsString := sL_InvFormat;
        FieldByName('Prefix').AsString := sL_Prefix;

        FieldByName('StartNum').AsString := Lpad( sL_StartNum, 8, '0' );
        FieldByName('EndNum').AsString := Lpad( sL_EndNum, 8, '0' );
        FieldByName('CurNum').AsString := Lpad( sL_CurNum, 8, '0' );
        FieldByName('LastInvDate').AsString := sL_LastInvDate;
        FieldByName('Useful').AsString := sL_Useful;
        FieldByName('Memo').AsString := sL_Memo;

        FieldByName('UptTime').AsString := DateTimeToStr(now);
        FieldByName('UptEn').AsString := dtmMain.getLoginUser;

        dsrInv099.DataSet.Post;

      end;

      fraYM.mseYM.Enabled := true;
      ChangeBtnStatus;
      setButtonCompetence;
      dtmMainJ.getInv099Data(sG_CompID,sG_QueryYear,sG_QueryYM);
      initialForm;

    end;
end;

procedure TfrmInvD07_1.medPrefixExit(Sender: TObject);
begin
  medPrefix.Text := UpperCase(Trim(medPrefix.Text));
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD07_1.medStartNumExit(Sender: TObject);
begin
  if dsrInv099.DataSet.State in [dsInsert] then
  begin
    medCurNum.Text := Trim(medStartNum.Text);
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD07_1.fraYMmseYMExit(Sender: TObject);
begin
  fraYM.mseYMExit(Sender);
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD07_1.fraYMExit(Sender: TObject);
var
    sL_YearMonth : String;
begin
    sL_YearMonth := TUstr.replaceStr(Trim(fraYM.getYM),'/','');

    if sL_YearMonth <> '' then
    begin
      sL_YearMonth := dtmMainJ.changePrefixYM(sL_YearMonth);
      fraYM.setYM(Copy(sL_YearMonth,1,4) + '/' + Copy(sL_YearMonth,5,2));

      fraLastInvDate.setYMD(Copy(sL_YearMonth,1,4) + '/' + Copy(sL_YearMonth,5,2) + '/01');
    end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD07_1.medEndNumExit(Sender: TObject);
begin
  if strtoint(medEndNum.Text) < strtoint(medCurNum.Text) then
  begin
    WarningMsg('?????s???????????p?? ?U?@?????s?? ?I' );
    medEndNum.SetFocus;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvD07_1.CheckBeforeDelete: Boolean;
var
  aDataSet: TADOQuery;
begin
  Result := False;
  aDataSet := TADOQuery( dsrInv099.DataSet );
  if ( aDataSet.FieldByName( 'STARTNUM' ).AsString <>
       aDataSet.FieldByName( 'CURNUM' ).AsString ) then
  begin
    WarningMsg( '???o?????w????, ???i?R???C' );
    Exit;
  end;     
  if ( aDataSet.FieldByName( 'USEFUL' ).AsString <> 'Y' ) then
  begin
    WarningMsg( '???o???????A???i???i???j, ???i?R???C' );
    Exit;
  end;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD07_1.btnRefreshClick(Sender: TObject);
begin
  dtmMainJ.getInv099Data(sG_CompID, sG_QueryYear, sG_QueryYM);
  initialForm;
  ChangeBtnStatus;
end;

{ ---------------------------------------------------------------------------- }

end.
