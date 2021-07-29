unit frmSO8F10_2U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, Grids, DBGrids, StdCtrls, Buttons, ExtCtrls, Mask, DBCtrls,
  DBClient;

const
    COMP_INFO_FILE_NAME = 'CompInfo.txt';
type
  TfrmSO8F10_2 = class(TForm)
    pnlMFunction: TPanel;
    spbMInsert: TSpeedButton;
    spbMEdit: TSpeedButton;
    spbMDelete: TSpeedButton;
    spbMCancel: TSpeedButton;
    spbMPost: TSpeedButton;
    spbExit: TSpeedButton;
    spbNext: TSpeedButton;
    spbLast: TSpeedButton;
    spbFirst: TSpeedButton;
    spbPrior: TSpeedButton;
    pnlMainMaster: TPanel;
    pnlRMaster: TPanel;
    pnlMaster: TPanel;
    pnlMainDetail: TPanel;
    pnlDFunction: TPanel;
    chbDContinueAppend: TCheckBox;
    dsrDetail: TDataSource;
    dsrMaster: TDataSource;
    DBGrid1: TDBGrid;
    DBEdit1: TDBEdit;
    DBEdit2: TDBEdit;
    DBEdit3: TDBEdit;
    StaticText1: TStaticText;
    StaticText2: TStaticText;
    StaticText3: TStaticText;
    rgpServiceType: TRadioGroup;
    DBRadioGroup1: TDBRadioGroup;
    pnlSingleDetail: TPanel;
    pnlMultiDetail: TPanel;
    dbgDetail: TDBGrid;
    DBEdit5: TDBEdit;
    StaticText4: TStaticText;
    StaticText5: TStaticText;
    DBRadioGroup3: TDBRadioGroup;
    stxTableName: TStaticText;
    dsrCompInfo: TDataSource;
    cobCompInfo: TComboBox;
    StaticText6: TStaticText;
    StaticText7: TStaticText;
    StaticText8: TStaticText;
    spbDInsert: TBitBtn;
    spbDEdit: TBitBtn;
    spbDDelete: TBitBtn;
    spbDCancel: TBitBtn;
    spbDPost: TBitBtn;
    cobCompInfo2: TComboBox;
    procedure FormShow(Sender: TObject);
    procedure spbMInsertClick(Sender: TObject);
    procedure spbMEditClick(Sender: TObject);
    procedure spbMDeleteClick(Sender: TObject);
    procedure spbMCancelClick(Sender: TObject);
    procedure spbMPostClick(Sender: TObject);
    procedure spbExitClick(Sender: TObject);
    procedure spbFirstClick(Sender: TObject);
    procedure spbPriorClick(Sender: TObject);
    procedure spbNextClick(Sender: TObject);
    procedure spbLastClick(Sender: TObject);
    procedure dsrMasterDataChange(Sender: TObject; Field: TField);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure cobCompInfoCloseUp(Sender: TObject);
    procedure spbDInsertClick(Sender: TObject);
    procedure spbDEditClick(Sender: TObject);
    procedure spbDDeleteClick(Sender: TObject);
    procedure spbDCancelClick(Sender: TObject);
    procedure spbDPostClick(Sender: TObject);
    procedure dsrDetailDataChange(Sender: TObject; Field: TField);
    procedure cobCompInfo2CloseUp(Sender: TObject);
  private
    { Private declarations }
    DetailDataSet, MasterDataSet : TDataSet;
    procedure setServiceTypeUI;
    procedure SetDetailKeyValues;
    procedure SaveDetailDataSet;
    procedure preAppendMData;
    procedure preAppendDData;    
    procedure ChangeMasterBtnState;
    procedure loadCompInfo(I_ComboBox : TComboBox);
  public
    { Public declarations }
    sG_ActiveTableName,sG_CobCompCode : String;
    procedure ChangeDetailBtnState;

  end;

var
  frmSO8F10_2: TfrmSO8F10_2;

implementation

uses dtmMain4U, UdateTimeu, Ustru, frmMainMenuU;

{$R *.dfm}

{ TfrmSO8F10_2 }

procedure TfrmSO8F10_2.preAppendMData;
begin
    MasterDataSet.Insert;
    MasterDataSet.FieldByName('STOPFLAG').AsInteger := 0;
    MasterDataSet.FieldByName('SERVICETYPE').AsString := 'C';    
    rgpServiceType.ItemIndex := 0;
    ChangeMasterBtnState;
    MasterDataSet.FieldByName('CODENO').ReadOnly := false;    
    MasterDataSet.FieldByName('CODENO').FocusControl;
end;

procedure TfrmSO8F10_2.FormShow(Sender: TObject);
var
    bL_IsEnable : Boolean;
begin

    MasterDataSet := dsrMaster.DataSet;
    DetailDataSet := dsrDetail.DataSet;
    setServiceTypeUI;

    ChangeMasterBtnState;
    ChangeDetailBtnState;

    self.Caption := frmMainMenu.setCaption('SO8F10','[二階代碼分類資料對照表]');

    //down, 設定權限...
    if (frmMainMenu.sG_IsSupervisor='Y') or (frmMainMenu.sG_NeedLogin='Y') then
    begin
      spbMInsert.Enabled := true;
      spbMEdit.Enabled := true;
      spbMDelete.Enabled := true;

      spbDInsert.Enabled := true;
      spbDEdit.Enabled := true;
      spbDDelete.Enabled := true;
    end
    else
    begin
      frmMainMenu.activePrivDataset(frmMainMenu.sG_CompCode,frmMainMenu.sG_UserID);
      bL_IsEnable := frmMainMenu.checkPriv('SO8F10');
      spbMInsert.Enabled := bL_IsEnable;
      spbMEdit.Enabled := bL_IsEnable;
      spbMDelete.Enabled := bL_IsEnable;

      spbDInsert.Enabled := bL_IsEnable;
      spbDEdit.Enabled := bL_IsEnable;
      spbDDelete.Enabled := bL_IsEnable;
    end;
    //up, 設定權限...

end;

procedure TfrmSO8F10_2.ChangeDetailBtnState;
begin
    if not Assigned(DetailDataSet) then Exit;
    if not DetailDataSet.active then Exit;
    if MasterDataSet.RecordCount=0 then
    begin
      spbDEdit.Enabled := False;
      spbDDelete.Enabled := False;
      dbgDetail.Options := [dgTitles,dgIndicator,dgColumnResize,dgColLines,dgRowLines,dgTabs,dgConfirmDelete,dgCancelOnExit];
      spbDInsert.Enabled := False;
      spbDCancel.Enabled := False;
      spbDPost.Enabled := False;
    end
    else
    begin
      with DetailDataSet do
      begin
        if state in [dsEdit, dsInsert] then
        begin
          dbgDetail.Options := [dgEditing,dgTitles,dgIndicator,dgColumnResize,dgColLines,dgRowLines,dgTabs,dgConfirmDelete,dgCancelOnExit];
          spbDInsert.Enabled := False;
          spbDEdit.Enabled := False;
          spbDDelete.Enabled := False;
          spbDCancel.Enabled := True;
          spbDPost.Enabled := True;
          pnlMFunction.Enabled := False;
          pnlRMaster.Enabled := False;
          pnlMaster.Enabled := False;
          pnlSingleDetail.Enabled := true;
          pnlMultiDetail.Enabled := false;
          chbDContinueAppend.Enabled := false;
          cobCompInfo.Enabled := false;
        end
        else
        begin
          if RecordCount=0 then
          begin
            spbDEdit.Enabled := False;
            spbDDelete.Enabled := False;
          end
          else
          begin
            spbDEdit.Enabled := True;
            spbDDelete.Enabled := True;
          end;
          pnlSingleDetail.Enabled := false;
          pnlMultiDetail.Enabled := true;
          pnlMFunction.Enabled := True;
          pnlRMaster.Enabled := True;
          pnlMaster.Enabled := False;
          dbgDetail.Options := [dgTitles,dgIndicator,dgColumnResize,dgColLines,dgRowLines,dgTabs,dgConfirmDelete,dgCancelOnExit];
          spbDInsert.Enabled := True;
          spbDCancel.Enabled := False;
          spbDPost.Enabled := False;
          chbDContinueAppend.Enabled := True;
          cobCompInfo.Enabled := True;

        end;
      end;
    end;

end;

procedure TfrmSO8F10_2.ChangeMasterBtnState;
begin
     with MasterDataSet do
     begin
       if state in [dsEdit, dsInsert] then
       begin
         spbFirst.Enabled := False;
         spbPrior.Enabled := False;
         spbNext.Enabled := False;
         spbLast.Enabled := False;

         pnlMainDetail.Enabled := False;
         pnlDFunction.Enabled := False;
         pnlMaster.Enabled := True;
         pnlRMaster.Enabled := False;
         spbMInsert.Enabled := False;
         spbMEdit.Enabled := False;
         spbMDelete.Enabled := False;
         spbMCancel.Enabled := True;
         spbMPost.Enabled := True;
       end
       else
       begin
         if RecordCount=0 then
         begin
           spbFirst.Enabled := False;
           spbPrior.Enabled := False;
           spbNext.Enabled := False;
           spbLast.Enabled := False;

           spbMEdit.Enabled := False;
           spbMDelete.Enabled := False;
         end
         else
         begin
           spbFirst.Enabled := True;
           spbPrior.Enabled := True;
           spbNext.Enabled := True;
           spbLast.Enabled := True;
           
           spbMEdit.Enabled := True;
           spbMDelete.Enabled := True;
         end;
         pnlMainDetail.Enabled := True;         
         pnlDFunction.Enabled := True;
         pnlMaster.Enabled := False;
         pnlRMaster.Enabled := True;         
         spbMInsert.Enabled := True;
         spbMCancel.Enabled := False;
         spbMPost.Enabled := False;
       end;
     end;

end;

procedure TfrmSO8F10_2.spbMInsertClick(Sender: TObject);
begin
     PreAppendMData;
end;

procedure TfrmSO8F10_2.spbMEditClick(Sender: TObject);
begin
    MasterDataSet.Edit;
    MasterDataSet.FieldByName('CODENO').ReadOnly := true;
    ChangeMasterBtnState;
    DBEdit2.SetFocus;
end;

procedure TfrmSO8F10_2.spbMDeleteClick(Sender: TObject);
var
    nL_DetailDataCounts : Integer;
    sL_TableName,sL_MasterCode : String;
begin
    sL_TableName := dtmMain4.adoCD067A.FieldByName('TableName').AsString;
    sL_MasterCode := dtmMain4.adoCD067B.FieldByName('CodeNo').AsString;
    nL_DetailDataCounts := dtmMain4.checkDetailDataCounts(sL_TableName,sL_MasterCode);


    if nL_DetailDataCounts = 0 then
    begin
      if MessageDlg('是否確認刪除?', mtConfirmation, [mbYes, mbNo], 0) = mrNo then Exit;

  //     if MessageDlg('是否一併刪除明細項?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
  //     begin

      dtmMain4.adoCD067BAfterScroll(MasterDataSet);
      MasterDataSet.Delete;
      if (DetailDataSet.State = dsInactive) or (DetailDataSet.RecordCount=0)then Exit;
       with DetailDataSet do
       begin

         First;
         while not IsEmpty do
         begin
           Delete;
           Next;
         end;
       end;
  //     end;


      ChangeMasterBtnState;
    end
    else
    begin
      MessageDlg('尚有代碼對應資料,不能刪除!', mtWarning, [mbOK],0);
    end;
end;

procedure TfrmSO8F10_2.spbMCancelClick(Sender: TObject);
begin
     if MasterDataSet.State in [dsEdit, dsInsert] then
       MasterDataSet.Cancel;
     ChangeMasterBtnState;

end;

procedure TfrmSO8F10_2.spbMPostClick(Sender: TObject);
var
    sL_UpdTime : String;
begin
     SaveDetailDataSet;
     if MasterDataSet.State in [dsInsert] then
       MasterDataSet.FieldByName('TABLENAME').AsString := sG_ActiveTableName;
     if MasterDataSet.State in [dsEdit, dsInsert] then
     begin
       case rgpServiceType.ItemIndex of
         0://有線電視
           MasterDataSet.FieldByName('SERVICETYPE').AsString := 'C';
         1: // CM
           MasterDataSet.FieldByName('SERVICETYPE').AsString := 'I';
       end;
       MasterDataSet.FieldByName('OPERATOR').AsString := frmMainMenu.sG_USerID;

//       sL_UpdTime := TUdateTime.CDateStr(date, 9) + ' ' + timeToStr(time);
//       if Copy(sL_UpdTime,1,1)='0' then
//         sL_UpdTime := Copy(sL_UpdTime,2,Length(sL_UpdTime));

       MasterDataSet.FieldByName('UPDTIME').AsString := FormatDateTime( 'yyyy/mm/dd hh:mm:ss', Now );
       MasterDataSet.Post;
     end;
     if MasterDataSet.RecordCount=1 then
       dtmMain4.adoCD067BAfterScroll(MasterDataSet);
     ChangeMasterBtnState;
     ChangeDetailBtnState;
end;

procedure TfrmSO8F10_2.SaveDetailDataSet;
var
    sL_UpdTime : String;
begin
     if DetailDataSet.State in [dsEdit, dsInsert] then
     begin
       SetDetailKeyValues;
       DetailDataSet.FieldByName('OPERATOR').AsString := frmMainMenu.sG_USerID;

//       sL_UpdTime := TUdateTime.CDateStr(date, 9) + ' ' + timeToStr(time);
//       if Copy(sL_UpdTime,1,1)='0' then
//         sL_UpdTime := Copy(sL_UpdTime,2,Length(sL_UpdTime));

       DetailDataSet.FieldByName('UPDTIME').AsString := FormatDateTime( 'yyyy/mm/dd hh:mm:ss', Now );


       DetailDataSet.Post;
       ChangeDetailBtnState;
     end;
end;

procedure TfrmSO8F10_2.SetDetailKeyValues;
begin
    DetailDataSet.FieldByName('TABLENAME').AsString := MasterDataSet.FieldByName('TABLENAME').AsString;
    DetailDataSet.FieldByName('MASTERCODENO').AsInteger := MasterDataSet.FieldByName('CODENO').AsInteger;
end;

procedure TfrmSO8F10_2.spbExitClick(Sender: TObject);
begin
     if MasterDataSet.State in [dsEdit, dsInsert] then
     begin
       if MessageDlg('資料有所異動,是否儲存?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
         MasterDataSet.Post;
     end;

     if DetailDataSet.State in [dsEdit, dsInsert] then
     begin
       if MessageDlg('資料有所異動,是否儲存?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
         DetailDataSet.Post;
     end;

     Close;
end;

procedure TfrmSO8F10_2.spbFirstClick(Sender: TObject);
begin
     SaveDetailDataSet;
     MasterDataSet.First;
     ChangeDetailBtnState;

end;

procedure TfrmSO8F10_2.spbPriorClick(Sender: TObject);
begin
     SaveDetailDataSet;
     MasterDataSet.Prior;
     ChangeDetailBtnState;

end;

procedure TfrmSO8F10_2.spbNextClick(Sender: TObject);
begin
     SaveDetailDataSet;
     MasterDataSet.Next;
     ChangeDetailBtnState;

end;

procedure TfrmSO8F10_2.spbLastClick(Sender: TObject);
begin
     SaveDetailDataSet;
     MasterDataSet.Last;
     ChangeDetailBtnState;

end;



procedure TfrmSO8F10_2.preAppendDData;
begin
    if DetailDataSet.State = dsInactive Then Exit;
    DetailDataSet.Insert;
    DetailDataSet.FieldByName('STOPFLAG').AsInteger := 0;
    DetailDataSet.FieldByName('COMPCODE').AsString := frmMainMenu.sG_CompCode;
    ChangeDetailBtnState;
    cobCompInfo2.SetFocus;


end;



procedure TfrmSO8F10_2.dsrMasterDataChange(Sender: TObject;
  Field: TField);
begin
    if MasterDataSet=nil then Exit;
    setServiceTypeUI;
end;

procedure TfrmSO8F10_2.setServiceTypeUI;
begin
    if MasterDataSet.State in [dsBrowse] then
    begin
      if MasterDataSet.FieldByName('ServiceType').AsString = 'C' then
        rgpServiceType.ItemIndex := 0
      else
        rgpServiceType.ItemIndex := 1;
    end;

end;

procedure TfrmSO8F10_2.loadCompInfo(I_ComboBox : TComboBox);
var
    ii : Integer;
    sL_FileName : String;
    L_TmpStrList, L_TmpStrList1 : TStringList;
    sL_ExeFileName, sL_ExePath : STring;
    sL_CompCode, sL_CompName : String;


begin

    sL_ExeFileName := Application.ExeName;
    sL_ExePath := ExtractFileDir(sL_ExeFileName);

    sL_FileName := sL_ExePath + '\' + COMP_INFO_FILE_NAME;

    if FileExists(sL_FileName) then
    begin
      if dtmMain4.cdsCompInfo.State = dsInactive then
        dtmMain4.cdsCompInfo.CreateDataSet;

      I_ComboBox.Items.Clear;
      I_ComboBox.Items.Add(ALL_COMPCODE + '-所有');
      I_ComboBox.ItemIndex := 0;

      L_TmpStrList := TStringList.Create;
      L_TmpStrList.LoadFromFile(sL_FileName);

      for ii:=0 to L_TmpStrList.Count -1 do
      begin
        if L_TmpStrList.Strings[ii] <>'' then
        begin
          L_TmpStrList1 := TUstr.ParseStrings(L_TmpStrList.Strings[ii],'=');
          if L_TmpStrList1.Count=2 then
          begin
            sL_CompCode := L_TmpStrList1.Strings[0];
            sL_CompName := L_TmpStrList1.Strings[1];
            I_ComboBox.Items.Add(sL_CompCode + '-' + sL_CompName);

            dtmMain4.cdsCompInfo.Append;
            dtmMain4.cdsCompInfo.FieldByName('CompCode').AsInteger := StrToInt(sL_CompCode);
            dtmMain4.cdsCompInfo.FieldByName('CompName').AsString := sL_CompCode + '-' + sL_CompName;
            dtmMain4.cdsCompInfo.FieldByName('Index').AsInteger := ii;
            dtmMain4.cdsCompInfo.Post;

          end;

          if Assigned(L_TmpStrList1) then
            L_TmpStrList1.Free;

        end;
      end;

      L_TmpStrList.Free;
    end
    else
    begin
      MessageDlg('找不到公司資訊檔(' + sL_FileName + ')程式即將結束.', mtError, [mbOK],0);
      Application.Terminate;
    end;
end;

procedure TfrmSO8F10_2.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
    dtmMain4.cdsCompInfo.Close;
end;

procedure TfrmSO8F10_2.FormCreate(Sender: TObject);
begin
    loadCompInfo(cobCompInfo);
    loadCompInfo(cobCompInfo2);
end;

procedure TfrmSO8F10_2.cobCompInfoCloseUp(Sender: TObject);
var
    sL_CompCode : String;
    L_StrList : TStringList;
begin
    L_StrList := TStringList.Create;
    L_StrList := TUstr.ParseStrings(cobCompInfo.Text,'-');

    sG_CobCompCode := L_StrList.Strings[0];
    dtmMain4.filterCd067C(sG_CobCompCode);

    L_StrList.Free;
end;



procedure TfrmSO8F10_2.spbDInsertClick(Sender: TObject);
begin
     PreAppendDData;
end;

procedure TfrmSO8F10_2.spbDEditClick(Sender: TObject);
begin
    if DetailDataSet.State = dsInactive Then Exit;
    DetailDataSet.Edit;
    ChangeDetailBtnState;
    cobCompInfo2.SetFocus;
end;

procedure TfrmSO8F10_2.spbDDeleteClick(Sender: TObject);
begin
    if DetailDataSet.State = dsInactive Then Exit;
    if DetailDataSet.RecordCount=0 then
    begin
      MessageDlg('沒有資料可供刪除!!',mtWarning,[mbOK],0);
      Exit;
    end;
    if MessageDlg('是否確認刪除?', mtConfirmation, [mbYes, mbNo], 0) = mrNo then Exit;
    DetailDataSet.Delete;
    ChangeDetailBtnState;
end;

procedure TfrmSO8F10_2.spbDCancelClick(Sender: TObject);
begin
    if DetailDataSet.State = dsInactive Then Exit;
    DetailDataSet.Cancel;
    ChangeDetailBtnState;
end;

procedure TfrmSO8F10_2.spbDPostClick(Sender: TObject);
begin
    if DetailDataSet.State = dsInactive Then Exit;
    SaveDetailDataSet;
    if chbDContinueAppend.Checked then
      PreAppendDData;
end;

procedure TfrmSO8F10_2.dsrDetailDataChange(Sender: TObject; Field: TField);
var
    sL_CompCode : String;
    nL_Ndx : Integer;
begin
    sL_CompCode := dtmMain4.adoCD067C.FieldByName('CompCode').AsString;
    if (sL_CompCode = ALL_COMPCODE) OR (sL_CompCode = '') then
    begin
      cobCompInfo2.ItemIndex := 0;
    end
    else if dtmMain4.cdsCompInfo.Locate('CompCode',sL_CompCode,[]) then
    begin
      nL_Ndx := dtmMain4.cdsCompInfo.FieldByName('Index').AsInteger;
      cobCompInfo2.ItemIndex := nL_Ndx + 1;
    end;

end;

procedure TfrmSO8F10_2.cobCompInfo2CloseUp(Sender: TObject);
var
    sL_SelectedCompCode : String;
    L_CompStrList : TStringList;
begin
     sL_SelectedCompCode := Trim(cobCompInfo2.Text);
     L_CompStrList := TStringList.Create;
     L_CompStrList := TUstr.ParseStrings(sL_SelectedCompCode,'-');
     DetailDataSet.FieldByName('CompCode').AsString := L_CompStrList.Strings[0];
     L_CompStrList.Free;
end;

end.
