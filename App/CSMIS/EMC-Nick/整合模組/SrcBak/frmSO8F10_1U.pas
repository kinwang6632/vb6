unit frmSO8F10_1U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, Buttons, Grids, DBGrids, DB, 
  Mask, DBCtrls, DBClient ,IniFiles;

type
  TfrmSO8F10_1 = class(TForm)
    pnlCommand: TPanel;
    pnlMultiData: TPanel;
    Splitter1: TSplitter;
    pnlSingleData: TPanel;
    dsrCD067A: TDataSource;
    DBGrid1: TDBGrid;
    StaticText5: TStaticText;
    StaticText7: TStaticText;
    DBEdit2: TDBEdit;
    btnAppend: TBitBtn;
    btnEdit: TBitBtn;
    btnDelete: TBitBtn;
    btnCancelExit: TBitBtn;
    btnSave: TBitBtn;
    DBEdit5: TDBEdit;
    btnEditData: TBitBtn;
    btnSynchronization: TBitBtn;
    Label1: TLabel;
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnAppendClick(Sender: TObject);
    procedure btnEditClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure btnCancelExitClick(Sender: TObject);
    procedure DBGrid1TitleClick(Column: TColumn);
    procedure DBGrid1DblClick(Sender: TObject);
    procedure btnEditDataClick(Sender: TObject);
    procedure btnSynchronizationClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FormCreate(Sender: TObject);

  private
    { Private declarations }
    sG_EmpSQL : String;
    G_UserNameStrList : TstringList;
    G_UserPasswdStrList : TstringList;
    function IsDataOK:boolean;
    procedure ChangeBtnState;

  public
    { Public declarations }
    function getCommDBIniInfo(var sI_Alias,sI_UserID,sI_Password : String) : WideString;
  end;

var
  frmSO8F10_1: TfrmSO8F10_1;

implementation

uses dtmMain4U, UdateTimeu, frmSO8F10_2U, frmMainMenuU, frmSO8F10_3U,
  frmSO8F10_4U, UCommonU;

{$R *.dfm}

procedure TfrmSO8F10_1.FormShow(Sender: TObject);
var
    bL_IsEnable : Boolean;
begin
    ChangeBtnState;

    Application.CreateForm(TfrmSO8F10_4,frmSO8F10_4);
    frmSO8F10_4.ShowModal;


    dtmMain4.connectToCommDB(frmSO8F10_4.bG_IsSuperUser);

    self.Caption := frmMainMenu.setCaption('SO8F10','[二階代碼分類資料維護]');

    if frmSO8F10_4.bG_IsSuperUser then
    begin
      Label1.Caption := '共用資料區';
      btnSynchronization.Enabled := true;
    end
    else
    begin
      Label1.Caption := frmMainMenu.sG_CompName;
      btnSynchronization.Enabled := false;
    end;

    //down, 設定權限...
    if (frmMainMenu.sG_IsSupervisor='Y') or (frmMainMenu.sG_NeedLogin='Y') then
    begin
      btnAppend.Enabled := true;
      btnEdit.Enabled := true;
      btnDelete.Enabled := true;
    end
    else
    begin
      frmMainMenu.activePrivDataset(frmMainMenu.sG_CompCode,frmMainMenu.sG_UserID);
      bL_IsEnable := frmMainMenu.checkPriv('SO8F10');
      btnAppend.Enabled := bL_IsEnable;
      btnEdit.Enabled := bL_IsEnable;
      btnDelete.Enabled := bL_IsEnable;
    end;
    //up, 設定權限...


    dtmMain4.activeDataSet(dtmMain4.adoCD067A);
    dtmMain4.activeDataSet(dtmMain4.adoCodeCD067A);
end;

procedure TfrmSO8F10_1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
    Action := caFree;
    G_UserNameStrList.Free;
    G_UserPasswdStrList.Free;
end;

procedure TfrmSO8F10_1.ChangeBtnState;
begin
     with dsrCD067A.DataSet do
     begin
       if state in [dsInactive] then
         Exit
       else if state in [dsEdit, dsInsert] then
       begin
          btnCancelExit.Caption := '取消';
          btnCancelExit.Enabled := TRUE;
          btnSave.Enabled := TRUE;
          btnDelete.Enabled := FALSE;
          btnEditData.Enabled := FALSE;
          btnEdit.Enabled := FALSE;
          btnAppend.Enabled := FALSE;
          pnlSingleData.Enabled := TRUE;
          pnlMultiData.Enabled := FALSE;
       end
       else
       begin
          if RecordCount>0 then
          begin
            btnEditData.Enabled := true;
            btnEdit.Enabled := TRUE;
            btnDelete.Enabled := TRUE;
          end
          else
          begin
            btnEditData.Enabled := FALSE;
            btnEdit.Enabled := FALSE;
            btnDelete.Enabled := FALSE;
          end;
          btnCancelExit.Caption := '結束';
          btnCancelExit.Enabled := TRUE;
          btnSave.Enabled := FALSE;
          btnAppend.Enabled := TRUE;
          pnlSingleData.Enabled := FALSE;
          pnlMultiData.Enabled := TRUE;
       end;
     end;
end;

procedure TfrmSO8F10_1.btnAppendClick(Sender: TObject);
begin
    if dsrCD067A.DataSet.State = dsInactive then
      dsrCD067A.DataSet.Active := True;
    dsrCD067A.DataSet.Append;
    ChangeBtnState;
    dsrCD067A.DataSet.FieldByName('TABLENAME').ReadOnly := false;
    dsrCD067A.DataSet.FieldByName('TABLENAME').FocusControl;
end;

procedure TfrmSO8F10_1.btnEditClick(Sender: TObject);
var
    nL_Seq : Integer;
begin
    if dsrCD067A.DataSet.State = dsInactive then Exit;
    if dsrCD067A.DataSet.RecordCount =0 then Exit;
    dsrCD067A.DataSet.Edit;
    ChangeBtnState;
    dsrCD067A.DataSet.FieldByName('TABLENAME').ReadOnly := true;
    dsrCD067A.DataSet.FieldByName('TABLEDESCRIPTION').FocusControl;
end;

procedure TfrmSO8F10_1.btnDeleteClick(Sender: TObject);
var
    nL_Seq,nL_MasterDataCounts,nL_DetailDataCounts : Integer;
    sL_TableName : String;
begin
    sL_TableName := dtmMain4.adoCD067A.FieldByName('TableName').AsString;
    nL_MasterDataCounts := dtmMain4.checkCodeDataCounts('Cd067B',sL_TableName);
    nL_DetailDataCounts := dtmMain4.checkCodeDataCounts('Cd067C',sL_TableName);

    if (nL_MasterDataCounts = 0) and (nL_DetailDataCounts = 0) then
    begin
      if dsrCD067A.DataSet.State = dsInactive then Exit;
      if dsrCD067A.DataSet.RecordCount =0 then Exit;
        if MessageDlg('是否確認刪除?',mtConfirmation, [mbYes, mbNo], 0) = mrNo then
          Exit;
        dsrCD067A.DataSet.Delete;
        ChangeBtnState;
    end
    else
    begin
      MessageDlg('尚有代碼對應資料,不能刪除!', mtWarning, [mbOK],0);
    end;
//      MessageDlg('此組單據號碼,已經有對應之回單資料,所以無法刪除!', mtWarning, [mbOK],0);
end;

procedure TfrmSO8F10_1.btnSaveClick(Sender: TObject);
var
    sL_UpdTime : String;
begin
    try
      if dsrCD067A.DataSet.State = dsInactive then Exit;
      if IsDataOK then
      begin
        with dsrCD067A.DataSet do
        begin
          FieldByName('OPERATOR').AsString := frmMainMenu.sG_USerID;

          //sL_UpdTime := TUdateTime.CDateStr(date, 9) + ' ' + timeToStr(time);
          //if Copy(sL_UpdTime,1,1)='0' then
          //  sL_UpdTime := Copy(sL_UpdTime,2,Length(sL_UpdTime));

          FieldByName('UPDTIME').AsString := FormatDateTime( 'yyyy/mm/dd hh:mm:ss', Now );
          Post;
        end;
        ChangeBtnState;
      end;
    except
      on E: Exception do
      begin
        if dsrCD067A.DataSet.State = dsInsert then
          dsrCD067A.DataSet.FieldByName('TABLENAME').FocusControl
        else if dsrCD067A.DataSet.State = dsEdit then
          dsrCD067A.DataSet.FieldByName('TABLEDESCRIPTION').FocusControl;
        MessageDlg(E.Message, mtWarning, [mbok],0);
      end;
    end;
end;

procedure TfrmSO8F10_1.btnCancelExitClick(Sender: TObject);
begin
    if dsrCD067A.DataSet.State in [dsEdit, dsInsert ] then
    begin
      dsrCD067A.DataSet.Cancel;
      ChangeBtnState;
    end
    else if dsrCD067A.DataSet.State in [dsInactive, dsBrowse] then
    begin
      G_UserNameStrList.Free;
      G_UserPasswdStrList.Free;
      Close;
    end;


end;

function TfrmSO8F10_1.IsDataOK: boolean;
var
    bL_Result : boolean;
begin
    bL_Result := true;
    if dsrCD067A.DataSet.FieldByName('TABLENAME').AsString='' then
    begin
      MessageDlg('請輸入 Table Name!', mtWarning, [mbOK],0);
      dsrCD067A.DataSet.FieldByName('TABLENAME').FocusControl;
      bL_Result := false;
    end
    else if dsrCD067A.DataSet.FieldByName('TABLEDESCRIPTION').AsString='' then
    begin
      MessageDlg('請輸入 Table Description!', mtWarning, [mbOK],0);
      dsrCD067A.DataSet.FieldByName('TABLEDESCRIPTION').FocusControl;
      bL_Result := false;
    end;
    
    result := bL_Result;
end;

procedure TfrmSO8F10_1.DBGrid1TitleClick(Column: TColumn);
var
   ii : Integer;
   bFlag : Boolean;
begin

     bFlag := False;
     for ii:=0 to dsrCD067A.DataSet.FieldCount-1 do
     begin
       if (dsrCD067A.DataSet.Fields[ii].FieldName=Column.FieldName) and
          (dsrCD067A.DataSet.Fields[ii].FieldKind=fkData) then
       begin
         bFlag := True;
         break;
       end;
     end;

     if not bFlag then
       Exit;//該FieldName不存在於FDataSet中
     if dsrCD067A.DataSet IS TClientDataSet then
     begin
       if TClientDataSet(dsrCD067A.DataSet).IndexName = 'TmpIndex' then
         TClientDataSet(dsrCD067A.DataSet).DeleteIndex('TmpIndex');
       TClientDataSet(dsrCD067A.DataSet).AddIndex('TmpIndex', Column.FieldName, [ixCaseInsensitive],'','',0);
       TClientDataSet(dsrCD067A.DataSet).IndexName := 'TmpIndex';
     end;
end;

procedure TfrmSO8F10_1.DBGrid1DblClick(Sender: TObject);
begin
//    btnEditClick(nil);
    btnEditDataClick(nil);
end;

procedure TfrmSO8F10_1.btnEditDataClick(Sender: TObject);
var
    sL_TableName : String;
begin
    sL_TableName := dsrCD067A.DataSet.fieldByName('TABLENAME').AsString;

    frmSO8F10_2 := TfrmSO8F10_2.Create(Application);
    dtmMain4.adoCD067C.Close;
    dtmMain4.activeCd067B(sL_TableName);
    frmSO8F10_2.stxTableName.Caption := '目前 active 的 table name : ' +  sL_TableName;
    frmSO8F10_2.sG_ActiveTableName := sL_TableName;
    frmSO8F10_2.ShowModal;
    frmSO8F10_2.Free;

end;

procedure TfrmSO8F10_1.btnSynchronizationClick(Sender: TObject);
begin
    Application.CreateForm(TfrmSO8F10_3,frmSO8F10_3);
    frmSO8F10_3.ShowModal;
end;

procedure TfrmSO8F10_1.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
    if (ssCtrl in Shift) and (Key=83) then //Ctrl + s
    begin
      Application.CreateForm(TfrmSO8F10_4,frmSO8F10_4);
      frmSO8F10_4.ShowModal;
    end;
end;

procedure TfrmSO8F10_1.FormCreate(Sender: TObject);
begin
    //資料同步除了通過密碼人員,否則不能同步資料
    btnSynchronization.Enabled := FALSE;
end;

function TfrmSO8F10_1.getCommDBIniInfo(var sI_Alias, sI_UserID,
  sI_Password: String): WideString;
var
    sL_ErrMsg,sL_ExeFileName,sL_ExePath,sL_IniFileName,sL_Tag : String;
    L_IniFile : TIniFile;
    nL_TotalDataArea,ii : Integer;
begin
    //將加密的ini檔案解密
    sL_ErrMsg := frmMainMenu.TransTmpIniFile(SYS_COMPDBINFO_INI_FILE_NAME);
    if sL_ErrMsg = '' then
    begin
//**********************************
      sL_ExeFileName := Application.ExeName;

      sL_ExePath := ExtractFileDir(sL_ExeFileName);

      sL_IniFileName := sL_ExePath + '\' + TMP_SYS_INI_FILE_NAME;

      if not FileExists(sL_IniFileName) then
      begin
        result := '-1: 讀取資料庫連線參數檔<'+sL_IniFileName +'>失敗';
        exit;
      end;
      L_IniFile := TIniFile.Create(sL_IniFileName);
      nL_TotalDataArea := L_IniFile.ReadInteger('SYSINFO','TotalDataAreaCounts',0);
      if nL_TotalDataArea=0 then
      begin
        result := '-1: 資料庫連線參數檔之資料區總數設定錯誤!<'+sL_IniFileName +'>';
        exit;
      end;

      try
        for ii:=0 to nL_TotalDataArea -1 do
        begin
          if ii = 0 then
          begin
            sL_Tag := DATA_AREA_HEADER + IntToStr(ii);

            sI_Alias := L_IniFile.ReadString(sL_Tag,'ALIAS','');;
            sI_UserID := L_IniFile.ReadString(sL_Tag,'USERID','');
            sI_Password := L_IniFile.ReadString(sL_Tag,'PASSWORD','');
          end;
        end;

        L_IniFile.Free;
      except
        result := '-1: 資料庫連線參數檔之資料區設定錯誤!<'+sL_IniFileName +'>';
        exit;
      end;

      result := '';


//**********************************
    end
    else
    begin
      showmessage(sL_ErrMsg);
    end;
end;

end.

