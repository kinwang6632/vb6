unit frmDbInfoU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, Grids, DBGrids, StdCtrls, Buttons, DB, IniFiles, Mask,
  DBCtrls, DBClient, fraYMDU;


const
    INSERT_MODE = 'I';
    EDIT_MODE = 'E';

type
  TfrmDbInfo = class(TForm)
    Panel1: TPanel;
    Label0: TLabel;
    Panel3: TPanel;
    btnAppend: TBitBtn;
    btnEdit: TBitBtn;
    btnDelete: TBitBtn;
    btnCancel: TBitBtn;
    btnSave: TBitBtn;
    Panel2: TPanel;
    pnlMultiData: TPanel;
    dbgParam: TDBGrid;
    dsrDbLinkSet: TDataSource;
    pnlSingleData: TPanel;
    OpenDialog1: TOpenDialog;
    Panel4: TPanel;
    medDbCount: TMaskEdit;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    medCompCode: TMaskEdit;
    medGroup: TMaskEdit;
    medAlias: TEdit;
    medPassword: TEdit;
    medUserID: TEdit;
    medCompName: TEdit;
    btnExit: TBitBtn;
    Label4: TLabel;
    medCasID: TEdit;
    SpeedButton1: TSpeedButton;
    medDataPath1: TEdit;
    medMail1: TEdit;
    medMail2: TEdit;
    medSrcCode: TEdit;
    medSrcIP: TEdit;
    medSrcType: TEdit;
    medDstIP: TEdit;
    medDstCode: TEdit;
    medDstType: TEdit;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    Label16: TLabel;
    Label17: TLabel;
    Label18: TLabel;
    Label19: TLabel;
    Label20: TLabel;
    Label21: TLabel;
    medVersion: TEdit;
    Label22: TLabel;
    medDataPath2: TEdit;
    SpeedButton2: TSpeedButton;
    fraYMD1: TfraYMD;
    procedure BitBtn1Click(Sender: TObject);
    procedure btnAppendClick(Sender: TObject);
    procedure btnEditClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure dsrDbLinkSetDataChange(Sender: TObject; Field: TField);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure SpeedButton2Click(Sender: TObject);
  private
    { Private declarations }
//    nG_DbCount : Integer;
    sG_IniFileName,sG_Mode : String;
    procedure ReadIni;
    procedure ChangeBtnStatus;
    function getMaxGroupNo : Integer;
    function isDataOK : Boolean;
    procedure saveTempIni(I_IniFile : TIniFile);
    procedure setLanguageString;
  public
    { Public declarations }
  end;

var
  frmDbInfo: TfrmDbInfo;

implementation

uses dtmMainU, ConstU, DevPower_TLB, Ustru;

{$R *.dfm}

procedure TfrmDbInfo.BitBtn1Click(Sender: TObject);
begin
    Close;
end;

procedure TfrmDbInfo.ChangeBtnStatus;
begin
     with dtmMain.cdsSGSParam do
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
          pnlSingleData.Enabled := TRUE;
          pnlMultiData.Enabled := FALSE;
          //dbtID.SetFocus;
       end
       else
       begin
          if dtmMain.cdsSGSParam.RecordCount>0 then
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
          pnlSingleData.Enabled := FALSE;
          pnlMultiData.Enabled := TRUE;
       end;
     end;

end;

procedure TfrmDbInfo.ReadIni;
var
    L_IniFile : TIniFile;
    ii, nL_CompCode, nL_Key ,nL_CompCounts : Integer;
    L_StrList : TStringList;
    sL_ProcessorIP,sL_DispatcherIP,sL_TcOrSc : String;
begin
    //解密
    dtmMain.TransTmpIniFile;

    L_StrList := TStringList.Create;

    sG_IniFileName := dtmMain.sG_ExePath + '\' + TMP_SYS_INI_FILE_NAME;

    if FileExists(sG_IniFileName) then
    begin
      L_IniFile := TIniFile.Create(sG_IniFileName);

      //判別簡體繁體
      sL_TcOrSc := L_IniFile.ReadString('CURRENT_LANGUAGE_TYPE','CURRENT_LANGUAGE_TYPE','');

      L_IniFile.ReadSectionValues(sL_TcOrSc,L_StrList);

      dtmMain.initialcdsSGSParamSet(L_StrList);

      nL_CompCounts := L_IniFile.ReadInteger(sL_TcOrSc,'COMP_COUNT',0);
      medDbCount.Text := IntToStr(nL_CompCounts);

      L_IniFile.Free;
    end;
    L_StrList.Free;
end;


procedure TfrmDbInfo.btnAppendClick(Sender: TObject);
var
    nL_MaxGroup : Integer;
begin
    nL_MaxGroup := getMaxGroupNo;
    dtmMain.cdsSGSParam.Append;
    sG_Mode := INSERT_MODE;

    medGroup.Text := IntToStr(nL_MaxGroup + 1);
    medCompCode.Text := '';
    medCompName.Text := '';
    medUserID.Text := '';
    medPassword.Text := '';
    medAlias.Text := '';
    
    ChangeBtnStatus;
    medCompCode.SetFocus;
end;

procedure TfrmDbInfo.btnEditClick(Sender: TObject);
begin
    dtmMain.cdsSGSParam.Edit;
    sG_Mode := EDIT_MODE;

    ChangeBtnStatus;
    medCompCode.SetFocus;
end;

procedure TfrmDbInfo.btnDeleteClick(Sender: TObject);
var
    sL_Group,sL_TcOrSc : String;
    L_IniFile : TIniFile;
begin
    if MessageDlg(dtmMain.getLanguageSettingInfo('frmDbInfo_Msg_1_Content'),mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      sL_Group := dtmMain.cdsSGSParam.FieldByName('Group').AsString;

      dtmMain.cdsSGSParam.Delete;
      if FileExists(sG_IniFileName) then
      begin
        L_IniFile := TIniFile.Create(sG_IniFileName);

        //判別簡體繁體
        sL_TcOrSc := L_IniFile.ReadString('CURRENT_LANGUAGE_TYPE','CURRENT_LANGUAGE_TYPE','');

        L_IniFile.DeleteKey(sL_TcOrSc, 'ALIAS_' + sL_Group);
        L_IniFile.DeleteKey(sL_TcOrSc, 'USERID_' + sL_Group);
        L_IniFile.DeleteKey(sL_TcOrSc, 'PASSWORD_' + sL_Group);
        L_IniFile.DeleteKey(sL_TcOrSc, 'COMPCODE_' + sL_Group);
        L_IniFile.DeleteKey(sL_TcOrSc, 'PROCESSORIP_' + sL_Group);
        L_IniFile.DeleteKey(sL_TcOrSc, 'COMPNAME_' + sL_Group);
      end;

      L_IniFile.Free;

      ChangeBtnStatus;
    end;
end;

procedure TfrmDbInfo.btnCancelClick(Sender: TObject);
begin
    dtmMain.cdsSGSParam.Cancel;
    ChangeBtnStatus;
end;

procedure TfrmDbInfo.btnExitClick(Sender: TObject);
begin
    Close;
end;

procedure TfrmDbInfo.dsrDbLinkSetDataChange(Sender: TObject;
  Field: TField);
begin
    with dtmMain.cdsSGSParam do
    begin
      if RecordCount > 0 then
      begin
        medGroup.Text := FieldByName('Group').AsString;
        medCompCode.Text := FieldByName('CompCode').AsString;
        medCompName.Text := FieldByName('CompName').AsString;
        medUserID.Text := FieldByName('UserID').AsString;
        medPassword.Text := FieldByName('Password').AsString;
        medAlias.Text := FieldByName('Alias').AsString;

        medDataPath1.Text := FieldByName('DataPath1').AsString;
        medDataPath2.Text := FieldByName('DataPath2').AsString;
        fraYMD1.setYMD(FieldByName('FirstDate').AsString);
        medMail1.Text := FieldByName('AdministratorMail1').AsString;
        medMail2.Text := FieldByName('AdministratorMail2').AsString;
        medVersion.Text := FieldByName('Version').AsString;
        medCasID.Text := FieldByName('CasID').AsString;
        medSrcCode.Text := FieldByName('SrcCode').AsString;
        medSrcIP.Text := FieldByName('SrcIP').AsString;
        medSrcType.Text := FieldByName('SrcType').AsString;
        medDstCode.Text := FieldByName('DstCode').AsString;
        medDstIP.Text := FieldByName('DstIP').AsString;
        medDstType.Text := FieldByName('DstType').AsString;

      end;

    end;
end;

procedure TfrmDbInfo.FormCreate(Sender: TObject);
begin
    ReadIni;
    ChangeBtnStatus;
    setLanguageString;
end;

function TfrmDbInfo.getMaxGroupNo: Integer;
begin
{
    if TClientDataSet(dtmMain.cdsDbLinkSet).IndexName = 'TmpIndex' then
      TClientDataSet(dtmMain.cdsDbLinkSet).DeleteIndex('TmpIndex');
    TClientDataSet(dtmMain.cdsDbLinkSet).AddIndex('TmpIndex', 'GROUP', [ixCaseInsensitive],'','',0);
    TClientDataSet(dtmMain.cdsDbLinkSet).IndexName := 'TmpIndex';

    dtmMain.cdsDbLinkSet.Last;
      Result := dtmMain.cdsDbLinkSet.FieldByName('Group').AsInteger;
}
    Result := dtmMain.cdsSGSParam.RecordCount;
end;

procedure TfrmDbInfo.FormShow(Sender: TObject);
begin
    if dtmMain.cdsSGSParam.RecordCount = 0 then
      btnAppend.Enabled := false;
end;

function TfrmDbInfo.isDataOK: Boolean;
begin
    if Trim(medDbCount.Text) = '' then
    begin
      medDbCount.SetFocus;
      Result := false;
      MessageDlg(dtmMain.getLanguageSettingInfo('frmDbInfo_Msg_2_Content'),mtError, [mbOK],0);
      exit;
    end;

    if Trim(medCompCode.Text) = '' then
    begin
      medCompCode.SetFocus;
      Result := false;
      MessageDlg(dtmMain.getLanguageSettingInfo('frmDbInfo_Msg_2_Content'),mtError, [mbOK],0);
      exit;
    end;

    if Trim(medCompName.Text) = '' then
    begin
      medCompName.SetFocus;
      Result := false;
      MessageDlg(dtmMain.getLanguageSettingInfo('frmDbInfo_Msg_2_Content'),mtError, [mbOK],0);
      exit;
    end;

    if Trim(medUserID.Text) = '' then
    begin
      medUserID.SetFocus;
      Result := false;
      MessageDlg(dtmMain.getLanguageSettingInfo('frmDbInfo_Msg_2_Content'),mtError, [mbOK],0);
      exit;
    end;

    if Trim(medPassword.Text) = '' then
    begin
      medPassword.SetFocus;
      Result := false;
      MessageDlg(dtmMain.getLanguageSettingInfo('frmDbInfo_Msg_2_Content'),mtError, [mbOK],0);
      exit;
    end;

    if Trim(medAlias.Text) = '' then
    begin
      medAlias.SetFocus;
      Result := false;
      MessageDlg(dtmMain.getLanguageSettingInfo('frmDbInfo_Msg_2_Content'),mtError, [mbOK],0);
      exit;
    end;


    if Trim(medCasID.Text) = '' then
    begin
      medCasID.SetFocus;
      Result := false;
      MessageDlg(dtmMain.getLanguageSettingInfo('frmDbInfo_Msg_2_Content'),mtError, [mbOK],0);
      exit;
    end;


    if Trim(TUstr.replaceStr(fraYMD1.getYMD,'/','')) = '' then
    begin
      fraYMD1.SetFocus;
      Result := false;
      MessageDlg(dtmMain.getLanguageSettingInfo('frmDbInfo_Msg_2_Content'),mtError, [mbOK],0);
      exit;
    end;

    if Trim(medDataPath1.Text) = '' then
    begin
      medDataPath1.SetFocus;
      Result := false;
      MessageDlg(dtmMain.getLanguageSettingInfo('frmDbInfo_Msg_2_Content'),mtError, [mbOK],0);
      exit;
    end;

    if Trim(medDataPath2.Text) = '' then
    begin
      medDataPath2.SetFocus;
      Result := false;
      MessageDlg(dtmMain.getLanguageSettingInfo('frmDbInfo_Msg_2_Content'),mtError, [mbOK],0);
      exit;
    end;

    if Trim(medVersion.Text) = '' then
    begin
      medVersion.SetFocus;
      Result := false;
      MessageDlg(dtmMain.getLanguageSettingInfo('frmDbInfo_Msg_2_Content'),mtError, [mbOK],0);
      exit;
    end;

    if Trim(medSrcCode.Text) = '' then
    begin
      medSrcCode.SetFocus;
      Result := false;
      MessageDlg(dtmMain.getLanguageSettingInfo('frmDbInfo_Msg_2_Content'),mtError, [mbOK],0);
      exit;
    end;

    if Trim(medSrcIP.Text) = '' then
    begin
      medSrcIP.SetFocus;
      Result := false;
      MessageDlg(dtmMain.getLanguageSettingInfo('frmDbInfo_Msg_2_Content'),mtError, [mbOK],0);
      exit;
    end;

    if Trim(medSrcType.Text) = '' then
    begin
      medSrcType.SetFocus;
      Result := false;
      MessageDlg(dtmMain.getLanguageSettingInfo('frmDbInfo_Msg_2_Content'),mtError, [mbOK],0);
      exit;
    end;

    if Trim(medDstCode.Text) = '' then
    begin
      medDstCode.SetFocus;
      Result := false;
      MessageDlg(dtmMain.getLanguageSettingInfo('frmDbInfo_Msg_2_Content'),mtError, [mbOK],0);
      exit;
    end;

    if Trim(medDstIP.Text) = '' then
    begin
      medDstIP.SetFocus;
      Result := false;
      MessageDlg(dtmMain.getLanguageSettingInfo('frmDbInfo_Msg_2_Content'),mtError, [mbOK],0);
      exit;
    end;

    if Trim(medDstType.Text) = '' then
    begin
      medDstType.SetFocus;
      Result := false;
      MessageDlg(dtmMain.getLanguageSettingInfo('frmDbInfo_Msg_2_Content'),mtError, [mbOK],0);
      exit;
    end;

    Result := true;
end;

procedure TfrmDbInfo.btnSaveClick(Sender: TObject);
var
    sL_Group,sL_CompCode,sL_CompName,sL_UserID : String;
    sL_Password,sL_Alias,sL_ProcessorIP,sL_TcOrSc : String;
    sL_DataPath1,sL_DataPath2,sL_FirstDate,sL_Mail1,sL_Mail2,sL_Version,sL_CasID : String;
    sL_SrcCode,sL_SrcIP,sL_SrcType,sL_DstCode,sL_DstIP,sL_DstType : String;
    L_IniFile : TIniFile;
    nL_DbCount : Integer;
    L_Intf : _Encrypt;
begin
    if isDataOK then
    begin
      nL_DbCount := StrToInt(Trim(medDbCount.Text));
      sL_Group := Trim(medGroup.Text);
      sL_CompCode := Trim(medCompCode.Text);
      sL_CompName := Trim(medCompName.Text);
      sL_UserID := Trim(medUserID.Text);
      sL_Password := Trim(medPassword.Text);
      sL_Alias := Trim(medAlias.Text);

      sL_DataPath1 := Trim(medDataPath1.Text);
      sL_DataPath2 := Trim(medDataPath2.Text);
      sL_FirstDate := Trim(fraYMD1.getYMD);
      sL_Mail1 := Trim(medMail1.Text);
      sL_Mail2 := Trim(medMail2.Text);
      sL_Version := Trim(medVersion.Text);
      sL_CasID := Trim(medCasID.Text);
      sL_SrcCode := Trim(medSrcCode.Text);
      sL_SrcIP := Trim(medSrcIP.Text);
      sL_SrcType := Trim(medSrcType.Text);
      sL_DstCode := Trim(medDstCode.Text);
      sL_DstIP := Trim(medDstIP.Text);
      sL_DstType := Trim(medDstType.Text);

      //新增或修改到 CDS
      with dtmMain.cdsSGSParam do
      begin

        if nL_DbCount = 0then
        begin
          medDbCount.SetFocus;
          MessageDlg(dtmMain.getLanguageSettingInfo('frmDbInfo_Msg_3_Content'),mtError, [mbOK],0);
          exit;
        end;

        if nL_DbCount>dtmMain.cdsSGSParam.RecordCount then
        begin
          medDbCount.SetFocus;
          MessageDlg(dtmMain.getLanguageSettingInfo('frmDbInfo_Msg_4_Content'),mtError, [mbOK],0);
          exit;
        end;




        if sG_Mode = INSERT_MODE then      //新增模式
          Append
        else if sG_Mode = EDIT_MODE then //修改模式
          Edit;

        FieldByName('Group').AsString := sL_Group;
        FieldByName('CompCode').AsString := sL_CompCode;
        FieldByName('CompName').AsString := sL_CompName;
        FieldByName('UserID').AsString := sL_UserID;
        FieldByName('Password').AsString := sL_Password;
        FieldByName('Alias').AsString := sL_Alias;

        FieldByName('DataPath1').AsString := sL_DataPath1;
        FieldByName('DataPath2').AsString := sL_DataPath2;
        FieldByName('FirstDate').AsString := sL_FirstDate;
        FieldByName('AdministratorMail1').AsString := sL_Mail1;
        FieldByName('AdministratorMail2').AsString := sL_Mail2;
        FieldByName('Version').AsString := sL_Version;
        FieldByName('CasID').AsString := sL_CasID;
        FieldByName('SrcCode').AsString := sL_SrcCode;
        FieldByName('SrcIP').AsString := sL_SrcIP;
        FieldByName('SrcType').AsString := sL_SrcType;
        FieldByName('DstCode').AsString := sL_DstCode;
        FieldByName('DstIP').AsString := sL_DstIP;
        FieldByName('DstType').AsString := sL_DstType;

        Post;
      end;


      //新增或修改到 Ini
      if FileExists(sG_IniFileName) then
      begin
        L_IniFile := TIniFile.Create(sG_IniFileName);

        //判別簡體繁體
        sL_TcOrSc := L_IniFile.ReadString('CURRENT_LANGUAGE_TYPE','CURRENT_LANGUAGE_TYPE','');

        L_IniFile.WriteString(sL_TcOrSc, 'COMP_COUNT',IntToStr(nL_DbCount));


        L_IniFile.WriteString(sL_TcOrSc, 'ALIAS_' + sL_Group,sL_Alias);
        L_IniFile.WriteString(sL_TcOrSc, 'USERID_' + sL_Group,sL_UserID);
        L_IniFile.WriteString(sL_TcOrSc, 'PASSWORD_' + sL_Group,sL_Password);
        L_IniFile.WriteString(sL_TcOrSc, 'COMPCODE_' + sL_Group,sL_CompCode);
        L_IniFile.WriteString(sL_TcOrSc, 'COMPNAME_' + sL_Group,sL_CompName);

        L_IniFile.WriteString(sL_TcOrSc, 'DATAPATH1_' + sL_Group,sL_DataPath1);
        L_IniFile.WriteString(sL_TcOrSc, 'DATAPATH2_' + sL_Group,sL_DataPath2);
        L_IniFile.WriteString(sL_TcOrSc, 'FIRSTDATE_' + sL_Group,sL_FirstDate);

        if sL_Mail1 <> '' then
          L_IniFile.WriteString(sL_TcOrSc, 'ADMINISTRATORMAIL1_' + sL_Group,sL_Mail1);

        if sL_Mail2 <> '' then
          L_IniFile.WriteString(sL_TcOrSc, 'ADMINISTRATORMAIL2_' + sL_Group,sL_Mail2);

        L_IniFile.WriteString(sL_TcOrSc, 'VERSION_' + sL_Group,sL_Version);
        L_IniFile.WriteString(sL_TcOrSc, 'CASID_' + sL_Group,sL_CasID);
        L_IniFile.WriteString(sL_TcOrSc, 'SRCCODE_' + sL_Group,sL_SrcCode);
        L_IniFile.WriteString(sL_TcOrSc, 'SRCIP_' + sL_Group,sL_SrcIP);
        L_IniFile.WriteString(sL_TcOrSc, 'SRCTYPE_' + sL_Group,sL_SrcType);
        L_IniFile.WriteString(sL_TcOrSc, 'DSTCODE_' + sL_Group,sL_DstCode);
        L_IniFile.WriteString(sL_TcOrSc, 'DSTIP_' + sL_Group,sL_DstIP);
        L_IniFile.WriteString(sL_TcOrSc, 'DSTTYPE_' + sL_Group,sL_DstType);
      end;

      L_IniFile.Free;

      ChangeBtnStatus;
    end;

    saveTempIni(L_IniFile);
end;

procedure TfrmDbInfo.SpeedButton1Click(Sender: TObject);
var
    sL_Path : String;
begin
    if OpenDialog1.Execute then
    begin
      sL_Path := ExtractFileDir(OpenDialog1.FileName);
      medDataPath1.Text := sL_Path;
      //ReadIni;
      //ChangeBtnStatus;
    end;
end;

procedure TfrmDbInfo.saveTempIni(I_IniFile: TIniFile);
var
    L_StrList, L_TmpStrList : TStringList;
    sL_ExeFileName, sL_ExePath ,sL_IniFileName,sL_TempIniFileName : String;
    ii : Integer;
    L_Intf : _Encrypt;
    sL_EncKey : WideString;
    sL_Temp : WideString;
begin
    sL_EncKey := 'CS';
    L_Intf := CoEncrypt.Create;

    L_StrList := TStringList.Create;
    L_TmpStrList := TStringList.Create;

    L_StrList.LoadFromFile(sG_IniFileName);
    for ii:=0 to L_StrList.Count-1 do
    begin
       if (Copy(L_StrList.Strings[ii],1,1)=';') or (L_StrList.Strings[ii]='')
          or (Copy(L_StrList.Strings[ii],1,8)='COMPCODE') or (Copy(L_StrList.Strings[ii],1,8)='COMPNAME')
          or (Copy(L_StrList.Strings[ii],1,8)='DATAPATH') or (Copy(L_StrList.Strings[ii],1,23)='[CURRENT_LANGUAGE_TYPE]')
          or (Copy(L_StrList.Strings[ii],1,21)='CURRENT_LANGUAGE_TYPE') or (Copy(L_StrList.Strings[ii],1,9)='[SYSINFO]')
          or (Copy(L_StrList.Strings[ii],1,19)='NIGHT_RUN_EXEC_TIME') or (Copy(L_StrList.Strings[ii],1,14)='CANCEL_CHANNEL')
          or (Copy(L_StrList.Strings[ii],1,9)='HTTP_PORT') or (Copy(L_StrList.Strings[ii],1,11)='WINZIP_PATH') then
       begin
         //不加密
         sL_Temp := L_StrList.Strings[ii];
         L_TmpStrList.Add(sL_Temp);
       end
       else
       begin
         //加密
         sL_Temp := L_StrList.Strings[ii];
         L_TmpStrList.Add(L_Intf.Encrypt(sL_Temp));
       end;

    end;
    L_TmpStrList.SaveToFile(dtmMain.sG_ExePath + '\' + SYS_INI_FILE_NAME);
    L_TmpStrList.Free;

    L_Intf := nil;
end;

procedure TfrmDbInfo.FormClose(Sender: TObject; var Action: TCloseAction);
begin
    //刪除TempIni檔案
    dtmMain.DeleteTmpIniFile;
end;

procedure TfrmDbInfo.setLanguageString;
begin
    Caption := dtmMain.getLanguageSettingInfo('frmDbInfo_Caption');
    Label0.Caption := dtmMain.getLanguageSettingInfo('frmDbInfo_Label0_Caption');
    btnAppend.Caption := dtmMain.getLanguageSettingInfo('frmDbInfo_btnAppend_Caption');
    btnEdit.Caption := dtmMain.getLanguageSettingInfo('frmDbInfo_btnEdit_Caption');
    btnDelete.Caption := dtmMain.getLanguageSettingInfo('frmDbInfo_btnDelete_Caption');
    btnCancel.Caption := dtmMain.getLanguageSettingInfo('frmDbInfo_btnCancel_Caption');
    btnSave.Caption := dtmMain.getLanguageSettingInfo('frmDbInfo_btnSave_Caption');
    btnExit.Caption := dtmMain.getLanguageSettingInfo('frmDbInfo_btnExit_Caption');

    Label4.Caption := dtmMain.getLanguageSettingInfo('frmDbInfo_Label4_Caption');
    Label5.Caption := dtmMain.getLanguageSettingInfo('frmDbInfo_Label5_Caption');
    Label6.Caption := dtmMain.getLanguageSettingInfo('frmDbInfo_Label6_Caption');
    Label7.Caption := dtmMain.getLanguageSettingInfo('frmDbInfo_Label7_Caption');
    Label8.Caption := dtmMain.getLanguageSettingInfo('frmDbInfo_Label8_Caption');
    Label9.Caption := dtmMain.getLanguageSettingInfo('frmDbInfo_Label9_Caption');
    Label10.Caption := dtmMain.getLanguageSettingInfo('frmDbInfo_Label10_Caption');
    Label11.Caption := dtmMain.getLanguageSettingInfo('frmDbInfo_Label11_Caption');
    Label12.Caption := dtmMain.getLanguageSettingInfo('frmDbInfo_Label12_Caption');
    Label13.Caption := dtmMain.getLanguageSettingInfo('frmDbInfo_Label13_Caption');
    Label14.Caption := dtmMain.getLanguageSettingInfo('frmDbInfo_Label14_Caption');
    Label21.Caption := dtmMain.getLanguageSettingInfo('frmDbInfo_Label21_Caption');
    Label15.Caption := dtmMain.getLanguageSettingInfo('frmDbInfo_Label15_Caption');
    Label16.Caption := dtmMain.getLanguageSettingInfo('frmDbInfo_Label16_Caption');
    Label17.Caption := dtmMain.getLanguageSettingInfo('frmDbInfo_Label17_Caption');
    Label18.Caption := dtmMain.getLanguageSettingInfo('frmDbInfo_Label18_Caption');
    Label19.Caption := dtmMain.getLanguageSettingInfo('frmDbInfo_Label19_Caption');
    Label20.Caption := dtmMain.getLanguageSettingInfo('frmDbInfo_Label20_Caption');
    Label22.Caption := dtmMain.getLanguageSettingInfo('frmDbInfo_Label22_Caption');

    dbgParam.Columns[0].Title.Caption := Label4.Caption;
    dbgParam.Columns[1].Title.Caption := Label8.Caption;    
    dbgParam.Columns[2].Title.Caption := Label7.Caption;
    dbgParam.Columns[3].Title.Caption := Label6.Caption;
    dbgParam.Columns[4].Title.Caption := Label9.Caption;
    dbgParam.Columns[5].Title.Caption := Label11.Caption;
    dbgParam.Columns[6].Title.Caption := Label12.Caption;
    dbgParam.Columns[7].Title.Caption := Label22.Caption;

    dbgParam.Columns[8].Title.Caption := Label13.Caption;
    dbgParam.Columns[9].Title.Caption := Label14.Caption;
    dbgParam.Columns[10].Title.Caption := Label21.Caption;
    dbgParam.Columns[11].Title.Caption := Label10.Caption;

    dbgParam.Columns[12].Title.Caption := Label15.Caption;
    dbgParam.Columns[13].Title.Caption := Label16.Caption;
    dbgParam.Columns[14].Title.Caption := Label17.Caption;

    dbgParam.Columns[15].Title.Caption := Label18.Caption;
    dbgParam.Columns[16].Title.Caption := Label19.Caption;
    dbgParam.Columns[17].Title.Caption := Label20.Caption;
end;

procedure TfrmDbInfo.SpeedButton2Click(Sender: TObject);
var
    sL_Path : String;
begin
    if OpenDialog1.Execute then
    begin
      sL_Path := ExtractFileDir(OpenDialog1.FileName);
      medDataPath2.Text := sL_Path;
    end;
end;

end.


