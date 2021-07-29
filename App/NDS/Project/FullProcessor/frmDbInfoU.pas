unit frmDbInfoU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, Grids, DBGrids, StdCtrls, Buttons, DB, IniFiles, Mask,
  DBCtrls, DBClient;


const
    INSERT_MODE = 'I';
    EDIT_MODE = 'E';

type
  TfrmDbInfo = class(TForm)
    Panel1: TPanel;
    Label7: TLabel;
    Panel3: TPanel;
    btnAppend: TBitBtn;
    btnEdit: TBitBtn;
    btnDelete: TBitBtn;
    btnCancel: TBitBtn;
    btnSave: TBitBtn;
    Panel2: TPanel;
    pnlMultiData: TPanel;
    DBGrid1: TDBGrid;
    dsrDbLinkSet: TDataSource;
    SpeedButton1: TSpeedButton;
    StaticText1: TStaticText;
    edtIniFileName: TEdit;
    pnlSingleData: TPanel;
    OpenDialog1: TOpenDialog;
    Panel4: TPanel;
    medDbCount: TMaskEdit;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label5: TLabel;
    Label4: TLabel;
    Label6: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    medCompCode: TMaskEdit;
    medGroup: TMaskEdit;
    medProcessorIP: TEdit;
    medAlias: TEdit;
    medPassword: TEdit;
    medUserID: TEdit;
    medCompName: TEdit;
    btnExit: TBitBtn;
    Label12: TLabel;
    edtProcessorIP: TEdit;
    Label13: TLabel;
    edtDispatcherIP: TEdit;
    procedure BitBtn1Click(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure btnAppendClick(Sender: TObject);
    procedure btnEditClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure dsrDbLinkSetDataChange(Sender: TObject; Field: TField);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
  private
    { Private declarations }
//    nG_DbCount : Integer;
    sG_IniFileName,sG_Mode : String;
    procedure ReadIni;
    procedure ChangeBtnStatus;
    function getMaxGroupNo : Integer;
    function isDataOK : Boolean;
  public
    { Public declarations }
  end;

var
  frmDbInfo: TfrmDbInfo;

implementation

uses dtmMainU;

{$R *.dfm}

procedure TfrmDbInfo.BitBtn1Click(Sender: TObject);
begin
    Close;
end;

procedure TfrmDbInfo.ChangeBtnStatus;
begin
     with dtmMain.cdsDbLinkSet do
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
          if dtmMain.cdsDbLinkSet.RecordCount>0 then
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
    ii, nL_CompCode, nL_Key ,nL_DbCount : Integer;
    L_StrList : TStringList;
    sL_ProcessorIP,sL_DispatcherIP : String;
begin
    L_StrList := TStringList.Create;
    sG_IniFileName := edtIniFileName.Text;

    if FileExists(sG_IniFileName) then
    begin
      L_IniFile := TIniFile.Create(sG_IniFileName);
      L_IniFile.ReadSectionValues('DBINFO',L_StrList);

      dtmMain.initialCdsDbLinkSet(L_StrList);

      nL_DbCount := L_IniFile.ReadInteger('DBINFO','DB_COUNT',0);
      medDbCount.Text := IntToStr(nL_DbCount);

      edtProcessorIP.Text := L_IniFile.ReadString('CA_DISPATCHER_PROCESSOR','PROCESSOR_IP','');
      edtDispatcherIP.Text := L_IniFile.ReadString('CA_DISPATCHER_PROCESSOR','DISPATCHER_IP','');

      L_IniFile.Free;
    end;
    L_StrList.Free;
end;

procedure TfrmDbInfo.SpeedButton1Click(Sender: TObject);
begin
    if OpenDialog1.Execute then
    begin
      edtIniFileName.Text := OpenDialog1.FileName;
      ReadIni;
      ChangeBtnStatus;
    end;
end;

procedure TfrmDbInfo.btnAppendClick(Sender: TObject);
var
    nL_MaxGroup : Integer;
begin
    nL_MaxGroup := getMaxGroupNo;
    dtmMain.cdsDbLinkSet.Append;
    sG_Mode := INSERT_MODE;

    medGroup.Text := IntToStr(nL_MaxGroup + 1);
    medCompCode.Text := '';
    medCompName.Text := '';
    medUserID.Text := '';
    medPassword.Text := '';
    medAlias.Text := '';
    medProcessorIP.Text := '';
    
    ChangeBtnStatus;
    medCompCode.SetFocus;
end;

procedure TfrmDbInfo.btnEditClick(Sender: TObject);
begin
    dtmMain.cdsDbLinkSet.Edit;
    sG_Mode := EDIT_MODE;

    ChangeBtnStatus;
    medCompCode.SetFocus;
end;

procedure TfrmDbInfo.btnDeleteClick(Sender: TObject);
var
    sL_Group : String;
    L_IniFile : TIniFile;
begin
    if MessageDlg('是否確認刪除?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      sL_Group := dtmMain.cdsDbLinkSet.FieldByName('Group').AsString;

      dtmMain.cdsDbLinkSet.Delete;
      if FileExists(sG_IniFileName) then
      begin
        L_IniFile := TIniFile.Create(sG_IniFileName);

        L_IniFile.DeleteKey('DBINFO', 'ALIAS_' + sL_Group);
        L_IniFile.DeleteKey('DBINFO', 'USERID_' + sL_Group);
        L_IniFile.DeleteKey('DBINFO', 'PASSWORD_' + sL_Group);
        L_IniFile.DeleteKey('DBINFO', 'COMPCODE_' + sL_Group);
        L_IniFile.DeleteKey('DBINFO', 'PROCESSORIP_' + sL_Group);
        L_IniFile.DeleteKey('DBINFO', 'COMPNAME_' + sL_Group);
      end;

      L_IniFile.Free;

      ChangeBtnStatus;
    end;
end;

procedure TfrmDbInfo.btnCancelClick(Sender: TObject);
begin
    dtmMain.cdsDbLinkSet.Cancel;
    ChangeBtnStatus;
end;

procedure TfrmDbInfo.btnExitClick(Sender: TObject);
begin
    Close;
end;

procedure TfrmDbInfo.dsrDbLinkSetDataChange(Sender: TObject;
  Field: TField);
begin
    with dtmMain.cdsDbLinkSet do
    begin
      if RecordCount > 0 then
      begin
        medGroup.Text := FieldByName('Group').AsString;
        medCompCode.Text := FieldByName('CompCode').AsString;
        medCompName.Text := FieldByName('CompName').AsString;
        medUserID.Text := FieldByName('UserID').AsString;
        medPassword.Text := FieldByName('Password').AsString;
        medAlias.Text := FieldByName('Alias').AsString;
        medProcessorIP.Text := FieldByName('ProcessorIP').AsString;
      end;

    end;
end;

procedure TfrmDbInfo.FormCreate(Sender: TObject);
begin
    if not dtmMain.cdsDbLinkSet.Active then
      dtmMain.cdsDbLinkSet.CreateDataSet;

    ChangeBtnStatus;
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
    Result := dtmMain.cdsDbLinkSet.RecordCount;
end;

procedure TfrmDbInfo.FormShow(Sender: TObject);
begin
    if Trim(edtIniFileName.Text)='' then
      btnAppend.Enabled := false;

    frmDbInfo.Caption := 'NDS GateWay --[資料庫連線設定]';;
end;

function TfrmDbInfo.isDataOK: Boolean;
begin
    if Trim(medDbCount.Text) = '' then
    begin
      medDbCount.SetFocus;
      Result := false;
      MessageDlg('請輸入要啟用幾組連線',mtError, [mbOK],0);
      exit;
    end;

    if Trim(medCompCode.Text) = '' then
    begin
      medCompCode.SetFocus;
      Result := false;
      MessageDlg('請輸入公司別',mtError, [mbOK],0);
      exit;
    end;

    if Trim(medCompName.Text) = '' then
    begin
      medCompName.SetFocus;
      Result := false;
      MessageDlg('請輸入公司名稱',mtError, [mbOK],0);
      exit;
    end;

    if Trim(medUserID.Text) = '' then
    begin
      medUserID.SetFocus;
      Result := false;
      MessageDlg('請輸入帳號',mtError, [mbOK],0);
      exit;
    end;

    if Trim(medPassword.Text) = '' then
    begin
      medPassword.SetFocus;
      Result := false;
      MessageDlg('請輸入密碼',mtError, [mbOK],0);
      exit;
    end;

    if Trim(medAlias.Text) = '' then
    begin
      medAlias.SetFocus;
      Result := false;
      MessageDlg('請輸入Alias',mtError, [mbOK],0);
      exit;
    end;

    if Trim(medProcessorIP.Text) = '' then
    begin
      medProcessorIP.SetFocus;
      Result := false;
      MessageDlg('請輸入ProcessorIP',mtError, [mbOK],0);
      exit;
    end;

    Result := true;
end;

procedure TfrmDbInfo.btnSaveClick(Sender: TObject);
var
    sL_Group,sL_CompCode,sL_CompName,sL_UserID : String;
    sL_Password,sL_Alias,sL_ProcessorIP : String;
    L_IniFile : TIniFile;
    nL_DbCount : Integer;
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
      sL_ProcessorIP := Trim(medProcessorIP.Text);


      //新增或修改到 CDS
      with dtmMain.cdsDbLinkSet do
      begin
        if (nL_DbCount = 0) or (nL_DbCount>StrToInt(sL_Group)) then
        begin
          medDbCount.SetFocus;
          MessageDlg('啟用資料庫連線數須介於 1 ~ ' + sL_Group + ' 之間',mtError, [mbOK],0);
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
        FieldByName('ProcessorIP').AsString := sL_ProcessorIP;
        Post;
      end;


      //新增或修改到 Ini
      if FileExists(sG_IniFileName) then
      begin
        L_IniFile := TIniFile.Create(sG_IniFileName);

        L_IniFile.WriteString('DBINFO', 'DB_COUNT',IntToStr(nL_DbCount));
        L_IniFile.WriteString('CA_DISPATCHER_PROCESSOR', 'PROCESSOR_IP',Trim(edtProcessorIP.Text));
        L_IniFile.WriteString('CA_DISPATCHER_PROCESSOR', 'DISPATCHER_IP',Trim(edtDispatcherIP.Text));

        L_IniFile.WriteString('DBINFO', 'ALIAS_' + sL_Group,sL_Alias);
        L_IniFile.WriteString('DBINFO', 'USERID_' + sL_Group,sL_UserID);
        L_IniFile.WriteString('DBINFO', 'PASSWORD_' + sL_Group,sL_Password);
        L_IniFile.WriteString('DBINFO', 'COMPCODE_' + sL_Group,sL_CompCode);
        L_IniFile.WriteString('DBINFO', 'COMPNAME_' + sL_Group,sL_CompName);
        L_IniFile.WriteString('DBINFO', 'PROCESSORIP_' + sL_Group,sL_ProcessorIP);

      end;

      L_IniFile.Free;

      ChangeBtnStatus;
    end;
end;

end.
