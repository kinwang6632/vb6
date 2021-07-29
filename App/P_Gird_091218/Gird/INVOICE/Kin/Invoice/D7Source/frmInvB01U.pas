unit frmInvB01U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, ExtCtrls, DB, Grids, DBGrids, Mask, xmldom,
  XMLIntf, msxmldom, XMLDoc, FileCtrl, ADODB,
  cxContainer, cxEdit,
  cxGroupBox, cxRadioGroup, cxPC, cxControls, Menus, cxLookAndFeelPainters,
  cxButtons, cxGraphics, cxTextEdit, cxMaskEdit, cxDropDownEdit,
  dxSkinsCore, dxSkinsDefaultPainters, dxSkinscxPCPainter;

type
  TfrmInvB01 = class(TForm)
    Panel1: TPanel;
    Label5: TLabel;
    pnlEdit: TPanel;
    pnlGrid: TPanel;
    DBGrid1: TDBGrid;
    Panel4: TPanel;
    dsrInv003: TDataSource;
    getXMLDocument: TXMLDocument;
    Stt_Show: TStaticText;
    OpenDialog1: TOpenDialog;
    MainPage: TcxPageControl;
    cxTabSheet1: TcxTabSheet;
    cxTabSheet2: TcxTabSheet;
    cxTabSheet3: TcxTabSheet;
    rgpUseBaseCompName: TcxRadioGroup;
    rgpExpAddrType: TcxRadioGroup;
    rgpPrintAddr: TcxRadioGroup;
    rgpPrintTitle: TcxRadioGroup;
    rgpIfPrintCheck: TcxRadioGroup;
    Bevel1: TBevel;
    btnAppend: TcxButton;
    btnExit: TcxButton;
    btnSave: TcxButton;
    btnCancel: TcxButton;
    btnDelete: TcxButton;
    btnEdit: TcxButton;
    Label11: TLabel;
    cmbCity: TcxComboBox;
    Label12: TLabel;
    cmbExportMode: TcxComboBox;
    Label13: TLabel;
    cmbImportNote: TcxComboBox;
    Label14: TLabel;
    txtEntryClass: TcxTextEdit;
    Label8: TLabel;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    medSysID: TMaskEdit;
    medCompName: TMaskEdit;
    cobCompType: TComboBox;
    btnRejectionInfoFile: TButton;
    medRejectionInfoFile: TMaskEdit;
    btnPInvoiceFilePath: TButton;
    medPInvoiceFilePath: TMaskEdit;
    btnInvoiceFilePath: TButton;
    medInvoiceFilePath: TMaskEdit;
    btnMediaFilePath: TButton;
    medMediaFilePath: TMaskEdit;
    btnBackupFilePath: TButton;
    medBackupFilePath: TMaskEdit;
    medCompID: TMaskEdit;
    procedure btnExitClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure btnRejectionInfoFileClick(Sender: TObject);
    procedure btnAppendClick(Sender: TObject);
    procedure btnPInvoiceFilePathClick(Sender: TObject);
    procedure btnInvoiceFilePathClick(Sender: TObject);
    procedure btnMediaFilePathClick(Sender: TObject);
    procedure btnBackupFilePathClick(Sender: TObject);
    procedure dsrInv003DataChange(Sender: TObject; Field: TField);
    procedure btnEditClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
  private
    { Private declarations }
    procedure ChangeBtnStatus;
    procedure setButtonCompetence;
    procedure initialForm;
    function LoadInv003InvParam(sI_InvParam : String) : String;
    function isDataOK : Boolean;
    procedure SetCmbCity(aValue: String);
    procedure SetCmbExportMode(aValue: String);
    procedure SetCmbImportNote(aValue: String);
  public
    { Public declarations }
  end;

var
  frmInvB01: TfrmInvB01;

implementation

uses frmMainU, dtmMainJU, dtmMainU, xmlU, cbUtilis;

{$R *.dfm}

procedure TfrmInvB01.btnExitClick(Sender: TObject);
begin
   Close;
end;

procedure TfrmInvB01.ChangeBtnStatus;
begin
     with dsrInv003.DataSet do
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
          { 第一頁籤 }
          medSysID.ReadOnly := False;
          medCompName.ReadOnly := False;
          cobCompType.Enabled := True;
          btnRejectionInfoFile.Enabled := True;
          btnPInvoiceFilePath.Enabled := True;
          btnInvoiceFilePath.Enabled := True;
          btnMediaFilePath.Enabled := True;
          btnBackupFilePath.Enabled := True;
          { 第二頁籤 }
          rgpUseBaseCompName.Properties.ReadOnly := False;
          rgpExpAddrType.Properties.ReadOnly := False;
          rgpPrintAddr.Properties.ReadOnly := False;
          rgpPrintTitle.Properties.ReadOnly := False;
          rgpIfPrintCheck.Properties.ReadOnly := False;
          { 第三頁籤 }
          cmbCity.Properties.ReadOnly := False;
          cmbExportMode.Properties.ReadOnly := False;
          cmbImportNote.Properties.ReadOnly := False;
          txtEntryClass.Properties.ReadOnly := False;
       end
       else
       begin
          if dsrInv003.DataSet.RecordCount>0 then
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
          { 第一頁籤 }
          medSysID.ReadOnly := True;
          medCompName.ReadOnly := True;
          cobCompType.Enabled := False;
          btnRejectionInfoFile.Enabled := False;
          btnPInvoiceFilePath.Enabled := False;
          btnInvoiceFilePath.Enabled := False;
          btnMediaFilePath.Enabled := False;
          btnBackupFilePath.Enabled := False;
          { 第二頁籤 }
          rgpUseBaseCompName.Properties.ReadOnly := True;
          rgpExpAddrType.Properties.ReadOnly := True;
          rgpPrintAddr.Properties.ReadOnly := True;
          rgpPrintTitle.Properties.ReadOnly := True;
          rgpIfPrintCheck.Properties.ReadOnly := True;
          { 第三頁籤 }
          cmbCity.Properties.ReadOnly := True;
          cmbExportMode.Properties.ReadOnly := True;
          cmbImportNote.Properties.ReadOnly := True;
          txtEntryClass.Properties.ReadOnly := True;
       end;
       medCompID.Enabled := ( state = dsInsert );
     end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB01.FormShow(Sender: TObject);
begin
  Self.Caption := frmMain.GetFormTitleString('B01','發票參數設定');
  dtmMainJ.getAllInv003Data;
  ChangeBtnStatus;
  setButtonCompetence;
  initialForm;
  MainPage.ActivePageIndex := 0;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB01.initialForm;
begin
  with dsrInv003.DataSet do
  begin
    FindField('CompId').DisplayLabel := '公司代號';
    FindField('CompTitle').DisplayLabel := '公司抬頭';
    FindField('InvCreating').DisplayLabel := '是否正在發票開立';
    FindField('UptTime').DisplayLabel := '異動時間';
    FindField('UptEn').DisplayLabel := '異動人員';
    FindField('CompTitle').DisplayWidth := 30;
    FindField('UptEn').DisplayWidth := 10;
    FindField('IdentifyId1').Visible := false;
    FindField('IdentifyId2').Visible := false;
    FindField('InvParam').Visible := false;
    FindField('IfPrintTitle').Visible := false;
    FindField('IfPrintAddr').Visible := false;
    FindField('IfPrintCheck').Visible := false;
    FindField('city').Visible := false;
    FindField('exportmode').Visible := false;
    FindField('importnote').Visible := false;
    FindField('entryclass').Visible := false;
    FindField('expaddrtype').Visible := false;
    if RecordCount = 0 then
    begin
      btnEdit.Enabled := False;
      btnDelete.Enabled := False;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB01.setButtonCompetence;
var
    sL_QueryCompetence,sL_InsertCompetence,sL_DeleteCompetence,sL_UpdateCompetence,sL_ExecuteCompetence : String;
begin
     //設定權限
     dtmMain.getCompetence('B01',sL_QueryCompetence,sL_InsertCompetence,sL_DeleteCompetence,sL_UpdateCompetence,sL_ExecuteCompetence);

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

procedure TfrmInvB01.btnRejectionInfoFileClick(Sender: TObject);
begin
   OpenDialog1.FileName := 'rejection1.txt' ;
   if OpenDialog1.Execute then
      medRejectionInfoFile.Text := OpenDialog1.FileName;
end;

procedure TfrmInvB01.btnAppendClick(Sender: TObject);
begin
    dtmMainJ.adoInv003Code.Append;

    rgpUseBaseCompName.ItemIndex := 0;
    rgpExpAddrType.ItemIndex := 0;
    rgpPrintAddr.ItemIndex := 0;
    rgpPrintTitle.ItemIndex := 0;

    ChangeBtnStatus;

    medCompID.Text := '';
    medSysID.Text := '';
    medCompName.Text := '';
    medRejectionInfoFile.Text := '';
    medPInvoiceFilePath.Text := '';
    medInvoiceFilePath.Text := '';
    medMediaFilePath.Text := '';
    medBackupFilePath.Text := '';

    cobCompType.ItemIndex := 0;
    rgpUseBaseCompName.ItemIndex := 0;
    rgpExpAddrType.ItemIndex := 0;
    rgpPrintAddr.ItemIndex := 0;
    rgpPrintTitle.ItemIndex := 0;

    if medSysID.CanFocus then medSysID.SetFocus;
end;

procedure TfrmInvB01.btnPInvoiceFilePathClick(Sender: TObject);
begin
   OpenDialog1.FileName := 'PCHARGE.txt' ;
   if OpenDialog1.Execute then
      medPInvoiceFilePath.Text := OpenDialog1.FileName;
end;

procedure TfrmInvB01.btnInvoiceFilePathClick(Sender: TObject);
begin
   OpenDialog1.FileName := 'LCHARGE.txt' ;
   if OpenDialog1.Execute then
      medInvoiceFilePath.Text := OpenDialog1.FileName;
end;

procedure TfrmInvB01.btnMediaFilePathClick(Sender: TObject);
var
  aPath: String;
begin
  aPath := medMediaFilePath.Text;
  if SelectDirectory( '媒體申報檔輸出路徑', EmptyStr, aPath ) then
    medMediaFilePath.Text := aPath;
end;

procedure TfrmInvB01.btnBackupFilePathClick(Sender: TObject);
var
  aPath : String;
begin
  aPath := medBackupFilePath.Text;
  if SelectDirectory( '檔案備份路徑', EmptyStr, aPath ) then
    medBackupFilePath.Text := aPath;
end;

procedure TfrmInvB01.dsrInv003DataChange(Sender: TObject; Field: TField);
var
  aInvParam : String;
begin
  if dtmMainJ.adoInv003Code.RecordCount > 0 then
  begin
    aInvParam := dtmMainJ.adoInv003Code.FieldByName('InvParam').AsString;
    LoadInv003InvParam( aInvParam );
    rgpIfPrintCheck.ItemIndex := 0;
    if ( dtmMainJ.adoInv003Code.FieldByName( 'IfPrintCheck' ).AsString = 'N' ) then
      rgpIfPrintCheck.ItemIndex := 1;
    medCompID.Text := dtmMainJ.adoInv003Code.FieldByName( 'compid' ).AsString;
    rgpExpAddrType.ItemIndex := 0;
    if ( dtmMainJ.adoInv003Code.FieldByName( 'ExpAddrType' ).AsInteger > 1 ) then
      rgpExpAddrType.ItemIndex := 1;
    {}
    SetCmbCity( dtmMainJ.adoInv003Code.FieldByName( 'city' ).AsString );
    SetCmbExportMode( dtmMainJ.adoInv003Code.FieldByName( 'exportmode' ).AsString );
    SetCmbImportNote( dtmMainJ.adoInv003Code.FieldByName( 'importnote' ).AsString );
    txtEntryClass.Text := dtmMainJ.adoInv003Code.FieldByName( 'entryclass' ).AsString;
  end
  else
  begin
    medCompID.Text := '';
    medSysID.Text := '';
    medRejectionInfoFile.Text := '';
    medPInvoiceFilePath.Text := '';
    medInvoiceFilePath.Text := '';
    medMediaFilePath.Text := '';
    medBackupFilePath.Text := '';

    rgpUseBaseCompName.ItemIndex := 0;
    rgpExpAddrType.ItemIndex := 0;
    rgpPrintAddr.ItemIndex := 0;
    rgpPrintTitle.ItemIndex := 0;
    rgpIfPrintCheck.ItemIndex := 0;

    medCompName.Text := '';
    cobCompType.ItemIndex := 0;

    {}
    cmbCity.ItemIndex := 0;
    cmbExportMode.ItemIndex := 0;
    cmbImportNote.ItemIndex := 0;
    txtEntryClass.Text := EmptyStr;
    {}

  end;
  Stt_Show.Caption := Format( '%d　/　%d', [
    dtmMainJ.adoInv003Code.RecNo,  dtmMainJ.adoInv003Code.RecordCount] );
end;

function TfrmInvB01.LoadInv003InvParam(sI_InvParam: String): String;
var
    sL_CompType,sL_CompName : String;
    L_MsgNodeList : IDOMNodeList;
    ii,jj : Integer;
    L_NodeList : IXmlNodeList;
    sL_SysID,sL_PrintTitle,sL_PrintAddr,sL_UseMis,sL_UseBaseCompName : String;
    sL_InvoiceFilePath,sL_PInvoiceFilePath,sL_MediaFilePath : String;
    sL_BackupFilePath,sL_RejectionInfoFile : String;
begin
    if sI_InvParam <> '' then
    begin
      getXMLDocument.Active := False;
      getXMLDocument.XML.Text := sI_InvParam;
      getXMLDocument.Active := True;
      try
        L_NodeList := getXMLDocument.ChildNodes;
        sL_SysID := TMyXML.getAttributeValue(L_NodeList.Nodes[1],'SysID');
        sL_PrintTitle := TMyXML.getAttributeValue(L_NodeList.Nodes[1],'PrintTitle');
        sL_PrintAddr := TMyXML.getAttributeValue(L_NodeList.Nodes[1],'PrintAddr');
        sL_UseMis := TMyXML.getAttributeValue(L_NodeList.Nodes[1],'UseMis');
        sL_UseBaseCompName := TMyXML.getAttributeValue(L_NodeList.Nodes[1],'UseBaseCompName');
        sL_InvoiceFilePath := TMyXML.getAttributeValue(L_NodeList.Nodes[1],'InvoiceFilePath');
        sL_PInvoiceFilePath := TMyXML.getAttributeValue(L_NodeList.Nodes[1],'PInvoiceFilePath');
        sL_MediaFilePath := TMyXML.getAttributeValue(L_NodeList.Nodes[1],'MediaFilePath');
        sL_BackupFilePath := TMyXML.getAttributeValue(L_NodeList.Nodes[1],'BackupFilePath');
        sL_RejectionInfoFile := TMyXML.getAttributeValue(L_NodeList.Nodes[1],'RejectionInfoFile');
        medSysID.Text := sL_SysID;
        medRejectionInfoFile.Text := sL_RejectionInfoFile;
        medPInvoiceFilePath.Text := sL_PInvoiceFilePath;
        medInvoiceFilePath.Text := sL_InvoiceFilePath;
        medMediaFilePath.Text := sL_MediaFilePath;
        medBackupFilePath.Text := sL_BackupFilePath;
        if sL_UseBaseCompName = 'Y' then
          rgpUseBaseCompName.ItemIndex := 0
        else
          rgpUseBaseCompName.ItemIndex := 1;
//        if sL_UseMis = 'Y' then
//          rgpUseMis.ItemIndex := 0
//        else
//          rgpUseMis.ItemIndex := 1;
        if sL_PrintAddr = 'Y' then
          rgpPrintAddr.ItemIndex := 0
        else
          rgpPrintAddr.ItemIndex := 1;
        if sL_PrintTitle = 'Y' then
          rgpPrintTitle.ItemIndex := 0
        else
          rgpPrintTitle.ItemIndex := 1;
        L_MsgNodeList := getXMLDocument.DOMDocument.getElementsByTagName('InvoiceParam');
        for ii:=0 to L_MsgNodeList.length-1 do
        begin
          for jj:=0 to L_MsgNodeList.item[ii].childNodes.length-1 do
          begin
            sL_CompType := TMyXML.getAttributeValue(L_MsgNodeList.item[ii].childNodes.item[jj],'Type');
            sL_CompName := TMyXML.getAttributeValue(L_MsgNodeList.item[ii].childNodes.item[jj],'Name');
            medCompName.Text := sL_CompName;
            if sL_CompType = '1' then
              cobCompType.ItemIndex := 0
            else if sL_CompType = '2' then
              cobCompType.ItemIndex := 1
            else if sL_CompType = '3' then
              cobCompType.ItemIndex := 2;
          end;
        end;
      finally
        getXMLDocument.Active := False;
      end;  
    end;

end;

procedure TfrmInvB01.btnEditClick(Sender: TObject);
begin
    dsrInv003.DataSet.Edit;
    ChangeBtnStatus;
end;

procedure TfrmInvB01.btnDeleteClick(Sender: TObject);
begin
    //if MessageDlg('請確任要刪除此筆資料?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    if ConfirmMsg( '確認刪除此筆資料?' ) then
    begin
      dsrInv003.DataSet.Delete;

      dtmMainJ.getAllInv003Data;
      ChangeBtnStatus;
    end;
    setButtonCompetence;
    initialForm;
end;

procedure TfrmInvB01.btnCancelClick(Sender: TObject);
begin
    dsrInv003.DataSet.Cancel;
    dsrInv003.DataSet.First;
    ChangeBtnStatus;

    setButtonCompetence;
    initialForm;
end;

procedure TfrmInvB01.btnSaveClick(Sender: TObject);
var
    sL_SysID,sL_PrintTitle,sL_PrintAddr,sL_UseMis,sL_UseBaseCompName : String;
    sL_InvoiceFilePath,sL_PInvoiceFilePath,sL_MediaFilePath : String;
    sL_BackupFilePath,sL_RejectionInfoFile,sL_CompType,sL_CompName : String;
    sL_InvParam, aIfPrintCheck, aCompId, aExpAddrType: String;
    aCity, aExportMode, aImportNote, aEntryClass: String;
    aComboParam: TComboBoxItemParam;
begin
    if isDataOK then
    begin
      sL_SysID := Trim(medSysID.Text);

      sL_InvoiceFilePath := Trim(medInvoiceFilePath.Text);
      sL_PInvoiceFilePath := Trim(medPInvoiceFilePath.Text);
      sL_MediaFilePath := Trim(medMediaFilePath.Text);
      sL_BackupFilePath := Trim(medBackupFilePath.Text);
      sL_RejectionInfoFile := Trim(medRejectionInfoFile.Text);
      sL_CompName := Trim(medCompName.Text);

      if cobCompType.ItemIndex = 0 then
        sL_CompType := '1'
      else if cobCompType.ItemIndex = 1 then
        sL_CompType := '2'
      else if cobCompType.ItemIndex = 2 then
        sL_CompType := '3';

      sL_UseBaseCompName := 'N';
      if rgpUseBaseCompName.ItemIndex = 0 then sL_UseBaseCompName := 'Y';


      sL_UseMis := 'N';
      //if rgpUseMis.ItemIndex = 0 then sL_UseMis := 'Y';

      sL_PrintAddr := 'N';
      if rgpPrintAddr.ItemIndex = 0 then  sL_PrintAddr := 'Y';

      sL_PrintTitle := 'N';
      if rgpPrintTitle.ItemIndex = 0 then sL_PrintTitle := 'Y';

      aIfPrintCheck := 'N';
      if rgpIfPrintCheck.ItemIndex = 0 then aIfPrintCheck := 'Y';

      aExpAddrType := '1';
      if ( rgpExpAddrType.ItemIndex > 0 ) then aExpAddrType := '2';

      sL_InvParam := '<?xml version="1.0" encoding="Big5" ?>' +
                     '<InvoiceParam SysID="' + sL_SysID + '" PrintTitle="' + sL_PrintTitle + '" PrintAddr="' + sL_PrintAddr + '" ' +
                     'UseMis="' + sL_UseMis + '" UseBaseCompName="' + sL_UseBaseCompName + '" ' +
                     'InvoiceFilePath="' + sL_InvoiceFilePath + '" ' +
                     'PInvoiceFilePath="' + sL_PInvoiceFilePath + '" ' +
                     'MediaFilePath="' + sL_MediaFilePath + '" ' +
                     'BackupFilePath="' + sL_BackupFilePath + '" ' +
                     'RejectionInfoFile="' + sL_RejectionInfoFile + '"> ' +
                     '<Comp Type="' + sL_CompType + '" Name="' + sL_CompName + '"></Comp></InvoiceParam>';

      aCompId := Trim( medCompID.Text );
      {}
      ZeroMemory( @aComboParam, SizeOf( aComboParam ) );
      GetCxComboBoxItemValue( cmbCity, aComboParam );
      aCity := aComboParam.KeyValue;
      {}
      ZeroMemory( @aComboParam, SizeOf( aComboParam ) );
      GetCxComboBoxItemValue( cmbExportMode, aComboParam );
      aExportMode := aComboParam.KeyValue;
      {}
      ZeroMemory( @aComboParam, SizeOf( aComboParam ) );
      GetCxComboBoxItemValue( cmbImportNote, aComboParam );
      aImportNote := aComboParam.KeyValue;
      {}
      aEntryClass := txtEntryClass.Text;
      {}
      dtmMainJ.adoInv003Code.FieldByName('InvParam').AsString := sL_InvParam;
      dtmMainJ.adoInv003Code.FieldByName('IdentifyId1').AsString := IDENTIFYID1;
      dtmMainJ.adoInv003Code.FieldByName('IdentifyId2').AsString := IDENTIFYID2;

      if dtmMainJ.adoInv003Code.State = dsInsert then
        dtmMainJ.adoInv003Code.FieldByName('CompId').AsString := aCompId;

      dtmMainJ.adoInv003Code.FieldByName('IfPrintTitle').AsString := sL_PrintTitle;
      dtmMainJ.adoInv003Code.FieldByName('IfPrintAddr').AsString := sL_PrintAddr;
      dtmMainJ.adoInv003Code.FieldByName('CompTitle').AsString := sL_CompName;

      dtmMainJ.adoInv003Code.FieldByName('UptTime').AsString := DateTimeToStr(now);
      dtmMainJ.adoInv003Code.FieldByName('UptEn').AsString := dtmMain.getLoginUser;
      dtmMainJ.adoInv003Code.FieldByName( 'IfPrintCheck' ).AsString := aIfPrintCheck;
      dtmMainJ.adoInv003Code.FieldByName( 'ExpAddrType' ).AsString := aExpAddrType;
      dtmMainJ.adoInv003Code.FieldByName( 'City' ).AsString := aCity;
      dtmMainJ.adoInv003Code.FieldByName( 'ExportMode' ).AsString := aExportMode;
      dtmMainJ.adoInv003Code.FieldByName( 'ImportNote' ).AsString := aImportNote;
      dtmMainJ.adoInv003Code.FieldByName( 'EntryClass' ).AsString := aEntryClass;
      {}

      dsrInv003.DataSet.Post;

      dtmMainJ.getAllInv003Data;
      ChangeBtnStatus;
      setButtonCompetence;
      initialForm;
    end;
end;

function TfrmInvB01.isDataOK: Boolean;
var
    sL_CompTitle : String;
begin
    Result := true;
    sL_CompTitle := Trim(medCompName.Text);

    if sL_CompTitle = '' then
    begin
      WarningMsg( '請輸入公司名稱' );
      MainPage.ActivePageIndex := 0;
      if medCompName.CanFocus then medCompName.SetFocus;
      Result := false;
      Exit;
    end;

    if Trim(medCompID.Text) = '' then
    begin
      WarningMsg( '請輸入公司代碼' );
      MainPage.ActivePageIndex := 0;
      if medCompID.CanFocus then medCompID.SetFocus;
      Result := False;
      Exit;
    end;

end;

procedure TfrmInvB01.SetCmbCity(aValue: String);
var
  aIndex: Integer;
begin
  cmbCity.ItemIndex := 0;
  if ( aValue <> EmptyStr ) then
    SetCxComboBoxItemValue( cmbCity, aValue );
end;

procedure TfrmInvB01.SetCmbExportMode(aValue: String);
begin
  cmbExportMode.ItemIndex := 0;
  if ( aValue <> EmptyStr ) then
    SetCxComboBoxItemValue( cmbExportMode, aValue );
end;

procedure TfrmInvB01.SetCmbImportNote(aValue: String);
begin
  cmbImportNote.ItemIndex := 0;
  if ( aValue <> EmptyStr ) then
    SetCxComboBoxItemValue( cmbImportNote, aValue );
end;

end.
