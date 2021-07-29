unit frmInvB01U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, ExtCtrls, DB, Grids, DBGrids, Mask, xmldom,
  XMLIntf, msxmldom, XMLDoc, FileCtrl, ADODB,
  cxContainer, cxEdit, cxCheckBox,
  cxGroupBox, cxRadioGroup, cxPC, cxControls, Menus, cxLookAndFeelPainters,
  cxButtons, cxGraphics, cxTextEdit, cxMaskEdit, cxDropDownEdit,
  dxSkinsCore, dxSkinsDefaultPainters, dxSkinscxPCPainter, cxLookAndFeels,
  dxSkinBlack, dxSkinBlue, dxSkinCaramel, dxSkinCoffee;

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
    Label15: TLabel;
    medEInvoiceFilePath: TMaskEdit;
    btnEInvoiceFilePath: TButton;
    Label16: TLabel;
    txtBranchNo: TcxTextEdit;
    chkPrintZeroTax: TcxCheckBox;
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
    procedure btnEInvoiceFilePathClick(Sender: TObject);
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
          btnEInvoiceFilePath.Enabled := dtmMain.GetStarEInvoice;
          btnMediaFilePath.Enabled := True;
          btnBackupFilePath.Enabled := True;
          { 第二頁籤 }
          rgpUseBaseCompName.Properties.ReadOnly := False;
          rgpExpAddrType.Properties.ReadOnly := False;
          rgpPrintAddr.Properties.ReadOnly := False;
          rgpPrintTitle.Properties.ReadOnly := False;
          rgpIfPrintCheck.Properties.ReadOnly := False;
          chkPrintZeroTax.Properties.ReadOnly := False;
          { 第三頁籤 }
          cmbCity.Properties.ReadOnly := False;
          cmbExportMode.Properties.ReadOnly := False;
          cmbImportNote.Properties.ReadOnly := False;
          txtEntryClass.Properties.ReadOnly := False;
          txtBranchNo.Properties.ReadOnly := False;
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
          btnEInvoiceFilePath.Enabled := False;
          btnMediaFilePath.Enabled := False;
          btnBackupFilePath.Enabled := False;
          { 第二頁籤 }
          rgpUseBaseCompName.Properties.ReadOnly := True;
          rgpExpAddrType.Properties.ReadOnly := True;
          rgpPrintAddr.Properties.ReadOnly := True;
          rgpPrintTitle.Properties.ReadOnly := True;
          rgpIfPrintCheck.Properties.ReadOnly := True;
          chkPrintZeroTax.Properties.ReadOnly := True;
          { 第三頁籤 }
          cmbCity.Properties.ReadOnly := True;
          cmbExportMode.Properties.ReadOnly := True;
          cmbImportNote.Properties.ReadOnly := True;
          txtEntryClass.Properties.ReadOnly := True;
          txtBranchNo.Properties.ReadOnly := True;
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
    FindField('branchno').Visible := false;

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
    sL_VerifyCompetence : String;
begin
     //設定權限
     dtmMain.getCompetence('B01',sL_QueryCompetence,sL_InsertCompetence,sL_DeleteCompetence,
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

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB01.btnRejectionInfoFileClick(Sender: TObject);
begin
   OpenDialog1.FileName := 'rejection1.txt' ;
   if OpenDialog1.Execute then
      medRejectionInfoFile.Text := OpenDialog1.FileName;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB01.btnAppendClick(Sender: TObject);
begin
  dtmMainJ.adoInv003Code.Append;
  rgpUseBaseCompName.ItemIndex := 0;
  rgpExpAddrType.ItemIndex := 0;
  rgpPrintAddr.ItemIndex := 0;
  rgpPrintTitle.ItemIndex := 0;
  ChangeBtnStatus;
  medCompID.Text := EmptyStr;
  medSysID.Text := EmptyStr;
  medCompName.Text := EmptyStr;
  medRejectionInfoFile.Text := EmptyStr;
  medPInvoiceFilePath.Text := EmptyStr;
  medInvoiceFilePath.Text := EmptyStr;
  medEInvoiceFilePath.Text := EmptyStr;
  medMediaFilePath.Text := EmptyStr;
  medBackupFilePath.Text := EmptyStr;
  cobCompType.ItemIndex := 0;
  rgpUseBaseCompName.ItemIndex := 0;
  rgpExpAddrType.ItemIndex := 0;
  rgpPrintAddr.ItemIndex := 0;
  rgpPrintTitle.ItemIndex := 0;

  if medSysID.CanFocus then medSysID.SetFocus;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB01.btnPInvoiceFilePathClick(Sender: TObject);
begin
   OpenDialog1.FileName := 'PCHARGE.txt' ;
   if OpenDialog1.Execute then
      medPInvoiceFilePath.Text := OpenDialog1.FileName;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB01.btnInvoiceFilePathClick(Sender: TObject);
begin
   OpenDialog1.FileName := 'LCHARGE.txt' ;
   if OpenDialog1.Execute then
      medInvoiceFilePath.Text := OpenDialog1.FileName;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB01.btnEInvoiceFilePathClick(Sender: TObject);
var
  aPath:String;
begin
  {
 OpenDialog1.FileName := 'ECHARGE.txt' ;
 if OpenDialog1.Execute then
    medEInvoiceFilePath.Text := OpenDialog1.FileName;
 }
  aPath := medEInvoiceFilePath.Text;
  if SelectDirectory( '電子發票拋檔路徑', EmptyStr, aPath ) then
    medEInvoiceFilePath.Text := aPath;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB01.btnMediaFilePathClick(Sender: TObject);
var
  aPath: String;
begin
  aPath := medMediaFilePath.Text;
  if SelectDirectory( '媒體申報檔輸出路徑', EmptyStr, aPath ) then
    medMediaFilePath.Text := aPath;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB01.btnBackupFilePathClick(Sender: TObject);
var
  aPath : String;
begin
  aPath := medBackupFilePath.Text;
  if SelectDirectory( '檔案備份路徑', EmptyStr, aPath ) then
    medBackupFilePath.Text := aPath;
end;

{ ---------------------------------------------------------------------------- }

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
    txtBranchNo.Text := dtmMainJ.adoInv003Code.FieldByName( 'branchno' ).AsString;
  end
  else
  begin
    medCompID.Text := EmptyStr;
    medSysID.Text := EmptyStr;
    medRejectionInfoFile.Text := EmptyStr;
    medPInvoiceFilePath.Text := EmptyStr;
    medInvoiceFilePath.Text := EmptyStr;
    medEInvoiceFilePath.Text := EmptyStr;
    medMediaFilePath.Text := EmptyStr;
    medBackupFilePath.Text := EmptyStr;

    rgpUseBaseCompName.ItemIndex := 0;
    rgpExpAddrType.ItemIndex := 0;
    rgpPrintAddr.ItemIndex := 0;
    rgpPrintTitle.ItemIndex := 0;
    rgpIfPrintCheck.ItemIndex := 0;

    medCompName.Text := EmptyStr;
    cobCompType.ItemIndex := 0;

    {}
    cmbCity.ItemIndex := 0;
    cmbExportMode.ItemIndex := 0;
    cmbImportNote.ItemIndex := 0;
    txtEntryClass.Text := EmptyStr;
    txtBranchNo.Text := EmptyStr;
    {}
  end;
  Stt_Show.Caption := Format( '%d　/　%d', [
    dtmMainJ.adoInv003Code.RecNo,  dtmMainJ.adoInv003Code.RecordCount] );
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvB01.LoadInv003InvParam(sI_InvParam: String): String;
var
  aCompType, aCompName : String;
  aMsgNodeList : IDOMNodeList;
  aIndex, aIndex2 : Integer;
  aNodeList : IXmlNodeList;
  aSysID, aPrintTitle, aPrintAddr, aUseMis, aUseBaseCompName,
  aInvoiceFilePath, aPInvoiceFilePath, aMediaFilePath, aBackupFilePath,
  aRejectionInfoFile, aEInvoiceFilePath: String;
begin
    if sI_InvParam <> '' then
    begin
      getXMLDocument.Active := False;
      getXMLDocument.XML.Text := sI_InvParam;
      getXMLDocument.Active := True;
      try
        aNodeList := getXMLDocument.ChildNodes;
        aSysID := TMyXML.getAttributeValue(aNodeList.Nodes[1],'SysID');
        aPrintTitle := TMyXML.getAttributeValue(aNodeList.Nodes[1],'PrintTitle');
        aPrintAddr := TMyXML.getAttributeValue(aNodeList.Nodes[1],'PrintAddr');
        aUseMis := TMyXML.getAttributeValue(aNodeList.Nodes[1],'UseMis');
        aUseBaseCompName := TMyXML.getAttributeValue(aNodeList.Nodes[1],'UseBaseCompName');
        aInvoiceFilePath := TMyXML.getAttributeValue(aNodeList.Nodes[1],'InvoiceFilePath');
        aPInvoiceFilePath := TMyXML.getAttributeValue(aNodeList.Nodes[1],'PInvoiceFilePath');
        aEInvoiceFilePath := TMyXML.getAttributeValue(aNodeList.Nodes[1],'EInvoiceFilePath');
        aMediaFilePath := TMyXML.getAttributeValue(aNodeList.Nodes[1],'MediaFilePath');
        aBackupFilePath := TMyXML.getAttributeValue(aNodeList.Nodes[1],'BackupFilePath');
        aRejectionInfoFile := TMyXML.getAttributeValue(aNodeList.Nodes[1],'RejectionInfoFile');
        chkPrintZeroTax.Checked := TMyXML.getAttributeValue(aNodeList.Nodes[1],'PrintZeroTax')='Y';

        medSysID.Text := aSysID;
        medRejectionInfoFile.Text := aRejectionInfoFile;

        medPInvoiceFilePath.Text := aPInvoiceFilePath;
        medInvoiceFilePath.Text := aInvoiceFilePath;
        medEInvoiceFilePath.Text := EmptyStr;
        if ( dtmMain.GetStarEInvoice ) then
          medEInvoiceFilePath.Text := aEInvoiceFilePath;

        medMediaFilePath.Text := aMediaFilePath;
        medBackupFilePath.Text := aBackupFilePath;

        if aUseBaseCompName = 'Y' then
          rgpUseBaseCompName.ItemIndex := 0
        else
          rgpUseBaseCompName.ItemIndex := 1;
        {}
        if aPrintAddr = 'Y' then
          rgpPrintAddr.ItemIndex := 0
        else
          rgpPrintAddr.ItemIndex := 1;
        {}
        if aPrintTitle = 'Y' then
          rgpPrintTitle.ItemIndex := 0
        else
          rgpPrintTitle.ItemIndex := 1;
        {}
        aMsgNodeList := getXMLDocument.DOMDocument.getElementsByTagName('InvoiceParam');

        for aIndex:=0 to aMsgNodeList.length-1 do
        begin
          for aIndex2:=0 to aMsgNodeList.item[aIndex].childNodes.length-1 do
          begin
            aCompType := TMyXML.getAttributeValue(aMsgNodeList.item[aIndex].childNodes.item[aIndex2],'Type');
            aCompName := TMyXML.getAttributeValue(aMsgNodeList.item[aIndex].childNodes.item[aIndex2],'Name');
            medCompName.Text := aCompName;
            if aCompType = '1' then
              cobCompType.ItemIndex := 0
            else if aCompType = '2' then
              cobCompType.ItemIndex := 1
            else if aCompType = '3' then
              cobCompType.ItemIndex := 2;
          end;
        end;
      finally
        getXMLDocument.Active := False;
      end;
    end;

end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB01.btnEditClick(Sender: TObject);
begin
  dsrInv003.DataSet.Edit;
  ChangeBtnStatus;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB01.btnDeleteClick(Sender: TObject);
begin
  if ConfirmMsg( '確認刪除此筆資料?' ) then
  begin
    dsrInv003.DataSet.Delete;

    dtmMainJ.getAllInv003Data;
    ChangeBtnStatus;
  end;
  setButtonCompetence;
  initialForm;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB01.btnCancelClick(Sender: TObject);
begin
    dsrInv003.DataSet.Cancel;
    dsrInv003.DataSet.First;
    ChangeBtnStatus;

    setButtonCompetence;
    initialForm;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB01.btnSaveClick(Sender: TObject);
var
  aSysId, aPrintTitle, aPrintAddr, aUseMis, aUseBaseCompName,
  aInvoiceFilePath, aPInvoiceFilePath, aMediaFilePath,
  aBackupFilePath, aRejectionInfoFile, aEInvoiceFilePath,
  aCompType, aCompName, aInvParam, aIfPrintCheck, aCompId, aExpAddrType: String;
  aCity, aExportMode, aImportNote, aEntryClass, aBranchNo, aPrintZeroTax: String;
  aComboParam: TComboBoxItemParam;
begin
    if isDataOK then
    begin
      aSysId := Trim(medSysID.Text);

      aInvoiceFilePath := Trim(medInvoiceFilePath.Text);
      aPInvoiceFilePath := Trim(medPInvoiceFilePath.Text);
      aEInvoiceFilePath := EmptyStr;
      if ( dtmMain.GetStarEInvoice ) then
        aEInvoiceFilePath := Trim(medEInvoiceFilePath.Text);

      aMediaFilePath := Trim(medMediaFilePath.Text);
      aBackupFilePath := Trim(medBackupFilePath.Text);
      aRejectionInfoFile := Trim(medRejectionInfoFile.Text);
      aCompName := Trim(medCompName.Text);

      if cobCompType.ItemIndex = 0 then
        aCompType := '1'
      else if cobCompType.ItemIndex = 1 then
        aCompType := '2'
      else if cobCompType.ItemIndex = 2 then
        aCompType := '3';

      aUseBaseCompName := 'N';
      if rgpUseBaseCompName.ItemIndex = 0 then aUseBaseCompName := 'Y';


      aUseMis := 'N';
      //if rgpUseMis.ItemIndex = 0 then sL_UseMis := 'Y';

      aPrintAddr := 'N';
      if rgpPrintAddr.ItemIndex = 0 then  aPrintAddr := 'Y';

      aPrintTitle := 'N';
      if rgpPrintTitle.ItemIndex = 0 then aPrintTitle := 'Y';

      aIfPrintCheck := 'N';
      if rgpIfPrintCheck.ItemIndex = 0 then aIfPrintCheck := 'Y';

      aExpAddrType := '1';
      if ( rgpExpAddrType.ItemIndex > 0 ) then aExpAddrType := '2';
      //#5764 增加 非營業人是否列印稅額 By Kin 2010/09/10
      aPrintZeroTax := 'N';
      if chkPrintZeroTax.Checked then aPrintZeroTax := 'Y';
      aInvParam := '<?xml version="1.0" encoding="Big5" ?>' +
                     '<InvoiceParam SysID="' + aSysId + '" PrintTitle="' + aPrintTitle + '" PrintAddr="' + aPrintAddr + '" ' +
                     'PrintZeroTax="' + aPrintZeroTax + '" ' +
                     'UseMis="' + aUseMis + '" UseBaseCompName="' + aUseBaseCompName + '" ' +
                     'InvoiceFilePath="' + aInvoiceFilePath + '" ' +
                     'PInvoiceFilePath="' + aPInvoiceFilePath + '" ' +
                     'EInvoiceFilePath="' + aEInvoiceFilePath + '" ' +
                     'MediaFilePath="' + aMediaFilePath + '" ' +
                     'BackupFilePath="' + aBackupFilePath + '" ' +
                     'RejectionInfoFile="' + aRejectionInfoFile + '"> ' +
                     '<Comp Type="' + aCompType + '" Name="' + aCompName + '"></Comp></InvoiceParam>';

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
      aBranchNo := txtBranchNo.Text;
      {}
      dtmMainJ.adoInv003Code.FieldByName('InvParam').AsString := aInvParam;
      dtmMainJ.adoInv003Code.FieldByName('IdentifyId1').AsString := IDENTIFYID1;
      dtmMainJ.adoInv003Code.FieldByName('IdentifyId2').AsString := IDENTIFYID2;

      if dtmMainJ.adoInv003Code.State = dsInsert then
        dtmMainJ.adoInv003Code.FieldByName('CompId').AsString := aCompId;

      dtmMainJ.adoInv003Code.FieldByName('IfPrintTitle').AsString := aPrintTitle;
      dtmMainJ.adoInv003Code.FieldByName('IfPrintAddr').AsString := aPrintAddr;
      dtmMainJ.adoInv003Code.FieldByName('CompTitle').AsString := aCompName;

      dtmMainJ.adoInv003Code.FieldByName('UptTime').AsString := DateTimeToStr(now);
      dtmMainJ.adoInv003Code.FieldByName('UptEn').AsString := dtmMain.getLoginUser;
      dtmMainJ.adoInv003Code.FieldByName( 'IfPrintCheck' ).AsString := aIfPrintCheck;
      dtmMainJ.adoInv003Code.FieldByName( 'ExpAddrType' ).AsString := aExpAddrType;
      dtmMainJ.adoInv003Code.FieldByName( 'City' ).AsString := aCity;
      dtmMainJ.adoInv003Code.FieldByName( 'ExportMode' ).AsString := aExportMode;
      dtmMainJ.adoInv003Code.FieldByName( 'ImportNote' ).AsString := aImportNote;
      dtmMainJ.adoInv003Code.FieldByName( 'EntryClass' ).AsString := aEntryClass;
      dtmMainJ.adoInv003Code.FieldByName( 'BranchNo' ).AsString := aBranchNo;

      {}

      dsrInv003.DataSet.Post;

      dtmMainJ.getAllInv003Data;
      ChangeBtnStatus;
      setButtonCompetence;
      initialForm;
    end;
end;

{ ---------------------------------------------------------------------------- }

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

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB01.SetCmbCity(aValue: String);
var
  aIndex: Integer;
begin
  cmbCity.ItemIndex := 0;
  if ( aValue <> EmptyStr ) then
    SetCxComboBoxItemValue( cmbCity, aValue );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB01.SetCmbExportMode(aValue: String);
begin
  cmbExportMode.ItemIndex := 0;
  if ( aValue <> EmptyStr ) then
    SetCxComboBoxItemValue( cmbExportMode, aValue );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB01.SetCmbImportNote(aValue: String);
begin
  cmbImportNote.ItemIndex := 0;
  if ( aValue <> EmptyStr ) then
    SetCxComboBoxItemValue( cmbImportNote, aValue );
end;

{ ---------------------------------------------------------------------------- }

end.
