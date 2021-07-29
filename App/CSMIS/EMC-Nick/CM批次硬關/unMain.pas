unit unMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, cxEdit, cxControls, cxContainer, cxTextEdit, StdCtrls, ExtCtrls,
  cxGraphics, Menus, cxLookAndFeelPainters, cxButtons, cxMaskEdit,
  cxDropDownEdit, GraphicEx, cxStyles, cxCustomData, cxFilter, cxData,
  cxDataStorage, DB, cxDBData, cxGridLevel, cxClasses, cxGridCustomView,
  cxGridCustomTableView, cxGridTableView, cxGridDBTableView, cxGrid, MemDS,
  DBAccess, Ora, OracleCI, cxMemo, cxBlobEdit, xmldom, XMLIntf, msxmldom,
  XMLDoc, cxProgressBar, cbSQL;

type
  TfmMain = class(TForm)
    cxDefaultEditStyleController1: TcxDefaultEditStyleController;
    Panel1: TPanel;
    Panel2: TPanel;
    Label1: TLabel;
    txtUserId: TcxTextEdit;
    Label2: TLabel;
    Label3: TLabel;
    txtPassword: TcxTextEdit;
    Label4: TLabel;
    cmbAliase: TcxComboBox;
    btnLogin: TcxButton;
    imgDB: TImage;
    Label5: TLabel;
    Bevel1: TBevel;
    Image2: TImage;
    btnQuery: TcxButton;
    Panel3: TPanel;
    Panel4: TPanel;
    gvCM: TcxGridDBTableView;
    glCm: TcxGridLevel;
    CmGrid: TcxGrid;
    OraSession: TOraSession;
    qrCMBatch: TOraQuery;
    dsCMBatch: TDataSource;
    Image1: TImage;
    Label7: TLabel;
    lblRecord: TLabel;
    btnExec: TcxButton;
    gvCMMSGCODE: TcxGridDBColumn;
    gvCMMSGNAME: TcxGridDBColumn;
    gvCMWORKID: TcxGridDBColumn;
    gvCMCUSSO: TcxGridDBColumn;
    gvCMCUSID: TcxGridDBColumn;
    gvCMCMMAC: TcxGridDBColumn;
    gvCMEMTAMAC: TcxGridDBColumn;
    gvCMEMTAPORT: TcxGridDBColumn;
    gvCMCPNO: TcxGridDBColumn;
    gvCMCLASSID: TcxGridDBColumn;
    gvCMPCIPNO: TcxGridDBColumn;
    gvCMPCIP: TcxGridDBColumn;
    gvCMOPERTYPE: TcxGridDBColumn;
    gvCMSTARTDATETIME: TcxGridDBColumn;
    gvCMENDDATETIME: TcxGridDBColumn;
    gvCMZONE: TcxGridDBColumn;
    gvCMLIN: TcxGridDBColumn;
    gvCMSECTION: TcxGridDBColumn;
    gvCMLANE: TcxGridDBColumn;
    gvCMALLEY: TcxGridDBColumn;
    gvCMSUBALLEY: TcxGridDBColumn;
    gvCMNO1: TcxGridDBColumn;
    gvCMNO2: TcxGridDBColumn;
    gvCMOPERRESULT: TcxGridDBColumn;
    gvCMQUERYRESULT: TcxGridDBColumn;
    gvCMFAULTREASON: TcxGridDBColumn;
    gvCMGROUPID: TcxGridDBColumn;
    gvCMCMSTATUS: TcxGridDBColumn;
    gvCMRCMIP: TcxGridDBColumn;
    gvCMRCMUPFREQ: TcxGridDBColumn;
    gvCMRCMDOWNFREQ: TcxGridDBColumn;
    gvCMRCMRECPW: TcxGridDBColumn;
    gvCMRCMTRANSPW: TcxGridDBColumn;
    gvCMRCMSNR: TcxGridDBColumn;
    gvCMRCMONLINETIMES: TcxGridDBColumn;
    gvCMRCMINOCTETS: TcxGridDBColumn;
    gvCMRCMOUTOCTETS: TcxGridDBColumn;
    gvCMRCMINERRORS: TcxGridDBColumn;
    gvCMRCMOUTERRORS: TcxGridDBColumn;
    gvCMRPCIP: TcxGridDBColumn;
    gvCMRCMTSRECPW: TcxGridDBColumn;
    gvCMRCMTSRXSNR: TcxGridDBColumn;
    gvCMRCMTSUPMOD: TcxGridDBColumn;
    gvCMRHFCTYPE: TcxGridDBColumn;
    gvCMRPCOFFLINETIME: TcxGridDBColumn;
    gvCMRPCONLINETIME: TcxGridDBColumn;
    gvCMROFFLINETIME: TcxGridDBColumn;
    gvCMRONLINETIME: TcxGridDBColumn;
    gvCMRDETCTTIME: TcxGridDBColumn;
    gvCMRCMMAC: TcxGridDBColumn;
    gvCMRCTIP: TcxGridDBColumn;
    gvCMRNODEID: TcxGridDBColumn;
    gvCMEXECENTRY: TcxGridDBColumn;
    gvCMEXECDATE: TcxGridDBColumn;
    gvCMHOSTNAME: TcxGridDBColumn;
    gvCMCMDSEQNO: TcxGridDBColumn;
    gvCMDIALACCOUNT: TcxGridDBColumn;
    gvCMDIALPASSWORD: TcxGridDBColumn;
    gvCMCUSTNAME: TcxGridDBColumn;
    gvCMFINTIME: TcxGridDBColumn;
    gvCMTEL1: TcxGridDBColumn;
    gvCMTEL2: TcxGridDBColumn;
    gvCMCOMMANDTYPE: TcxGridDBColumn;
    gvCMXMLDATA: TcxGridDBColumn;
    gvCMID: TcxGridDBColumn;
    XmlDoc: TXMLDocument;
    OraSQL: TOraSQL;
    PBar: TcxProgressBar;
    btnSQL: TcxButton;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnQueryClick(Sender: TObject);
    procedure btnLoginClick(Sender: TObject);
    procedure btnExecClick(Sender: TObject);
    procedure btnSQLClick(Sender: TObject);
  private
    { Private declarations }
    FOraAliase: TStringList;
    FIsLogin: Boolean;
    function DoDbLogin(aUser, aPass, aAliase: String): Boolean;
    function DoExec(var aMsg: String): Boolean;
  public
    { Public declarations }
  end;

var
  fmMain: TfmMain;

implementation

uses cbUtilis;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TfmMain.FormCreate(Sender: TObject);
begin
  FIsLogin := False;
  FOraAliase := OracleAliasList;
  cmbAliase.Properties.Items.AddStrings( FOraAliase );
  btnQuery.Enabled := False;
  btnExec.Enabled := False;
  btnSQL.Enabled := False;
  lblRecord.Caption := '����:0';
  PBar.Visible := False;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.FormShow(Sender: TObject);
begin

end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.FormDestroy(Sender: TObject);
begin

end;

{ ---------------------------------------------------------------------------- }

function TfmMain.DoDbLogin(aUser, aPass, aAliase: String): Boolean;
var
  aMsg: String;
begin
  Result := False;
  aMsg := EmptyStr;
  if Trim( aUser ) = EmptyStr then
  begin
    aMsg := '�п�J�ϥΪ̱b���C'#13;
  end;
  if Trim( aPass ) = EmptyStr then
  begin
    aMsg := '�п�J�K�X�C'#13;
  end;
  if Trim( aPass ) = EmptyStr then
  begin
    aMsg := '�п��NET�n�J�O�W�C'#13;
  end;
  Result := ( aMsg = EmptyStr );
  if not Result then
  begin
    WarningMsg( aMsg );
    Exit;
  end;
  OraSession.Close;
  OraSession.LoginPrompt := False;
  OraSession.ConnectString := Format( '%s/%s@%s', [aUser,aPass,aAliase] );
  OraSession.Open;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.btnQueryClick(Sender: TObject);
begin
  if ( OraSession.InTransaction ) then
  begin
    OraSession.Rollback;
    Exit;
  end;
  {}
  if ( qrCMBatch.SQL.Text = EmptyStr ) then
  begin
    WarningMsg( '�п�JSQL���O' );
    Exit;
  end;
  {}
  Screen.Cursor := crSQLWait;
  try
    qrCMBatch.DisableControls;
    try
      qrCMBatch.Close;
      qrCMBatch.Open;
      qrCMBatch.Last;
      qrCMBatch.First;
      lblRecord.Caption := Format( '����:%d�C', [qrCMBatch.RecordCount] );
      if ( qrCMBatch.RecordCount > 0 ) then
      begin
        btnExec.Enabled := True;
      end;
    finally
      qrCMBatch.EnableControls;
    end;  
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.btnLoginClick(Sender: TObject);
begin
  Screen.Cursor := crSQLWait;
  try
    if ( not FIsLogin ) then
    begin
      if not DoDbLogin( txtUserId.Text, txtPassword.Text, cmbAliase.Text ) then
        Exit;
      btnQuery.Enabled := True;
      btnSQL.Enabled := True;
      FIsLogin := True;
      btnLogin.Caption := '�n�X';
      txtUserId.Enabled := False;
      txtPassword.Enabled := False;
      cmbAliase.Enabled := False;
    end else
    begin
      if ( OraSession.InTransaction ) then
      begin
        WarningMsg( '�w���ʹL���, �b�n�X�e�Х��T�{�O�n commit �� rollback �C' );
        Exit;
      end;
      qrCMBatch.Close;
      OraSession.Close;
      lblRecord.Caption := '����:0�C';
      btnQuery.Enabled := False;
      btnExec.Enabled := False;
      btnSQL.Enabled := False;
      FIsLogin := False;
      btnLogin.Caption := '�n�J';
      txtUserId.Enabled := True;
      txtPassword.Enabled := True;
      cmbAliase.Enabled := True;
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.btnExecClick(Sender: TObject);
var
  aExcResult: Boolean;
  aMsg: String;
begin
  Screen.Cursor := crSQLWait;
  try
    btnExec.Enabled := False;
    qrCMBatch.DisableControls;
    try
      OraSession.StartTransaction;
      try
        PBar.Position := 0;
        PBar.Properties.Max := qrCMBatch.RecordCount;
        PBar.Visible := True;
        aExcResult := DoExec( aMsg );
        PBar.Position := 0;
      except
        on E: Exception do
        begin
          PBar.Visible := False;
          ErrorMsg( Format( '���O�Ǹ� = %s ������~, �T��:%s�C', [
          qrCMBatch.FieldByName( 'cmdseqno' ).AsString,
           E.Message] ) );
        end;
      end;
      PBar.Visible := False;
      if not aExcResult then
      begin
        OraSession.Rollback;
      end else
      begin
        OraSession.Commit;
        if ( aMsg = EmptyStr ) then
          InfoMsg( '���槹��, ���ˮָ�ƬO�_���T�C' )
        else
          WarningMsg( aMsg );  
      end;
    finally
      btnExec.Enabled := True;
      qrCMBatch.EnableControls;
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.DoExec(var aMsg: String): Boolean;
var
  aXmlData: String;
  aOperaResult, aFaultReason, aRCTIP: String;
  aNodeList: IDOMNodeList;
  aNode: IDOMNode;
  aIndex: Integer;
begin
  Result := False;
  aMsg := EmptyStr;
  qrCMBatch.First;
  while not qrCMBatch.Eof do
  begin
    {}
    aFaultReason := EmptyStr;
    aOperaResult := EmptyStr;
    aRCTIP := EmptyStr;
    {}
    aXmlData := qrCMBatch.FieldByName( 'xmldata' ).AsString;
    if ( aXmlData = EmptyStr ) then
    begin
      qrCMBatch.Next;
      Continue;
    end;
    aXmlData := StringReplace( aXmlData, 'encoding="utf-8"?', 'encoding="Big5"?', [] );
    XmlDoc.LoadFromXML( aXmlData );
    try
      aNode := XmlDoc.Node.ChildNodes.Nodes['DataSet'].DOMNode;
      aNodeList := ( aNode as IDOMNodeSelect ).selectNodes( '//*' );
      {}
      for aIndex := 0 to aNodeList.length - 1 do
      begin
        if ( aNodeList.item[aIndex].nodeName = 'OperResult' ) then
        begin
          if Assigned( aNodeList.item[aIndex].firstChild ) then
            aOperaResult := aNodeList.item[aIndex].firstChild.nodeValue;
        end;
        if ( aNodeList.item[aIndex].nodeName = 'FaultReason' ) then
        begin
          if Assigned( aNodeList.item[aIndex].firstChild ) then
            aFaultReason := aNodeList.item[aIndex].firstChild.nodeValue;
        end;
        if ( aNodeList.item[aIndex].nodeName = 'CTIP' ) then
        begin
          if Assigned( aNodeList.item[aIndex].firstChild ) then
            aRCTIP := aNodeList.item[aIndex].firstChild.nodeValue;
        end;
      end;
    except
      on E: Exception do
      begin
        if ( aMsg = EmptyStr ) then
          aMsg := '����L�{��,������Ʀ��~�L�k��s�B���z���A,�U�C�������D�����O�Ǹ�:'#13#10;
        aMsg := ( aMsg + qrCMBatch.FieldByName( 'cmdseqno' ).AsString + #13#10 );
      end;
    end;
    {}
    if ( aOperaResult <> EmptyStr ) then
    begin
      OraSQL.SQL.Text := Format(
        ' update so311 set operresult = ''%s'',     ' +
        '   faultreason = ''%s'', rctip = ''%s''    ' +
        '  where cmdseqno = ''%s''                  ',
        [aOperaResult, aFaultReason, aRCTIP,
         qrCMBatch.FieldByName( 'cmdseqno' ).AsString] );
      OraSQL.Execute;
      PBar.Position := ( PBar.Position + 1 );
      Application.ProcessMessages; 
    end; 
    qrCMBatch.Next;
  end;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.btnSQLClick(Sender: TObject);
begin
  fmSQL := TfmSQL.Create( nil );
  try
    fmSQL.txtSQL.Text := qrCMBatch.SQL.Text;
    if ( fmSQL.ShowModal = mrOk ) then
    begin
      qrCMBatch.SQL.Text := fmSQL.txtSQL.Text;
    end;
  finally
    fmSQL.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

end.
