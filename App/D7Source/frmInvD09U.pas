unit frmInvD09U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, dxSkinsCore, dxSkinBlack, dxSkinBlue, dxSkinCaramel,
  dxSkinCoffee, dxSkinsDefaultPainters, cxGraphics, cxMaskEdit,
  cxDropDownEdit, cxMemo, cxTextEdit, cxControls, cxContainer, cxEdit,
  cxLabel, DB, Grids, DBGrids, StdCtrls, Buttons, ExtCtrls, Menus,
  cxLookAndFeelPainters, cxButtons, ADODB;

type
  TfrmINVD09 = class(TForm)
    Panel1: TPanel;
    Panel3: TPanel;
    btnAppend: TBitBtn;
    btnEdit: TBitBtn;
    btnDelete: TBitBtn;
    btnCancel: TBitBtn;
    btnSave: TBitBtn;
    btnExit: TBitBtn;
    pnlGrid: TPanel;
    DBGrid1: TDBGrid;
    dsrInv040: TDataSource;
    pnlEdit: TPanel;
    cxLabel1: TcxLabel;
    cxLabel2: TcxLabel;
    cxLabel3: TcxLabel;
    cxLabel4: TcxLabel;
    edtEmailHead: TcxTextEdit;
    cboSendType: TcxComboBox;
    cboComId: TcxComboBox;
    edtContext: TcxMemo;
    cxLabel5: TcxLabel;
    edtBatHead: TcxTextEdit;
    cxLabel6: TcxLabel;
    btnBatHead: TcxButton;
    OpenDialog1: TOpenDialog;
    edtBatModule: TcxTextEdit;
    btnBatModule: TcxButton;
    cxLabel7: TcxLabel;
    lblUptEn: TcxLabel;
    cxLabel8: TcxLabel;
    lblUptTime: TcxLabel;
    cxLabel9: TcxLabel;
    edtdonatetxt: TcxMemo;
    procedure FormShow(Sender: TObject);
    procedure btnAppendClick(Sender: TObject);
    procedure btnEditClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure dsrInv040DataChange(Sender: TObject; Field: TField);
    procedure btnSaveClick(Sender: TObject);
    procedure cboSendTypePropertiesChange(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure btnBatHeadClick(Sender: TObject);
    procedure btnBatModuleClick(Sender: TObject);
    procedure edtEmailHeadPropertiesChange(Sender: TObject);
    procedure edtContextPropertiesChange(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
    procedure ChangeBtnStatus;
    procedure initialForm;
    function isDataOK: Boolean;
    function GetSendType(aStr : String ):Integer;
//    procedure setButtonCompetence;
  public
    { Public declarations }
  end;

var
  frmINVD09: TfrmINVD09;

implementation

uses cbUtilis, frmMainU, dtmMainJU, dtmMainU, Math;
{$R *.dfm}

procedure TfrmINVD09.ChangeBtnStatus;
begin
  with dtmMainJ.adoInv040Code do
   begin
     if state in [dsInactive] then
       Exit
     else if state in [dsEdit, dsInsert] then
     begin
        btnCancel.Enabled := True;
        btnSave.Enabled := True;
        btnDelete.Enabled := False;
        btnEdit.Enabled := False;
        btnExit.Enabled := False;
        btnAppend.Enabled := False;
        pnlGrid.Enabled := False;
        pnlEdit.Enabled := True;
     end
     else
     begin
        if dtmMainJ.adoInv040Code.RecordCount>0 then
        begin
          btnAppend.Enabled :=True;
          btnEdit.Enabled := True;
          btnDelete.Enabled := True;
        end
        else
        begin
          btnEdit.Enabled := False;
          btnDelete.Enabled := False;
        end;
        btnCancel.Enabled := False;
        btnSave.Enabled := False;
        btnAppend.Enabled := True;
        btnExit.Enabled := True;
        pnlGrid.Enabled := True;
        pnlEdit.Enabled := False;
     end;
   end;
end;

procedure TfrmINVD09.FormShow(Sender: TObject);
begin

  dtmMainJ.adoInvD09.Connected := False;
  {
  dtmMainJ.adoInvD09.ConnectionString :=Format(
     'Provider=OraOLEDB.Oracle.1;Password=%s;Persist Security Info=True;' +
      'User ID=%s;Data Source=%s' ,
     [dtmMain.getDbPassword , dtmMain.getDbUserID , dtmMain.getDbSID] );
  dtmMainJ.adoInvD09.Connected := True;
  }
  if dtmMain.GetStarEmail then
    cboSendType.Properties.Items.Add( 'E-Mail' );
  if dtmMain.GetStarMessage then
    cboSendType.Properties.Items.Add( '簡訊');
  if dtmMain.GetStarTVmail then
    cboSendType.Properties.Items.Add( 'TVMail');
  if dtmMain.GetStarCMSend then
    cboSendType.Properties.Items.Add( 'CM 導流');

  if dtmMainJ.getAllInv001Data > 0 then
  begin
    while not dtmMainJ.adoInv001Code.Eof do
    begin
      cboComId.Properties.Items.Add( dtmMainJ.adoInv001Code.FieldByName('CompId').AsString );
      dtmMainJ.adoInv001Code.Next;
    end;

  end;
  if dtmMainJ.getAllInv040Data = 0 then
    WarningMsg(  '無資料。' );
  ChangeBtnStatus;
//  setButtonCompetence;
  initialForm;
end;

procedure TfrmINVD09.initialForm;
begin
  with dsrInv040.DataSet do
  begin
    FindField('CompId').DisplayLabel := '公司代碼';
    FindField('SendType2').DisplayLabel := '傳送類別';
    FindField( 'EmailHead' ).DisplayLabel := 'Email主旨';
    FindField( 'EmailHead' ).DisplayWidth := 30;
    FindField( 'SendContext' ).DisplayLabel := '傳送內容';
    FindField( 'BatHeadFile' ).DisplayLabel := 'E-MAIL整批主旨文件檔名';
    FindField( 'BatHeadFile' ).DisplayWidth := 30;
    FindField( 'BatModuleFile' ).DisplayLabel := 'E-MAIL整批模板文件檔名';
    FindField( 'BatModuleFile' ).DisplayWidth := 30;
    FindField( 'UPTEN' ).DisplayLabel := '異動人員';
    FindField( 'UPTEN' ).DisplayWidth := 10;
    FindField( 'UPTTIME' ).DisplayLabel := '異動時間';
    FindField( 'SendDonatetext' ).DisplayLabel := '發票捐贈內容';
    FindField( 'SendDonatetext' ).DisplayWidth :=30;
    FindField( 'SENDTYPE' ).Visible := False;
    if RecordCount = 0 then
    begin
      btnEdit.Enabled := False;
      btnDelete.Enabled := False;
    end;
  end;


end;

procedure TfrmINVD09.btnAppendClick(Sender: TObject);
begin
  dtmMainJ.adoInv040Code.Append;
  if cboSendType.Properties.Items.Count > 0 then
    cboSendType.ItemIndex := 0;
  if cboComId.Properties.Items.Count > 0 then
    cboComId.Text := dtmMain.getCompID;
  ChangeBtnStatus;
  cboComId.SetFocus;
  edtContext.Text := EmptyStr;
  edtdonatetxt.Text := EmptyStr;

end;

procedure TfrmINVD09.btnEditClick(Sender: TObject);
begin
  dtmMainJ.adoInv040Code.Edit;
  ChangeBtnStatus;
  edtContext.SetFocus;
  cboComId.Enabled := False;
  cboSendType.Enabled := False;
end;

procedure TfrmINVD09.btnDeleteClick(Sender: TObject);
begin
  if ConfirmMsg( '確認刪除此筆資料?' ) then
  begin
    dtmMainJ.adoInv040Code.Delete;
    dtmMainJ.getAllInv040Data;
    ChangeBtnStatus;
  end;
  //setButtonCompetence;
  initialForm;
end;

procedure TfrmINVD09.btnCancelClick(Sender: TObject);
begin
  if  dsrInv040.DataSet.State in [dsEdit] then
    cboComId.Enabled := True;
    cboSendType.Enabled := True;
  dtmMainJ.adoInv040Code.Cancel;
  dtmMainJ.adoInv040Code.First;
  ChangeBtnStatus;

  //setButtonCompetence;
  initialForm;
end;

function TfrmINVD09.isDataOK: Boolean;
begin
  Result := True;
  if cboComId.Text = EmptyStr then
  begin
    WarningMsg('請輸入公司代碼');
    cboComId.SetFocus;
    Result := False;
    Exit;
  end;
  if  UpperCase( cboSendType.Text  ) = 'E-MAIL' then
  begin
    if Trim( edtEmailHead.Text ) = EmptyStr then
    begin
      WarningMsg( '請輸入EMAIL 主旨');
      edtEmailHead.SetFocus;
      Result := False;
      Exit;
    end;
  end;
  if Trim(edtContext.Text)= EmptyStr then
  begin
    WarningMsg( '請輸入傳送內容' );
    edtContext.SetFocus;
    Result := False;
    Exit;
  end;
end;

procedure TfrmINVD09.dsrInv040DataChange(Sender: TObject; Field: TField);
begin
  with dsrInv040.DataSet do
  begin
    if RecordCount > 0 then
    begin
      cboComId.Text := FieldByName('CompID').AsString;
      cboSendType.ItemIndex := FieldByName('SendType').AsInteger -1;

      case FieldByName('SendType').AsInteger of
        1: cboSendType.Text :='E-Mail' ;
        2: cboSendType.Text :='簡訊' ;
        3: cboSendType.Text :='TVMail';
        4: cboSendType.Text :='CM 導流';
      end;

      edtEmailHead.Text := FieldByName('EmailHead').AsString;
      edtBatHead.Text := FieldByName( 'BatHeadFile' ).AsString;
      edtBatModule.Text := FieldByName( 'BatModuleFile' ).AsString;
      edtContext.Text := FieldByName('SendContext').AsString;
      lblUptEn.Caption := FieldByName( 'UptEN' ).AsString;
      lblUptTime.Caption := FieldByName( 'UPTTIME' ).AsString;
      //增加發票捐贈內容 By Kin 2011/03/29
      edtdonatetxt.Text := FieldByName( 'SendDonatetext' ).AsString;
    end else
    begin
      cboComId.ItemIndex := -1;
      cboSendType.ItemIndex := -1;
      edtEmailHead.Text := EmptyStr;
      edtContext.Text := EmptyStr;
    end;
  end;
end;

procedure TfrmINVD09.btnSaveClick(Sender: TObject);
var
  aSQL : String;
  aHavePKValue :Boolean;
  aBook : Pointer;
  aSendType : Integer;
begin

  if UpperCase( cboSendType.Text )= 'E-MAIL' then
    aSendType :=1;
  if UpperCase( cboSendType.Text )='簡訊' then
    aSendType :=2;
  if UpperCase( cboSendType.Text )='TVMAIL' then
    aSendType := 3;
  if Trim(UpperCase( cboSendType.Text )) ='CM導流' then
    aSendType := 4;

  aSQL := Format(
      ' select * from inv040 where CompId = ''%s''' +
      ' and SendType = %d',
      [cboComId.Text, aSendType ]);
  if  dsrInv040.DataSet.State in [dsInsert] then
    aHavePKValue := dtmMainJ.checkPK(aSQL)
  else
    aHavePKValue := False;
  if aHavePKValue then
  begin
    WarningMsg( '違反唯一值條件' );
    cboComId.SetFocus;
  end else
  begin
      dsrInv040.OnDataChange := nil;
      aBook := nil;
      if  dsrInv040.DataSet.State in [dsEdit ] then
      begin
        aBook := dsrInv040.DataSet.GetBookmark;
      end;
      try
        with dsrInv040.DataSet do
        begin
          FieldByName('CompID').AsString := cboComId.Text;
          FieldByName('SendType').AsInteger := aSendType;             // GetSendType(cboSendType.Text);
          if edtEmailHead.Enabled then
          begin
            if edtEmailHead.Text <> EmptyStr then
            begin
              FieldByName('EmailHead').AsString := edtEmailHead.Text;
            end else
            begin
              FieldByName('EmailHead').AsString := EmptyStr;
            end;
          end else
          begin
            FieldByName('EmailHead').AsString := EmptyStr;
          end;
          FieldByName( 'BatHeadFile').AsString := edtBatHead.Text;
          FieldByName( 'BatModuleFile' ).AsString := edtBatModule.Text;
          FieldByName('SendContext').AsString := edtContext.Text;
          FieldByName('SendDonatetext').AsString := edtdonatetxt.Text;
          FieldByName( 'UptEN' ).AsString := dtmMain.getLoginUserName;
          FieldByName( 'UptTime' ).AsDateTime := Now;
//          FieldByName( 'SendDonatetext' ).AsString := edtdonatetxt.Text;
          dsrInv040.DataSet.Post;
        end;
      finally
        dsrInv040.OnDataChange := dsrInv040DataChange;
      end;
  end;
  dtmMainJ.getAllInv040Data;
  {
  try
    if  Assigned(aBook) then
      dsrInv040.DataSet.GotoBookmark(aBook);
  finally
    dsrInv040.DataSet.FreeBookmark( aBook );
  end;
  }
  ChangeBtnStatus;
  cboSendType.Enabled := True;
  cboComId.Enabled := True;
  //setButtonCompetence;
  initialForm;
end;

function TfrmINVD09.GetSendType(aStr: String): Integer;
begin
  if UpperCase( aStr )= 'E-MAIL' then
    Result :=1;
  if UpperCase( aStr )='簡訊' then
    Result :=2;
  if UpperCase( aStr )='TVMAIL' then
    Result := 3;
  if Trim(UpperCase( aStr )) ='CM導流' then
    Result := 4;
end;

procedure TfrmINVD09.cboSendTypePropertiesChange(Sender: TObject);
begin

  if UpperCase(cboSendType.Text)='E-MAIL' then
  begin
    edtEmailHead.Enabled := True;
    btnBatHead.Enabled := True;
    btnBatModule.Enabled := True;
  end else
  begin
    edtEmailHead.Enabled := False;
    edtEmailHead.Text := EmptyStr;
    btnBatHead.Enabled := False;
    btnBatModule.Enabled := False;
  end;
  
end;

procedure TfrmINVD09.btnExitClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmINVD09.btnBatHeadClick(Sender: TObject);
begin
   OpenDialog1.FilterIndex := 1;
   if OpenDialog1.Execute then
   begin
      edtBatHead.Text := OpenDialog1.FileName;
      edtEmailHead.Text := EmptyStr;
   end;
end;

procedure TfrmINVD09.btnBatModuleClick(Sender: TObject);
begin
  OpenDialog1.Files.Clear;
  OpenDialog1.FileName := EmptyStr;
  OpenDialog1.FilterIndex :=2;
  if OpenDialog1.Execute then
  begin
    edtBatModule.Text := OpenDialog1.FileName;
    edtContext.Text := EmptyStr;
  end;
end;

procedure TfrmINVD09.edtEmailHeadPropertiesChange(Sender: TObject);
begin
  if edtEmailHead.Text <> EmptyStr then
    edtBatHead.Text := EmptyStr;
end;

procedure TfrmINVD09.edtContextPropertiesChange(Sender: TObject);
begin
  if edtContext.Text <> EmptyStr then
  begin
    edtBatModule.Text := EmptyStr;
  end;
end;

procedure TfrmINVD09.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  dtmMainJ.adoInvD09.Connected := False;
  dtmMainJ.adoInvD09.Close;
end;

end.
