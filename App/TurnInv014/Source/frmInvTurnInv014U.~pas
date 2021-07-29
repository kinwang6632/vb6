unit frmInvTurnInv014U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, dxSkinsCore, dxSkinBlack, dxSkinBlue, dxSkinCaramel,
  dxSkinCoffee, dxSkinsDefaultPainters, Menus, cxLookAndFeelPainters,
  StdCtrls, cxButtons, cxCheckBox, cxControls, cxContainer, cxEdit,
  cxProgressBar, ExtCtrls, DB, cxTextEdit, ShellApi ;

type
  TfrmTurnInv014 = class(TForm)
    Panel1: TPanel;
    cxProgressBar1: TcxProgressBar;
    chkBackup: TcxCheckBox;
    btnOk: TcxButton;
    btnCancel: TcxButton;
    edtBackUp: TcxTextEdit;
    procedure btnOkClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
  private
    { Private declarations }
    FCount:Integer;
    FStringList : TStrings;
    FLogFile : string;
    function TurnInv014(var aErrMsg : string) : Boolean;
    function BackupTbl(const aTblName : string) : Boolean;
  public
    { Public declarations }
  end;

var
  frmTurnInv014: TfrmTurnInv014 ;

implementation
uses dtmInv014AU ;

{$R *.dfm}

procedure TfrmTurnInv014.btnOkClick(Sender: TObject);
var
  aErrMsg : String;
begin
  try
    cxProgressBar1.Position := 0;
    Screen.Cursor := crSQLWait;
    btnCancel.Enabled := False;
    btnOk.Enabled := False;
    if BackupTbl('INV014_BAK') then
    begin
      if not TurnInv014( aErrMsg ) then
      begin
        MessageDlg( '轉檔失敗！'+ #13 +'請看轉檔記錄檔' ,mtWarning,[mbOK],0);
        ShellExecute(0,PChar('OPEN'),PChar('NotePad'),PChar(FLogFile),nil,SW_SHOW);
      end  else
        MessageDlg( '轉檔成功！' + #13 +'筆數:'+ IntToStr(FCount),mtInformation,[mbOK],0);
      FStringList.SaveToFile(FLogFile);
    end;
  finally
    Screen.Cursor := crDefault;
    btnCancel.Enabled := True;
    btnOk.Enabled := True;
  end;

end;

function TfrmTurnInv014.TurnInv014( var aErrMsg : string) : Boolean;
var
  aAllowance : string;
begin
  FCount := 0;
  FStringList.Clear;
  try
    with dtmInv014 do
    begin
      adoInv014.Close;
      adoInv014A.Close;
      adoInv014.SQL.Text := Format('Select * From %s.Inv014',[getMisDbOwner]);
      adoInv014A.SQL.Text := Format('Select * From %s.Inv014A',[getMisDbOwner]);
      adoInv014.Open;
      adoInv014A.Open;
      cxProgressBar1.Properties.Min := 0;
      cxProgressBar1.Properties.Max := adoInv014.RecordCount;
      if adoInv014.IsEmpty then
      begin
        aErrMsg := '無任何資料！';
        FStringList.Add(aErrMsg);
        Result := False;
        Exit;
      end;
      adoInv014.First;
      InvConnection.BeginTrans;
      adoComm.Close;
      adoComm.SQL.Text :=Format('Delete  %s.INV014A',[dtmInv014.getMisDbOwner]);
      adoComm.ExecSQL;
      while  not adoInv014.Eof  do
      begin
        try
          Application.ProcessMessages;
          aAllowance := adoInv014.FieldByName('AllowanceNo').AsString;
          adoInv014A.Append;
          if ( not VarIsNull( adoInv014.FieldByName('ServiceTypeID').Value ) ) then
            adoInv014A.FieldByName('ServiceTypeID').AsString := adoInv014.FieldByName('ServiceTypeID').AsString;
          adoInv014A.FieldByName('PaperNo').AsString := adoInv014.FieldByName('PaperNo').AsString;
          adoInv014A.FieldByName('Seq').AsInteger := adoInv014.FieldByName('Seq').AsInteger;
          adoInv014A.FieldByName('ItemId').AsString := adoInv014.FieldByName('ItemId').AsString;
          adoInv014A.FieldByName('Description').AsString := adoInv014.FieldByName('Description').AsString;
          if ( not VarIsNull( adoInv014.FieldByName('PayTypeId').Value )) then
            adoInv014A.FieldByName('PayTypeId').AsString := adoInv014.FieldByName('PayTypeId').AsString;
          if ( not VarIsNull( adoInv014.FieldByName('PayTypeDesc').Value )) then
            adoInv014A.FieldByName('PayTypeDesc').AsString := adoInv014.FieldByName('PayTypeDesc').AsString;
          adoInv014A.FieldByName('InvAmount').AsInteger := adoInv014.FieldByName('InvAmount').AsInteger;
          adoInv014A.FieldByName('UptTime').AsDateTime := adoInv014.FieldByName('UptTime').AsDateTime;
          adoInv014A.FieldByName('UptEn').AsString := adoInv014.FieldByName('UptEn').AsString;
          adoInv014A.FieldByName('AllowanceNo').AsString := adoInv014.FieldByName('AllowanceNo').AsString;
          adoInv014A.Post;
          adoInv014.Next;
          Inc(FCount);
          cxProgressBar1.Position := cxProgressBar1.Position + 1;
          FStringList.Add(Format('折讓單號 : %s 成功！',[aAllowance]));
        except
          on E : Exception do
          begin
            aErrMsg := E.Message;
            FStringList.Add(Format('折讓單號 : %s 失敗！ 失敗原因：%s ',[aAllowance,aErrMsg]));
            Break;
          end;

        end;
      end;
      if ( aErrMsg = EmptyStr ) then
        InvConnection.CommitTrans
      else
        InvConnection.RollbackTrans;
    end;
  except
    on E: Exception do
    begin
      if dtmInv014.InvConnection.InTransaction then
        dtmInv014.InvConnection.RollbackTrans;
      aErrMsg := E.Message;
      FStringList.Add( '轉檔失敗！失敗原因：' + aErrMsg );
    end;
  end;
  Result := ( aErrMsg = EmptyStr );
end;

procedure TfrmTurnInv014.FormCreate(Sender: TObject);
begin
  FStringList := TStringList.Create;
  FLogFile :=IncludeTrailingPathDelimiter(
    ExtractFilePath( Application.ExeName ) ) + TurnLog;
end;

procedure TfrmTurnInv014.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
  FStringList.Free;
end;

function TfrmTurnInv014.BackupTbl(const aTblName : string): Boolean;
begin
  if not chkBackup.Checked then
  begin
    Result := True;
    Exit;
  end;

  if ( chkBackup.Checked ) and ( aTblName = EmptyStr ) then
  begin
    MessageDlg( '請輸入備份檔名！',mtError,[mbOK],0);
    Result := True;
    Exit;
  end;
  try
    dtmInv014.adoComm.Close;
    dtmInv014.adoComm.SQL.Text := Format('DROP TABLE %s.%s',[dtmInv014.getMisDbOwner,aTblName]);
    dtmInv014.adoComm.ExecSQL;
  except
    dtmInv014.adoComm.Close;
  end;
  try
    with dtmInv014 do
    begin

      adoComm.Close;
      adoComm.SQL.Text :=Format(' Create Table %s.%s       ' +
          ' As Select * From %s.Inv014                     ' ,
          [getMisDbOwner,aTblName,getMisDbOwner] );
      adoComm.ExecSQL;
    end;
    Result := True;
  except
    Result := False;
  end;

end;

end.
