{#6629 �W�[�פJ�����M��\�� By Kin 2013/11/04}
unit frmInvB06_3U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, Buttons, StdCtrls, DB, ADODB;

type
  TfrmInvB06_3 = class(TForm)
    pnl1: TPanel;
    grp1: TGroupBox;
    edtImportPrize: TEdit;
    btnImportPrize: TButton;
    btnOk: TBitBtn;
    btnExit: TBitBtn;
    dlgOpen1: TOpenDialog;
    adoInv007: TADOQuery;
    procedure btnImportPrizeClick(Sender: TObject);
    procedure btnOkClick(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmInvB06_3: TfrmInvB06_3;

implementation
uses cbUtilis,dtmMainU,dtmMainJU,dtmSOU,dtmMainHU;
{$R *.dfm}

procedure TfrmInvB06_3.btnImportPrizeClick(Sender: TObject);
begin
  {#6629 �W�[�פJ�o�������M�U By Kin 2013/11/04}
  dlgOpen1.Filter := 'Text files (*.*)|*.*';
  if dlgOpen1.Execute then
    edtImportPrize.Text := dlgOpen1.FileName;
  if ( not FileExists( edtImportPrize.Text )) and ( edtImportPrize.Text <> EmptyStr ) then
  begin
    WarningMsg( '�ҫ��w���פJ�����M�U����ɤ��s�b�C' );
    Exit;
  end;
end;

procedure TfrmInvB06_3.btnOkClick(Sender: TObject);
var
  T_Import : TStringList;
  iCount,iSuccess,iExecute : Integer;
  aSQL : string;

begin
  if ( edtImportPrize.Text <> EmptyStr ) then
  begin
    if not FileExists( edtImportPrize.Text ) then
    begin
      WarningMsg( '�ҫ��w���פJ�����M�U����ɤ��s�b�C' );
      Exit;
    end;
    T_Import := TStringList.Create;
    T_Import.LoadFromFile( edtImportPrize.Text );
    iSuccess := 0;
    iExecute := 0;

    try

      dtmMain.InvConnection.BeginTrans;
      try
        for iCount := 0 to T_Import.Count -1  do
        begin
          Application.ProcessMessages;
          if (T_Import.Strings[iCount] <> EmptyStr) and
            ( Copy( T_Import.Strings[iCount],1,1) <> 'F' ) and
            (Length(T_Import.Strings[iCount])>= 446) then
          begin
            if (copy( T_Import.Strings[iCount],424,1) <> EmptyStr ) then
            begin
              {#6716 �W�[�פJDepositMK�PDataType By Kin 2014/02/10}
              aSQL := Format( 'UPDATE INV007 SET PrizeType = ''%s'',        ' +
                        'DepositMK = ''%s'',                                ' +
                        'DataType = ''%s''                                  ' +
                        ' WHERE INVID = ''%s'' AND COMPID = ''%s''          ',
                        [UpperCase(copy( T_Import.Strings[iCount],424,1)),
                        UpperCase(copy( T_Import.Strings[iCount],445,1)),
                        UpperCase(copy( T_Import.Strings[iCount],446,1)),
                        UpperCase(copy( T_Import.Strings[iCount],16,10)),
                        dtmMain.getCompID] );

              adoInv007.Close;
              adoInv007.SQL.Clear;
              adoInv007.SQL.Text := aSQL;
              iExecute:= adoInv007.ExecSQL;
              iSuccess := iExecute + iSuccess;
              adoInv007.Close;
            end;
          end;
        end;
        dtmMain.InvConnection.CommitTrans;
        MessageDlg('�פJ�����I' + #10#13 + Format( '�פJ����:%d' , [iSuccess] )
         ,mtInformation,[mbOK],0);
      except
        on E: Exception do
        begin
          dtmMain.InvConnection.RollbackTrans;
          ErrorMsg( Format( '�פJ����, �T��:%s', [E.Message] ) );
          Exit;
        end;
      end;

    finally
      T_Import.Free;
      adoInv007.Close;
    end;
  end else
  begin
    WarningMsg( '�Ы��w�פJ�����M�U�ɮ׸��|�C' );
  end;

end;

procedure TfrmInvB06_3.btnExitClick(Sender: TObject);
begin
  Close;
end;

end.
