unit frmTL0040U;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, ExtCtrls, ComCtrls;

type
  TfrmTL0040 = class(TForm)
    Label1: TLabel;
    Label2: TLabel;
    rgpLang: TRadioGroup;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    pgbStatus: TProgressBar;
    Label3: TLabel;
    lblCount: TLabel;
    procedure BitBtn2Click(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmTL0040: TfrmTL0040;

implementation

uses dtmTL0040U;

{$R *.DFM}

procedure TfrmTL0040.BitBtn2Click(Sender: TObject);
begin
    Close;
end;

procedure TfrmTL0040.BitBtn1Click(Sender: TObject);
var
    sL_Lang,sL_Src : String;
    sL_Body, sL_Eng, sL_China, sL_Japan : String;
    nL_StatusPos : Integer;
begin
    case rgpLang.ItemIndex of
     0:
      sL_Lang := '英文';
     1:
      sL_Lang := '簡體字';
     2:
      sL_Lang := '日文';
     3:
      sL_Lang := '所有';
    end;

    if MessageDlg('是否確定開始轉換入'+sL_Lang+'譯文?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      dtmTL0040.ActiveDB;
      nL_StatusPos := 0;
      pgbStatus.Min := nL_StatusPos;
      pgbStatus.Max := dtmTL0040.adoRWordsInfo.RecordCount;
      dtmTL0040.PrepareExecSQL;
      with dtmTL0040.adoRWordsInfo do
      begin
        while not Eof do
        begin
          INC(nL_StatusPos);
          pgbStatus.Position := nL_StatusPos;
          sL_Src := FieldByName('sSrcWords').AsString;
          sL_Body := FieldByName('sBody').AsString;
          sL_Eng := FieldByName('sEngWords').AsString;
          sL_China := FieldByName('sChinaWords').AsString;
          sL_Japan := FieldByName('sJapanWords').AsString;
          dtmTL0040.UpdateWordsInfo(rgpLang.ItemIndex,sL_Src, sL_Body, sL_Eng, sL_China, sL_Japan);
          Next;
        end;
      end;
      dtmTL0040.PrepareExecSQL;
      lblCount.Caption := IntToStr(nL_StatusPos);
      MessageDlg('轉入完成!!',mtInformation,[mbOK],0);      
    end;

end;

end.
