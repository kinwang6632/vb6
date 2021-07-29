unit frmMainU;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Db, ADODB, StdCtrls, Buttons, ComCtrls, Provider, DBClient, DBTables;

type
  TfrmMain = class(TForm)
    adoSrc: TADOConnection;
    qrySrc: TADOQuery;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    pgbStatus: TProgressBar;
    dbDest: TDatabase;
    qryDest: TQuery;
    usqDest: TUpdateSQL;
    qryCommon: TQuery;
    procedure BitBtn2Click(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;

implementation

{$R *.DFM}

procedure TfrmMain.BitBtn2Click(Sender: TObject);
begin
    Application.Terminate;
end;

procedure TfrmMain.BitBtn1Click(Sender: TObject);
var
    nL_BeTransCount, nL_ProgPos : Integer;
    sL_Caption, sL_EngCaption : String;
begin
    nL_BeTransCount := 0;

    if not adoSrc.Connected then
      adoSrc.Connected := True;

    if not qrySrc.Active then
      qrySrc.Open;

    if not dbDest.Connected then
      dbDest.Connected := True;

    if not qryDest.Active then
      qryDest.Open;

    qryDest.First;
    nL_ProgPos := 0;
    pgbStatus.Min := nL_ProgPos;
    pgbStatus.Max := qryDest.RecordCount;
    Screen.Cursor := crHourGlass;
    while not qryDest.Eof do
    begin
      Inc(nL_ProgPos);
      pgbStatus.Position := nL_ProgPos;
      sL_Caption := Trim(qryDest.fieldByName('Caption').AsString);
      if qrySrc.Locate('sBody',sL_Caption,[]) then
      begin
        Inc(nL_BeTransCount);
        while not qrySrc.Eof do
        begin
          sL_EngCaption := Trim(qrySrc.fieldByName('sEngWords').AsString);
          if sL_EngCaption<> ''then
            break;
          qrySrc.next;  
        end;
        if sL_EngCaption='' then
          sL_EngCaption := ' ';
        qryCommon.Close;  
        qryCommon.ParamByName('EngCaption').AsString := sL_EngCaption;
        qryCommon.ParamByName('Caption').AsString := sL_Caption;
        qryCommon.ExecSQL;
        {
        qryDest.Edit;
        qryDest.FieldByName('EngCaption').AsString := sL_EngCaption;
        qryDest.Post;
        }
      end;
      qryDest.Next;
    end;
    {
    if qryDest.UpdatesPending then
      qryDest.ApplyUpdates;
    }
    Screen.Cursor := crDefault;
    MessageDlg('轉入完成!!' + '(共 ' + IntToStr(nL_BeTransCount) + ' 組)', mtInformation,[mbOK],0);
end;

end.
