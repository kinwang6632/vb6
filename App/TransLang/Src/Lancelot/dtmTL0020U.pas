unit dtmTL0020U;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  dtmBasicU, Db, ADODB, Provider, DBClient, DBTables;

type
  TdtmTL0020 = class(TdtmBasic)
    adoWordsInfo: TADOQuery;
    adoCheckExist: TADOQuery;
  private
    { Private declarations }
  public
    { Public declarations }
    procedure EmptyWordsInfo;
    procedure AppendWordsInfo(nI_Length, nI_HeadSpaceCount,nI_TailSpaceCount:Integer; sI_SrcWords, sI_Body, sI_Head, sI_Tail:String);


  end;

var
  dtmTL0020: TdtmTL0020;

implementation

{$R *.DFM}

{ TdtmTL0020 }

procedure TdtmTL0020.AppendWordsInfo(nI_Length, nI_HeadSpaceCount,nI_TailSpaceCount: Integer; sI_SrcWords, sI_Body,
  sI_Head, sI_Tail: String);
var
    bL_AppendData : Boolean;
begin
    with adoCheckExist do
    begin

      Close;
      Parameters.Clear;
      SQL.Clear;
      SQL.Add('select * from WordsInfo ');
      SQL.Add('where sSrcWords=:sSrcWords ');
      SQL.Add('and nPos=:nPos ');

      if sI_SrcWords='' then
      begin
        Parameters.Items[0].DataType := ftString;
        Parameters.Items[0].Direction := pdInput;
        Parameters.Items[0].Value := ' ';//若給NULL==>會當掉
      end
      else
        Parameters.Items[0].Value := sI_SrcWords;

      Parameters.Items[1].Value := 0;
      Open;
      if RecordCount=0 then
        bL_AppendData := True
      else
        bL_AppendData := False;
      close;
    end;

    if bL_AppendData then
    begin
      with adoWordsInfo do
      begin
        Append;
        FieldByName('sSrcWords').AsString := sI_SrcWords;
        FieldByName('nSrcLength').AsInteger := nI_Length;
        FieldByName('nHeadSpaceCount').AsInteger := nI_HeadSpaceCount;
        FieldByName('nTailSpaceCount').AsInteger := nI_TailSpaceCount;        
        FieldByName('sBody').AsString := Trim(sI_Body);
        FieldByName('sHead').AsString := Trim(sI_Head);
        FieldByName('sTail').AsString := Trim(sI_Tail);
        FieldByName('nPos').AsInteger := 0;//表示程式中用到的字串        
        Post;
      end;
    end;
end;

procedure TdtmTL0020.EmptyWordsInfo;
begin
    with adoWordsInfo do
    begin
      Close;
      Open;
    end;
end;

end.
