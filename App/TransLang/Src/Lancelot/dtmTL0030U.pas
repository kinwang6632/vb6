unit dtmTL0030U;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  dtmBasicU, Db, ADODB;

type
  TdtmTL0030 = class(TdtmBasic)
    adoTransLanguage: TADOQuery;
  private
    { Private declarations }
  public
    { Public declarations }
    function TransLanguage(nI_Type:Integer; sI_SrcWords, sI_Body, sI_Head, sI_Tail:String):String;
  end;

var
  dtmTL0030: TdtmTL0030;

implementation

uses Ustru;

{$R *.DFM}

{ TdtmTL0030 }

function TdtmTL0030.TransLanguage(nI_Type:Integer; sI_SrcWords,  sI_Body, sI_Head,
  sI_Tail: String): String;
var
    nL_HeadSpaceCount,nL_TailSpaceCount, nL_Length : Integer;
    sL_Result,sL_SrcWords : String;
    sL_EngWords, sL_SimChinaWords, sL_JapanWords :String;
begin
    with adoTransLanguage do
    begin
      Close;
      SQL.Clear;
      Parameters.Clear;
//      SQL.Add('select nWordsSeq, nSecLength from SrcWordsInfo ');
      SQL.Add('select * from WordsInfo ');
      SQL.Add('where sSrcWords=:sSrcWords ');
      SQL.Add('and nPos in (0,2) ');//找出在程式中(nPos=0)或程式與report中(nPos=2)均出現之字串

      if sI_SrcWords='' then
      begin
        Parameters.Items[0].DataType := ftString;
        Parameters.Items[0].Direction := pdInput;
        Parameters.Items[0].Value := ' ';//若給NULL==>會當掉
      end
      else
        Parameters.Items[0].Value := sI_SrcWords;

      Open;
      if RecordCount=1 then
      begin
        nL_Length := FieldByName('nSrcLength').AsInteger;
        nL_HeadSpaceCount := FieldByName('nHeadSpaceCount').AsInteger;
        nL_TailSpaceCount := FieldByName('nTailSpaceCount').AsInteger;        
        sL_SrcWords := FieldByName('sSrcWords').AsString;
        sL_EngWords := FieldByName('sEngWords').AsString;
        sL_SimChinaWords := FieldByName('sChinaWords').AsString;
        sL_JapanWords := FieldByName('sJapanWords').AsString;


        case nI_Type of
         0://Eng
          sL_Result := sL_EngWords;
         1://China
          sL_Result := sL_SimChinaWords;
         2://Japan
          sL_Result := sL_JapanWords;
        end;
        
        if Trim(sL_Result) = '' then
          sL_Result := sL_SrcWords
        else
          sL_Result := sI_Head + TUstr.Space(nL_HeadSpaceCount) + sL_Result + TUstr.Space(nL_TailSpaceCount) + sI_Tail;
          
        if nL_Length > Length(sL_Result) then
          sL_Result := sL_Result + TUstr.Space(nL_Length - Length(sL_Result))
      end;
      Close;

    end;
    Result := sL_Result;
end;

end.
