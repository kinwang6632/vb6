unit dtmTL0040U;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  dtmBasicU, Db, ADODB;

type
  TdtmTL0040 = class(TdtmBasic)
    adoRWordsInfo: TADOQuery;
    adoUpdateWordsInfo: TADOQuery;
  private
    { Private declarations }
  public
    { Public declarations }
    procedure PrepareExecSQL;
    procedure ActiveDB;
    procedure UpdateWordsInfo(nI_Index : Integer;sI_Src, sI_Body, sI_Eng, sI_China, sI_Japan:String); 
  end;

var
  dtmTL0040: TdtmTL0040;

implementation

{$R *.DFM}

{ TdtmTL0040 }

procedure TdtmTL0040.ActiveDB;
begin
    adoRWordsInfo.Active := False;
    adoRWordsInfo.Active := True;

end;

procedure TdtmTL0040.PrepareExecSQL;
begin
    adoUpdateWordsInfo.SQL.Clear;
end;

procedure TdtmTL0040.UpdateWordsInfo(nI_Index: Integer; sI_Src, sI_Body, sI_Eng,
  sI_China, sI_Japan: String);
var
    sL_SQL : String;  
begin

    with adoUpdateWordsInfo do
    begin
      if SQL.Text = '' then
      begin
        case nI_Index of
          0 ://�^��
             sL_SQL := 'update WordsInfo set sEngWords=:sEngWords ';
          1 ://²��r
             sL_SQL := 'update WordsInfo set sChinaWords=:sChinaWords ';
          2 ://���
             sL_SQL := 'update WordsInfo set sJapanWords=:sJapanWords ';
          3://all
           begin
             sL_SQL := 'update WordsInfo set sEngWords=:sEngWords ';
             sL_SQL := sL_SQL + ', sChinaWords=:sChinaWords';
             sL_SQL := sL_SQL + ', sJapanWords=:sJapanWords';
           end;
        end;
        sL_SQL := sL_SQL + ' where sBody=:sBody ';

        SQL.Text := sL_SQL;
      end;

      Parameters.ParamByName('sBody').Value := sI_Body;

      if Parameters.FindParam('sEngWords') <> nil then
      begin
        if sI_Eng='' then
          sI_Eng := sI_Src;//�]��sI_Eng����O�Ū�(�_�hExecSQL�|error),�Gassign�@��space
        Parameters.ParamByName('sEngWords').Value := sI_Eng;
      end;

      if Parameters.FindParam('sChinaWords') <> nil then
      begin
        if sI_China='' then
          sI_China := sI_Src;//�]��sI_China����O�Ū�(�_�hExecSQL�|error),�Gassign�@��space
        Parameters.ParamByName('sChinaWords').Value := sI_China;
      end;

      if Parameters.FindParam('sJapanWords') <> nil then
      begin
        if sI_Japan='' then
          sI_Japan := sI_Src;//�]��sI_Japan����O�Ū�(�_�hExecSQL�|error),�Gassign�@��space
        Parameters.ParamByName('sJapanWords').Value := sI_Japan;
      end;

      ExecSQL;          
    end;    
end;

end.
