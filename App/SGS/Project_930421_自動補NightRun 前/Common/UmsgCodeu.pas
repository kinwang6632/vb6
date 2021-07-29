unit UmsgCodeu;

interface

uses
  Windows, Messages, SysUtils, Classes, Controls, Forms ;

Type
  TUmsgCode = class
  public
    class function MessageStr(asMsgNo:string):string;
    class procedure ShowLastErrorMessage ;
  end;


implementation

{ TUmsgCode }

class function TUmsgCode.MessageStr(asMsgNo: string): string;
begin

end;

class procedure TUmsgCode.ShowLastErrorMessage;
var
  PMsgBuf : PChar;
begin
  pMsgBuf := nil ;
  FormatMessage(
    FORMAT_MESSAGE_ALLOCATE_BUFFER or FORMAT_MESSAGE_FROM_SYSTEM,
    nil,
    GetLastError(),
    LOCALE_USER_DEFAULT,
    PMsgBuf,
    0,
    nil );
  Application.MessageBox(PMsgBuf, 'Error', MB_OK or MB_ICONERROR or MB_SYSTEMMODAL);
  LocalFree(HLOCAL(PMsgBuf));
end;

end.
