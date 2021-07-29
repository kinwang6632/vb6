unit Unit2;

{$WARN SYMBOL_PLATFORM OFF}

interface

uses
  ComObj, ActiveX, Project1_TLB, StdVcl, SysUtils;

type
  Tabc = class(TAutoObject, Iabc)
  protected

    procedure Method1(sI_Data: OleVariant; out sI_Result, sI_ErrorCode,
      sI_ErrorMsg: OleVariant); safecall;
    { Protected declarations }
  end;

implementation

uses ComServ, Unit1;

procedure Tabc.Method1(sI_Data: OleVariant; out sI_Result, sI_ErrorCode,
  sI_ErrorMsg: OleVariant);
begin
    Form1.addData(sI_Data);
    Inc(Form1.nG_Result);
    
    sI_Result := IntToStr(Form1.nG_Result);
    if (Form1.nG_Result mod 2) = 0 then
    begin
      sI_ErrorCode := '¿ù»~¥N½X';
      sI_ErrorMsg := '¿ù»~°T®§';
    end  
    else
    begin
      sI_ErrorCode := '';
      sI_ErrorMsg := '';
    end;

end;

initialization
  TAutoObjectFactory.Create(ComServer, Tabc, Class_abc,
    ciMultiInstance, tmApartment);
end.
