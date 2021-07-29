unit NDSGatewayImpU;

{$WARN SYMBOL_PLATFORM OFF}

interface

uses
  ComObj, ActiveX, prjNDSgateway_TLB, StdVcl;

type
  TNDSGateway = class(TAutoObject, INDSGateway)
  protected
    procedure F100202(vI_Operator: OleVariant; out vO_Mode,
      vO_Error: OleVariant); safecall;
    procedure F060202(vI_Operator, vI_CardID: OleVariant; out vO_SCID,
      vO_SCInfo, vO_Error: OleVariant); safecall;
    procedure F060203(vI_Operator, vI_SCID: OleVariant;
      out vO_Error: OleVariant); safecall;
    procedure F060236(vI_Operator, vI_SCID, vI_CardID: OleVariant;
      out vO_Error: OleVariant); safecall;
    procedure F060218(vI_Operator, vI_SCID, vI_CardID, vI_Type: OleVariant;
      out vO_Error: OleVariant); safecall;
    procedure F060205(vI_Operator, vI_SCID: OleVariant; out vO_RDate,
      vO_Error: OleVariant); safecall;
    procedure F060206(vI_Operator, vI_SCID: OleVariant;
      out vO_Error: OleVariant); safecall;
    procedure F060204(vI_Operator, vI_SCID: OleVariant;
      out vO_Error: OleVariant); safecall;
    procedure F060228(vI_Operator, vI_SCID, vI_Control, vI_UserNumber,
      vI_PINControl, vI_PINumber, vI_CBalance: OleVariant;
      out vO_Error: OleVariant); safecall;
    procedure F060211(vI_Operator, vI_SCID, vI_Status: OleVariant;
      out vO_Error: OleVariant); safecall;
    procedure F060222(vI_Operator: OleVariant; out vO_Error: OleVariant);
      safecall;
    procedure F060233(vI_Operator, vI_SCID, vI_Mode: OleVariant;
      out vO_Error: OleVariant); safecall;
    procedure F060234(vI_Operator, vI_SCID: OleVariant;
      out vO_Error: OleVariant); safecall;
    procedure F060210(vI_Operator, vI_SCID: OleVariant;
      out vO_Error: OleVariant); safecall;
    procedure F060209(vI_Operator, vI_SCID: OleVariant;
      out vO_Error: OleVariant); safecall;
    procedure F060215(vI_Operator, vI_SCID, vI_InfoCode: OleVariant;
      out vO_Info, vO_Error: OleVariant); safecall;
    procedure F060221(vI_Operator, vI_SCID: OleVariant;
      out vO_Error: OleVariant); safecall;
    procedure F060212(vI_Operator, vI_SCID, sI_ServiceOppvID: OleVariant;
      out vO_Error: OleVariant); safecall;
    { Protected declarations }
  end;

implementation

uses ComServ;


procedure TNDSGateway.F100202(vI_Operator: OleVariant; out vO_Mode,
  vO_Error: OleVariant);
var
    sL_ActionCode : String;
begin
    sL_ActionCode := 'D';

    vO_Mode := 'O';
    vO_Error := 'EZ02';
end;

procedure TNDSGateway.F060202(vI_Operator, vI_CardID: OleVariant;
  out vO_SCID, vO_SCInfo, vO_Error: OleVariant);
var
    sL_ActionCode : String;
begin
    sL_ActionCode := 'C';

    vO_SCID := '123456';
    vO_SCInfo := 'howard';

    vO_Error := '';
end;

procedure TNDSGateway.F060203(vI_Operator, vI_SCID: OleVariant;
  out vO_Error: OleVariant);
var
    sL_ActionCode : String;
begin
    sL_ActionCode := 'D';

    vO_Error := 'ES01';
end;

procedure TNDSGateway.F060236(vI_Operator, vI_SCID, vI_CardID: OleVariant;
  out vO_Error: OleVariant);
var
    sL_ActionCode : String;
begin
    sL_ActionCode := 'g';
    vO_Error := 'ES01';

end;

procedure TNDSGateway.F060218(vI_Operator, vI_SCID, vI_CardID,
  vI_Type: OleVariant; out vO_Error: OleVariant);
var
    sL_ActionCode : String;
begin
    sL_ActionCode := 'N';
    vO_Error := 'ES01';

end;

procedure TNDSGateway.F060205(vI_Operator, vI_SCID: OleVariant;
  out vO_RDate, vO_Error: OleVariant);
var
    sL_ActionCode : String;
begin
    sL_ActionCode := 'Q';
    vO_RDate := '';
    vO_Error := 'ER11';

end;

procedure TNDSGateway.F060206(vI_Operator, vI_SCID: OleVariant;
  out vO_Error: OleVariant);
var
    sL_ActionCode : String;
begin
    sL_ActionCode := 'R';
    vO_Error := 'ES01';

end;

procedure TNDSGateway.F060204(vI_Operator, vI_SCID: OleVariant;
  out vO_Error: OleVariant);
var
    sL_ActionCode : String;
begin
    sL_ActionCode := 'S';
    vO_Error := 'ES01';

end;


procedure TNDSGateway.F060228(vI_Operator, vI_SCID, vI_Control,
  vI_UserNumber, vI_PINControl, vI_PINumber, vI_CBalance: OleVariant;
  out vO_Error: OleVariant);
var
    sL_ActionCode : String;
begin
    sL_ActionCode := 'q';
    vO_Error := 'ES01';

end;

procedure TNDSGateway.F060211(vI_Operator, vI_SCID, vI_Status: OleVariant;
  out vO_Error: OleVariant);
var
    sL_ActionCode : String;
begin
    sL_ActionCode := 'j';
    vO_Error := 'ER11';

end;

procedure TNDSGateway.F060222(vI_Operator: OleVariant;
  out vO_Error: OleVariant);
var
    sL_ActionCode : String;
begin
    sL_ActionCode := 'x';
    vO_Error := 'ES01';

end;

procedure TNDSGateway.F060233(vI_Operator, vI_SCID, vI_Mode: OleVariant;
  out vO_Error: OleVariant);
var
    sL_ActionCode : String;
begin
    sL_ActionCode := 'Y';
    vO_Error := 'ES01';

end;

procedure TNDSGateway.F060234(vI_Operator, vI_SCID: OleVariant;
  out vO_Error: OleVariant);
var
    sL_ActionCode : String;
begin
    sL_ActionCode := 'y';
    vO_Error := 'ES01';

end;


procedure TNDSGateway.F060210(vI_Operator, vI_SCID: OleVariant;
  out vO_Error: OleVariant);
var
    sL_ActionCode : String;
begin
    sL_ActionCode := 'J';
    vO_Error := 'ES01';

end;

procedure TNDSGateway.F060209(vI_Operator, vI_SCID: OleVariant;
  out vO_Error: OleVariant);
var
    sL_ActionCode : String;
begin
    sL_ActionCode := 'a';
    vO_Error := 'ES01';

end;

procedure TNDSGateway.F060215(vI_Operator, vI_SCID,
  vI_InfoCode: OleVariant; out vO_Info, vO_Error: OleVariant);
var
    sL_ActionCode : String;
begin
    sL_ActionCode := 'G';
    vO_Info := '';
    vO_Error := 'ES01';

end;

procedure TNDSGateway.F060221(vI_Operator, vI_SCID: OleVariant;
  out vO_Error: OleVariant);
var
    sL_ActionCode : String;
begin
    sL_ActionCode := 'U';
    vO_Error := 'ES01';

end;

procedure TNDSGateway.F060212(vI_Operator, vI_SCID,
  sI_ServiceOppvID: OleVariant; out vO_Error: OleVariant);
var
    sL_ActionCode : String;
begin
    sL_ActionCode := 'X';
    vO_Error := 'ER14';

end;

initialization
  TAutoObjectFactory.Create(ComServer, TNDSGateway, Class_NDSGateway,
    ciMultiInstance, tmApartment);
end.
