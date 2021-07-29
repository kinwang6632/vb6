unit prjNDSgateway_TLB;

// ************************************************************************ //
// WARNING                                                                    
// -------                                                                    
// The types declared in this file were generated from data read from a       
// Type Library. If this type library is explicitly or indirectly (via        
// another type library referring to this type library) re-imported, or the   
// 'Refresh' command of the Type Library Editor activated while editing the   
// Type Library, the contents of this file will be regenerated and all        
// manual modifications will be lost.                                         
// ************************************************************************ //

// PASTLWTR : $Revision:   1.130  $
// File generated on 2001/7/26 ¤U¤È 04:36:26 from Type Library described below.

// ************************************************************************  //
// Type Lib: D:\APP\Start-TV\Src\prjNDSgateway.tlb (1)
// LIBID: {BC754026-CF84-4DCD-B36B-27F9DB786ACF}
// LCID: 0
// Helpfile: 
// DepndLst: 
//   (1) v2.0 stdole, (C:\WINNT\System32\stdole2.tlb)
//   (2) v4.0 StdVCL, (C:\WINNT\System32\stdvcl40.dll)
// ************************************************************************ //
{$TYPEDADDRESS OFF} // Unit must be compiled without type-checked pointers. 
{$WARN SYMBOL_PLATFORM OFF}
{$WRITEABLECONST ON}

interface

uses ActiveX, Classes, Graphics, StdVCL, Variants, Windows;
  

// *********************************************************************//
// GUIDS declared in the TypeLibrary. Following prefixes are used:        
//   Type Libraries     : LIBID_xxxx                                      
//   CoClasses          : CLASS_xxxx                                      
//   DISPInterfaces     : DIID_xxxx                                       
//   Non-DISP interfaces: IID_xxxx                                        
// *********************************************************************//
const
  // TypeLibrary Major and minor versions
  prjNDSgatewayMajorVersion = 1;
  prjNDSgatewayMinorVersion = 0;

  LIBID_prjNDSgateway: TGUID = '{BC754026-CF84-4DCD-B36B-27F9DB786ACF}';

  IID_INDSGateway: TGUID = '{758DA26D-4B93-43E6-A473-CB2FD6247981}';
  CLASS_NDSGateway: TGUID = '{82A64670-18A7-4EA4-91F0-156D7A87497D}';
type

// *********************************************************************//
// Forward declaration of types defined in TypeLibrary                    
// *********************************************************************//
  INDSGateway = interface;
  INDSGatewayDisp = dispinterface;

// *********************************************************************//
// Declaration of CoClasses defined in Type Library                       
// (NOTE: Here we map each CoClass to its Default Interface)              
// *********************************************************************//
  NDSGateway = INDSGateway;


// *********************************************************************//
// Interface: INDSGateway
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {758DA26D-4B93-43E6-A473-CB2FD6247981}
// *********************************************************************//
  INDSGateway = interface(IDispatch)
    ['{758DA26D-4B93-43E6-A473-CB2FD6247981}']
    procedure F100202(vI_Operator: OleVariant; out vO_Mode: OleVariant; out vO_Error: OleVariant); safecall;
    procedure F060202(vI_Operator: OleVariant; vI_CardID: OleVariant; out vO_SCID: OleVariant; 
                      out vO_SCInfo: OleVariant; out vO_Error: OleVariant); safecall;
    procedure F060203(vI_Operator: OleVariant; vI_SCID: OleVariant; out vO_Error: OleVariant); safecall;
    procedure F060236(vI_Operator: OleVariant; vI_SCID: OleVariant; vI_CardID: OleVariant; 
                      out vO_Error: OleVariant); safecall;
    procedure F060218(vI_Operator: OleVariant; vI_SCID: OleVariant; vI_CardID: OleVariant; 
                      vI_Type: OleVariant; out vO_Error: OleVariant); safecall;
    procedure F060205(vI_Operator: OleVariant; vI_SCID: OleVariant; out vO_RDate: OleVariant; 
                      out vO_Error: OleVariant); safecall;
    procedure F060206(vI_Operator: OleVariant; vI_SCID: OleVariant; out vO_Error: OleVariant); safecall;
    procedure F060204(vI_Operator: OleVariant; vI_SCID: OleVariant; out vO_Error: OleVariant); safecall;
    procedure F060228(vI_Operator: OleVariant; vI_SCID: OleVariant; vI_Control: OleVariant; 
                      vI_UserNumber: OleVariant; vI_PINControl: OleVariant; 
                      vI_PINumber: OleVariant; vI_CBalance: OleVariant; out vO_Error: OleVariant); safecall;
    procedure F060211(vI_Operator: OleVariant; vI_SCID: OleVariant; vI_Status: OleVariant; 
                      out vO_Error: OleVariant); safecall;
    procedure F060222(vI_Operator: OleVariant; out vO_Error: OleVariant); safecall;
    procedure F060233(vI_Operator: OleVariant; vI_SCID: OleVariant; vI_Mode: OleVariant; 
                      out vO_Error: OleVariant); safecall;
    procedure F060234(vI_Operator: OleVariant; vI_SCID: OleVariant; out vO_Error: OleVariant); safecall;
    procedure F060210(vI_Operator: OleVariant; vI_SCID: OleVariant; out vO_Error: OleVariant); safecall;
    procedure F060209(vI_Operator: OleVariant; vI_SCID: OleVariant; out vO_Error: OleVariant); safecall;
    procedure F060215(vI_Operator: OleVariant; vI_SCID: OleVariant; vI_InfoCode: OleVariant; 
                      out vO_Info: OleVariant; out vO_Error: OleVariant); safecall;
    procedure F060221(vI_Operator: OleVariant; vI_SCID: OleVariant; out vO_Error: OleVariant); safecall;
    procedure F060212(vI_Operator: OleVariant; vI_SCID: OleVariant; sI_ServiceOppvID: OleVariant; 
                      out vO_Error: OleVariant); safecall;
  end;

// *********************************************************************//
// DispIntf:  INDSGatewayDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {758DA26D-4B93-43E6-A473-CB2FD6247981}
// *********************************************************************//
  INDSGatewayDisp = dispinterface
    ['{758DA26D-4B93-43E6-A473-CB2FD6247981}']
    procedure F100202(vI_Operator: OleVariant; out vO_Mode: OleVariant; out vO_Error: OleVariant); dispid 2;
    procedure F060202(vI_Operator: OleVariant; vI_CardID: OleVariant; out vO_SCID: OleVariant; 
                      out vO_SCInfo: OleVariant; out vO_Error: OleVariant); dispid 3;
    procedure F060203(vI_Operator: OleVariant; vI_SCID: OleVariant; out vO_Error: OleVariant); dispid 1;
    procedure F060236(vI_Operator: OleVariant; vI_SCID: OleVariant; vI_CardID: OleVariant; 
                      out vO_Error: OleVariant); dispid 4;
    procedure F060218(vI_Operator: OleVariant; vI_SCID: OleVariant; vI_CardID: OleVariant; 
                      vI_Type: OleVariant; out vO_Error: OleVariant); dispid 5;
    procedure F060205(vI_Operator: OleVariant; vI_SCID: OleVariant; out vO_RDate: OleVariant; 
                      out vO_Error: OleVariant); dispid 6;
    procedure F060206(vI_Operator: OleVariant; vI_SCID: OleVariant; out vO_Error: OleVariant); dispid 7;
    procedure F060204(vI_Operator: OleVariant; vI_SCID: OleVariant; out vO_Error: OleVariant); dispid 8;
    procedure F060228(vI_Operator: OleVariant; vI_SCID: OleVariant; vI_Control: OleVariant; 
                      vI_UserNumber: OleVariant; vI_PINControl: OleVariant; 
                      vI_PINumber: OleVariant; vI_CBalance: OleVariant; out vO_Error: OleVariant); dispid 9;
    procedure F060211(vI_Operator: OleVariant; vI_SCID: OleVariant; vI_Status: OleVariant; 
                      out vO_Error: OleVariant); dispid 10;
    procedure F060222(vI_Operator: OleVariant; out vO_Error: OleVariant); dispid 11;
    procedure F060233(vI_Operator: OleVariant; vI_SCID: OleVariant; vI_Mode: OleVariant; 
                      out vO_Error: OleVariant); dispid 12;
    procedure F060234(vI_Operator: OleVariant; vI_SCID: OleVariant; out vO_Error: OleVariant); dispid 13;
    procedure F060210(vI_Operator: OleVariant; vI_SCID: OleVariant; out vO_Error: OleVariant); dispid 14;
    procedure F060209(vI_Operator: OleVariant; vI_SCID: OleVariant; out vO_Error: OleVariant); dispid 15;
    procedure F060215(vI_Operator: OleVariant; vI_SCID: OleVariant; vI_InfoCode: OleVariant; 
                      out vO_Info: OleVariant; out vO_Error: OleVariant); dispid 16;
    procedure F060221(vI_Operator: OleVariant; vI_SCID: OleVariant; out vO_Error: OleVariant); dispid 17;
    procedure F060212(vI_Operator: OleVariant; vI_SCID: OleVariant; sI_ServiceOppvID: OleVariant; 
                      out vO_Error: OleVariant); dispid 18;
  end;

// *********************************************************************//
// The Class CoNDSGateway provides a Create and CreateRemote method to          
// create instances of the default interface INDSGateway exposed by              
// the CoClass NDSGateway. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoNDSGateway = class
    class function Create: INDSGateway;
    class function CreateRemote(const MachineName: string): INDSGateway;
  end;

implementation

uses ComObj;

class function CoNDSGateway.Create: INDSGateway;
begin
  Result := CreateComObject(CLASS_NDSGateway) as INDSGateway;
end;

class function CoNDSGateway.CreateRemote(const MachineName: string): INDSGateway;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_NDSGateway) as INDSGateway;
end;

end.
