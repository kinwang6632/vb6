unit prjE2C_TLB;

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

// PASTLWTR : $Revision:   1.130.1.0.1.0.1.6  $
// File generated on 2003/04/23 13:32:57 from Type Library described below.

// ************************************************************************  //
// Type Lib: D:\App\EMC\MSS\Project\Server\prjE2C.tlb (1)
// LIBID: {46295CEB-E5BF-4658-8B2F-F67CA31A98A6}
// LCID: 0
// Helpfile: 
// DepndLst: 
//   (1) v2.0 stdole, (C:\WINNT\System32\stdole2.tlb)
//   (2) v4.0 StdVCL, (C:\WINNT\System32\stdvcl40.dll)
// ************************************************************************ //
{$TYPEDADDRESS OFF} // Unit must be compiled without type-checked pointers. 
{$WARN SYMBOL_PLATFORM OFF}
{$WRITEABLECONST ON}
{$VARPROPSETTER ON}
interface

uses Windows, ActiveX, Classes, Graphics, StdVCL, Variants;
  

// *********************************************************************//
// GUIDS declared in the TypeLibrary. Following prefixes are used:        
//   Type Libraries     : LIBID_xxxx                                      
//   CoClasses          : CLASS_xxxx                                      
//   DISPInterfaces     : DIID_xxxx                                       
//   Non-DISP interfaces: IID_xxxx                                        
// *********************************************************************//
const
  // TypeLibrary Major and minor versions
  prjE2CMajorVersion = 1;
  prjE2CMinorVersion = 0;

  LIBID_prjE2C: TGUID = '{46295CEB-E5BF-4658-8B2F-F67CA31A98A6}';

  IID_IE2C: TGUID = '{581C140B-CE91-4741-A9F4-304C6AA11565}';
  CLASS_E2C: TGUID = '{E4AB6476-E03F-46DE-90FD-0EBD4C2E0CCC}';
type

// *********************************************************************//
// Forward declaration of types defined in TypeLibrary                    
// *********************************************************************//
  IE2C = interface;
  IE2CDisp = dispinterface;

// *********************************************************************//
// Declaration of CoClasses defined in Type Library                       
// (NOTE: Here we map each CoClass to its Default Interface)              
// *********************************************************************//
  E2C = IE2C;


// *********************************************************************//
// Interface: IE2C
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {581C140B-CE91-4741-A9F4-304C6AA11565}
// *********************************************************************//
  IE2C = interface(IDispatch)
    ['{581C140B-CE91-4741-A9F4-304C6AA11565}']
    procedure E2C_DLL1(vI_Data1: OleVariant; vI_Data2: OleVariant); safecall;
    procedure E2C_DLL2(vI_Data1: OleVariant; vI_Data2: OleVariant; vI_Data3: OleVariant); safecall;
    function Get_E2C_DLL1_RESULT: WideString; safecall;
    procedure Set_E2C_DLL1_RESULT(const Value: WideString); safecall;
    function Get_E2C_DLL2_RESULT: WideString; safecall;
    procedure Set_E2C_DLL2_RESULT(const Value: WideString); safecall;
    procedure E2C_DLL3(vI_Data1: OleVariant); safecall;
    function Get_E2C_DLL3_RESULT: WideString; safecall;
    procedure Set_E2C_DLL3_RESULT(const Value: WideString); safecall;
    property E2C_DLL1_RESULT: WideString read Get_E2C_DLL1_RESULT write Set_E2C_DLL1_RESULT;
    property E2C_DLL2_RESULT: WideString read Get_E2C_DLL2_RESULT write Set_E2C_DLL2_RESULT;
    property E2C_DLL3_RESULT: WideString read Get_E2C_DLL3_RESULT write Set_E2C_DLL3_RESULT;
  end;

// *********************************************************************//
// DispIntf:  IE2CDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {581C140B-CE91-4741-A9F4-304C6AA11565}
// *********************************************************************//
  IE2CDisp = dispinterface
    ['{581C140B-CE91-4741-A9F4-304C6AA11565}']
    procedure E2C_DLL1(vI_Data1: OleVariant; vI_Data2: OleVariant); dispid 1;
    procedure E2C_DLL2(vI_Data1: OleVariant; vI_Data2: OleVariant; vI_Data3: OleVariant); dispid 2;
    property E2C_DLL1_RESULT: WideString dispid 5;
    property E2C_DLL2_RESULT: WideString dispid 3;
    procedure E2C_DLL3(vI_Data1: OleVariant); dispid 4;
    property E2C_DLL3_RESULT: WideString dispid 6;
  end;

// *********************************************************************//
// The Class CoE2C provides a Create and CreateRemote method to          
// create instances of the default interface IE2C exposed by              
// the CoClass E2C. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoE2C = class
    class function Create: IE2C;
    class function CreateRemote(const MachineName: string): IE2C;
  end;

implementation

uses ComObj;

class function CoE2C.Create: IE2C;
begin
  Result := CreateComObject(CLASS_E2C) as IE2C;
end;

class function CoE2C.CreateRemote(const MachineName: string): IE2C;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_E2C) as IE2C;
end;

end.
