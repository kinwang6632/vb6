unit Project1_TLB;

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
// File generated on 2003/09/22 16:25:42 from Type Library described below.

// ************************************************************************  //
// Type Lib: D:\App\NDS\Project\簡易開關機程式\Testing\Server\Project1.tlb (1)
// LIBID: {3796FE8A-72AD-4F96-8B4B-F7F43B5C545A}
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
  Project1MajorVersion = 1;
  Project1MinorVersion = 0;

  LIBID_Project1: TGUID = '{3796FE8A-72AD-4F96-8B4B-F7F43B5C545A}';

  IID_Iabc: TGUID = '{A05FA713-F449-4B92-A646-1867142821AE}';
  CLASS_abc: TGUID = '{A9629298-BD38-4EB2-9841-3B26665E8450}';
type

// *********************************************************************//
// Forward declaration of types defined in TypeLibrary                    
// *********************************************************************//
  Iabc = interface;
  IabcDisp = dispinterface;

// *********************************************************************//
// Declaration of CoClasses defined in Type Library                       
// (NOTE: Here we map each CoClass to its Default Interface)              
// *********************************************************************//
  abc = Iabc;


// *********************************************************************//
// Interface: Iabc
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {A05FA713-F449-4B92-A646-1867142821AE}
// *********************************************************************//
  Iabc = interface(IDispatch)
    ['{A05FA713-F449-4B92-A646-1867142821AE}']
    procedure Method1(sI_Data: OleVariant; out sI_Result: OleVariant; out sI_ErrorCode: OleVariant; 
                      out sI_ErrorMsg: OleVariant); safecall;
  end;

// *********************************************************************//
// DispIntf:  IabcDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {A05FA713-F449-4B92-A646-1867142821AE}
// *********************************************************************//
  IabcDisp = dispinterface
    ['{A05FA713-F449-4B92-A646-1867142821AE}']
    procedure Method1(sI_Data: OleVariant; out sI_Result: OleVariant; out sI_ErrorCode: OleVariant; 
                      out sI_ErrorMsg: OleVariant); dispid 1;
  end;

// *********************************************************************//
// The Class Coabc provides a Create and CreateRemote method to          
// create instances of the default interface Iabc exposed by              
// the CoClass abc. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  Coabc = class
    class function Create: Iabc;
    class function CreateRemote(const MachineName: string): Iabc;
  end;

implementation

uses ComObj;

class function Coabc.Create: Iabc;
begin
  Result := CreateComObject(CLASS_abc) as Iabc;
end;

class function Coabc.CreateRemote(const MachineName: string): Iabc;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_abc) as Iabc;
end;

end.
