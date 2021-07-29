unit LibCM2CT_TLB;

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

// $Rev: 8291 $
// File generated on 2008/1/14 ¤U¤È 04:29:03 from Type Library described below.

// ************************************************************************  //
// Type Lib: D:\App\EMC\MSS\Project\CM2CT\LibCM2CT.tlb (1)
// LIBID: {63450A41-75E4-463A-8E9B-D78924C6FE28}
// LCID: 0
// Helpfile: 
// HelpString: CM2CT Library
// DepndLst: 
//   (1) v2.0 stdole, (C:\WINDOWS\system32\stdole2.tlb)
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
  LibCM2CTMajorVersion = 1;
  LibCM2CTMinorVersion = 0;

  LIBID_LibCM2CT: TGUID = '{63450A41-75E4-463A-8E9B-D78924C6FE28}';

  IID_ICM2CT: TGUID = '{DEAF2F30-F94A-4027-BAD1-08E478D28373}';
  CLASS_CM2CT: TGUID = '{A0C1E91F-BFDC-44EF-8141-628D55AD1018}';
type

// *********************************************************************//
// Forward declaration of types defined in TypeLibrary                    
// *********************************************************************//
  ICM2CT = interface;
  ICM2CTDisp = dispinterface;

// *********************************************************************//
// Declaration of CoClasses defined in Type Library                       
// (NOTE: Here we map each CoClass to its Default Interface)              
// *********************************************************************//
  CM2CT = ICM2CT;


// *********************************************************************//
// Interface: ICM2CT
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {DEAF2F30-F94A-4027-BAD1-08E478D28373}
// *********************************************************************//
  ICM2CT = interface(IDispatch)
    ['{DEAF2F30-F94A-4027-BAD1-08E478D28373}']
    function CM2CT_DLL1(aData1: OleVariant; aData2: OleVariant): WideString; safecall;
    function CM2CT_DLL2(aData1: OleVariant): WideString; safecall;
    function CM2CT_DLL3(aData1: OleVariant): WideString; safecall;
    function CM2CT_DLL5(aData1: OleVariant): WideString; safecall;
  end;

// *********************************************************************//
// DispIntf:  ICM2CTDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {DEAF2F30-F94A-4027-BAD1-08E478D28373}
// *********************************************************************//
  ICM2CTDisp = dispinterface
    ['{DEAF2F30-F94A-4027-BAD1-08E478D28373}']
    function CM2CT_DLL1(aData1: OleVariant; aData2: OleVariant): WideString; dispid 201;
    function CM2CT_DLL2(aData1: OleVariant): WideString; dispid 202;
    function CM2CT_DLL3(aData1: OleVariant): WideString; dispid 203;
    function CM2CT_DLL5(aData1: OleVariant): WideString; dispid 204;
  end;

// *********************************************************************//
// The Class CoCM2CT provides a Create and CreateRemote method to          
// create instances of the default interface ICM2CT exposed by              
// the CoClass CM2CT. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoCM2CT = class
    class function Create: ICM2CT;
    class function CreateRemote(const MachineName: string): ICM2CT;
  end;

implementation

uses ComObj;

class function CoCM2CT.Create: ICM2CT;
begin
  Result := CreateComObject(CLASS_CM2CT) as ICM2CT;
end;

class function CoCM2CT.CreateRemote(const MachineName: string): ICM2CT;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_CM2CT) as ICM2CT;
end;

end.
