unit DevPower_TLB;

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
// File generated on 2004/05/24 13:55:24 from Type Library described below.

// ************************************************************************  //
// Type Lib: D:\App\EMC\Snapshot\Project\Encrypt\DevPower.dll (1)
// LIBID: {3CC1715A-AE02-4295-B5C0-44437F82D3D2}
// LCID: 0
// Helpfile: 
// DepndLst: 
//   (1) v2.0 stdole, (C:\WINNT\system32\stdole2.tlb)
//   (2) v4.0 StdVCL, (C:\WINNT\system32\stdvcl40.dll)
// ************************************************************************ //
// *************************************************************************//
// NOTE:                                                                      
// Items guarded by $IFDEF_LIVE_SERVER_AT_DESIGN_TIME are used by properties  
// which return objects that may need to be explicitly created via a function 
// call prior to any access via the property. These items have been disabled  
// in order to prevent accidental use from within the object inspector. You   
// may enable them by defining LIVE_SERVER_AT_DESIGN_TIME or by selectively   
// removing them from the $IFDEF blocks. However, such items must still be    
// programmatically created via a method of the appropriate CoClass before    
// they can be used.                                                          
{$TYPEDADDRESS OFF} // Unit must be compiled without type-checked pointers. 
{$WARN SYMBOL_PLATFORM OFF}
{$WRITEABLECONST ON}
{$VARPROPSETTER ON}
interface

uses Windows, ActiveX, Classes, Graphics, OleServer, StdVCL, Variants;
  

// *********************************************************************//
// GUIDS declared in the TypeLibrary. Following prefixes are used:        
//   Type Libraries     : LIBID_xxxx                                      
//   CoClasses          : CLASS_xxxx                                      
//   DISPInterfaces     : DIID_xxxx                                       
//   Non-DISP interfaces: IID_xxxx                                        
// *********************************************************************//
const
  // TypeLibrary Major and minor versions
  DevPowerMajorVersion = 1;
  DevPowerMinorVersion = 0;

  LIBID_DevPower: TGUID = '{3CC1715A-AE02-4295-B5C0-44437F82D3D2}';

  IID__Encrypt: TGUID = '{BA5D7910-3C72-43B5-B103-A695035D3193}';
  CLASS_Encrypt: TGUID = '{73E8D9AA-393D-4F6F-97D8-1BD2FB84C8AF}';
type

// *********************************************************************//
// Forward declaration of types defined in TypeLibrary                    
// *********************************************************************//
  _Encrypt = interface;
  _EncryptDisp = dispinterface;

// *********************************************************************//
// Declaration of CoClasses defined in Type Library                       
// (NOTE: Here we map each CoClass to its Default Interface)              
// *********************************************************************//
  Encrypt = _Encrypt;


// *********************************************************************//
// Interface: _Encrypt
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {BA5D7910-3C72-43B5-B103-A695035D3193}
// *********************************************************************//
  _Encrypt = interface(IDispatch)
    ['{BA5D7910-3C72-43B5-B103-A695035D3193}']
    function Encrypt(var OriginalString: WideString): WideString; safecall;
    function Decrypt(var EncryptionString: WideString): WideString; safecall;
  end;

// *********************************************************************//
// DispIntf:  _EncryptDisp
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {BA5D7910-3C72-43B5-B103-A695035D3193}
// *********************************************************************//
  _EncryptDisp = dispinterface
    ['{BA5D7910-3C72-43B5-B103-A695035D3193}']
    function Encrypt(var OriginalString: WideString): WideString; dispid 1610809344;
    function Decrypt(var EncryptionString: WideString): WideString; dispid 1610809345;
  end;

// *********************************************************************//
// The Class CoEncrypt provides a Create and CreateRemote method to          
// create instances of the default interface _Encrypt exposed by              
// the CoClass Encrypt. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoEncrypt = class
    class function Create: _Encrypt;
    class function CreateRemote(const MachineName: string): _Encrypt;
  end;

implementation

uses ComObj;

class function CoEncrypt.Create: _Encrypt;
begin
  Result := CreateComObject(CLASS_Encrypt) as _Encrypt;
end;

class function CoEncrypt.CreateRemote(const MachineName: string): _Encrypt;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_Encrypt) as _Encrypt;
end;

end.
