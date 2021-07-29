unit cbIphlpapi;

interface


uses Windows, Messages, SysUtils, Classes;


procedure GetMACAddress(aMACList: TStringList);


implementation

const
  MAX_ADAPTER_NAME_LENGTH = 256;
  MAX_ADAPTER_DESCRIPTION_LENGTH = 128;
  MAX_ADAPTER_ADDRESS_LENGTH = 8;

type

  PIP_ADDRESS_STRING = ^IP_ADDRESS_STRING;
  IP_ADDRESS_STRING = array[0..15] of Char;

  PIP_ADDR_STRING = ^IP_ADDR_STRING;
  IP_ADDR_STRING = record
    Next: PIP_ADDR_STRING;
    IpAddress: IP_ADDRESS_STRING;
    IpMask: IP_ADDRESS_STRING;
    Context: DWORD;
  end;

  PIP_ADAPTER_INFO = ^IP_ADAPTER_INFO;
  IP_ADAPTER_INFO = record
    Next: PIP_ADAPTER_INFO;
    ComboIndex: DWORD;
    AdapterName: array[1..MAX_ADAPTER_NAME_LENGTH + 4] of  char;
    Description: array[1..MAX_ADAPTER_DESCRIPTION_LENGTH + 4] of char;
    AddressLength: UINT;
    Address: array[1..MAX_ADAPTER_ADDRESS_LENGTH] of byte;
    Index: DWORD;
    aType: UINT;
    DHCPEnabled: UINT;
    CurrentIPAddress: PIP_ADDR_STRING;
    IPAddressList: IP_ADDR_STRING;
    GatewayList: IP_ADDR_STRING;
    DHCPServer: IP_ADDR_STRING;
    HaveWINS: BOOL;
    PrimaryWINSServer: IP_ADDR_STRING;
    SecondaryWINSServer: IP_ADDR_STRING;
    LeaseObtained: LongInt;
    LeaseExpires: LongInt;
  end;

function  GetAdaptersInfo(pAdapterInfo: PIP_ADAPTER_INFO;  pOutBufLen: PULONG): DWORD; stdcall;
  external 'Iphlpapi.dll' name 'GetAdaptersInfo';

  
procedure GetMACAddress(aMACList: TStringList);
var
  aCode: Cardinal;
  aBufferSize: PULONG;
  aAdapters, aNext: PIP_ADAPTER_INFO;
  aIndex: Integer;
  aMACText: String;
begin
  if not Assigned( aMACList ) then Exit;
  aMACList.Clear;
  New( aBufferSize );
  try
    GetAdaptersInfo( nil, aBufferSize );
    aAdapters := AllocMem( aBufferSize^ );
    try
       aCode := GetAdaptersInfo( aAdapters, aBufferSize );
       if ( aCode = 0 ) then
       begin
         aNext := aAdapters;
         repeat
           aMACText := EmptyStr;
           for aIndex := 1 to aNext.AddressLength do
             aMACText := ( aMACText + IntToHex( aNext.Address[aIndex], 2 ) );
           aMACList.Add( aMACText );  
           aNext := aNext.Next;
         until ( not Assigned( aNext ) );
       end;
    finally
      FreeMem( aAdapters, aBufferSize^ );
    end;
  finally
    Dispose( aBufferSize );
  end;
end;


end.
 