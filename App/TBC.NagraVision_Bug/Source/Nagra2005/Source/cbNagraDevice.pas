unit cbNagraDevice;

interface

uses SysUtils, Classes;

type

  { 連線送出的資料結構 }

  PDeviceCall = ^TDeviceCall;
  TDeviceCall = record
    Len: array [1..2] of Byte;
    OpMode: Byte;
    ObNameLen: Byte;
    ObName: array [1..32] of Char;
  end;

  { Nagra CAS 會回應的資料結構 }

  PDeviceAnswer = ^TDeviceAnswer;
  TDeviceAnswer = record
    Len: array [1..2] of Byte;
    Status: Byte;
    Len2: array [1..2] of Byte;
    AnswerCode: Byte;
  end;

  { 接收的資料結構 }

  PDeviceBuffer = ^TDeviceBuffer;
  TDeviceBuffer = record
    Data: array [0..1023] of Char;
  end;



implementation

end.
 