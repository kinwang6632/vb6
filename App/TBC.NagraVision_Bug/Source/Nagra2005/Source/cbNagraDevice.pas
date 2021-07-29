unit cbNagraDevice;

interface

uses SysUtils, Classes;

type

  { �s�u�e�X����Ƶ��c }

  PDeviceCall = ^TDeviceCall;
  TDeviceCall = record
    Len: array [1..2] of Byte;
    OpMode: Byte;
    ObNameLen: Byte;
    ObName: array [1..32] of Char;
  end;

  { Nagra CAS �|�^������Ƶ��c }

  PDeviceAnswer = ^TDeviceAnswer;
  TDeviceAnswer = record
    Len: array [1..2] of Byte;
    Status: Byte;
    Len2: array [1..2] of Byte;
    AnswerCode: Byte;
  end;

  { ��������Ƶ��c }

  PDeviceBuffer = ^TDeviceBuffer;
  TDeviceBuffer = record
    Data: array [0..1023] of Char;
  end;



implementation

end.
 