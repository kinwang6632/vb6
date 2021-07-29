{ ---------------------------------------------------------------------------- }
{                                                                              }
{ PC HOME ONLINE Copyright (c) 2002-2003                                       }
{                                                                              }
{ Project: PC home online �����a�x( EC2 ) ERP Program                          }
{ Unit: �`�Ʃw�q                                                               }
{ Author: Bug                                                                  }
{ Date: 2003/07/09                                                             }
{                                                                              }
{ ---------------------------------------------------------------------------- }

unit ApConst;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, cxEdit;
  

   { ------------------------------------------------------------ }

   const

   { ��s�{���ѧO�W�� }

        { ERP }

        AP_NAME = 'EC2ERP';

        { Channel }

        CNL_NAME = 'EC2CNL';

   { �{������, ��s�{����, ���ܧ󦹱`�� }

        { ERP }

        AP_VERSION = '2.0.6';

        { Channel }

        CNL_VERSION = '1.5.9';

   { ApServer �s�u�]�w�� }

        { ERP }

        CFG_FILE_NAME = 'EC2ERP.CFG';

        { Channel }

        CNL_CFG_FILE_NAME = 'EC2CNL.CFG';

   { �I�s����s�{�� }

   UPDATE_EXE = 'LOCUPD.EXE';

   { ------------------------------------------------------------ }


   { MDI Form �إ߰T��, �q�� Main Form, ��s WindowMenu �� }

   WM_CHILDCREATE = WM_USER + 1;

   { MDI Form �����T��, �q�� Main Form, ��s WindowMenu �� }

   WM_CHILDCLOSE = WM_USER + 2;


   { ------------------------------------------------------------ }


   { ���i�s���Ʈ��C�� }

   NONE_COLOR = TColor( $00E0E3E7 );

   { �s�説�A���C�� }

   EDIT_COLOR = clWindow;

   { �s�説�A�� Focus ���C�� }

   FOCUS_COLOR = clInfoBk;

   DATE_TITLE_BACK_COLOR = TColor( $00787878 );

   { Remark Color $00E2A496 }

   { ������ Mask }

   FULL_DATE_MASK = '####/##/##;0;_';

   { ���(�ܤ��)��� Mask }

   MONTH_DATE_MASK  = '####/##;0;_';


   { ------------------------------------------------------------ }


   { New custom class, published the protected member, so we can access it }

   type
     TcxPublicProperty = class( TcxCustomEdit )
     public
       procedure PostEditValue;
     published
       property Properties;
       property InnerEdit;
       property Style;
       property OnExit;
     end;

   type
     TcxPublicProperties = class( TcxCustomEditProperties )
     published
       property OnButtonClick;
       property OnValidate;
     end;


   { ------------------------------------------------------------ }

   { �s���ƪ��A }

   type
     TDMLMode =
       ( dmlNone, dmlInsert, dmlUpdate, dmlDelete, dmlPrint, dmlQuery, dmlClear,
         dmlSave, dmlCancel, dmlQuit );

   { �s���ƨƥ� }

     TDMLActionEvent =
       procedure(const ADMLMode: TDMLMode; var Aborted: Boolean) of object;

   { ��ƽs�� Button }

     TDMLButtonSet = set of TDMLMode;

   { ��ƽs�� ShortCut Key }

     TDMLKeySet = set of TDMLMode;

   { Editor �ˬd Rule �ƥ� }

     TValidateEvent = function(Sender: TObject): Boolean of object;


   { ------------------------------------------------------------ }
   
   const
   
   { �L�s�W����v���T�� }

   DML_NO_INSERT = '�藍�_�A�z�L�s�W��ƪ��v��! �Ь��ߨt�ΤH���C';

   { �L�ק����v���T�� }

   DML_NO_UPDATE = '�藍�_�A�z�L�ק��ƪ��v��! �Ь��ߨt�ΤH���C';

   { �L�R������v���T�� }

   DML_NO_DELETE = '�藍�_�A�z�L�R����ƪ��v��! �Ь��ߨt�ΤH���C';

   { �L�C�L�����v���T�� }

   DML_NO_PRINT = '�藍�_�A�z�L�C�L�����v��! �Ь��ߨt�ΤH���C';

   { �L�d�߸���v���T�� }

   DML_NO_QUERY = '�藍�_�A�z�L�d�߸�ƪ��v��! �Ь��ߨt�ΤH���C';


   { ------------------------------------------------------------ }


implementation

{ TcxPublicProperty }

{ ---------------------------------------------------------------------------- }

procedure TcxPublicProperty.PostEditValue;
begin
  inherited PostEditValue;
end;

{ ---------------------------------------------------------------------------- }

end.
