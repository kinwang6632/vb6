{ ---------------------------------------------------------------------------- }
{                                                                              }
{ PC HOME ONLINE Copyright (c) 2002-2003                                       }
{                                                                              }
{ Project: PC home online 網路家庭( EC2 ) ERP Program                          }
{ Unit: 常數定義                                                               }
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

   { 更新程式識別名稱 }

        { ERP }

        AP_NAME = 'EC2ERP';

        { Channel }

        CNL_NAME = 'EC2CNL';

   { 程式版本, 更新程式時, 請變更此常數 }

        { ERP }

        AP_VERSION = '2.0.6';

        { Channel }

        CNL_VERSION = '1.5.9';

   { ApServer 連線設定檔 }

        { ERP }

        CFG_FILE_NAME = 'EC2ERP.CFG';

        { Channel }

        CNL_CFG_FILE_NAME = 'EC2CNL.CFG';

   { 呼叫的更新程式 }

   UPDATE_EXE = 'LOCUPD.EXE';

   { ------------------------------------------------------------ }


   { MDI Form 建立訊息, 通知 Main Form, 更新 WindowMenu 用 }

   WM_CHILDCREATE = WM_USER + 1;

   { MDI Form 關閉訊息, 通知 Main Form, 更新 WindowMenu 用 }

   WM_CHILDCLOSE = WM_USER + 2;


   { ------------------------------------------------------------ }


   { 不可編輯資料時顏色 }

   NONE_COLOR = TColor( $00E0E3E7 );

   { 編輯狀態時顏色 }

   EDIT_COLOR = clWindow;

   { 編輯狀態時 Focus 的顏色 }

   FOCUS_COLOR = clInfoBk;

   DATE_TITLE_BACK_COLOR = TColor( $00787878 );

   { Remark Color $00E2A496 }

   { 日期顯示 Mask }

   FULL_DATE_MASK = '####/##/##;0;_';

   { 日期(至月份)顯示 Mask }

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

   { 編輯資料狀態 }

   type
     TDMLMode =
       ( dmlNone, dmlInsert, dmlUpdate, dmlDelete, dmlPrint, dmlQuery, dmlClear,
         dmlSave, dmlCancel, dmlQuit );

   { 編輯資料事件 }

     TDMLActionEvent =
       procedure(const ADMLMode: TDMLMode; var Aborted: Boolean) of object;

   { 資料編輯 Button }

     TDMLButtonSet = set of TDMLMode;

   { 資料編輯 ShortCut Key }

     TDMLKeySet = set of TDMLMode;

   { Editor 檢查 Rule 事件 }

     TValidateEvent = function(Sender: TObject): Boolean of object;


   { ------------------------------------------------------------ }
   
   const
   
   { 無新增資料權限訊息 }

   DML_NO_INSERT = '對不起，您無新增資料的權限! 請洽詢系統人員。';

   { 無修改資料權限訊息 }

   DML_NO_UPDATE = '對不起，您無修改資料的權限! 請洽詢系統人員。';

   { 無刪除資料權限訊息 }

   DML_NO_DELETE = '對不起，您無刪除資料的權限! 請洽詢系統人員。';

   { 無列印報表權限訊息 }

   DML_NO_PRINT = '對不起，您無列印報表的權限! 請洽詢系統人員。';

   { 無查詢資料權限訊息 }

   DML_NO_QUERY = '對不起，您無查詢資料的權限! 請洽詢系統人員。';


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
