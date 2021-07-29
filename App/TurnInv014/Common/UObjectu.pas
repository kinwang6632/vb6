unit UObjectu;


interface


type
  T4470Obj = class(TObject)
    State : String;//共有12 bytes, '1' => 有效值; '0'=>無效值
    Month1 : String;//'1' => 有效值; '0'=>無效值
    Month2 : String;//'1' => 有效值; '0'=>無效值
    Month3 : String;//'1' => 有效值; '0'=>無效值
    Month4 : String;//'1' => 有效值; '0'=>無效值
    Month5 : String;//'1' => 有效值; '0'=>無效值
    Month6 : String;//'1' => 有效值; '0'=>無效值
    Month7 : String;//'1' => 有效值; '0'=>無效值
    Month8 : String;//'1' => 有效值; '0'=>無效值
    Month9 : String;//'1' => 有效值; '0'=>無效值
    Month10 : String;//'1' => 有效值; '0'=>無效值
    Month11 : String;//'1' => 有效值; '0'=>無效值
    Month12 : String;//'1' => 有效值; '0'=>無效值
  end;
  
  TAccObj = class(TObject) //內帳使用之會計科目特性物件
      AccCodeNo : String; //會計科目代碼
      AccCodeName : STring; //會計科目名稱
      DeCrCode : String;  //(借貸條件)
      ClearCode : String; //(懸記帳條件)
      IDCode : String;   //(戶號條件)
      RefCode : String;  //(參考號碼條件)
      UnitCode : String; //(幣別條件)
      DueDateCode : String; //(到期日條件)
      PrvlgCode : String; //(機帳條件)

  end;

  TAccBObj = class(TObject) //外帳使用之會計科目特性物件
      AccCodeNo : String; //會計科目代碼
      AccCodeName : STring; //會計科目名稱      
      DeCrCode : String;  //(借貸條件)
      ClearCode : String; //(懸記帳條件)
      IDCode : String;   //(戶號條件)
      RefCode : String;  //(參考號碼條件)
      UnitCode : String; //(幣別條件)
      DueDateCode : String; //(到期日條件)
      PrvlgCode : String; //(機帳條件)

  end;

  TSysObj = class(TObject)
     UserID : String;
     UserName : String;
     Passwd : String;
     GroupID : String;
     CompCode : String;
     CompName : String;
     PrgName : STring;
     Version : String;
     IsSysOp : boolean;
     TempRptPath : String;
     IfUseChineseFormat : boolean;
     IfShowDateHeader  : boolean;


     AccType : Integer;
     VouTypeCode : String;
     VouKeyType : Integer;
     AccCodeType : String;
     AccCodeLength : Integer;
     DepCode : String;
     VoyRptForm : String;
     CopyRightRate : double;
     OpenYear : Integer;
     TaxRate : double;
     DayLimit : Integer;
     RefNoCheck : Integer;
     ForeignCheck : Integer;

  end;

  TNormalObj = class(TObject)
     s_Code : String;
     s_Desc : String;
  end;

  TDeptObj = class(TObject)
     s_Code : String;
     s_Desc : String;
     s_MainDepCode : String;
  end;
    
  TQryCondObj = class(TObject)
     s_Desc : String;
     s_FieldName : String;
  end;


implementation



end.
