unit UObjectu;


interface


type
  T4470Obj = class(TObject)
    State : String;//�@��12 bytes, '1' => ���ĭ�; '0'=>�L�ĭ�
    Month1 : String;//'1' => ���ĭ�; '0'=>�L�ĭ�
    Month2 : String;//'1' => ���ĭ�; '0'=>�L�ĭ�
    Month3 : String;//'1' => ���ĭ�; '0'=>�L�ĭ�
    Month4 : String;//'1' => ���ĭ�; '0'=>�L�ĭ�
    Month5 : String;//'1' => ���ĭ�; '0'=>�L�ĭ�
    Month6 : String;//'1' => ���ĭ�; '0'=>�L�ĭ�
    Month7 : String;//'1' => ���ĭ�; '0'=>�L�ĭ�
    Month8 : String;//'1' => ���ĭ�; '0'=>�L�ĭ�
    Month9 : String;//'1' => ���ĭ�; '0'=>�L�ĭ�
    Month10 : String;//'1' => ���ĭ�; '0'=>�L�ĭ�
    Month11 : String;//'1' => ���ĭ�; '0'=>�L�ĭ�
    Month12 : String;//'1' => ���ĭ�; '0'=>�L�ĭ�
  end;
  
  TAccObj = class(TObject) //���b�ϥΤ��|�p��دS�ʪ���
      AccCodeNo : String; //�|�p��إN�X
      AccCodeName : STring; //�|�p��ئW��
      DeCrCode : String;  //(�ɶU����)
      ClearCode : String; //(�a�O�b����)
      IDCode : String;   //(�ḹ����)
      RefCode : String;  //(�ѦҸ��X����)
      UnitCode : String; //(���O����)
      DueDateCode : String; //(��������)
      PrvlgCode : String; //(���b����)

  end;

  TAccBObj = class(TObject) //�~�b�ϥΤ��|�p��دS�ʪ���
      AccCodeNo : String; //�|�p��إN�X
      AccCodeName : STring; //�|�p��ئW��      
      DeCrCode : String;  //(�ɶU����)
      ClearCode : String; //(�a�O�b����)
      IDCode : String;   //(�ḹ����)
      RefCode : String;  //(�ѦҸ��X����)
      UnitCode : String; //(���O����)
      DueDateCode : String; //(��������)
      PrvlgCode : String; //(���b����)

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
