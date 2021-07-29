object TestMain: TTestMain
  Left = 332
  Top = 166
  Caption = 'TestMain'
  ClientHeight = 487
  ClientWidth = 634
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 14
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 165
    Height = 487
    Align = alLeft
    BevelOuter = bvNone
    TabOrder = 0
    object Button1: TButton
      Left = 8
      Top = 26
      Width = 150
      Height = 25
      Caption = 'CM2CT_DLL1 ( '#20837#24235' )'
      TabOrder = 0
      OnClick = Button1Click
    end
    object Button2: TButton
      Left = 8
      Top = 61
      Width = 150
      Height = 25
      Caption = 'CM2CT_DLL1 ( '#39511#36864' )'
      TabOrder = 1
      OnClick = Button2Click
    end
    object Button3: TButton
      Left = 8
      Top = 97
      Width = 150
      Height = 25
      Caption = 'CM2CT_DLL2 ( '#35519#25765' )'
      TabOrder = 2
      OnClick = Button3Click
    end
    object Button4: TButton
      Left = 8
      Top = 132
      Width = 150
      Height = 25
      Caption = 'CM2CT_DLL3 ( '#38936#36864#26009' )'
      TabOrder = 3
      OnClick = Button4Click
    end
    object Button5: TButton
      Left = 8
      Top = 168
      Width = 150
      Height = 25
      Caption = 'CM2CT_DLL5 ( '#32173#20462#36865#22238' )'
      TabOrder = 4
      OnClick = Button5Click
    end
  end
  object Panel2: TPanel
    Left = 165
    Top = 0
    Width = 469
    Height = 487
    Align = alClient
    BevelOuter = bvNone
    TabOrder = 1
    object PageControl1: TPageControl
      Left = 0
      Top = 0
      Width = 469
      Height = 487
      ActivePage = TabSheet2
      Align = alClient
      Style = tsFlatButtons
      TabOrder = 0
      object TabSheet1: TTabSheet
        Caption = #20837#24235' XML'
        object Memo1: TMemo
          Left = 0
          Top = 0
          Width = 461
          Height = 455
          Align = alClient
          Font.Charset = ANSI_CHARSET
          Font.Color = clWindowText
          Font.Height = -13
          Font.Name = 'Courier New'
          Font.Style = []
          Lines.Strings = (
            '<?xml version="1.0" encoding="Big5"?>'
            '<'#20837#24235'>'
            #9'<'#25209#34399'>AB1234567890'
            #9#9'<'#26085#26399'>0960801</'#26085#26399'>'
            #9#9'<'#26178#38291'>1005</'#26178#38291'>'
            #9#9'<'#24288#21830#20195#30908'>123</'#24288#21830#20195#30908'>'
            #9#9'<'#24288#21830#21517#31281'>'#38283#21338'</'#24288#21830#21517#31281'>'
            #9#9'<'#20844#21496#20195#30908'>0</'#20844#21496#20195#30908'>'
            #9#9'<'#26009#34399' '#25976#37327'="2" '#21697#21517'="'#32412#32218#25976#25818#27231'" '#35215#26684'="20*40*30" '#21015#24115#21934#20301'="1">AB1234567890'
            #9#9#9'<'#35373#20633' '#20027#24207#34399'="1">'
            #9#9#9#9'<HFCMAC'#24207#34399'>999000000001</HFCMAC'#24207#34399'>'
            #9#9#9#9'<MTAMAC'#24207#34399'>999000000001</MTAMAC'#24207#34399'>'
            #9#9#9#9'<'#22522#26495#24207#34399'>123456780</'#22522#26495#24207#34399'>'
            #9#9#9#9'<'#27231#22120#24207#34399'>65536</'#27231#22120#24207#34399'>'
            #9#9#9#9'<'#27231#31278'>'#27161#28310#22411'</'#27231#31278'>'
            #9#9#9#9'<'#22411#24335'>0</'#22411#24335'>'
            #9#9#9#9'<'#20445#22266#26399#38480'>0951231</'#20445#22266#26399#38480'>'
            #9#9#9#9'<'#23458#26381#21697#21517'>EMTA('#21407#20126#22826')</'#23458#26381#21697#21517'>'
            #9#9#9#9'<'#29289#26009#22411#34399'>CM550A</'#29289#26009#22411#34399'>'
            #9#9#9#9'<'#30828#39636#29256#26412'>1.0A</'#30828#39636#29256#26412'>'
            #9#9#9'</'#35373#20633'>'
            #9#9#9'<'#35373#20633' '#20027#24207#34399'="1">'
            #9#9#9#9'<HFCMAC'#24207#34399'>999000000002</HFCMAC'#24207#34399'>'
            #9#9#9#9'<MTAMAC'#24207#34399'>999000000002</MTAMAC'#24207#34399'>'
            #9#9#9#9'<'#22522#26495#24207#34399'>123456781</'#22522#26495#24207#34399'>'
            #9#9#9#9'<'#27231#22120#24207#34399'>65537</'#27231#22120#24207#34399'>'
            #9#9#9#9'<'#27231#31278'>'#27161#28310#22411'</'#27231#31278'>'
            #9#9#9#9'<'#22411#24335'>0</'#22411#24335'>'
            #9#9#9#9'<'#20445#22266#26399#38480'>0951231</'#20445#22266#26399#38480'>'
            #9#9#9#9'<'#23458#26381#21697#21517'>EMTA('#21407#20126#22826')</'#23458#26381#21697#21517'>'
            #9#9#9#9'<'#29289#26009#22411#34399'>CM550A</'#29289#26009#22411#34399'>'
            #9#9#9#9'<'#30828#39636#29256#26412'>1.0A</'#30828#39636#29256#26412'>'
            #9#9#9'</'#35373#20633'>'
            #9#9'</'#26009#34399'>'
            #9'</'#25209#34399'>'
            '</'#20837#24235'>')
          ParentFont = False
          ScrollBars = ssBoth
          TabOrder = 0
          WordWrap = False
        end
      end
      object TabSheet2: TTabSheet
        Caption = #39511#36864' XML'
        ImageIndex = 1
        object Memo2: TMemo
          Left = 0
          Top = 0
          Width = 461
          Height = 455
          Align = alClient
          Font.Charset = ANSI_CHARSET
          Font.Color = clWindowText
          Font.Height = -13
          Font.Name = 'Courier New'
          Font.Style = []
          Lines.Strings = (
            '<?xml version="1.0" encoding="Big5"?>'
            '<'#39511#36864' '#21934#34399'="1234" '#20844#21496#20195#30908'="0">'
            '   <'#35373#20633' '#25976#37327'="2">'
            '      <HFCMAC'#24207#34399'>999000000001</HFCMAC'#24207#34399'>'
            '      <HFCMAC'#24207#34399'>999000000002</HFCMAC'#24207#34399'>'
            '   </'#35373#20633'>'
            '</'#39511#36864'>')
          ParentFont = False
          ScrollBars = ssBoth
          TabOrder = 0
          WordWrap = False
          ExplicitLeft = 1
        end
      end
      object TabSheet3: TTabSheet
        Caption = #35519#25765' XML'
        ImageIndex = 2
        object Memo3: TMemo
          Left = 0
          Top = 0
          Width = 461
          Height = 455
          Align = alClient
          Font.Charset = ANSI_CHARSET
          Font.Color = clWindowText
          Font.Height = -13
          Font.Name = 'Courier New'
          Font.Style = []
          Lines.Strings = (
            '<?xml version="1.0" encoding="Big5"?>'
            '<'#35519#25765'>'
            #9'<'#35373#20633' '#25976#37327'="2" '#35519#20837#20844#21496'="6" '#35519#20986#20844#21496'="0">'
            #9#9'<HFCMAC'#24207#34399'>999000000001</HFCMAC'#24207#34399'>'
            #9#9'<HFCMAC'#24207#34399'>999000000002</HFCMAC'#24207#34399'>'
            #9'</'#35373#20633'>'
            '</'#35519#25765'>')
          ParentFont = False
          ScrollBars = ssBoth
          TabOrder = 0
          WordWrap = False
        end
      end
      object TabSheet4: TTabSheet
        Caption = #38936#36864#26009' XML'
        ImageIndex = 3
        object Memo4: TMemo
          Left = 0
          Top = 0
          Width = 461
          Height = 455
          Align = alClient
          Font.Charset = ANSI_CHARSET
          Font.Color = clWindowText
          Font.Height = -13
          Font.Name = 'Courier New'
          Font.Style = []
          Lines.Strings = (
            '<?xml version="1.0" encoding="Big5"?>'
            '<'#38936#36864#26009'>'
            #9'<'#35373#20633' '#25976#37327'="2" '#20316#26989'="1" '#20844#21496#20195#30908'="7">'
            #9#9'<HFCMAC'#24207#34399'>000103A278ED</HFCMAC'#24207#34399'>'
            #9#9'<HFCMAC'#24207#34399'>000103A278EE</HFCMAC'#24207#34399'>'
            #9'</'#35373#20633'>'
            #9'<'#35373#20633' '#25976#37327'="2" '#20316#26989'="2" '#20844#21496#20195#30908'="7">'
            #9#9'<HFCMAC'#24207#34399'>100013A27800</HFCMAC'#24207#34399'>'
            #9#9'<HFCMAC'#24207#34399'>700103A27890</HFCMAC'#24207#34399'>'
            #9'</'#35373#20633'>'
            '</'#38936#36864#26009'>')
          ParentFont = False
          ScrollBars = ssBoth
          TabOrder = 0
          WordWrap = False
        end
      end
      object TabSheet5: TTabSheet
        Caption = #32173#20462#36865#22238
        ImageIndex = 4
        object Memo5: TMemo
          Left = 0
          Top = 0
          Width = 461
          Height = 455
          Align = alClient
          Font.Charset = ANSI_CHARSET
          Font.Color = clWindowText
          Font.Height = -13
          Font.Name = 'Courier New'
          Font.Style = []
          Lines.Strings = (
            '<?xml version="1.0" encoding="Big5"?>'
            '<'#32173#20462#36865#22238'>'
            #9'<'#35373#20633' '#25976#37327'="1" '#20844#21496#20195#30908'="16">'
            #9#9'<'#24207#34399'>'
            #9#9#9'<'#36865#20462' HFCMAC="000E9B673575" MTAMAC="000E9B673577"/>'
            #9#9#9'<'#38936#22238' HFCMAC="000000000100" MTAMAC="000E9B673577"/>'
            #9#9'</'#24207#34399'>'
            #9'</'#35373#20633'>'
            '</'#32173#20462#36865#22238'>')
          ParentFont = False
          ScrollBars = ssBoth
          TabOrder = 0
          WordWrap = False
        end
      end
      object TabSheet6: TTabSheet
        Caption = #21152#35299#23494
        ImageIndex = 5
        object Label1: TLabel
          Left = 39
          Top = 36
          Width = 36
          Height = 14
          Caption = #21152#23494#21069
        end
        object Label2: TLabel
          Left = 39
          Top = 65
          Width = 36
          Height = 14
          Caption = #21152#23494#24460
        end
        object Label3: TLabel
          Left = 39
          Top = 94
          Width = 36
          Height = 14
          Caption = #35299#23494#24460
        end
        object Edit1: TEdit
          Left = 86
          Top = 33
          Width = 177
          Height = 22
          TabOrder = 0
        end
        object Edit2: TEdit
          Left = 86
          Top = 61
          Width = 177
          Height = 22
          TabOrder = 1
        end
        object Button6: TButton
          Left = 39
          Top = 142
          Width = 90
          Height = 25
          Caption = #21152'/'#35299#23494
          TabOrder = 2
          OnClick = Button6Click
        end
        object Edit3: TEdit
          Left = 86
          Top = 90
          Width = 177
          Height = 22
          TabOrder = 3
        end
      end
    end
  end
end
