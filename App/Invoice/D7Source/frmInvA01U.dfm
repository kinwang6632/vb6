object frmInvA01: TfrmInvA01
  Left = 473
  Top = 273
  BorderStyle = bsDialog
  Caption = #30332#31080#38283#31435
  ClientHeight = 533
  ClientWidth = 606
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 14
  object Panel2: TPanel
    Left = 0
    Top = 0
    Width = 606
    Height = 533
    Align = alClient
    BevelOuter = bvNone
    BorderWidth = 3
    TabOrder = 0
    object Panel1: TPanel
      Left = 3
      Top = 3
      Width = 600
      Height = 41
      Align = alTop
      BevelOuter = bvNone
      BorderWidth = 3
      TabOrder = 0
      object Bevel3: TBevel
        Left = 3
        Top = 35
        Width = 594
        Height = 3
        Align = alBottom
        Shape = bsBottomLine
      end
      object rdoAction1: TcxRadioButton
        Left = 18
        Top = 8
        Width = 90
        Height = 17
        Caption = #30332#31080#32080#36681
        Checked = True
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
        TabStop = True
        OnClick = rdoAction1Click
      end
      object rdoAction2: TcxRadioButton
        Left = 117
        Top = 8
        Width = 93
        Height = 17
        Caption = #30332#31080#38283#31435
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        OnClick = rdoAction2Click
      end
    end
    object ActionPageControl: TcxPageControl
      Left = 3
      Top = 44
      Width = 600
      Height = 486
      ActivePage = TransInv1
      Align = alClient
      Style = 9
      TabOrder = 1
      ClientRectBottom = 486
      ClientRectRight = 600
      ClientRectTop = 21
      object TransInv1: TcxTabSheet
        Caption = '  '#32080#36681#30059#38754'  '
        OnShow = TransInv1Show
        object BkImage: TImage
          Left = 0
          Top = 0
          Width = 600
          Height = 465
          Align = alClient
          Picture.Data = {
            0A544A504547496D61676589140000FFD8FFE000104A46494600010101006000
            600000FFDB004300080606070605080707070909080A0C140D0C0B0B0C191213
            0F141D1A1F1E1D1A1C1C20242E2720222C231C1C2837292C30313434341F2739
            3D38323C2E333432FFDB0043010909090C0B0C180D0D1832211C213232323232
            3232323232323232323232323232323232323232323232323232323232323232
            32323232323232323232323232FFC200110801CD021503012200021101031101
            FFC4001A000101010101010100000000000000000000010203040506FFC40017
            010101010100000000000000000000000000010203FFDA000C03010002100310
            000001FDF8000000000000000000000000000000000000000000000000000000
            00000025C6E02800000000000000000000000000000000000000000000004CE7
            35BE697B31BD6428000000000000000000000000000000000000000000072B31
            A09406F0B3A8DE4000000000000000000000025000000000000002663690D334
            A90C131BB2C165250DDE7D359A3500000000000000000006635225590D5C8D22
            CA280000000018DE63031B4A2EF95B0B9955002A0DEF8DB3A6703B396ACD8D40
            00000000006634CD28A67588A965816A11ACDAD0B00000000000E4D679ED144A
            040532AACEA52A2051BCF4B90DC000000000019D48C91ABACC4DE60A250A0916
            2B699B373325D5C7402C0000128E79EBCF1A4A960000092B52D8CDA2594DEA5D
            E028000000000067523329A800160A04B21528A1286F3A642800006752319D4C
            6E012894228C6F1DF53934962A2551B9ADE42C0000000000244942582800160A
            204AA202B56564280000019D239E7AF3CEA4A9575BD4E7BAB96759358D8E4EB9
            CD5AD40A00000000000063723226A0A2C0002A003512281A5B02C000000019D2
            2280A000CEB3628A0000000000000004A30DE65825020A10A0594006928B0000
            000000000000063680A0000000000000000008A8CCD8C4E85E4D08D530DC32DD
            252C000000000000000000000000000000000000000000025202800000000000
            00000000000000000000000000000000000012C2800000000000000000000000
            000000000000000000000000012800000000000000000000000000010BE77973
            7DFBE3DAC0A0000000000000000000000000000000000000000000064D79EF92
            566E73AF67A7C7EAD6742C000000000000000000000000000000000000000000
            F3FA3C7BCDF4F9FD1E6B3C4D4CEE5CD35D397A0F68DE00000000000000000000
            00000000000000000000003C5CBD1E6C6FD3BF12CB8D62526AA7D0F0FD34D396
            2E7D0F16D7D42C000000000000000000000000000000000000000F1F0F7F871B
            8BA33C7B71B34B25BACD587437EECEB5CC280000000000000000000000000000
            000000000038F97DFE3CEB9DB995CBAF234CD0BB971EAE9DB590D64000000000
            00000000000000000000000000000006363C19FA1E7CEBCDCFA773C97D5CE5E7
            BEFD4D68DE00000000000000000000000000000000000000000000038F6E5D60
            280000000000000000000000000000000000000000000000038F6E7D20280000
            000000000000000000000000000000000000000000039F4E7D20280000000000
            000000000000000000000000000000000000039ED88E82800000000000000000
            0000000000000000000000000000001CFA039750000000000000000000000000
            0000000000000000000000000C4E9CE3A39F4A00000000000000000000000000
            000000000000000000000073E838F6000000000000000000000000000000003F
            FFC4002810000102060104020203000000000000000102110003102050604012
            21313230411380224243FFDA0008010100010502FD19341A01B0688F9C7D0CF9
            BDE1F26F0F47A9CE9B47C42AF913C47CE77BDB847E6163F17EA9E4DA38679CDC
            21C338A6CB3C3FC6076B9F84FCF3F18F0467DA1A828D9C6B4D8D97686D2CE2DB
            93E398DA4B434343437EDAAD6D03BA7405A9A92FD73F34D65AB3C4B01312C4F5
            1A4AF6CD7534CA4DF5A7DD25FBE6667B21749BE2D943F966667B3C26644CF436
            CA0C9CCCC6A9B443811D698334419A611D672D3032AC37A10F97507153E2E4A5
            C80C32CB2450C0A7DDB2DF32A4F4D3ECC0B1A1280334A1D40A4A63E89ED624B1
            19B21E0CB82853A6542A547E354095012067BFD7421EFA127CE849D0C79D0BFB
            E84AD113A19EDA20EC74221C24BE864424F50D054582030DB7FFC4001C110002
            0203010100000000000000000000011110400020506080FFDA0008010301013F
            01F14B820718DA12B558B1708D25B83C737843C72FC93C7A3F6E070408379609
            37862930AE0D8DC1A9379C3C63E04FFFC4001C11000203010101010000000000
            00000000011110405020007080FFDA0008010201013F01C91824E0932EDBF3EC
            1B465F2FCFCEF991648C13D0E05A22170BE9C4E098179E1B9178E11E40BCA57E
            04FFC40027100000050205030500000000000000000000011121603150021030
            516112707120224191A0FFDA0008010100063F02FDF8B0280A1560499F4DF972
            5BF1EDA257B43D05BE21FAFCC0EA2A18305336812FC5DD351205C7705826E3DC
            185039DFCB8281E281E23819F9819F619768226D0358227D413981A95608F539
            77FFC40028100002010205040203010100000000000001110021311020405060
            3041516171817091B1D1E1FFDA0008010100013F21FC28EBC0DD7815981700ED
            8E0B4237FD8E5143177C6CE03D85EC40BC581B601A3689090A0F28C79C2CC860
            B65010D61CEF4A252E3A3760C4BF56338B75C18CE6CE77C443682382991C1A25
            A38FA834E40831ED58023CE3CB4A05405C3D1BB29F4CAC96CB4EEAE0990E1270
            9C78868ECEB0BC709EA90E1CC071392F6968C47EF02A6CC55EA14FA2E02F3054
            C0168CE8C74C4B171C78DA385097023319F11FA31E20A0674427463A443818AF
            510F110F0221E04FAC1388C0492A31954034446901EA11E2238D7044C1E71076
            C3C0E07C22C00D2DB480750DB15583CF2D15C874C3A303AC9122CE2856A4E885
            F500C7B80B0F54475D3D5055E8EAD605145D216DE9389418098886FA7F038E07
            DF819E556177986C7AE01697C7CA125DE1BF89E009007BDE7CC71C28B77DF806
            28554518F7CA134825DF8848053AEF54426AC4923B986086CDF50A7F518BE838
            0053EE110CEF8B3D1BD19FAE0FF761A2F78099DF1AAF7AB7A62443B5F1A82843
            10C1803207986E00422902C73B60128FF417DDAA2AF3BC548A1C00530171179C
            2A7D4BB7FD6EEFE14B561101C1ED2B2D8131422841082DBBA5214EE7DE0F6948
            B26D041026956F0402118C7A31D20AC5D2C71CBE00894A550D4EF4141872BFB8
            EC9731D04180847084C0242DEC008873CEFA304148AB611E04EF0FD381FF00EA
            007725986FC43F90E07FCDC0FF0078B81FF7703A43EDF03693C8E074290E046B
            287C0EAC3C6FF1C13E1781A952A81B2FC0DFD81B1954488A11C08EF019ED0C53
            70CF2EFFDA000C03010002000300000010F3CF3CF3CF3CF3CF3CF3CF3CF3CF3C
            F3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF31F3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF
            3CF3CF3CF3CF3CF382F97FCF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3
            CF1A40195CF3CF3CF3CF3CF3CF3CF3CE3CF3CF3CF3CF3CF3DDD33A6543843F3C
            F3CF3CF3CF3CF3CF2B3A6BB7F3CF3CF3CF35051FCA8D30DE7D6D3CF3CF3CF3C8
            C3F256C26FCF3CE3CF3CF2809B4966B5983CCB73CF3CF3CF3C4C4DCD080FE5F4
            523CF3CF3CF93BF39E58C39DFF003CF3CF3CF3C4D0F0CB005A4C10FF00CF3CF0
            95A77FFEEEE038ACF3CF3CF3CF2C442F7962C08C0A0CFCF3CF3C799FC8BCB1F7
            4FCF3CF3CF3CF3C3C2F6481033C58293CF3CF3CF167BCF3C5FCF3CF3CF3CF3CF
            3CF25082A06428019D3CF3CF3CF3CF3CF3C6BCF3CF3CF3CF3CF3CF3C8FF2B62B
            5EBCF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF32F3CF3CF3CF3CF3C
            F3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF
            3CF3CF3CF3CF3CF3CF3CF3CD3CF3CF3CF3CF3CF3CF3CF3CF3CF3F9FCF3CF3CF3
            CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CD3C17953CF3CF3CF3CF3CF3CF3CF3C
            F3CF3CF3CF3CF3CF3CF3C5FF006D77FCF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3
            CF3CF3CF3C8DD39DABB7BCF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3C3AD
            553B389FCF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CDA27AF2C2F3CF3CF
            3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF04652FBCF3CF3CF3CF3CF3CF3CF3
            CF3CF3CF3CF3CF3CF3CF3CF3C5FCF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3C
            F3CF3CF3CF3CDFCF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF
            FCF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3C7FCF3CF3CF3CF3
            CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CB1CF3CF3CF3CF3CF3CF3CF3CF3C
            F3CF3CF3CF3CF3CF3CF3CF3CF3EF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF
            3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CF3CFFFC4002011010002
            0104030101000000000000000001001110202130314041505160FFDA00080103
            01013F10F9295F04F6897129F3E8CA0F7E585CA6230EE06801CC172A54A95C0B
            7CB6DE0DE86CCAC59C21729C19A8F5C2378B952E30BF795C27784B815A1EA05C
            21DB559AAF0ED1EF84EF89D257B857AC5E5541B32A7AE10389D43515E1FCC530
            FC8350FDC5E21AE0AC2EBB96E87F7947814E51E6B65B2B259FDC3D8C767E025E
            F834F9CECB276978E9E7748EFB32FD4A628EFCC764BC1862B7CC6759302EC79C
            7EE584047E04FC27BF84FC2F5F0DFE6FFFC40020110100020104030101000000
            000000000001001110202130314041505160FFDA0008010201013F10F92AFE0F
            A20D306CF3ECEB223CB69290462ED7170E11E658B8B97C02CC942A543150A117
            03C578741C2959B8170EF40DF7C2E0D24B970EF551A037953697836385E23487
            D46FDE36CEE8ECE43EFC33525CA307EE0047F625C640389E20D68328D072A78A
            F354A95A2BFB7F410EBE02DB1D3CE37531EB0C3BF39EE1E908E164BF30538340
            57996C18B9ECF39FC4088C17C03E11D7C23E174FC261FCDFFFC4002C10010002
            020103020504020300000000000100112131411051614071205060819130A1B1
            F0D1E170C1F1FFDA0008010100013F10FF0085383E83BE706CFA016A2A83B5AC
            14A8239F9F08988A911B150BB5DCB358EF16CF9F5A94B1B5CCC1D2A2418F93F1
            06CF9DD18C4E678EA46583CFC814201F4B52BE0E23AA04BE27897C730EF09712
            07B4F72FDA16A3EB56BA3156E2C1BF408F0CA1CC1A5A5338A1ED3F3044B21BBF
            6847BF4179086BA2BA6DD33DE39F58A960DC65F51A8F641B2FF5CD3673D37D1D
            665BA752CC0C710894C746B38F241F1CCD4312A1BF225CF31F9991031E7D5397
            A2E2CF687458EFF5DEE04D473A83399CEE0F88448E0CC2881B8B4246EDD71122
            48F17CC7B32C34DB6F3716F2C76306B106186A1853B82259E812E2A920CE605E
            AE09752F3197C42336CA9BBF15FE8BA979E8B2DED2BA0A63431949CF99A17B89
            6430171C2749DB03A5E618CB5F7817D1420370338161D7F080180F41C41CF1D5
            3D900474541CE62418908BDA0317130EAB0888B6073FA2A1B495EBF6979C7F13
            ED30F111D33046259A839A899C46D607710FEDCA034B971D06A18037369ECA9E
            678222E2057A1E3E1352EF6CE65F41B2A69A84B962C541E2208AE0E9B217028A
            FD1014CA93B4A070FE627695E667BCD32C5129ED71398302966FF99788D5E628
            1D896E431E798770BF68727EE4AF11E7BCBA95551F43C45C4A9CCFBCAE9ACF4B
            E9CC632BA730DC580C307E9A099D4C8AC9EF07767133C473A96983F8943EFDC8
            35BD7786723B95622D4A1962DB198077994D62A5BDAA5BCCB820428F45C4250F
            4C4AE9C4E20C671097F1105E20EFFA2EA2EB0466140965E3A347A3697E60E225
            2CFC403EFDA1A30A18A18AB67BD77859488F6489EDFD997FF44AB1716EA5BB27
            2C230B4AF40B45CA4B4C75AF8F8847A11F1025F1D34FD211143C329E94787E21
            DA7E2538FD888FF826381F89CC4774735B8DC020D9DE57DF6E54F7E976D41BA9
            4E533E8AAD6A574BEAEBAEA5E481F05C0952A1702A63F4EF302E25948CB8B4AF
            DE01EFEF046058BB5F621A422094EA5C1DC7F13718572F115C89D04330035E8D
            04CC4B574AB952BA9398425F4BF1D065CB970DCE47F5102B9BD4A827020E1C98
            15AF832078DFB4112CE882405B30C2EB3E9459D399CFC169848C36427DFA7DE6
            7BFC34E5FD6F7A586A00D7C7EC3B3D4DAB113AD75AE86E1C4BE84597062E668F
            50D61A6480434FA949CA47E167338952A07C0762039815EA058790F7F5759497
            97ED2E3A8952BA54F7E9CF535A7AA41DFC8103316B311DE080AE7A14F311C424
            01A3E79A7D0696306CBFA0F0B3E83D7BFE83EEEDF42051F267BB94DF6788E836
            A2DF31F9FA8154036C518982F64405B3B5630094203DFE80451DDEC8DBB2B595
            63659516945CBCFF00BF9F6A2025EAA8B44FE258D8163DA54DA886822EC6EE3D
            C8210168E5F9D6AB01BD0D19EAA200E43B9FDA9860D44E166566A1E26558C154
            0CD7E1F9D587707FEA5F1175566FC7472EECA19EE7FA85651C73300A8AD6F883
            7537882D6D71F7FEBF3A6E14003DF9820C2D9AF1054C3C1FE708559407F7ED30
            16F238BE8194C0A712DD12891591EDC7CE9A258B41913C26B71292F72E6171A1
            7043C17F98114788EEE595A9DD2747DE061D8C508F521F617F89B67CDC112A19
            DDCB1251EF48B1F8F9B0D0542EFCCBA8A284D78DC6BDEBBCB025E8896EE55FE1
            1AAABCB2D70A2386F530B55ACBACF69A2B08CCA4BF301451AF9B69CB722F7971
            D5F78D3156EBED02B7CCBB79EF1053F301071EF0070BE19C1B7B4396D86DB1B7
            C1079A1F372314B4DFF7DE2AB2ECBFF900F02789901C30C6D671E62175544B6E
            AE26A02E9E8CEB11C1C0DF74FEBF3844023B196639D0D6FDE36C0F35DA60A3BC
            3FF51D25082B0ED9EE8FBAB83BC0612095EC666856EE60F6F9D29DDC35A95058
            6E8731E0155B258FC2AFB46855C41A77E263BC0D6060F2BAC577897812EAF5F3
            BA321D980AB5FF004DC739058CD9E6509C7611DC1BD2ACBFE138417CD310AC8B
            C7117B1BBBB735F3EDD3080F2B5F41EDBC04FC7D06159CFED7FC0CC41BFF0060
            FA0F6056B7BA3FEFE83261E5DFD9C3FB44208D8EBE820044B1D91511B70BF1C7
            D07600C6153FAC4DFB7D076BDBE53C78FA0CD5A8EC4E189474B40FE4F1F41BB1
            3ED07C3E267564F609F41217EB0179875D20BB5FD5DFFFD9}
          Stretch = True
        end
        object Label5: TLabel
          Left = 37
          Top = 112
          Width = 252
          Height = 16
          Caption = #35531#30906#35469#24744#30340#38928#38283#25910#36027#36039#26009#27284#23384#25918#26044' : '
          Font.Charset = CHINESEBIG5_CHARSET
          Font.Color = clTeal
          Font.Height = -16
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          ParentFont = False
          Transparent = True
        end
        object Label6: TLabel
          Left = 37
          Top = 164
          Width = 252
          Height = 16
          Caption = #35531#30906#35469#24744#30340#24460#38283#25910#36027#36039#26009#27284#23384#25918#26044' : '
          Font.Charset = CHINESEBIG5_CHARSET
          Font.Color = clTeal
          Font.Height = -16
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          ParentFont = False
          Transparent = True
        end
        object Label8: TLabel
          Left = 224
          Top = 24
          Width = 144
          Height = 24
          Caption = #25910#36027#36039#26009#32080#36681
          Font.Charset = CHINESEBIG5_CHARSET
          Font.Color = clBackground
          Font.Height = -24
          Font.Name = #27161#26999#39636
          Font.Style = []
          ParentFont = False
          Transparent = True
        end
        object radPreCreated: TRadioButton
          Left = 232
          Top = 244
          Width = 81
          Height = 17
          Caption = #38928#38283
          Color = clWhite
          Font.Charset = CHINESEBIG5_CHARSET
          Font.Color = clTeal
          Font.Height = -16
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          ParentColor = False
          ParentFont = False
          TabOrder = 0
        end
        object radCreated: TRadioButton
          Left = 328
          Top = 244
          Width = 65
          Height = 17
          Caption = #24460#38283
          Checked = True
          Color = clWhite
          Font.Charset = CHINESEBIG5_CHARSET
          Font.Color = clTeal
          Font.Height = -16
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          ParentColor = False
          ParentFont = False
          TabOrder = 1
          TabStop = True
        end
        object btnExit1: TcxButton
          Left = 343
          Top = 298
          Width = 85
          Height = 28
          Caption = #32080#26463
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -13
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 3
          OnClick = btnExit1Click
          Glyph.Data = {
            76010000424D7601000000000000760000002800000020000000100000000100
            04000000000000010000130B0000130B00001000000000000000000000000000
            800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
            FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
            333333333333333333333333333333333333333FFF33FF333FFF339993370733
            999333777FF37FF377733339993000399933333777F777F77733333399970799
            93333333777F7377733333333999399933333333377737773333333333990993
            3333333333737F73333333333331013333333333333777FF3333333333910193
            333333333337773FF3333333399000993333333337377737FF33333399900099
            93333333773777377FF333399930003999333337773777F777FF339993370733
            9993337773337333777333333333333333333333333333333333333333333333
            3333333333333333333333333333333333333333333333333333}
          NumGlyphs = 2
        end
        object BitRun: TcxButton
          Left = 200
          Top = 298
          Width = 85
          Height = 28
          Caption = #22519#34892
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -13
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 2
          OnClick = BitRunClick
          Glyph.Data = {
            76010000424D7601000000000000760000002800000020000000100000000100
            04000000000000010000120B0000120B00001000000000000000000000000000
            800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
            FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00555555555555
            555555555555555555555555555555555555555555FF55555555555559055555
            55555555577FF5555555555599905555555555557777F5555555555599905555
            555555557777FF5555555559999905555555555777777F555555559999990555
            5555557777777FF5555557990599905555555777757777F55555790555599055
            55557775555777FF5555555555599905555555555557777F5555555555559905
            555555555555777FF5555555555559905555555555555777FF55555555555579
            05555555555555777FF5555555555557905555555555555777FF555555555555
            5990555555555555577755555555555555555555555555555555}
          NumGlyphs = 2
        end
        object edtPreCreated: TcxButtonEdit
          Left = 296
          Top = 108
          Properties.Buttons = <
            item
              Default = True
              Kind = bkEllipsis
            end>
          Properties.OnButtonClick = edtPreCreatedPropertiesButtonClick
          Style.StyleController = dtmMain.cxEditStyle
          TabOrder = 4
          Width = 260
        end
        object edtCreated: TcxButtonEdit
          Left = 296
          Top = 160
          Properties.Buttons = <
            item
              Default = True
              Kind = bkEllipsis
            end>
          Properties.OnButtonClick = edtCreatedPropertiesButtonClick
          Style.StyleController = dtmMain.cxEditStyle
          TabOrder = 5
          Width = 260
        end
      end
      object TransInv2: TcxTabSheet
        Caption = '  '#32080#36681#32080#26524'  '
        ImageIndex = 1
        OnShow = TransInv1Show
        object Label11: TLabel
          Left = 58
          Top = 311
          Width = 65
          Height = 18
          Caption = #34389#29702#36914#24230':'
          Font.Charset = ANSI_CHARSET
          Font.Color = clGreen
          Font.Height = -15
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
          Transparent = True
          Layout = tlCenter
        end
        object lblTransSuccess: TLabel
          Left = 68
          Top = 74
          Width = 120
          Height = 16
          Caption = 'lblTransSuccess'
          Color = clActiveCaption
          Font.Charset = CHINESEBIG5_CHARSET
          Font.Color = clInfoText
          Font.Height = -16
          Font.Name = #27161#26999#39636
          Font.Style = []
          ParentColor = False
          ParentFont = False
          Transparent = True
        end
        object Label10: TLabel
          Left = 67
          Top = 41
          Width = 100
          Height = 19
          Alignment = taCenter
          Caption = #32080#36681#25104#21151#65306
          Font.Charset = CHINESEBIG5_CHARSET
          Font.Color = clGreen
          Font.Height = -19
          Font.Name = #27161#26999#39636
          Font.Style = []
          ParentFont = False
          Transparent = True
          Layout = tlCenter
        end
        object lblMsg: TLabel
          Left = 135
          Top = 354
          Width = 39
          Height = 18
          Caption = 'lblMsg'
          Font.Charset = ANSI_CHARSET
          Font.Color = clGreen
          Font.Height = -15
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
          Transparent = True
          Layout = tlCenter
        end
        object Label1: TLabel
          Left = 58
          Top = 353
          Width = 65
          Height = 18
          Caption = #34389#29702#35338#24687':'
          Font.Charset = ANSI_CHARSET
          Font.Color = clGreen
          Font.Height = -15
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
          Transparent = True
          Layout = tlCenter
        end
        object Label15: TLabel
          Left = 68
          Top = 148
          Width = 80
          Height = 19
          Alignment = taCenter
          Caption = #26410#32080#36681#65306
          Font.Charset = CHINESEBIG5_CHARSET
          Font.Color = clGreen
          Font.Height = -19
          Font.Name = #27161#26999#39636
          Font.Style = []
          ParentFont = False
          Transparent = True
          Layout = tlCenter
        end
        object lblTransError: TLabel
          Left = 68
          Top = 184
          Width = 104
          Height = 16
          Caption = 'lblTransError'
          Color = clActiveCaption
          Font.Charset = CHINESEBIG5_CHARSET
          Font.Color = clInfoText
          Font.Height = -16
          Font.Name = #27161#26999#39636
          Font.Style = []
          ParentColor = False
          ParentFont = False
          Transparent = True
        end
        object BitBtn1: TcxButton
          Left = 486
          Top = 308
          Width = 81
          Height = 26
          Caption = #32080#26463
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -13
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
          OnClick = btnExit1Click
          Glyph.Data = {
            76010000424D7601000000000000760000002800000020000000100000000100
            04000000000000010000130B0000130B00001000000000000000000000000000
            800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
            FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
            333333333333333333333333333333333333333FFF33FF333FFF339993370733
            999333777FF37FF377733339993000399933333777F777F77733333399970799
            93333333777F7377733333333999399933333333377737773333333333990993
            3333333333737F73333333333331013333333333333777FF3333333333910193
            333333333337773FF3333333399000993333333337377737FF33333399900099
            93333333773777377FF333399930003999333337773777F777FF339993370733
            9993337773337333777333333333333333333333333333333333333333333333
            3333333333333333333333333333333333333333333333333333}
          NumGlyphs = 2
        end
        object btnDuplicateReport: TcxButton
          Left = 150
          Top = 143
          Width = 30
          Height = 27
          Hint = #30332#31080#37325#35206#32080#36681#28165#21934
          ParentShowHint = False
          ShowHint = True
          TabOrder = 1
          OnClick = btnDuplicateReportClick
          Colors.Normal = clWhite
          Colors.Pressed = clWhite
          Glyph.Data = {
            36040000424D3604000000000000360000002800000010000000100000000100
            2000000000000004000000000000000000000000000000000000FFFFFF00FFFF
            FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
            FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
            FF00000000000000000000000000000000000000000000000000000000000000
            00000000000000000000000000000000000000000000FFFFFF00FFFFFF00FFFF
            FF0084848400C6C6C600C6C6C600C6C6C600C6C6C600C6C6C600C6C6C600C6C6
            C600C6C6C600C6C6C600C6C6C600C6C6C60000000000FFFFFF00FFFFFF00FFFF
            FF0084848400FFFFFF00C6C6C600FFFFFF00C6C6C600FFFFFF00C6C6C600FFFF
            FF00C6C6C600FFFFFF00C6C6C600C6C6C60000000000FFFFFF00FFFFFF00FFFF
            FF0084848400FFFFFF00FFFFFF00000084000000840000008400000084000000
            84000000840000008400FFFFFF00C6C6C60000000000FFFFFF00FFFFFF00FFFF
            FF0084848400FFFFFF00C6C6C600FFFFFF00C6C6C600FFFFFF00C6C6C600FFFF
            FF00C6C6C600FFFFFF00C6C6C600C6C6C60000000000FFFFFF00FFFFFF00FFFF
            FF0084848400FFFFFF00FFFFFF00000084000000840000008400000084000000
            84000000840000008400FFFFFF00C6C6C60000000000FFFFFF00FFFFFF00FFFF
            FF0084848400FFFFFF00C6C6C600FFFFFF00C6C6C600FFFFFF00C6C6C600FFFF
            FF00C6C6C600FFFFFF00C6C6C600C6C6C60000000000FFFFFF00FFFFFF00FFFF
            FF0084848400FFFFFF00FFFFFF00000084000000840000008400000084000000
            84000000840000008400FFFFFF00C6C6C60000000000FFFFFF00FFFFFF00FFFF
            FF0084848400FFFFFF00C6C6C600FFFFFF00C6C6C600FFFFFF00C6C6C600FFFF
            FF00C6C6C600FFFFFF00C6C6C600C6C6C60000000000FFFFFF00FFFFFF00FFFF
            FF0084848400FFFFFF00FFFFFF00000084000000840000008400000084000000
            84000000840000008400FFFFFF00C6C6C60000000000FFFFFF00FFFFFF00FFFF
            FF0084848400FFFFFF00C6C6C600FFFFFF00C6C6C600FFFFFF00C6C6C600FFFF
            FF00C6C6C600FFFFFF00C6C6C600C6C6C60000000000FFFFFF00FFFFFF00FFFF
            FF0084848400FFFFFF00FFFFFF00000084000000840000008400000084000000
            8400FFFFFF0000000000000000000000000000000000FFFFFF00FFFFFF00FFFF
            FF0084848400FFFFFF00C6C6C600FFFFFF00C6C6C600FFFFFF00C6C6C600FFFF
            FF00C6C6C600C6C6C600FFFFFF0084848400FFFFFF00FFFFFF00FFFFFF00FFFF
            FF0084848400FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
            FF00FFFFFF00C6C6C60084848400FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
            FF00848484008484840084848400848484008484840084848400848484008484
            84008484840084848400FFFFFF00FFFFFF00FFFFFF00FFFFFF00}
          LookAndFeel.Kind = lfUltraFlat
        end
        object ProgressBar2: TcxProgressBar
          Left = 131
          Top = 311
          ParentColor = False
          Style.Color = clWhite
          Style.StyleController = dtmMain.cxEditStyle
          TabOrder = 2
          Width = 337
        end
      end
      object CreateInv1: TcxTabSheet
        Caption = '  '#38283#31435#26597#35426#30059#38754'  '
        ImageIndex = 2
        OnShow = TransInv1Show
        DesignSize = (
          600
          465)
        object Label2: TLabel
          Left = 112
          Top = 108
          Width = 80
          Height = 16
          Caption = #25910#36027#26085#26399#65306
          Font.Charset = CHINESEBIG5_CHARSET
          Font.Color = clTeal
          Font.Height = -16
          Font.Name = #32048#26126#39636
          Font.Style = []
          ParentFont = False
          Transparent = True
        end
        object lblInvDate: TLabel
          Left = 112
          Top = 165
          Width = 80
          Height = 16
          Caption = #30332#31080#26085#26399#65306
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clTeal
          Font.Height = -16
          Font.Name = #32048#26126#39636
          Font.Style = []
          ParentFont = False
          Transparent = True
        end
        object Label13: TLabel
          Left = 8
          Top = 247
          Width = 80
          Height = 16
          Caption = #21487#29992#23383#36556#65306
          Color = clRed
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clRed
          Font.Height = -16
          Font.Name = #32048#26126#39636
          Font.Style = []
          ParentColor = False
          ParentFont = False
          Transparent = True
        end
        object Label14: TLabel
          Left = 312
          Top = 247
          Width = 80
          Height = 16
          Caption = #36984#29992#23383#36556#65306
          Color = clRed
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clRed
          Font.Height = -16
          Font.Name = #32048#26126#39636
          Font.Style = []
          ParentColor = False
          ParentFont = False
          Transparent = True
        end
        object Label4: TLabel
          Left = 232
          Top = 16
          Width = 112
          Height = 27
          Caption = #30332#31080#38283#31435
          Font.Charset = CHINESEBIG5_CHARSET
          Font.Color = clGreen
          Font.Height = -27
          Font.Name = #27161#26999#39636
          Font.Style = []
          ParentFont = False
          Transparent = True
        end
        object Label24: TLabel
          Left = 490
          Top = 144
          Width = 64
          Height = 16
          Caption = #29151#26989#21029#65306
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clTeal
          Font.Height = -16
          Font.Name = #32048#26126#39636
          Font.Style = []
          ParentFont = False
          Transparent = True
          Visible = False
        end
        object lbl1: TLabel
          Left = 79
          Top = 56
          Width = 112
          Height = 16
          Caption = #30332#31080#38283#31435#31278#39006#65306
          Font.Charset = CHINESEBIG5_CHARSET
          Font.Color = clTeal
          Font.Height = -16
          Font.Name = #32048#26126#39636
          Font.Style = []
          ParentFont = False
          Transparent = True
        end
        object tmeStartDate: TMaskEdit
          Left = 208
          Top = 100
          Width = 104
          Height = 27
          Color = clCream
          EditMask = '!9999/99/99;1;_'
          Font.Charset = CHINESEBIG5_CHARSET
          Font.Color = clWindowText
          Font.Height = -19
          Font.Name = #32048#26126#39636
          Font.Style = []
          MaxLength = 10
          ParentFont = False
          TabOrder = 2
          Text = '    /  /  '
          OnExit = tmeStartDateExit
        end
        object tmeEndInvDate: TMaskEdit
          Left = 344
          Top = 100
          Width = 111
          Height = 27
          Color = clCream
          EditMask = '!9999/99/99;1;_'
          Font.Charset = CHINESEBIG5_CHARSET
          Font.Color = clWindowText
          Font.Height = -19
          Font.Name = #32048#26126#39636
          Font.Style = []
          MaxLength = 10
          ParentFont = False
          TabOrder = 3
          Text = '    /  /  '
        end
        object tmeInvDate: TMaskEdit
          Left = 208
          Top = 157
          Width = 107
          Height = 27
          Color = clCream
          EditMask = '!9999/99/99;1;_'
          Font.Charset = CHINESEBIG5_CHARSET
          Font.Color = clWindowText
          Font.Height = -19
          Font.Name = #32048#26126#39636
          Font.Style = []
          MaxLength = 10
          ParentFont = False
          TabOrder = 4
          Text = '    /  /  '
          OnExit = tmeInvDateExit
        end
        object DBGrid1: TDBGrid
          Left = 8
          Top = 271
          Width = 297
          Height = 127
          TabStop = False
          Color = clCream
          Ctl3D = False
          DataSource = dtmMainJ.DataSource1
          FixedColor = clWhite
          Font.Charset = CHINESEBIG5_CHARSET
          Font.Color = clInfoText
          Font.Height = -13
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          Options = [dgTitles, dgColumnResize, dgColLines, dgRowLines, dgConfirmDelete, dgCancelOnExit]
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 8
          TitleFont.Charset = CHINESEBIG5_CHARSET
          TitleFont.Color = clMenuText
          TitleFont.Height = -13
          TitleFont.Name = #26032#32048#26126#39636
          TitleFont.Style = []
          OnDblClick = DBGrid1DblClick
          Columns = <
            item
              Alignment = taCenter
              Expanded = False
              FieldName = 'PREFIX'
              Font.Charset = ANSI_CHARSET
              Font.Color = clWindowText
              Font.Height = -16
              Font.Name = #32048#26126#39636
              Font.Style = []
              Title.Alignment = taCenter
              Title.Caption = #23383#36556
              Title.Font.Charset = DEFAULT_CHARSET
              Title.Font.Color = clWindowText
              Title.Font.Height = -13
              Title.Font.Name = #32048#26126#39636
              Title.Font.Style = []
              Width = 30
              Visible = True
            end
            item
              Alignment = taCenter
              Expanded = False
              FieldName = 'LASTINVDATE'
              Font.Charset = ANSI_CHARSET
              Font.Color = clWindowText
              Font.Height = -16
              Font.Name = #32048#26126#39636
              Font.Style = []
              Title.Alignment = taCenter
              Title.Caption = #26368#24460#30332#31080#38283#31435#26085
              Title.Font.Charset = DEFAULT_CHARSET
              Title.Font.Color = clWindowText
              Title.Font.Height = -13
              Title.Font.Name = #32048#26126#39636
              Title.Font.Style = []
              Width = 100
              Visible = True
            end
            item
              Expanded = False
              FieldName = 'STARTNUM'
              Title.Caption = #36215#22987#34399#30908
              Visible = False
            end
            item
              Expanded = False
              FieldName = 'CURNUM'
              Title.Caption = #30446#21069#21487#29992#34399#30908
              Width = 80
              Visible = True
            end
            item
              Expanded = False
              FieldName = 'Memo'
              Title.Caption = #20633#35387
              Width = 60
              Visible = True
            end
            item
              Expanded = False
              FieldName = 'UploadFlag'
              Title.Caption = #19978#20659#35387#35352
              Visible = True
            end>
        end
        object RadioButton1: TRadioButton
          Left = 208
          Top = 134
          Width = 81
          Height = 17
          Caption = #38928#38283
          Color = clWhite
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clTeal
          Font.Height = -16
          Font.Name = #32048#26126#39636
          Font.Style = []
          ParentColor = False
          ParentFont = False
          TabOrder = 5
          OnClick = RadioButton1Click
        end
        object RadioButton2: TRadioButton
          Left = 304
          Top = 134
          Width = 65
          Height = 17
          Caption = #24460#38283
          Checked = True
          Color = clWhite
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clTeal
          Font.Height = -16
          Font.Name = #32048#26126#39636
          Font.Style = []
          ParentColor = False
          ParentFont = False
          TabOrder = 6
          TabStop = True
          OnClick = RadioButton2Click
        end
        object chbInvAndCharge: TCheckBox
          Left = 112
          Top = 197
          Width = 225
          Height = 25
          Caption = #30332#31080#26085#26399#21516#25910#36027#26085#26399
          Color = clWhite
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clTeal
          Font.Height = -16
          Font.Name = #32048#26126#39636
          Font.Style = []
          ParentColor = False
          ParentFont = False
          TabOrder = 7
          OnClick = chbInvAndChargeClick
        end
        object DBGrid2: TDBGrid
          Left = 312
          Top = 271
          Width = 265
          Height = 127
          Color = clCream
          Ctl3D = False
          DataSource = dsrPrefix
          FixedColor = clMoneyGreen
          Font.Charset = ANSI_CHARSET
          Font.Color = clWindowText
          Font.Height = -13
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          ParentCtl3D = False
          ParentFont = False
          PopupMenu = ppmPrefix
          TabOrder = 9
          TitleFont.Charset = DEFAULT_CHARSET
          TitleFont.Color = clWindowText
          TitleFont.Height = -11
          TitleFont.Name = 'MS Sans Serif'
          TitleFont.Style = []
        end
        object btnQuery: TcxButton
          Left = 129
          Top = 421
          Width = 85
          Height = 29
          Anchors = [akTop, akRight]
          Caption = #26597#35426
          Default = True
          TabOrder = 10
          OnClick = BtnQueryClick
          Glyph.Data = {
            36040000424D3604000000000000360000002800000010000000100000000100
            2000000000000004000000000000000000000000000000000000000000000000
            0000000000000000000000000000000000000000000000000000000000000000
            0000000000000000000000000000000000000000000000000000000000000000
            0000000000000000000000000000000000000000000000000000000000000000
            0000000000000000FF018B8A7AC385827DA26B6B6B3F00000000000000000000
            0000000000000000000000000000000000000000000000000000000000000000
            00000000FF01094FFF974392F6FFEEE9DFFF86827FA700000000000000000000
            0000000000000000000000000000000000000000000000000000000000000000
            FF020D51FF9B439AFFFF6ADAFFFF5DADF5FF998F7ED900000000000000000000
            00000000000000000000000000000000000000000000000000006B7E690F0041
            FC9F469FFFFF6FDAFFFF50ACFFFF1357FFB9002AFF0A00000000000000000000
            000000000000000000006164682E515459524C4E523F1617180698928B9E7091
            B2FF61D3FFFF4EAAFFFF1657FFB4001FFF070000000000000000000000000000
            000076777A257B7E80C8B3A081FED2B588FDC3AA83FD83817AE566686CEAFFF7
            F0FF6B93BDFF084AFEAF0028FF06000000000000000000000000000000007B7D
            7F188D8A82E6F5CB84FEF5CB84FFF1C885FFF5CE8EFFFCD08CFFAE9E85FE7071
            75F3A9A193BB3956B20A0000000000000000000000000000000000000000767A
            7F95ECCB8EFFF3D192FFEECE92FFEDCC8EFFECC784FFEDC687FFFDD28FFF8783
            7CDB000000010000000000000000000000000000000000000000000000009894
            8BE1FEDFA1FFF2DAA5FFF2DBA7FFF1D79FFFEFD097FFECC886FFF6D298FFC1A9
            87FF5254592A000000000000000000000000000000000000000000000000A7A1
            91F4FFEBB9FFF8ECC6FFF7EBC0FFF5E3B2FFF2DAA3FFEECF94FFF4D093FFCFB6
            8DFF5153583D0000000000000000000000000000000000000000000000009996
            92D0FFF7CBFFFDFAE6FFFDF9E7FFF8EDC5FFF4E0B0FFF0D49BFFFAD594FFAF9F
            85FE6264671A000000000000000000000000000000000000000000000000A0A1
            A56FD5CCB1FFFFFFF2FFFFFFF2FFFBF3D2FFF6E2B4FFF7D99EFFF6D393FF7D7D
            81B800000000000000000000000000000000000000000000000000000000D9D9
            DA04A1A1A2B8D6CFB6FFFFFFDCFFFFF6CAFFFFEBB1FFEDD49CFE8C8983E47677
            7A20000000000000000000000000000000000000000000000000000000000000
            0000EBECED059C9CA07B979893E0A8A397FE97948CF07679809E7C7D7F1A0000
            0000000000000000000000000000000000000000000000000000000000000000
            0000000000000000000000000000000000000000000000000000000000000000
            0000000000000000000000000000000000000000000000000000}
        end
        object btnReset: TcxButton
          Left = 257
          Top = 421
          Width = 85
          Height = 29
          Anchors = [akTop, akRight]
          Caption = #37325#35373
          Default = True
          TabOrder = 11
          OnClick = BtnResetClick
          Glyph.Data = {
            36040000424D3604000000000000360000002800000010000000100000000100
            2000000000000004000000000000000000000000000000000000000000000000
            0000000000000000000000000000000000000000000000000000000000000000
            0000000000000000000000000000000000000000000000000000000000000399
            0697079B0AF50FA516FC0EA416FB10A518FB09A00CF700000000000000008A0F
            001F71040065781300AB801D0FD77000002B0000000000000000000000000000
            0000029304AB48DF6CFF67FF9AFF76FFB1FF11A818FD00000000A53A0BCCCB7D
            50FEE8A676FFFFC189FFD28359FE700500E8760500280000000000000000007F
            003C1BAE2CF950EB78FF56F083FF6CFFA2FF10A816FD00000000B55A24EBFFCD
            9AFFF2B27FFFEAA36FFFD88C60FF89200CFF710B008A0000000000000000027D
            06C33EDB5CFF46E068FF41D862FF15A71EFB10A615FD00000000B75716E4F1AE
            7DFFE69D6AFFE1905BFFDE9366FF9B3014FF791200A800000000000000000A8B
            11F235D34DFF2CC242FF008A009900940024029202DD00000000BF5D11DCE495
            66FFE19668FFE39D72FFEDB089FFB95A38FF831C01C600000000000000000276
            06CB2BCC40FF008000AB000000000000000000AA001800000000DA6800ABEFAC
            6EFFFCC498FFC3A363FF94883FFFD27B4AFF8B2607E300000000000000000052
            506C07800DFE03398CE00C1EBEE40103A0D300008E9C0000932C00000000D65E
            005BC26113D9775D08D108810BFE4448008B871D002700000000000000000436
            FADC70C2A5FF7BC7CFFF65A2FFFF2E6DFFFF0133FFFF00009CDA00AA00180000
            000000000000008000AB2BCC40FF027606CB000000000000000000000000001A
            F27B82BCFAFF94D2FFFF5993FFFF1D54FFFF0022E7FF00009084029302DD0094
            0024008A00992CC242FF35D34DFF0B8B12F20000000000000000000000000025
            FF0D1943ECEBA9E9FFFF5792FFFF0C46FFFF0003ACED00009A1010A615FD15A8
            1EFB41D862FF46E068FF3EDB5CFF027D06C30000000000000000000000000000
            00000010EF787AB1F4FF5995FFFF0025E5FF00009F780000000010A816FD6CFF
            A2FF56F083FF50EB78FF1CAE2CF9007F003C0000000000000000000000000000
            00000013FF0C1C3DD8E95A9EFFFF0000AEE60000B80A0000000011A818FD76FF
            B1FF67FF9AFF48DF6CFF039305AB000000000000000000000000000000000000
            000000000000000FD7741C3CD4FF0000AC6C000000000000000009A00CF710A6
            18FB0EA516FB0FA616FC079B0AF5049906970000000000000000000000000000
            0000000000000000ED0A000CB0BE0000C1060000000000000000000000000000
            0000000000000000000000000000000000000000000000000000000000000000
            0000000000000000000000000000000000000000000000000000000000000000
            0000000000000000000000000000000000000000000000000000}
        end
        object btnExit: TcxButton
          Left = 384
          Top = 421
          Width = 85
          Height = 29
          Anchors = [akTop, akRight]
          Caption = #32080#26463
          ModalResult = 2
          TabOrder = 12
          OnClick = btnExitClick
          Glyph.Data = {
            36030000424D3603000000000000360000002800000010000000100000000100
            18000000000000030000120B0000120B00000000000000000000FF00FFFF00FF
            FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
            FFFF00FFFF00FFFF00FFFF00FF0000A40206B0030AAE0000A6000098FF00FFFF
            00FFFF00FFFF00FF0000A700009A0000A7FF00FFFF00FFFF00FFFF00FF0000A9
            1844F6194DF81031D20102AB0000B6FF00FFFF00FF0000B10928D7092ED70313
            B30000ACFF00FFFF00FFFF00FF0103B32451F91F52FF1D4FFF1744E8040BB000
            00B00000AC0D2EDD1142F90D3DF50B3BF0041ABC0000A5FF00FFFF00FF0000AE
            1832DB285BFF2456FF2253FF1B4BF1050DB10F30DD164AFE1344F91041F60E3E
            F60A3CF000009FFF00FFFF00FF0000BE1F37DD3A6FFF2C5EFF295AFF2657FF20
            52FC1C4FFF194AFD1646FA1445FA0F3DF2020AB10000A8FF00FFFF00FFFF00FF
            0000C8121DC83D6AFB3567FF2C5DFF2859FF2253FF1D4EFF1A4DFF123DED0002
            AC0000BAFF00FFFF00FFFF00FFFF00FFFF00FF0000CC0000B62E4EE73668FF2E
            5EFF2859FF2254FF163DEA0000A80000ABFF00FFFF00FFFF00FFFF00FFFF00FF
            FF00FFFF00FF0000BF253FDF3B6DFF3464FF2E5EFF2759FF1B46EA0001AC0000
            A9FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF0203C84B7CFF4170FF3B
            6BFF396CFF2D5EFF2558FF1336D70000B6FF00FFFF00FFFF00FFFF00FFFF00FF
            FF00FF0000D9263CDB5080FF4575FF3662FA0D14C33C6DFF2A5BFF2053FD0B1D
            C40000C0FF00FFFF00FFFF00FFFF00FFFF00FF0000CB527CFA5081FF4B7DFF0B
            13C90000BB0E15C7386AFF2456FF1A4AF20207B30000B5FF00FFFF00FFFF00FF
            FF00FF131CDD6A9CFF5788FF2B46E70000CDFF00FF0000CD0F1BCB3065FF1F51
            FF1439DD0000B1FF00FFFF00FFFF00FFFF00FF0000DE3A52E45782FB0101D0FF
            00FFFF00FFFF00FF0000CC1426D6265AFF0F2EE30103B8FF00FFFF00FFFF00FF
            FF00FFFF00FF0000CF0000C00000CEFF00FFFF00FFFF00FFFF00FF0000C40001
            B80000B5000077FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
            00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF}
        end
        object cobBusinessId: TComboBox
          Left = 432
          Top = 186
          Width = 167
          Height = 22
          ItemHeight = 14
          TabOrder = 13
          Visible = False
          OnExit = cobBusinessIdExit
          Items.Strings = (
            '0.'#29151#26989#20154
            '1.'#38750#29151#26989#20154)
        end
        object chkNormalInv: TCheckBox
          Left = 201
          Top = 54
          Width = 154
          Height = 21
          Caption = '0.'#38651#23376#35336#31639#27231#30332#31080
          Checked = True
          Font.Charset = ANSI_CHARSET
          Font.Color = clTeal
          Font.Height = -16
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          ParentFont = False
          State = cbChecked
          TabOrder = 0
          OnClick = chkNormalInvClick
        end
        object chkElectronInv: TCheckBox
          Left = 367
          Top = 54
          Width = 98
          Height = 21
          Caption = '1.'#38651#23376#30332#31080
          Checked = True
          Font.Charset = ANSI_CHARSET
          Font.Color = clTeal
          Font.Height = -16
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          ParentFont = False
          State = cbChecked
          TabOrder = 1
          OnClick = chkElectronInvClick
        end
      end
      object CreateInv2: TcxTabSheet
        Caption = '  '#26597#35426#32080#26524'  '
        ImageIndex = 3
        OnShow = TransInv1Show
        DesignSize = (
          600
          465)
        object Label7: TLabel
          Left = 232
          Top = 14
          Width = 144
          Height = 24
          Caption = #30332#31080#38283#31435#26126#32048
          Font.Charset = CHINESEBIG5_CHARSET
          Font.Color = clGreen
          Font.Height = -24
          Font.Name = #27161#26999#39636
          Font.Style = []
          ParentFont = False
          Transparent = True
        end
        object lblUser2: TLabel
          Left = 24
          Top = 40
          Width = 64
          Height = 16
          Caption = #25805#20316#32773#65306
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlue
          Font.Height = -16
          Font.Name = #32048#26126#39636
          Font.Style = []
          ParentFont = False
          Transparent = True
        end
        object Label9: TLabel
          Left = 8
          Top = 74
          Width = 80
          Height = 16
          Caption = #26597#35426#26781#20214#65306
          Color = clActiveCaption
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlue
          Font.Height = -16
          Font.Name = #32048#26126#39636
          Font.Style = []
          ParentColor = False
          ParentFont = False
          Transparent = True
        end
        object lblQuery: TLabel
          Left = 88
          Top = 74
          Width = 8
          Height = 16
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clInfoText
          Font.Height = -16
          Font.Name = #32048#26126#39636
          Font.Style = []
          ParentFont = False
          Transparent = True
        end
        object lblInvCount: TLabel
          Left = 87
          Top = 196
          Width = 72
          Height = 16
          Alignment = taCenter
          Caption = #30332#31080#24373#25976':'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clTeal
          Font.Height = -16
          Font.Name = #32048#26126#39636
          Font.Style = []
          ParentFont = False
          Transparent = True
        end
        object lblSaleAmount: TLabel
          Left = 87
          Top = 219
          Width = 72
          Height = 16
          Caption = #32317#37559#21806#38989':'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clTeal
          Font.Height = -16
          Font.Name = #32048#26126#39636
          Font.Style = []
          ParentFont = False
          Transparent = True
        end
        object lblTaxAmount: TLabel
          Left = 87
          Top = 242
          Width = 72
          Height = 16
          Caption = #32317' '#31237' '#38989':'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clTeal
          Font.Height = -16
          Font.Name = #32048#26126#39636
          Font.Style = []
          ParentFont = False
          Transparent = True
        end
        object lblInvAmount: TLabel
          Left = 87
          Top = 265
          Width = 72
          Height = 16
          Caption = #32317' '#37329' '#38989':'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clTeal
          Font.Height = -16
          Font.Name = #32048#26126#39636
          Font.Style = []
          ParentFont = False
          Transparent = True
        end
        object lblHint: TLabel
          Left = 195
          Top = 327
          Width = 200
          Height = 22
          Alignment = taCenter
          AutoSize = False
          Caption = #30332#31080#26597#35426#23436#25104
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clRed
          Font.Height = -16
          Font.Name = #32048#26126#39636
          Font.Style = []
          ParentFont = False
          Transparent = True
        end
        object lblValidCounts: TLabel
          Left = 87
          Top = 289
          Width = 72
          Height = 16
          Caption = #21487#29992#30332#31080':'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clTeal
          Font.Height = -16
          Font.Name = #32048#26126#39636
          Font.Style = []
          ParentFont = False
          Transparent = True
        end
        object lblInvCount2: TLabel
          Left = 339
          Top = 196
          Width = 72
          Height = 16
          Alignment = taCenter
          Caption = #30332#31080#24373#25976':'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clTeal
          Font.Height = -16
          Font.Name = #32048#26126#39636
          Font.Style = []
          ParentFont = False
          Transparent = True
        end
        object lblSaleAmount2: TLabel
          Left = 339
          Top = 219
          Width = 72
          Height = 16
          Caption = #32317#37559#21806#38989':'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clTeal
          Font.Height = -16
          Font.Name = #32048#26126#39636
          Font.Style = []
          ParentFont = False
          Transparent = True
        end
        object lblTaxAmount2: TLabel
          Left = 339
          Top = 242
          Width = 72
          Height = 16
          Caption = #32317' '#31237' '#38989':'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clTeal
          Font.Height = -16
          Font.Name = #32048#26126#39636
          Font.Style = []
          ParentFont = False
          Transparent = True
        end
        object lblInvAmount2: TLabel
          Left = 339
          Top = 265
          Width = 72
          Height = 16
          Caption = #32317' '#37329' '#38989':'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clTeal
          Font.Height = -16
          Font.Name = #32048#26126#39636
          Font.Style = []
          ParentFont = False
          Transparent = True
        end
        object Bevel1: TBevel
          Left = 23
          Top = 317
          Width = 550
          Height = 7
          Shape = bsTopLine
        end
        object Label18: TLabel
          Left = 101
          Top = 154
          Width = 128
          Height = 16
          Caption = #38928#35336#30332#31080#38283#31435#32113#35336
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlue
          Font.Height = -13
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          ParentFont = False
          Transparent = True
        end
        object Label19: TLabel
          Left = 354
          Top = 154
          Width = 112
          Height = 16
          Caption = #19981#38283#31435#30332#31080#32113#35336
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlue
          Font.Height = -13
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          ParentFont = False
          Transparent = True
        end
        object Bevel2: TBevel
          Left = 22
          Top = 186
          Width = 550
          Height = 9
          Shape = bsTopLine
        end
        object btnListCreateInv: TcxButton
          Left = 232
          Top = 150
          Width = 30
          Height = 24
          Hint = #26126#32048#36039#26009
          ParentShowHint = False
          ShowHint = True
          TabOrder = 0
          OnClick = btnListCreateInvClick
          Colors.Normal = clWhite
          Colors.Pressed = clWhite
          Glyph.Data = {
            36040000424D3604000000000000360000002800000010000000100000000100
            2000000000000004000000000000000000000000000000000000FFFFFF00FFFF
            FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
            FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF008484
            8400000000000000000000000000000000000000000000000000000000000000
            0000000000000000000000000000000000000000000000000000FFFFFF008484
            8400FFFFFF00C6C6C600FFFFFF00C6C6C600FFFFFF00C6C6C600FFFFFF00C6C6
            C600FFFFFF00C6C6C600FFFFFF00C6C6C600FFFFFF0000000000FFFFFF008484
            8400FFFFFF00FFFFFF00C6C6C600FFFFFF00C6C6C600FFFFFF00C6C6C600FFFF
            FF00C6C6C600FFFFFF00C6C6C600FFFFFF00C6C6C60000000000FFFFFF008484
            8400FFFFFF00C6C6C600FFFFFF00C6C6C600FFFFFF00C6C6C600FFFFFF00C6C6
            C600FFFFFF00C6C6C600FFFFFF00C6C6C600FFFFFF0000000000FFFFFF008484
            8400FFFFFF00FFFFFF00C6C6C600FFFFFF00C6C6C600FFFFFF00C6C6C600FFFF
            FF00C6C6C600FFFFFF00C6C6C600FFFFFF00C6C6C60000000000FFFFFF008484
            8400FFFFFF00C6C6C600FFFFFF00C6C6C600FFFFFF00C6C6C600FFFFFF00C6C6
            C600FFFFFF00C6C6C600FFFFFF00C6C6C600FFFFFF0000000000FFFFFF008484
            8400FFFFFF00FFFFFF00C6C6C600FFFFFF00C6C6C600FFFFFF00C6C6C600FFFF
            FF00C6C6C600FFFFFF00C6C6C600FFFFFF00C6C6C60000000000FFFFFF008484
            8400FFFFFF00C6C6C600FFFFFF00C6C6C600FFFFFF00C6C6C600FFFFFF00C6C6
            C600FFFFFF00C6C6C600FFFFFF00C6C6C600FFFFFF0000000000FFFFFF008484
            8400FFFFFF00FFFFFF00C6C6C600FFFFFF00C6C6C600FFFFFF00C6C6C600FFFF
            FF00C6C6C600FFFFFF00C6C6C600FFFFFF00C6C6C60000000000FFFFFF008484
            8400FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
            FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF0000000000FFFFFF008484
            8400840000008400000084000000840000008400000084000000840000008400
            0000840000008400000084000000840000008400000000000000FFFFFF008484
            8400FFFFFF008400000084000000840000008400000084000000840000008400
            00008400000084000000FFFFFF0084000000FFFFFF0000000000FFFFFF008484
            8400848484008484840084848400848484008484840084848400848484008484
            8400848484008484840084848400848484008484840084848400FFFFFF00FFFF
            FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
            FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
            FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
            FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00}
          LookAndFeel.Kind = lfUltraFlat
        end
        object btnRptCreateInv: TcxButton
          Left = 265
          Top = 150
          Width = 30
          Height = 24
          Hint = #32113#35336#34920
          ParentShowHint = False
          ShowHint = True
          TabOrder = 1
          OnClick = btnRptCreateInvClick
          Colors.Normal = clWhite
          Colors.Pressed = clWhite
          Glyph.Data = {
            36040000424D3604000000000000360000002800000010000000100000000100
            2000000000000004000000000000000000000000000000000000FFFFFF00FFFF
            FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
            FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
            FF00000000000000000000000000000000000000000000000000000000000000
            00000000000000000000000000000000000000000000FFFFFF00FFFFFF00FFFF
            FF0084848400C6C6C600C6C6C600C6C6C600C6C6C600C6C6C600C6C6C600C6C6
            C600C6C6C600C6C6C600C6C6C600C6C6C60000000000FFFFFF00FFFFFF00FFFF
            FF0084848400FFFFFF00C6C6C600FFFFFF00C6C6C600FFFFFF00C6C6C600FFFF
            FF00C6C6C600FFFFFF00C6C6C600C6C6C60000000000FFFFFF00FFFFFF00FFFF
            FF0084848400FFFFFF00FFFFFF00000084000000840000008400000084000000
            84000000840000008400FFFFFF00C6C6C60000000000FFFFFF00FFFFFF00FFFF
            FF0084848400FFFFFF00C6C6C600FFFFFF00C6C6C600FFFFFF00C6C6C600FFFF
            FF00C6C6C600FFFFFF00C6C6C600C6C6C60000000000FFFFFF00FFFFFF00FFFF
            FF0084848400FFFFFF00FFFFFF00000084000000840000008400000084000000
            84000000840000008400FFFFFF00C6C6C60000000000FFFFFF00FFFFFF00FFFF
            FF0084848400FFFFFF00C6C6C600FFFFFF00C6C6C600FFFFFF00C6C6C600FFFF
            FF00C6C6C600FFFFFF00C6C6C600C6C6C60000000000FFFFFF00FFFFFF00FFFF
            FF0084848400FFFFFF00FFFFFF00000084000000840000008400000084000000
            84000000840000008400FFFFFF00C6C6C60000000000FFFFFF00FFFFFF00FFFF
            FF0084848400FFFFFF00C6C6C600FFFFFF00C6C6C600FFFFFF00C6C6C600FFFF
            FF00C6C6C600FFFFFF00C6C6C600C6C6C60000000000FFFFFF00FFFFFF00FFFF
            FF0084848400FFFFFF00FFFFFF00000084000000840000008400000084000000
            84000000840000008400FFFFFF00C6C6C60000000000FFFFFF00FFFFFF00FFFF
            FF0084848400FFFFFF00C6C6C600FFFFFF00C6C6C600FFFFFF00C6C6C600FFFF
            FF00C6C6C600FFFFFF00C6C6C600C6C6C60000000000FFFFFF00FFFFFF00FFFF
            FF0084848400FFFFFF00FFFFFF00000084000000840000008400000084000000
            8400FFFFFF0000000000000000000000000000000000FFFFFF00FFFFFF00FFFF
            FF0084848400FFFFFF00C6C6C600FFFFFF00C6C6C600FFFFFF00C6C6C600FFFF
            FF00C6C6C600C6C6C600FFFFFF0084848400FFFFFF00FFFFFF00FFFFFF00FFFF
            FF0084848400FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
            FF00FFFFFF00C6C6C60084848400FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
            FF00848484008484840084848400848484008484840084848400848484008484
            84008484840084848400FFFFFF00FFFFFF00FFFFFF00FFFFFF00}
          LookAndFeel.Kind = lfUltraFlat
        end
        object btnListNoInv: TcxButton
          Left = 472
          Top = 150
          Width = 30
          Height = 24
          Hint = #26126#32048#36039#26009
          ParentShowHint = False
          ShowHint = True
          TabOrder = 2
          OnClick = btnListNoInvClick
          Colors.Normal = clWhite
          Glyph.Data = {
            36040000424D3604000000000000360000002800000010000000100000000100
            2000000000000004000000000000000000000000000000000000FFFFFF00FFFF
            FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
            FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF008484
            8400000000000000000000000000000000000000000000000000000000000000
            0000000000000000000000000000000000000000000000000000FFFFFF008484
            8400FFFFFF00C6C6C600FFFFFF00C6C6C600FFFFFF00C6C6C600FFFFFF00C6C6
            C600FFFFFF00C6C6C600FFFFFF00C6C6C600FFFFFF0000000000FFFFFF008484
            8400FFFFFF00FFFFFF00C6C6C600FFFFFF00C6C6C600FFFFFF00C6C6C600FFFF
            FF00C6C6C600FFFFFF00C6C6C600FFFFFF00C6C6C60000000000FFFFFF008484
            8400FFFFFF00C6C6C600FFFFFF00C6C6C600FFFFFF00C6C6C600FFFFFF00C6C6
            C600FFFFFF00C6C6C600FFFFFF00C6C6C600FFFFFF0000000000FFFFFF008484
            8400FFFFFF00FFFFFF00C6C6C600FFFFFF00C6C6C600FFFFFF00C6C6C600FFFF
            FF00C6C6C600FFFFFF00C6C6C600FFFFFF00C6C6C60000000000FFFFFF008484
            8400FFFFFF00C6C6C600FFFFFF00C6C6C600FFFFFF00C6C6C600FFFFFF00C6C6
            C600FFFFFF00C6C6C600FFFFFF00C6C6C600FFFFFF0000000000FFFFFF008484
            8400FFFFFF00FFFFFF00C6C6C600FFFFFF00C6C6C600FFFFFF00C6C6C600FFFF
            FF00C6C6C600FFFFFF00C6C6C600FFFFFF00C6C6C60000000000FFFFFF008484
            8400FFFFFF00C6C6C600FFFFFF00C6C6C600FFFFFF00C6C6C600FFFFFF00C6C6
            C600FFFFFF00C6C6C600FFFFFF00C6C6C600FFFFFF0000000000FFFFFF008484
            8400FFFFFF00FFFFFF00C6C6C600FFFFFF00C6C6C600FFFFFF00C6C6C600FFFF
            FF00C6C6C600FFFFFF00C6C6C600FFFFFF00C6C6C60000000000FFFFFF008484
            8400FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
            FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF0000000000FFFFFF008484
            8400840000008400000084000000840000008400000084000000840000008400
            0000840000008400000084000000840000008400000000000000FFFFFF008484
            8400FFFFFF008400000084000000840000008400000084000000840000008400
            00008400000084000000FFFFFF0084000000FFFFFF0000000000FFFFFF008484
            8400848484008484840084848400848484008484840084848400848484008484
            8400848484008484840084848400848484008484840084848400FFFFFF00FFFF
            FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
            FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
            FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
            FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00}
          LookAndFeel.Kind = lfUltraFlat
        end
        object cxGroupBox1: TcxGroupBox
          Left = 20
          Top = 349
          Caption = ' '#25490#24207#26041#24335' '
          ParentFont = False
          Style.Font.Charset = ANSI_CHARSET
          Style.Font.Color = clWindowText
          Style.Font.Height = -13
          Style.Font.Name = 'Tahoma'
          Style.Font.Style = []
          Style.IsFontAssigned = True
          TabOrder = 3
          Height = 56
          Width = 550
          object rdoOrder0: TcxRadioButton
            Left = 21
            Top = 28
            Width = 81
            Height = 17
            Caption = #25910#36027#26085#26399
            Font.Charset = ANSI_CHARSET
            Font.Color = clWindowText
            Font.Height = -12
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
            TabOrder = 0
          end
          object rdoOrder1: TcxRadioButton
            Left = 119
            Top = 28
            Width = 81
            Height = 17
            Caption = #37109#36958#21312#34399
            Font.Charset = ANSI_CHARSET
            Font.Color = clWindowText
            Font.Height = -12
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
            TabOrder = 1
          end
          object rdoOrder2: TcxRadioButton
            Left = 219
            Top = 28
            Width = 145
            Height = 17
            Caption = #25910#36027#26085#26399' + '#37109#36958#21312#34399
            Checked = True
            Font.Charset = ANSI_CHARSET
            Font.Color = clWindowText
            Font.Height = -12
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
            TabOrder = 2
            TabStop = True
          end
          object rdoOrder3: TcxRadioButton
            Left = 368
            Top = 28
            Width = 174
            Height = 17
            Caption = #25910#36027#26085#26399' + '#37109#36958#21312#34399' + '#23458#32232
            Font.Charset = ANSI_CHARSET
            Font.Color = clWindowText
            Font.Height = -12
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
            TabOrder = 3
          end
        end
        object btnAssign: TcxButton
          Left = 197
          Top = 421
          Width = 85
          Height = 29
          Anchors = [akTop, akRight]
          Caption = #38283#31435
          Default = True
          TabOrder = 4
          OnClick = btnAssignClick
          Glyph.Data = {
            36060000424D3606000000000000360000002800000020000000100000000100
            18000000000000060000F00A0000F00A00000000000000000000FF00FFFF00FF
            FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
            FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
            00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
            FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
            FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
            00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
            FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
            FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
            00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
            FF00FFFF00FF008B00008200038805008A00FF00FFFF00FFFF00FFFF00FFFF00
            FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF97796797796797796797
            7967FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
            FF00FF00A300008E0013AC2715B32B06940E009300FF00FFFF00FFFF00FFFF00
            FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF977967977967FF00FFFF00FF97
            7967977967FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
            00B10000940016AE2E17B23114AD2A13B12906920B009500FF00FFFF00FFFF00
            FFFF00FFFF00FFFF00FFFF00FFFF00FF977967977967FF00FFFF00FFFF00FFFF
            00FF977967977967FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF00A900
            009A0019B2331CB63A18B33434BC4D17B03013B02A06910B009500FF00FFFF00
            FFFF00FFFF00FFFF00FFFF00FF977967977967FF00FFFF00FFFF00FFFF00FFFF
            00FFFF00FF977967977967FF00FFFF00FFFF00FFFF00FFFF00FFFF00FF0CA316
            1FB73E20BA421FB84011A824018E0162C77119B23213B12B06910B009500FF00
            FFFF00FFFF00FFFF00FFFF00FF977967FF00FFFF00FFFF00FFB89888977967FF
            00FFFF00FFFF00FF977967977967FF00FFFF00FFFF00FFFF00FFFF00FF28B338
            42C8651FBB4615B32C00A40000920000930061C87119B13113B22C06900B0095
            00FF00FFFF00FFFF00FFFF00FF977967FF00FFFF00FFB8988897796797796797
            7967FF00FFFF00FFFF00FF977967977967FF00FFFF00FFFF00FFFF00FF0EB924
            43C45636BC4E00AF08FF00FFFF00FF00AA0000920063CA7217B13114B12C0691
            09009B00FF00FFFF00FFFF00FF977967FF00FFFF00FF977967977967FF00FFB8
            9888977967FF00FFFF00FFFF00FF977967977967FF00FFFF00FFFF00FFFF00FF
            10D2240DBE20009700FF00FFFF00FFFF00FF00A80000930163CA7217B13114B3
            2D069009008F00FF00FFFF00FFFF00FF977967977967977967FF00FFFF00FFFF
            00FFB89888977967FF00FFFF00FFFF00FF977967977967FF00FFFF00FFFF00FF
            FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF00AD0000940064CB7515B1
            3116B330028904FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
            00FFFF00FFB89888977967FF00FFFF00FFFF00FF977967FF00FFFF00FFFF00FF
            FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF00AB0001970366CD
            782ABC46008A02FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
            00FFFF00FFFF00FFB89888977967FF00FFFF00FF977967FF00FFFF00FFFF00FF
            FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF009A000094
            01049306008400FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
            00FFFF00FFFF00FFFF00FF977967977967977967977967FF00FFFF00FFFF00FF
            FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
            FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
            00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
            FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
            FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
            00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF}
          NumGlyphs = 2
        end
        object btnRequery: TcxButton
          Left = 321
          Top = 421
          Width = 85
          Height = 29
          Anchors = [akTop, akRight]
          Caption = #37325#26032#26597#35426
          Default = True
          TabOrder = 5
          OnClick = btnRequeryClick
          Glyph.Data = {
            36040000424D3604000000000000360000002800000010000000100000000100
            2000000000000004000000000000000000000000000000000000000000000000
            0000000000000000000000000000000000000000000000000000000000000000
            0000000000000000000000000000000000000000000000000000000000000000
            0000000000000000000000000000000000000000000000000000000000000000
            0000000000000000FF018B8A7AC385827DA26B6B6B3F00000000000000000000
            0000000000000000000000000000000000000000000000000000000000000000
            00000000FF01094FFF974392F6FFEEE9DFFF86827FA700000000000000000000
            0000000000000000000000000000000000000000000000000000000000000000
            FF020D51FF9B439AFFFF6ADAFFFF5DADF5FF998F7ED900000000000000000000
            00000000000000000000000000000000000000000000000000006B7E690F0041
            FC9F469FFFFF6FDAFFFF50ACFFFF1357FFB9002AFF0A00000000000000000000
            000000000000000000006164682E515459524C4E523F1617180698928B9E7091
            B2FF61D3FFFF4EAAFFFF1657FFB4001FFF070000000000000000000000000000
            000076777A257B7E80C8B3A081FED2B588FDC3AA83FD83817AE566686CEAFFF7
            F0FF6B93BDFF084AFEAF0028FF06000000000000000000000000000000007B7D
            7F188D8A82E6F5CB84FEF5CB84FFF1C885FFF5CE8EFFFCD08CFFAE9E85FE7071
            75F3A9A193BB3956B20A0000000000000000000000000000000000000000767A
            7F95ECCB8EFFF3D192FFEECE92FFEDCC8EFFECC784FFEDC687FFFDD28FFF8783
            7CDB000000010000000000000000000000000000000000000000000000009894
            8BE1FEDFA1FFF2DAA5FFF2DBA7FFF1D79FFFEFD097FFECC886FFF6D298FFC1A9
            87FF5254592A000000000000000000000000000000000000000000000000A7A1
            91F4FFEBB9FFF8ECC6FFF7EBC0FFF5E3B2FFF2DAA3FFEECF94FFF4D093FFCFB6
            8DFF5153583D0000000000000000000000000000000000000000000000009996
            92D0FFF7CBFFFDFAE6FFFDF9E7FFF8EDC5FFF4E0B0FFF0D49BFFFAD594FFAF9F
            85FE6264671A000000000000000000000000000000000000000000000000A0A1
            A56FD5CCB1FFFFFFF2FFFFFFF2FFFBF3D2FFF6E2B4FFF7D99EFFF6D393FF7D7D
            81B800000000000000000000000000000000000000000000000000000000D9D9
            DA04A1A1A2B8D6CFB6FFFFFFDCFFFFF6CAFFFFEBB1FFEDD49CFE8C8983E47677
            7A20000000000000000000000000000000000000000000000000000000000000
            0000EBECED059C9CA07B979893E0A8A397FE97948CF07679809E7C7D7F1A0000
            0000000000000000000000000000000000000000000000000000000000000000
            0000000000000000000000000000000000000000000000000000000000000000
            0000000000000000000000000000000000000000000000000000}
        end
      end
      object CreateInv3: TcxTabSheet
        Caption = '  '#38283#31435#32080#26524'  '
        ImageIndex = 4
        OnShow = TransInv1Show
        object Label3: TLabel
          Left = 19
          Top = 6
          Width = 560
          Height = 21
          Alignment = taCenter
          AutoSize = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clRed
          Font.Height = -19
          Font.Name = #32048#26126#39636
          Font.Style = []
          ParentFont = False
          Transparent = True
          Layout = tlCenter
        end
        object pnlTotal: TPanel
          Left = 0
          Top = 0
          Width = 600
          Height = 221
          Align = alTop
          TabOrder = 0
          object lblQuery2: TLabel
            Left = 30
            Top = 11
            Width = 537
            Height = 32
            Alignment = taCenter
            AutoSize = False
            Caption = #38283#31435#32080#26524#65306
            Font.Charset = CHINESEBIG5_CHARSET
            Font.Color = clGreen
            Font.Height = -24
            Font.Name = #27161#26999#39636
            Font.Style = []
            ParentFont = False
            Transparent = True
            Layout = tlCenter
          end
          object lblSaleAmount3: TLabel
            Left = 88
            Top = 41
            Width = 414
            Height = 32
            Alignment = taCenter
            AutoSize = False
            Caption = #37559#36008#32317#38989#65306
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clMaroon
            Font.Height = -16
            Font.Name = #32048#26126#39636
            Font.Style = []
            ParentFont = False
            Transparent = True
            Layout = tlCenter
          end
          object lblTaxAmount3: TLabel
            Left = 146
            Top = 66
            Width = 300
            Height = 32
            Alignment = taCenter
            AutoSize = False
            Caption = #31237#37329#32317#38989#65306
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clMaroon
            Font.Height = -16
            Font.Name = #32048#26126#39636
            Font.Style = []
            ParentFont = False
            Transparent = True
            Layout = tlCenter
          end
          object lblInvAmount3: TLabel
            Left = 146
            Top = 93
            Width = 300
            Height = 32
            Alignment = taCenter
            AutoSize = False
            Caption = #32317#35336#65306
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clMaroon
            Font.Height = -16
            Font.Name = #32048#26126#39636
            Font.Style = []
            ParentFont = False
            Transparent = True
            Layout = tlCenter
          end
          object lblAssignInvCount: TLabel
            Left = 165
            Top = 121
            Width = 300
            Height = 32
            Alignment = taCenter
            AutoSize = False
            Caption = #30332#31080#38283#31435#24373#25976#65306'    '#24373
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clMaroon
            Font.Height = -16
            Font.Name = #32048#26126#39636
            Font.Style = []
            ParentFont = False
            Transparent = True
            Layout = tlCenter
          end
          object lblStartInvID: TLabel
            Left = 189
            Top = 150
            Width = 300
            Height = 32
            Alignment = taCenter
            AutoSize = False
            Caption = #30332#31080#38283#31435#36215#22987#34399#30908#65306'        '
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clMaroon
            Font.Height = -16
            Font.Name = #32048#26126#39636
            Font.Style = []
            ParentFont = False
            Transparent = True
            Layout = tlCenter
          end
          object lblStopInvID: TLabel
            Left = 189
            Top = 177
            Width = 300
            Height = 34
            Alignment = taCenter
            AutoSize = False
            Caption = #30332#31080#38283#31435#25130#27490#34399#30908#65306'        '
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clMaroon
            Font.Height = -16
            Font.Name = #32048#26126#39636
            Font.Style = []
            ParentFont = False
            Transparent = True
            Layout = tlCenter
          end
        end
        object pnlTotal3: TPanel
          Left = 0
          Top = 374
          Width = 600
          Height = 91
          Align = alBottom
          TabOrder = 1
          DesignSize = (
            600
            91)
          object Label12: TLabel
            Left = 245
            Top = 28
            Width = 105
            Height = 32
            Alignment = taCenter
            AutoSize = False
            Caption = #38283#31435#36914#24230#65306
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clGreen
            Font.Height = -16
            Font.Name = #32048#26126#39636
            Font.Style = []
            ParentFont = False
            Transparent = True
            Layout = tlCenter
            Visible = False
          end
          object lblCount: TLabel
            Left = 317
            Top = 36
            Width = 65
            Height = 24
            Alignment = taRightJustify
            AutoSize = False
            Caption = '0'
            Font.Charset = CHINESEBIG5_CHARSET
            Font.Color = clTeal
            Font.Height = -16
            Font.Name = #32048#26126#39636
            Font.Style = []
            ParentFont = False
            Transparent = True
            Visible = False
          end
          object cxButton1: TcxButton
            Left = 265
            Top = 58
            Width = 85
            Height = 29
            Anchors = [akTop, akRight]
            Caption = #32080#26463
            ModalResult = 2
            TabOrder = 0
            OnClick = btnExitClick
            Glyph.Data = {
              36030000424D3603000000000000360000002800000010000000100000000100
              18000000000000030000120B0000120B00000000000000000000FF00FFFF00FF
              FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
              FFFF00FFFF00FFFF00FFFF00FF0000A40206B0030AAE0000A6000098FF00FFFF
              00FFFF00FFFF00FF0000A700009A0000A7FF00FFFF00FFFF00FFFF00FF0000A9
              1844F6194DF81031D20102AB0000B6FF00FFFF00FF0000B10928D7092ED70313
              B30000ACFF00FFFF00FFFF00FF0103B32451F91F52FF1D4FFF1744E8040BB000
              00B00000AC0D2EDD1142F90D3DF50B3BF0041ABC0000A5FF00FFFF00FF0000AE
              1832DB285BFF2456FF2253FF1B4BF1050DB10F30DD164AFE1344F91041F60E3E
              F60A3CF000009FFF00FFFF00FF0000BE1F37DD3A6FFF2C5EFF295AFF2657FF20
              52FC1C4FFF194AFD1646FA1445FA0F3DF2020AB10000A8FF00FFFF00FFFF00FF
              0000C8121DC83D6AFB3567FF2C5DFF2859FF2253FF1D4EFF1A4DFF123DED0002
              AC0000BAFF00FFFF00FFFF00FFFF00FFFF00FF0000CC0000B62E4EE73668FF2E
              5EFF2859FF2254FF163DEA0000A80000ABFF00FFFF00FFFF00FFFF00FFFF00FF
              FF00FFFF00FF0000BF253FDF3B6DFF3464FF2E5EFF2759FF1B46EA0001AC0000
              A9FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF0203C84B7CFF4170FF3B
              6BFF396CFF2D5EFF2558FF1336D70000B6FF00FFFF00FFFF00FFFF00FFFF00FF
              FF00FF0000D9263CDB5080FF4575FF3662FA0D14C33C6DFF2A5BFF2053FD0B1D
              C40000C0FF00FFFF00FFFF00FFFF00FFFF00FF0000CB527CFA5081FF4B7DFF0B
              13C90000BB0E15C7386AFF2456FF1A4AF20207B30000B5FF00FFFF00FFFF00FF
              FF00FF131CDD6A9CFF5788FF2B46E70000CDFF00FF0000CD0F1BCB3065FF1F51
              FF1439DD0000B1FF00FFFF00FFFF00FFFF00FF0000DE3A52E45782FB0101D0FF
              00FFFF00FFFF00FF0000CC1426D6265AFF0F2EE30103B8FF00FFFF00FFFF00FF
              FF00FFFF00FF0000CF0000C00000CEFF00FFFF00FFFF00FFFF00FF0000C40001
              B80000B5000077FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
              00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF}
          end
          object ProgressBar1: TProgressBar
            Left = 195
            Top = 9
            Width = 225
            Height = 21
            Step = 1
            TabOrder = 1
            TabStop = True
            Visible = False
          end
        end
        object pnlTotal2: TPanel
          Left = 0
          Top = 221
          Width = 600
          Height = 153
          Align = alClient
          TabOrder = 2
          object pnlTotal4: TPanel
            Left = 1
            Top = 1
            Width = 598
            Height = 151
            Align = alClient
            TabOrder = 0
            Visible = False
            object Label23: TLabel
              Left = 20
              Top = 117
              Width = 91
              Height = 29
              AutoSize = False
              Caption = #38283#31435#24373#25976#65306
              Font.Charset = DEFAULT_CHARSET
              Font.Color = clMaroon
              Font.Height = -16
              Font.Name = #32048#26126#39636
              Font.Style = []
              ParentFont = False
              Transparent = True
              Layout = tlCenter
            end
            object Label22: TLabel
              Left = 52
              Top = 91
              Width = 56
              Height = 29
              AutoSize = False
              Caption = #32317#35336#65306
              Font.Charset = DEFAULT_CHARSET
              Font.Color = clMaroon
              Font.Height = -16
              Font.Name = #32048#26126#39636
              Font.Style = []
              ParentFont = False
              Transparent = True
              Layout = tlCenter
            end
            object Label21: TLabel
              Left = 20
              Top = 63
              Width = 104
              Height = 32
              AutoSize = False
              Caption = #31237#37329#32317#38989#65306
              Font.Charset = DEFAULT_CHARSET
              Font.Color = clMaroon
              Font.Height = -16
              Font.Name = #32048#26126#39636
              Font.Style = []
              ParentFont = False
              Transparent = True
              Layout = tlCenter
            end
            object Label20: TLabel
              Left = 20
              Top = 36
              Width = 104
              Height = 32
              AutoSize = False
              Caption = #37559#36008#32317#38989#65306
              Font.Charset = DEFAULT_CHARSET
              Font.Color = clMaroon
              Font.Height = -16
              Font.Name = #32048#26126#39636
              Font.Style = []
              ParentFont = False
              Transparent = True
              Layout = tlCenter
            end
            object Label16: TLabel
              Left = 131
              Top = 1
              Width = 141
              Height = 32
              Alignment = taCenter
              AutoSize = False
              Caption = #38651#23376#35336#31639#27231
              Font.Charset = CHINESEBIG5_CHARSET
              Font.Color = clGreen
              Font.Height = -24
              Font.Name = #27161#26999#39636
              Font.Style = []
              ParentFont = False
              Transparent = True
              Layout = tlCenter
            end
            object Label17: TLabel
              Left = 393
              Top = 1
              Width = 141
              Height = 32
              Alignment = taCenter
              AutoSize = False
              Caption = #38651#23376#30332#31080
              Font.Charset = CHINESEBIG5_CHARSET
              Font.Color = clGreen
              Font.Height = -24
              Font.Name = #27161#26999#39636
              Font.Style = []
              ParentFont = False
              Transparent = True
              Layout = tlCenter
            end
            object lblNoESaleAmt: TLabel
              Left = 154
              Top = 36
              Width = 104
              Height = 32
              Alignment = taCenter
              AutoSize = False
              Caption = '0'
              Font.Charset = DEFAULT_CHARSET
              Font.Color = clMaroon
              Font.Height = -16
              Font.Name = #32048#26126#39636
              Font.Style = []
              ParentFont = False
              Transparent = True
              Layout = tlCenter
            end
            object lblESaleAmt: TLabel
              Left = 420
              Top = 36
              Width = 104
              Height = 32
              Alignment = taCenter
              AutoSize = False
              Caption = '0'
              Font.Charset = DEFAULT_CHARSET
              Font.Color = clMaroon
              Font.Height = -16
              Font.Name = #32048#26126#39636
              Font.Style = []
              ParentFont = False
              Transparent = True
              Layout = tlCenter
            end
            object lblNoETaxAmt: TLabel
              Left = 154
              Top = 62
              Width = 104
              Height = 32
              Alignment = taCenter
              AutoSize = False
              Caption = '0'
              Font.Charset = DEFAULT_CHARSET
              Font.Color = clMaroon
              Font.Height = -16
              Font.Name = #32048#26126#39636
              Font.Style = []
              ParentFont = False
              Transparent = True
              Layout = tlCenter
            end
            object lblNoEInvAmt: TLabel
              Left = 154
              Top = 90
              Width = 104
              Height = 32
              Alignment = taCenter
              AutoSize = False
              Caption = '0'
              Font.Charset = DEFAULT_CHARSET
              Font.Color = clMaroon
              Font.Height = -16
              Font.Name = #32048#26126#39636
              Font.Style = []
              ParentFont = False
              Transparent = True
              Layout = tlCenter
            end
            object lblNoEInvCount: TLabel
              Left = 154
              Top = 116
              Width = 104
              Height = 32
              Alignment = taCenter
              AutoSize = False
              Caption = '0'
              Font.Charset = DEFAULT_CHARSET
              Font.Color = clMaroon
              Font.Height = -16
              Font.Name = #32048#26126#39636
              Font.Style = []
              ParentFont = False
              Transparent = True
              Layout = tlCenter
            end
            object lblETaxAmt: TLabel
              Left = 420
              Top = 62
              Width = 104
              Height = 32
              Alignment = taCenter
              AutoSize = False
              Caption = '0'
              Font.Charset = DEFAULT_CHARSET
              Font.Color = clMaroon
              Font.Height = -16
              Font.Name = #32048#26126#39636
              Font.Style = []
              ParentFont = False
              Transparent = True
              Layout = tlCenter
            end
            object lblEInvAmt: TLabel
              Left = 420
              Top = 90
              Width = 104
              Height = 32
              Alignment = taCenter
              AutoSize = False
              Caption = '0'
              Font.Charset = DEFAULT_CHARSET
              Font.Color = clMaroon
              Font.Height = -16
              Font.Name = #32048#26126#39636
              Font.Style = []
              ParentFont = False
              Transparent = True
              Layout = tlCenter
            end
            object lblEInvCount: TLabel
              Left = 420
              Top = 116
              Width = 104
              Height = 32
              Alignment = taCenter
              AutoSize = False
              Caption = '0'
              Font.Charset = DEFAULT_CHARSET
              Font.Color = clMaroon
              Font.Height = -16
              Font.Name = #32048#26126#39636
              Font.Style = []
              ParentFont = False
              Transparent = True
              Layout = tlCenter
            end
          end
        end
      end
    end
  end
  object OpenDialog1: TOpenDialog
    DefaultExt = 'txt'
    Filter = #30332#31080#25991#23383#27284'(Text File)|*.txt|'#25152#26377#27284#26696'(*.*)|*.*'
    Title = #38283#21855#27284#26696
    Left = 60
    Top = 484
  end
  object dsrPrefix: TDataSource
    DataSet = dtmMainJ.cdsPrefix
    Left = 24
    Top = 448
  end
  object ppmPrefix: TPopupMenu
    AutoHotkeys = maManual
    BiDiMode = bdLeftToRight
    ParentBiDiMode = False
    Left = 24
    Top = 484
    object N1: TMenuItem
      Caption = #37325#35373
      OnClick = N1Click
    end
  end
end
