Public Class InvTypeEnum
    Public Enum INVTYPE
        NormalInvAll = 0
        NormalInvOV = 1
        OnlyDiscount = 2
        OnlyDiscountOV = 3
        OnlyInvNum = 4
        OnlyNotUseInvNum = 5
        DestroyInv = 6
        DestroyReLoadInv = 7
        OnlyNormalInv = 8
    End Enum
    Public Enum InvFileType
        A0401 = 0
        A0501 = 1
        B0301 = 2
        B0401 = 3
        C0401 = 4
        C0501 = 5
        D0401 = 6
        D0501 = 7
        C0701 = 8
        E0402 = 9
        X0401 = 10
        B08 = 11
        UNKnow = 99
    End Enum
    Public Enum CreateInvoiceType
        icPrev = 1          '拋檔預開
        icAfter = 2         '拋檔後開
        icLocale = 3        '現場開立
        icNormal = 4     '一般開立
        icAll = 5              '全部
    End Enum
    Public Enum UploadSource
        B07 = 0
        Gateway = 1
    End Enum
End Class
Public Structure GatewayRunCommand
    Public Property Command As InvType.InvTypeEnum.InvFileType
    Public Property RunFrequency As Integer
    Public Property ExportElectronPath As String
End Structure
