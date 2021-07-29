Imports System
Imports System.Net
Imports System.Net.Sockets
Imports CableSoft_Nagra_BuilderCmd
Module mdlCommon
    Public FIPAddress As String = "127.0.0.1"
    Public FPort As String = "0"
    Public FSendTimeOut As String = "30"
    Public FRcvTimeOut As String = "30"
    Public FLowCmd As String = "0851"
    Public FEncrypted_asset_flag As String = "1"
    Public FValidity_Flag As String = "0"
    Public Fnb_of_R_VOD_ENT As String = String.Empty
    Public Fchipset_type_flag As String = String.Empty
    Public Fadditional_address_flag As String = String.Empty
    Public FDMM_creation_fate_flag As String = String.Empty
    Public FDMM_creation_date As String = String.Empty
    Public FSTBSNo As String = String.Empty
    Public FOrderDate As String = String.Empty
    Public Fviewing_duration As String = String.Empty
    Public Fsource_id As String = "0000"
    Public Fdest_id As String = "0000"
    Public FMOP_PPID As String = "00000"
    Public FSocket As Socket = Nothing
    Public FAsset As String = String.Empty
    Public FProductId As String = String.Empty
    Public Fexpiration_date As String = String.Empty
    Public FUrl As String = "Http://127.0.0.1"
    Public FUrlTimeOut As String = 30
    Public FUrlDuration As String = 30
    Public FUrlMetho As String = "POST"
    Public FUrlRepeat As String = "2"

    Public Sub InitialData()
        FIPAddress = My.Settings.IPAddress
        FPort = My.Settings.Port
        If String.IsNullOrEmpty(FPort) Then
            FPort = "0"
        End If
        FSendTimeOut = My.Settings.SendTimeOut
        FRcvTimeOut = My.Settings.RcvTimeOut
        FLowCmd = My.Settings.LowCmd
        FEncrypted_asset_flag = My.Settings.Encrypted_asset_flag
        FValidity_Flag = My.Settings.Validity_Flag
        Fnb_of_R_VOD_ENT = My.Settings.nb_of_R_VOD_ENT
        Fchipset_type_flag = My.Settings.chipset_type_flag
        Fadditional_address_flag = My.Settings.additional_address_flag
        FDMM_creation_fate_flag = My.Settings.DMM_creation_fate_flag
        FDMM_creation_date = Date.Now.ToString 'My.Settings.DMM_creation_date
        FSTBSNo = My.Settings.STBSNo
        FOrderDate = Date.Now.ToString  'My.Settings.OrderDate
        Fviewing_duration = My.Settings.viewing_duration
        Fsource_id = My.Settings.source_id
        Fdest_id = My.Settings.dest_id
        FMOP_PPID = My.Settings.MOP_PPID
        FProductId = My.Settings.ProductId
        FAsset = My.Settings.Asset
        Fexpiration_date = Date.Now.ToString 'My.Settings.expiration_date
        If Not String.IsNullOrEmpty(My.Settings.Url) Then
            FUrl = My.Settings.Url
        End If
        If Not String.IsNullOrEmpty(My.Settings.UrlDuration) Then
            FUrlDuration = My.Settings.UrlDuration
        End If
        If Not String.IsNullOrEmpty(My.Settings.UrlMetho) Then
            FUrlMetho = My.Settings.UrlMetho
        End If
        If Not String.IsNullOrEmpty(My.Settings.urlTimeOut) Then
            FUrlTimeOut = My.Settings.urlTimeOut
        End If
        If Not String.IsNullOrEmpty(My.Settings.UrlRepeat) Then
            FUrlRepeat = My.Settings.UrlRepeat
        End If
        NagraCommand.Dest_Id = Fdest_id
        NagraCommand.ErrCodeLst = Nothing
        NagraCommand.IniPath = CableSoft.Gateway.Common.csCommon.GetAppPath
        NagraCommand.MOP_PPID = FMOP_PPID
        NagraCommand.SourceId = Fsource_id

    End Sub
    Public Sub DisconnectSock()
        Try
            If FSocket IsNot Nothing Then
                FSocket.Shutdown(SocketShutdown.Both)
                FSocket.Disconnect(False)
                FSocket = Nothing
            End If
        Catch ex As Exception

        End Try
    End Sub
    Public Sub InitialSocket()
        Try
            Try
                If FSocket IsNot Nothing Then
                    FSocket.Shutdown(SocketShutdown.Both)
                    FSocket.Disconnect(False)
                    FSocket = Nothing
                End If
                
                FSocket = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)

            Catch ex As Exception

            End Try

        Catch ex As Exception
            MessageBox.Show(ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Module
