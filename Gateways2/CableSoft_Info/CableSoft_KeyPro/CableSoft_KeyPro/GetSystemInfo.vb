Imports System
Imports System.Net
Imports System.Net.NetworkInformation
Imports Cablesoft.CAS.CryptUtil
Imports System.Windows
Imports System.Windows.Forms
Imports System.IO
Imports System.Text
Imports System.Xml
Imports System.Globalization

Public Class GetSystemInfo
    Public Const FKeyProName As String = "KeyPro.lic"
    Private Shared chkDate As DateTime = DateTime.MinValue
    Private Shared chkBoolean As Boolean = False
    Private Shared blnFirst As Boolean = True
    Private Const FSplitChar As Char = Chr(0)
    Private Shared intLckReg As Int32 = 0
    Private Structure SYSTEM_INFO
        Dim dwOemID As Integer
        Dim wProcessorArchitecture As Integer
        Dim wReserved As Integer
        Dim dwPageSize As Integer
        Dim lpMinimumApplicationAddress As Integer
        Dim lpMaximumApplicationAddress As Integer
        Dim dwActiveProcessorMask As Integer
        Dim dwNumberOrfProcessors As Integer
        Dim dwProcessorType As Integer
        Dim dwAllocationGranularity As Short
        Dim wProcessLevel As Short
        Dim wProcessorRevision As Short
    End Structure
    Private Declare Sub GetSystemInfo Lib "kernel32" (ByRef lpSystemInfo As SYSTEM_INFO)
    '取出CPU序號
    Public Shared Function GetCPUCode(Optional ByVal blnEncry As Boolean = False) As String
        Dim CPUInfo As SYSTEM_INFO
        GetSystemInfo(CPUInfo)
        If blnEncry Then
            Return CryptUtil.Encrypt(CPUInfo.dwProcessorType)
        Else
            Return CPUInfo.dwProcessorType
        End If

    End Function
    '取得硬碟序號
    Public Shared Function GetHardDriveCode(Optional ByVal DrvIdx As Byte = 1, Optional ByVal blnEncry As Boolean = False) As String
        Dim aDrvIndex As Byte = DrvIdx
        Dim aRet As String = String.Empty
        If aDrvIndex = 0 Then
            aDrvIndex += 1
        End If
        If System.IO.DriveInfo.GetDrives.Length < aDrvIndex Then
            aDrvIndex = 1
        End If

        aDrvIndex -= 1
        Dim WMI As Object = GetObject("winmgmts:")
        Dim strCls As String = "Win32_PhysicalMedia"
        Dim strKey As String = strCls & ".Tag=""\\\\.\\PHYSICALDRIVE" & aDrvIndex & """"
        aRet = WMI.InstancesOf(strCls)(strKey).SerialNumber.ToString.Trim

        System.Runtime.InteropServices.Marshal.ReleaseComObject(WMI)
        If blnEncry Then
            aRet = CryptUtil.Encrypt(aRet)
        End If
        Return aRet
    End Function
    '取得主機板序號
    Public Shared Function GetMotherBoardCode(Optional ByVal blnEncry As Boolean = False) As String
        Try
            Dim WMI As Object = GetObject("winmgmts:")
            Dim strCls As String = "Win32_BaseBoard"
            Dim strKey As String = strCls & ".Tag=""Base Board"""
            Dim aRet As String = String.Empty
            aRet = WMI.InstancesOf(strCls)(strKey).SerialNumber.ToString.Trim
            System.Runtime.InteropServices.Marshal.ReleaseComObject(WMI)
            If blnEncry Then
                aRet = CryptUtil.Encrypt(aRet)
            End If
            Return aRet
        Catch ex As Exception
            Return "0000000"
        End Try




        '        ManagementClass mc = new ManagementClass("WIN32_BaseBoard");
        '04	    ManagementObjectCollection moc = mc.GetInstances();
        '05	    string SerialNumber = "";
        '06	    foreach (ManagementObject mo in moc)
        '07	    {
        '08	        SerialNumber= mo["SerialNumber"].ToString();
        '09	        break;
        '10	    }
        '11	    return SerialNumber;
    End Function
    '取出網路卡Mac
    Public Shared Function GetMacAddress(ByVal MacIdx As Byte, Optional ByVal blnEncry As Boolean = False) As String
        Dim MacAry() As NetworkInterface = NetworkInterface.GetAllNetworkInterfaces
        Dim aGetIndex As Byte = MacIdx
        Dim aGetAll As Boolean = aGetIndex = 0
        Dim aRetString As String = String.Empty
        If (Not aGetAll) AndAlso (MacAry.Length < aGetIndex) Then
            aGetAll = True
        End If
        If aGetAll Then
            For Each objMac As NetworkInterface In MacAry
                If Not String.IsNullOrEmpty(objMac.GetPhysicalAddress.ToString) Then
                    If String.IsNullOrEmpty(aRetString) Then
                        aRetString = objMac.GetPhysicalAddress().ToString
                    Else
                        aRetString = aRetString & ";" & objMac.GetPhysicalAddress.ToString
                    End If
                End If
            Next
        Else
            Dim objMac As NetworkInterface = MacAry(aGetIndex - 1)
            aRetString = objMac.GetPhysicalAddress.ToString
        End If
        If String.IsNullOrEmpty(aRetString) Then
            aRetString = "00000000"
        End If
        Erase MacAry
        If blnEncry Then
            aRetString = CryptUtil.Encrypt(aRetString)
        End If
        Return aRetString

    End Function
    '取出所有資訊
    Public Shared Function GetAllInfo(Optional ByVal blnEncry As Boolean = False,
                                      Optional ExeName As String = Nothing) As String
        Dim aRet As String = String.Empty
        aRet = Date.Today.ToString(CultureInfo.CreateSpecificCulture("zh-TW")) & FSplitChar & _
                GetCPUCode() & Environment.NewLine & GetHardDriveCode(1) & _
                Environment.NewLine & GetMotherBoardCode() & _
                Environment.NewLine & GetMacAddress(1) & _
                IIf(String.IsNullOrEmpty(ExeName),
                    Environment.NewLine & Path.GetFileName(Application.ExecutablePath),
                    Environment.NewLine & ExeName)

        If blnEncry Then
            aRet = CryptUtil.Encrypt(aRet)
        End If
        Return aRet
    End Function
    Public Shared Function EncryKeyPro(ByVal aSourceStr As String) As String
        Try
            Return CryptUtil.Encrypt(aSourceStr)

        Catch ex As Exception
            Return String.Empty
        End Try
    End Function
    Public Overloads Shared Function DecryKeyPro(ByVal aFile As String) As String
        Dim aString As String = String.Empty
        Using stm As New System.IO.FileStream(aFile, IO.FileMode.Open, IO.FileAccess.Read)
            aString = DecryKeyPro(stm)
            stm.Close()
        End Using
        Return aString
    End Function
    Public Overloads Shared Function DecryKeyPro() As String
        Dim aFile As String = String.Empty
        Dim aDirectory As String = Application.StartupPath
        If Not aDirectory.EndsWith("\") Then
            aDirectory = aDirectory & "\"
        End If
        aFile = aDirectory & FKeyProName
        Return DecryKeyPro(aFile)
    End Function
    Public Overloads Shared Function DecryKeyPro(ByVal aStream As System.IO.Stream) As String
        Dim aString As String = String.Empty
        Try
            Using stm As New System.IO.StreamReader(aStream)
                aString = stm.ReadToEnd
                stm.Close()
            End Using
            aString = CryptUtil.Decrypt(aString)
            Return aString
        Catch ex As Exception
            Return String.Empty
        End Try
    End Function
    Private Shared Function ProcRegister(ByVal aSourceStr As Array, Optional ByVal ExeName As String = Nothing) As Boolean
        Try

            SyncLock intLckReg.GetType
                Dim aGetString() As String
                Dim aSysInfo() As String
                Dim aSourceLst As New System.Collections.Generic.Dictionary(Of String, String)
                Dim aDesLst As New System.Collections.Generic.Dictionary(Of String, String)
                aSourceLst.Clear()
                aDesLst.Clear()
                aGetString = aSourceStr
                aSysInfo = GetAllInfo(False, ExeName).Split(FSplitChar).GetValue(1).ToString.Split(Environment.NewLine)
                If aGetString.Length <> 7 Then

                    Return False
                End If
                If aSysInfo.Length <> 5 Then

                    Return False
                End If
                aSourceLst.Add("CPU", aGetString(0).Trim(vbCrLf).Trim)
                aSourceLst.Add("HD", aGetString(1).Trim(vbCrLf).Trim)
                aSourceLst.Add("MB", aGetString(2).Trim(vbCrLf).Trim)
                aSourceLst.Add("MAC", aGetString(3).Trim(vbCrLf).Trim)
                aSourceLst.Add("EXE", aGetString(4).Trim(vbCrLf).Trim)
                aSourceLst.Add("INSDATE", aGetString(5).Trim(vbCrLf).Trim)
                aSourceLst.Add("USEDATE", aGetString(6).Trim(vbCrLf).Trim)
                aDesLst.Add("CPU", aSysInfo(0).Trim(vbCrLf).Trim)
                aDesLst.Add("HD", aSysInfo(1).Trim(vbCrLf).Trim)
                aDesLst.Add("MB", aSysInfo(2).Trim(vbCrLf).Trim)
                aDesLst.Add("MAC", aSysInfo(3).Trim(vbCrLf).Trim)
                aDesLst.Add("EXE", aSysInfo(4).Trim(vbCrLf).Trim)
                Dim aDate As DateTime = Nothing
                If DateTime.TryParse(aSourceLst.Item("USEDATE"), _
                                     CultureInfo.CreateSpecificCulture("zh-TW"), _
                                     DateTimeStyles.AssumeLocal, _
                                     aDate) Then
                    If aDate < DateTime.Now.Date Then
                        Return False
                    End If
                Else
                    Return False
                End If

                If Not aDesLst.Item("CPU").Equals(aSourceLst.Item("CPU")) Then
                    Return False
                End If
                If Not aDesLst.Item("HD").Equals(aSourceLst.Item("HD")) Then
                    Return False
                End If
                If Not aDesLst.Item("MB").Equals(aSourceLst.Item("MB")) Then
                    Return False
                End If
                Dim aryMac() As String = GetMacAddress(0).Split(";")
                If Not aryMac.Contains(aDesLst.Item("MAC")) Then
                    Return False
                End If

                If Not aDesLst.Item("EXE").ToUpper.Equals(aSourceLst.Item("EXE").ToUpper) Then
                    Return False
                End If
                Return True
            End SyncLock
        Catch ex As Exception
            Return False
        End Try
    End Function
    Public Overloads Shared Function IsRegister(ByVal aFile As String,
                                                ByVal ExeName As String,
                                                Optional ByVal aShowForm As ShowForm = ShowForm.Yes
                                                ) As Boolean
        Try

            If (blnFirst) OrElse (Date.Today <> chkDate) Then
                blnFirst = False
                chkBoolean = False
                If Not File.Exists(aFile) Then
                    If aShowForm = ShowForm.Yes Then
                        Dim aFrm As New frmRegister()
                        aFrm.ExeName = ExeName
                        aFrm.ShowDialog()
                        chkBoolean = aFrm.RegisterOK
                        Return chkBoolean
                    Else
                        chkBoolean = False
                        Return False
                    End If

                End If
                Dim aGetString() As String
                'Dim aSysInfo() As String
                Dim aSourceLst As New System.Collections.Generic.Dictionary(Of String, String)
                Dim aDesLst As New System.Collections.Generic.Dictionary(Of String, String)
                aSourceLst.Clear()
                aDesLst.Clear()
                aGetString = DecryKeyPro(aFile).Split(Environment.NewLine)
                chkBoolean = ProcRegister(aGetString, ExeName)

                'aSysInfo = GetAllInfo.Split(Environment.NewLine)
                'If aGetString.Length <> 7 Then

                '    Return False
                'End If
                'If aSysInfo.Length <> 5 Then

                '    Return False
                'End If
                'aSourceLst.Add("CPU", aGetString(0))
                'aSourceLst.Add("HD", aGetString(1))
                'aSourceLst.Add("MB", aGetString(2))
                'aSourceLst.Add("MAC", aGetString(3))
                'aSourceLst.Add("EXE", aGetString(4))
                'aSourceLst.Add("INSDATE", aGetString(5))
                'aSourceLst.Add("USEDATE", aGetString(6))
                'aDesLst.Add("CPU", aSysInfo(0))
                'aDesLst.Add("HD", aSysInfo(1))
                'aDesLst.Add("MB", aSysInfo(2))
                'aDesLst.Add("MAC", aSysInfo(3))
                'aDesLst.Add("EXE", aSysInfo(4))
                'Dim aDate As DateTime = Nothing
                'If DateTime.TryParse(aSourceLst.Item("USEDATE"), aDate) Then
                '    If aDate < DateTime.Now Then
                '        Return False
                '    End If
                'Else
                '    Return False
                'End If

                'If Not aDesLst.Item("CPU").Equals(aSourceLst.Item("CPU")) Then
                '    Return False
                'End If
                'If Not aDesLst.Item("HD").Equals(aSourceLst.Item("HD")) Then
                '    Return False
                'End If
                'If Not aDesLst.Item("MB").Equals(aSourceLst.Item("MB")) Then
                '    Return False
                'End If
                'If Not aDesLst.Item("MAC").Equals(aSourceLst.Item("MAC")) Then
                '    Return False
                'End If
                'If Not aDesLst.Item("EXE").Equals(aSourceLst.Item("EXE")) Then
                '    Return False
                'End If
                'chkBoolean = True
                'chkDate = DateTime.Today
                'Return True
                Return chkBoolean
            Else
                Return chkBoolean
            End If
        Catch ex As Exception
            Return False
        End Try

    End Function



    Public Overloads Shared Function IsRegister(ByVal aFile As String,
                                                Optional ByVal aShowForm As ShowForm = ShowForm.Yes
                                                ) As Boolean
        Try

            If (blnFirst) OrElse (Date.Today <> chkDate) Then
                blnFirst = False
                chkBoolean = False
                If Not File.Exists(aFile) Then
                    If aShowForm = ShowForm.Yes Then
                        Dim aFrm As New frmRegister()
                        aFrm.ShowDialog()
                        chkBoolean = aFrm.RegisterOK
                        Return chkBoolean
                    Else
                        chkBoolean = False
                        Return False
                    End If

                End If
                Dim aGetString() As String
                'Dim aSysInfo() As String
                Dim aSourceLst As New System.Collections.Generic.Dictionary(Of String, String)
                Dim aDesLst As New System.Collections.Generic.Dictionary(Of String, String)
                aSourceLst.Clear()
                aDesLst.Clear()
                aGetString = DecryKeyPro(aFile).Split(Environment.NewLine)
                chkBoolean = ProcRegister(aGetString)

                'aSysInfo = GetAllInfo.Split(Environment.NewLine)
                'If aGetString.Length <> 7 Then

                '    Return False
                'End If
                'If aSysInfo.Length <> 5 Then

                '    Return False
                'End If
                'aSourceLst.Add("CPU", aGetString(0))
                'aSourceLst.Add("HD", aGetString(1))
                'aSourceLst.Add("MB", aGetString(2))
                'aSourceLst.Add("MAC", aGetString(3))
                'aSourceLst.Add("EXE", aGetString(4))
                'aSourceLst.Add("INSDATE", aGetString(5))
                'aSourceLst.Add("USEDATE", aGetString(6))
                'aDesLst.Add("CPU", aSysInfo(0))
                'aDesLst.Add("HD", aSysInfo(1))
                'aDesLst.Add("MB", aSysInfo(2))
                'aDesLst.Add("MAC", aSysInfo(3))
                'aDesLst.Add("EXE", aSysInfo(4))
                'Dim aDate As DateTime = Nothing
                'If DateTime.TryParse(aSourceLst.Item("USEDATE"), aDate) Then
                '    If aDate < DateTime.Now Then
                '        Return False
                '    End If
                'Else
                '    Return False
                'End If

                'If Not aDesLst.Item("CPU").Equals(aSourceLst.Item("CPU")) Then
                '    Return False
                'End If
                'If Not aDesLst.Item("HD").Equals(aSourceLst.Item("HD")) Then
                '    Return False
                'End If
                'If Not aDesLst.Item("MB").Equals(aSourceLst.Item("MB")) Then
                '    Return False
                'End If
                'If Not aDesLst.Item("MAC").Equals(aSourceLst.Item("MAC")) Then
                '    Return False
                'End If
                'If Not aDesLst.Item("EXE").Equals(aSourceLst.Item("EXE")) Then
                '    Return False
                'End If
                'chkBoolean = True
                'chkDate = DateTime.Today
                'Return True
                Return chkBoolean
            Else
                Return chkBoolean
            End If
        Catch ex As Exception
            Return False
        End Try

    End Function
    Public Overloads Shared Function IsRegister(ByVal stm As Stream,
                                                ByVal ExeName As String) As Boolean
        Dim aSource() As String = DecryKeyPro(stm).Split(Environment.NewLine)
        Return ProcRegister(aSource, ExeName)
    End Function
    Public Overloads Shared Function IsRegister(ByVal stm As Stream) As Boolean
        Try
            'Dim aSource() As String = DecryKeyPro(stm).Split(Environment.NewLine)
            'Return ProcRegister(aSource)
            Return IsRegister(stm, Nothing)
        Catch ex As Exception
            Return False
        End Try
    End Function
    Public Shared Function RegisterApplication(ByVal ExeName As String) As Boolean
        Try
            Dim aFile As String = String.Empty
            Dim aDirectory As String = Application.StartupPath
            If Not aDirectory.EndsWith("\") Then
                aDirectory = aDirectory & "\"
            End If
            aFile = aDirectory & FKeyProName
            Return IsRegister(aFile, ExeName, ShowForm.Yes)

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Function



    Public Overloads Shared Function IsRegister(Optional ByVal aShowFrom As ShowForm = ShowForm.Yes) As Boolean
        Try
            Dim aFile As String = String.Empty
            Dim aDirectory As String = Application.StartupPath
            If Not aDirectory.EndsWith("\") Then
                aDirectory = aDirectory & "\"
            End If
            aFile = aDirectory & FKeyProName
            Return IsRegister(aFile, aShowFrom)
        Catch ex As Exception
            Return False
        End Try
    End Function

End Class
Public Enum ShowForm
    Yes = 0
    No = 1
End Enum