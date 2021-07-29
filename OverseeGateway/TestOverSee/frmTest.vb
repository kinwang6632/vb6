Imports System.Threading.Tasks

Public Class frmTest
    Private testSO As Cablesoft.Oversee.Gateway.ExecuteSO.ExecuteSO
    Private act As Action(Of Cablesoft.Oversee.Gateway.SODB.SO.StatusOfProcess,
                              Integer, Integer, String, String, String)
    Private Sub btnStart_Click(sender As System.Object, e As System.EventArgs) Handles btnStart.Click
        btnStop.Enabled = True
        btnStart.Enabled = False
        testSO = New Cablesoft.Oversee.Gateway.ExecuteSO.ExecuteSO
        AddHandler testSO.StatusChange, AddressOf procStatus


        act = AddressOf updUi


        testSO.ExecuteAll()

    End Sub

    Private Sub btnStop_Click(sender As System.Object, e As System.EventArgs) Handles btnStop.Click
        RemoveHandler testSO.StatusChange, AddressOf procStatus
        btnStop.Enabled = False
        testSO.StopAll()
        testSO.Dispose()
        btnStart.Enabled = True
    End Sub

    Private Sub procStatus(StatusOK As Cablesoft.Oversee.Gateway.SODB.SO.StatusOfProcess,
                              ByVal sequence As Integer,
                              ByVal comId As Integer,
                              ByVal compName As String,
                              ByVal prgName As String,
                              ByVal msg As String)
     
        If Me.InvokeRequired Then
            Dim objary As New List(Of Object)
            objary.Add(StatusOK)
            objary.Add(sequence)
            objary.Add(comId)
            objary.Add(compName)
            objary.Add(prgName)
            objary.Add(msg)
            Try
                Me.Invoke(act, objary.ToArray)
            Catch ex As Exception
            Finally
                objary.Clear()
                objary = Nothing
            End Try

        End If
    End Sub

    Private Sub frmTest_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        btnStop.Enabled = False
        If testSO IsNot Nothing Then
            testSO.Dispose()
            testSO = Nothing
        End If
    End Sub

    Private Sub updUi(StatusOK As Cablesoft.Oversee.Gateway.SODB.SO.StatusOfProcess,
                              ByVal sequence As Integer,
                              ByVal comId As Integer,
                              ByVal compName As String,
                              ByVal prgName As String,
                              ByVal msg As String)
        Try
            If lstEvent.Items.Count > 100 Then
                lstEvent.Items.Clear()
            End If
            lstEvent.Items.Add(compName & ":" & msg)
        Catch ex As Exception
        Finally
            compName = Nothing
            prgName = Nothing
            msg = Nothing
        End Try
    End Sub

End Class
