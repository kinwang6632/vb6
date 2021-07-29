
Public Class Form1

    Private Sub SimpleButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SimpleButton1.Click
        MsgBox(CableSoft_KeyPro.GetSystemInfo.GetMotherBoardCode())
        'If Not CableSoft_KeyPro.GetSystemInfo.IsRegister() Then
        '    MsgBox("不合法")
        'Else
        '    MsgBox("合法")
        'End If

    End Sub

    Private Sub SimpleButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SimpleButton2.Click
        'TextBox1.Text = CableSoft_KeyPro.GetSystemInfo.GetAllInfo(False)
        If CableSoft_KeyPro.GetSystemInfo.IsRegister(CableSoft_KeyPro.ShowForm.Yes) Then
            MsgBox("成功")
        Else
            MsgBox("失敗")
        End If
        'TextBox1.Text = CableSoft_KeyPro.GetSystemInfo.GrandKeyPro("bBkayM6ItQZzlRIv7VIEkQAZ5Hqq6Uzgt2JKPCd2W+kmRlJIy7AQzvHDTgKE72AJOr0rLHqcPR6dn98Zk9ED3XgBjU+MwVr8I0ymiJnAHOu8VL5YGlrgwDy6X36FLvZz63jZOvzN7PY4srx6F7pSsA==")
    End Sub
End Class
