Imports WIALib
Public Class Form1


    Private Sub НастройкаToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles НастройкаToolStripMenuItem.Click
        Dialog1.Show()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim wiaman = New Wia
        My.Settings.Reload()
        Dim i As Integer = 0
        Dim array(wiaman.Devices.Count) As WIALib.Item
        For i = 1 To wiaman.Devices.Count
            array(i) = wiaman.Create(wiaman.Devices.Item(i))
            Dim deviceWiaRoot
            Dim wiapics As WIALib.Collection
            deviceWiaRoot = wiaman.Create(wiaman.Devices.Item(i))
            wiapics = deviceWiaRoot.GetItemsFromUI()
            Dim cnt
            cnt = 0
            For cnt = 0 To wiapics.Count
                Dim wiaitem As WIALib.Item = wiapics.Item(cnt)
                Dim lastphotonumber As Integer
                lastphotonumber = My.Computer.FileSystem.GetFiles(My.Settings.ToPath).Count
                On Error GoTo errorka
                wiaitem.Transfer(My.Settings.ToPath + "\Фото" + CStr(lastphotonumber) + ".jpg")

            Next
            
        Next
errorka:
        If My.Settings.OpenOnEnd = True Then Shell("explorer " + My.Settings.ToPath, AppWinStyle.NormalFocus)

    End Sub
End Class
