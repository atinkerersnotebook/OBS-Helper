Imports Microsoft.Office.Interop.PowerPoint
Imports OBS_Helper.My.Resources
Imports OBSWebsocketDotNet
Imports OBSWebsocketDotNet.Types
Imports OBSWebsocketDotNet.OBSWebsocket


Public Class OBSControl

    Protected _obs As OBSWebsocket
    Protected _scenes As List(Of OBSScene)

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs)
    End Sub

    Private Sub RichTextBox1_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub ScriptingControl_Load(sender As Object, e As EventArgs) Handles Me.Load
        'LoadSceneList()
    End Sub

    Private Sub SceneList_SelectedValueChanged(sender As Object, e As EventArgs) Handles SceneList.SelectedValueChanged
        Globals.ThisAddIn.Scene = SceneList.Items(SceneList.SelectedIndex).ToString()
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click
        LoadSceneList()
    End Sub

    Private Sub LoadSceneList()

        Try
            SceneList.Items.Clear()
            SceneList.Items.Add("-")

            _obs = New OBSWebsocket()
            _obs.Connect(Globals.ThisAddIn.server, Globals.ThisAddIn.password)

            If _obs.IsConnected Then
                Dim _scenes = _obs.ListScenes()
                For Each scene In _scenes
                    SceneList.Items.Add(scene.Name.ToString())
                Next

                _obs.Disconnect()
            End If
        Catch ex As Exception
            Exit Sub
        End Try


    End Sub
End Class
