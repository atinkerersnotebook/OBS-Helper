Imports System.Diagnostics
Imports Microsoft.Office.Interop.PowerPoint
Imports OBSWebsocketDotNet
Imports Microsoft.VisualBasic.Devices.Keyboard
Imports Newtonsoft.Json.Linq
Imports System.IO
Imports System.Text

Public Class ThisAddIn
    Dim sOBSProcessName As String
    Dim OBSProcessID As Integer
    Private _obs As OBSWebsocket
    Public password As String = ""
    Public server As String = "ws://127.0.0.1:4444"
    Dim isInitialized As Boolean = False
    Public mOBSAutomationEnabled As Boolean
    Private myScriptPane As OBSControl
    Private myScriptTaskPane As Microsoft.Office.Tools.CustomTaskPane
    Private mScene As String
    Private mCurrentSlide As Slide

    Private Sub ThisAddIn_Startup() Handles Me.Startup

        myScriptPane = New OBSControl()
        myScriptTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(myScriptPane, "OBS Studio")
        myScriptTaskPane.Width = 320

        myScriptTaskPane.Visible = OBSAutomationEnabled
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        Try
            _obs.Disconnect()
        Catch ex As Exception

        End Try
    End Sub

    Public Property OBSAutomationEnabled() As Boolean
        Get
            Return mOBSAutomationEnabled
        End Get
        Set(value As Boolean)
            mOBSAutomationEnabled = value
            myScriptTaskPane.Visible = value
        End Set
    End Property

    Public Property Scene() As String
        Get
            'Return Application.ActiveWindow.Selection.SlideRange.Tags("Scene")
            Try
                Return mCurrentSlide.Tags("Scene")

            Catch ex As Exception
                Return ""

            End Try
        End Get
        Set(value As String)
            'Application.ActiveWindow.Selection.SlideRange.Tags.Add("Scene", value)
            'CallOBS(Application.ActivePresentation.Slides(Application.ActiveWindow.Selection.SlideRange(1).SlideIndex))

            mCurrentSlide.Tags.Add("Scene", value)
            CallOBS()
        End Set
    End Property

    Public Property Script() As String
        Get
            Try
                Return Application.ActiveWindow.Selection.SlideRange.Tags("Script")

            Catch exception As Exception
                Return ""
            End Try
        End Get
        Set(value As String)
            Try
                Application.ActiveWindow.Selection.SlideRange.Tags.Add("Script", value)
                'Globals.ThisAddIn.myScriptPane.TextBox1.Text = value
                'Globals.ThisAddIn.myScriptPane.Refresh()

            Catch exception As Exception
            End Try
        End Set
    End Property

    Private Sub Application_SlideSelectionChanged(SldRange As SlideRange) Handles Application.SlideSelectionChanged
        Try
            mCurrentSlide = Application.ActivePresentation.Slides(SldRange(1).SlideIndex)

            Globals.ThisAddIn.myScriptPane.SceneList.Text = Scene
            Globals.ThisAddIn.myScriptPane.Refresh()
            CallOBS()
            Exit Try
        Catch exception As Exception
            Exit Sub
        End Try
    End Sub

    Private Sub Application_SlideShowOnNext(Wn As SlideShowWindow) Handles Application.SlideShowOnNext
        Dim i As Integer

        Exit Sub

        Try
            i = Application.ActivePresentation.SlideShowWindow.View.CurrentShowPosition
            mCurrentSlide = Application.ActivePresentation.Slides(i)
            CallOBS()
            Exit Try
        Catch exception As Exception
            Exit Sub
        End Try
    End Sub

    Private Sub Application_SlideShowOnPrevious(Wn As SlideShowWindow) Handles Application.SlideShowOnPrevious
        Dim i As Integer

        Exit Sub

        Try
            i = Application.ActivePresentation.SlideShowWindow.View.CurrentShowPosition
            mCurrentSlide = Application.ActivePresentation.Slides(i)
            CallOBS()
            Exit Try
        Catch exception As Exception
            Exit Sub
        End Try
    End Sub

    Private Sub Application_SlideShowNextClick(Wn As SlideShowWindow, nEffect As Effect) Handles Application.SlideShowNextClick
        Dim i As Integer

        Try
            i = Application.ActivePresentation.SlideShowWindow.View.CurrentShowPosition
            mCurrentSlide = Application.ActivePresentation.Slides(i)
            CallOBS()
            Exit Try
        Catch exception As Exception
            Exit Sub
        End Try
    End Sub

    Private Sub Application_SlideShowNextSlide(Wn As SlideShowWindow) Handles Application.SlideShowNextSlide
        Dim i As Integer

        Try
            i = Application.ActivePresentation.SlideShowWindow.View.CurrentShowPosition
            mCurrentSlide = Application.ActivePresentation.Slides(i)
            CallOBS()
            Exit Try
        Catch exception As Exception
            Exit Sub
        End Try
    End Sub

    Private Sub Application_SlideShowNextBuild(Wn As SlideShowWindow) Handles Application.SlideShowNextBuild
        'RunMacro()
    End Sub

    Private Sub CallOBS()
        If Not OBSAutomationEnabled Then
            Exit Sub
        End If

        '        Try
        '        i = Application.ActivePresentation.SlideShowWindow.View.CurrentShowPosition
        '        Exit Try
        '        Catch exception As Exception
        '        Exit Sub
        '        End Try

        Try
            If Not isInitialized Then
                _obs = New OBSWebsocket()
                _obs.WSTimeout = New TimeSpan(0, 0, 0, 3)
                _obs.Connect(server, password)

                If _obs.IsConnected Then
                    isInitialized = True
                Else
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            Exit Sub
        End Try

        Try
            _obs.SetCurrentScene(Scene)

        Catch ex As Exception

        End Try
    End Sub

End Class
