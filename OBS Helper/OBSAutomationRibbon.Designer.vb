Partial Class OBSAutomationRibbon
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.OBSTab = Me.Factory.CreateRibbonTab
        Me.grpOBS = Me.Factory.CreateRibbonGroup
        Me.tbOBSAutomation = Me.Factory.CreateRibbonToggleButton
        Me.OBSTab.SuspendLayout()
        Me.grpOBS.SuspendLayout()
        Me.SuspendLayout()
        '
        'OBSTab
        '
        Me.OBSTab.Groups.Add(Me.grpOBS)
        Me.OBSTab.Label = "OBS"
        Me.OBSTab.Name = "OBSTab"
        '
        'grpOBS
        '
        Me.grpOBS.Items.Add(Me.tbOBSAutomation)
        Me.grpOBS.Label = "OBS"
        Me.grpOBS.Name = "grpOBS"
        '
        'tbOBSAutomation
        '
        Me.tbOBSAutomation.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.tbOBSAutomation.Label = "OBS Automation"
        Me.tbOBSAutomation.Name = "tbOBSAutomation"
        Me.tbOBSAutomation.ShowImage = True
        '
        'OBSAutomationRibbon
        '
        Me.Name = "OBSAutomationRibbon"
        Me.RibbonType = "Microsoft.PowerPoint.Presentation"
        Me.Tabs.Add(Me.OBSTab)
        Me.OBSTab.ResumeLayout(False)
        Me.OBSTab.PerformLayout()
        Me.grpOBS.ResumeLayout(False)
        Me.grpOBS.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents OBSTab As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents grpOBS As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents tbOBSAutomation As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property OBSAutomationRibbon() As OBSAutomationRibbon
        Get
            Return Me.GetRibbon(Of OBSAutomationRibbon)()
        End Get
    End Property
End Class
