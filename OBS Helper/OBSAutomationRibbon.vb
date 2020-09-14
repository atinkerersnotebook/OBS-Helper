Imports Microsoft.Office.Tools.Ribbon

Public Class OBSAutomationRibbon

    Private Sub OBSAutomationRibbon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        Globals.ThisAddIn.OBSAutomationEnabled = False
    End Sub

    Private Sub tbOBSAutomation_Click(sender As Object, e As RibbonControlEventArgs) Handles tbOBSAutomation.Click
        Globals.ThisAddIn.OBSAutomationEnabled = tbOBSAutomation.Checked
    End Sub
End Class
