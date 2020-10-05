Imports System.Runtime.InteropServices
Imports ExcelDna.Integration.CustomUI
Imports just4Excel.Core.TaskPane

<ComVisible(True)>
Public Class MyCustomRibbon : Inherits ExcelRibbon



    ''' <summary>
    ''' Use this function to return the custom ribbon XML.
    ''' </summary>
    ''' <param name="RibbonID"></param>
    ''' <returns></returns>
    Public Overrides Function GetCustomUI(RibbonID As String) As String

        Return "
    <customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
      <ribbon>
        <tabs>
          <tab id='myTab' label='My Tab' visible='true'>
            <group id='myGroup' label='My Group' visible='true'>
              <button id='myButton' label='My Button' size='large' imageMso='HappyFace' onAction='OnMyButton_Action' visible='true'/>
            </group >
          </tab>
        </tabs>
      </ribbon>
    </customUI>"

    End Function




    ''' <summary>
    ''' This procedure needs to have the same name as the property onAction.
    ''' </summary>
    ''' <param name="control"></param>
    Public Sub OnMyButton_Action(control As IRibbonControl)


        If myTaskPane Is Nothing Then
            myTaskPane = myUserControl.ShowInTaskPane("My user control")
        Else
            myTaskPane.Visible = True
        End If


    End Sub


    Private myUserControl As New MyUserControl
    Private myTaskPane As WpfTaskPane


End Class
