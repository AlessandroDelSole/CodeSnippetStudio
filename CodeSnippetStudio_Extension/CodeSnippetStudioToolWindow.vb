Imports System
Imports System.Collections
Imports System.ComponentModel
Imports System.Drawing
Imports System.Data
Imports System.Windows
Imports System.Runtime.InteropServices
Imports Microsoft.VisualStudio.Shell.Interop
Imports Microsoft.VisualStudio.Shell

''' <summary>
''' This class implements the tool window exposed by this package and hosts a user control.
''' </summary>
''' <remarks>
''' In Visual Studio tool windows are composed of a frame (implemented by the shell) and a pane, 
''' usually implemented by the package implementer.
''' <para>
''' This class derives from the ToolWindowPane class provided from the MPF in order to use its 
''' implementation of the IVsUIElementPane interface.
''' </para>
''' </remarks>
<Guid("7490573b-4962-47d3-9567-5459fc01d724")>
Public Class CodeSnippetStudioToolWindow
    Inherits ToolWindowPane

    ''' <summary>
    ''' Initializes a new instance of the <see cref="CodeSnippetStudioToolWindow"/> class.
    ''' </summary>
    Public Sub New()
        MyBase.New(Nothing)
        Me.Caption = "Code Snippet Studio"

        'This is the user control hosted by the tool window; Note that, even if this class implements IDisposable,
        'we are not calling Dispose on this object. This is because ToolWindowPane calls Dispose on 
        'the object returned by the Content property.
        Me.Content = New CodeSnippetStudioToolWindowControl()
    End Sub

End Class
