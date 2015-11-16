Imports System
Imports System.ComponentModel.Design
Imports System.Globalization
Imports Microsoft.VisualStudio.Shell
Imports Microsoft.VisualStudio.Shell.Interop

''' <summary>
''' Command handler
''' </summary>
Public NotInheritable Class CodeSnippetStudioToolWindowCommand

    ''' <summary>
    ''' Command ID.
    ''' </summary>
    Public Const CommandId As Integer = 256

    ''' <summary>
    ''' Command menu group (command set GUID).
    ''' </summary>
    Public Shared ReadOnly CommandSet As New Guid("976a5281-d250-47ee-b483-b19a2f3149b5")

    ''' <summary>
    ''' VS Package that provides this command, not null.
    ''' </summary>
    Private ReadOnly package As package

    ''' <summary>
    ''' Initializes a new instance of the <see cref="CodeSnippetStudioToolWindowCommand"/> class.
    ''' Adds our command handlers for menu (the commands must exist in the command table file)
    ''' </summary>
    ''' <param name="package">Owner package, not null.</param>
    Private Sub New(package As package)
        If package Is Nothing Then
            Throw New ArgumentNullException("package")
        End If

        Me.package = package
        Dim commandService As OleMenuCommandService = Me.ServiceProvider.GetService(GetType(IMenuCommandService))
        If commandService IsNot Nothing Then
            Dim menuCommandId = New CommandID(CommandSet, CommandId)
            Dim menuCommand = New MenuCommand(AddressOf Me.ShowToolWindow, menuCommandId)
            commandService.AddCommand(menuCommand)
        End If
    End Sub

    ''' <summary>
    ''' Gets the instance of the command.
    ''' </summary>
    Public Shared Property Instance As CodeSnippetStudioToolWindowCommand

    ''' <summary>
    ''' Get service provider from the owner package.
    ''' </summary>
    Private ReadOnly Property ServiceProvider As IServiceProvider
        Get
            Return Me.package
        End Get
    End Property

    ''' <summary>
    ''' Initializes the singleton instance of the command.
    ''' </summary>
    ''' <param name="package">Owner package, Not null.</param>
    Public Shared Sub Initialize(package As package)
        Instance = New CodeSnippetStudioToolWindowCommand(package)
    End Sub

    ''' <summary>
    ''' Shows the tool window when the menu item is clicked.
    ''' </summary>
    ''' <param name="sender">The event sender.</param>
    ''' <param name="e">The event args.</param>
    Private Sub ShowToolWindow(sender As Object, e As EventArgs)

        '' Get the instance number 0 of this tool window. This window Is single instance so this instance
        '' Is actually the only one.
        '' The last flag Is set to true so that if the tool window does Not exists it will be created.
        Dim window As ToolWindowPane = Me.package.FindToolWindow(GetType(CodeSnippetStudioToolWindow), 0, True)
        If window Is Nothing OrElse window.Frame Is Nothing Then
            Throw New NotSupportedException("Cannot create tool window")
        End If

        Dim windowFrame As IVsWindowFrame = window.Frame
        Microsoft.VisualStudio.ErrorHandler.ThrowOnFailure(windowFrame.Show())
    End Sub
End Class
