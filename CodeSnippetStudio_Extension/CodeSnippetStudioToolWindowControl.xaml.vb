
Imports System, System.Linq
Imports System.Diagnostics
Imports DelSole.VSIX
Imports Microsoft.Win32
'''<summary>
''' Interaction logic for CodeSnippetStudioToolWindowControl.xaml
'''</summary>
Partial Public Class CodeSnippetStudioToolWindowControl
    Inherits System.Windows.Controls.UserControl

    Private theData As VSIXPackage

    Private Sub ResetPkg()
        Me.theData = New VSIXPackage

        Me.DataContext = Me.theData
        Me.PackageTab.Focus()
    End Sub

    Private Sub MyControl_Loaded(sender As Object, e As Windows.RoutedEventArgs) Handles Me.Loaded
        ResetPkg()
    End Sub


    Private Sub AddSnippetsButton_Click(sender As Object, e As Windows.RoutedEventArgs)
        Dim dlg As New Microsoft.Win32.OpenFileDialog
        With dlg
            .Multiselect = True
            .Filter = "Snippet files (*.snippet)|*.snippet|All files|*.*"

            .Title = "Select code snippets"
            If .ShowDialog = True Then

                For Each item In .FileNames
                    Dim sninfo As New SnippetInfo
                    sninfo.SnippetFileName = IO.Path.GetFileName(item)
                    sninfo.SnippetPath = IO.Path.GetDirectoryName(item)
                    sninfo.SnippetLanguage = SnippetInfo.GetSnippetLanguage(item)
                    sninfo.SnippetDescription = SnippetInfo.GetSnippetDescription(item)
                    Me.theData.CodeSnippets.Add(sninfo)
                Next

            Else
                Exit Sub
            End If
        End With
    End Sub

    Private Sub BuildVsixButton_Click(sender As Object, e As Windows.RoutedEventArgs)
        If Me.theData.HasErrors Then
            System.Windows.MessageBox.Show("The metadata information is incomplete. Fix errors before compiling.", "Snippet Package Builder", Windows.MessageBoxButton.OK, Windows.MessageBoxImage.Error)
            Exit Sub
        End If

        If theData.CodeSnippets.Any = False Then
            System.Windows.MessageBox.Show("The code snippet list is empty. Please add at least one before proceding.", "Snippet Package Builder", Windows.MessageBoxButton.OK, Windows.MessageBoxImage.Error)
            Exit Sub
        End If

        Dim testLang = Me.theData.TestLanguageGroup()

        If testLang = False Then
            System.Windows.MessageBox.Show("You have added code snippets of different programming languages. " + Environment.NewLine + "VSIX packages offer the best customer experience possible with snippets of only one language." +
                                           "For this reason, leave snippets of only one language and remove others before building the package.", "Snippet Package Builder", Windows.MessageBoxButton.OK, Windows.MessageBoxImage.Warning)
            Exit Sub
        End If

        Try
            Dim dlg As New Microsoft.Win32.SaveFileDialog
            With dlg
                .Title = "Specify the .vsix name and location"
                .OverwritePrompt = True
                .Filter = "VSIX packages|*.vsix"
                If .ShowDialog = True Then
                    Me.theData.Build(.FileName)
                    Dim result = System.Windows.MessageBox.Show("Package " + IO.Path.GetFileName(.FileName) + " created. Would you like to install the package for testing now?", "Snippet Package Builder", Windows.MessageBoxButton.YesNo, Windows.MessageBoxImage.Question)
                    If result = Windows.MessageBoxResult.No Then
                        Exit Sub
                    Else
                        Process.Start(.FileName)
                        Exit Sub
                    End If
                End If
            End With
        Catch ex As Exception
            System.Windows.MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub GuideButton_Click(sender As Object, e As Windows.RoutedEventArgs) Handles GuideButton.Click
        Dim dlg As New OpenFileDialog
        With dlg
            .Title = "Select documentation file"
            .Filter = "All supported files (*.doc, *.docx, *.rtf, *.txt, *.htm, *.html)|*.doc;*.docx;*.rtf;*.htm;*.html;*.txt|All files|*.*"
            If .ShowDialog = True Then
                Me.theData.GettingStartedGuide = .FileName
            End If
        End With
    End Sub

    Private Sub IconButton_Click(sender As Object, e As Windows.RoutedEventArgs) Handles IconButton.Click
        Dim dlg As New OpenFileDialog
        With dlg
            .Title = "Select icon file"
            .Filter = "All supported files (*.jpg, *.png, *.ico, *.bmp, *.tif, *.tiff, *.gif)|*.jpg;*.png;*.ico;*.bmp;*.tiff;*.tif;*.gif|All files|*.*"
            If .ShowDialog = True Then
                Me.theData.IconPath = .FileName
            End If
        End With
    End Sub

    Private Sub ImageButton_Click(sender As Object, e As Windows.RoutedEventArgs) Handles ImageButton.Click
        Dim dlg As New OpenFileDialog
        With dlg
            .Title = "Select preview image"
            .Filter = "All supported files (*.jpg, *.png, *.ico, *.bmp, *.tif, *.tiff, *.gif)|*.jpg;*.png;*.ico;*.bmp;*.tiff;*.tif;*.gif|All files|*.*"
            If .ShowDialog = True Then
                Me.theData.PreviewImagePath = .FileName
            End If
        End With
    End Sub

    Private Sub LicenseButton_Click(sender As Object, e As Windows.RoutedEventArgs) Handles LicenseButton.Click
        Dim dlg As New OpenFileDialog
        With dlg
            .Title = "Select documentation file"
            .Filter = "All supported files (*.rtf, *.txt)|*.rtf;*.txt|All files|*.*"
            If .ShowDialog = True Then
                Me.theData.License = .FileName
            End If
        End With
    End Sub

    Private Sub RelNotesButton_Click(sender As Object, e As Windows.RoutedEventArgs) Handles RelNotesButton.Click
        Dim dlg As New OpenFileDialog
        With dlg
            .Title = "Select documentation file"
            .Filter = "All supported files (*.rtf, *.txt, *.htm, *.html)|*.rtf;*.htm;*.html;*.txt|All files|*.*"
            If .ShowDialog = True Then
                Me.theData.ReleaseNotes = .FileName
            End If
        End With
    End Sub

    Private Sub RemoveSnippetButton_Click(sender As Object, e As Windows.RoutedEventArgs)
        Try

            For Each item In Me.CodeSnippetsDataGrid.SelectedItems.Cast(Of SnippetInfo).ToList
                Me.theData.CodeSnippets.Remove(TryCast(item, SnippetInfo))
            Next
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ResetPkgButton_Click(sender As Object, e As Windows.RoutedEventArgs)
        ResetPkg()
    End Sub

    Private Sub AboutButton_Click(sender As Object, e As Windows.RoutedEventArgs)
        Process.Start("http://visualstudiogallery.msdn.microsoft.com/de44d368-bab1-43cb-9167-701ac668a09c?SRC=Home")
    End Sub

    Private Sub VsiButton_Click(sender As Object, e As Windows.RoutedEventArgs)
        Dim result = System.Windows.MessageBox.Show("Warning: only .vsi archives containing exclusively code snippets will be converted.
Important: the conversion tool uses the information provided in the Package Metadata tab, so make sure you have
provided the required information.", "Warning", Windows.MessageBoxButton.OKCancel, Windows.MessageBoxImage.Warning)

        If result = Windows.MessageBoxResult.Cancel Then
            Exit Sub
        End If

        Dim dlg As New OpenFileDialog
        Dim inputFile As String, outputFile As String

        With dlg
            .Title = "Select .vsi file"
            .Filter = "Visual Studio 2005 - 2010 .vsi files (*.vsi)|*.vsi|All files|*.*"
            If Not .ShowDialog = True Then
                Exit Sub
            End If
            inputFile = .FileName
        End With

        Dim dlg2 As New SaveFileDialog
        With dlg2
            .OverwritePrompt = True
            .Title = "Output .vsix file"
            .Filter = ".vsix files (*.vsix)|*.vsix|All files|*.*"
            If Not .ShowDialog = True Then
                Exit Sub
            End If
            outputFile = .FileName
        End With

        VSIXPackage.Vsi2Vsix(inputFile, outputFile, theData.SnippetFolderName, theData.PackageAuthor, theData.ProductName, theData.PackageDescription, theData.IconPath, theData.PreviewImagePath, theData.MoreInfoURL)
    End Sub
End Class