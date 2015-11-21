Imports DelSole.VSIX
Imports Microsoft.Win32
Imports Syncfusion.Windows.Edit
Imports Microsoft.WindowsAPICodePack.Dialogs
Imports Syncfusion.UI.Xaml.Grid
Imports DelSole.VSIX.VsiTools, DelSole.VSIX.SnippetTools
Imports System.Reflection
Imports Syncfusion.UI.Xaml.Grid.Helpers

Class MainWindow
    Private theData As VSIXPackage
    Private snippetProperties As CodeSnippet
    Private Property [Imports] As [Imports]
    Private Property References As References
    Private Property Declarations As Declarations

    Private Sub ResetPkg()
        Me.theData = New VSIXPackage

        Me.DataContext = Me.theData
        Me.PackageTab.Focus()
    End Sub

    Private Sub Hyperlink_RequestNavigate(ByVal sender As Object, ByVal e As RequestNavigateEventArgs)
        Process.Start(New ProcessStartInfo(e.Uri.AbsoluteUri))
        e.Handled = True
    End Sub

    Private Sub MyControl_Loaded(sender As Object, e As Windows.RoutedEventArgs) Handles Me.Loaded
        ResetPkg()
        Me.RootTabControl.SelectedIndex = 0
        Me.editControl1.DocumentLanguage = Languages.VisualBasic
        Me.LanguageCombo.SelectedIndex = 0

        'Me.SnippetPropertyGrid.DescriptionPanelVisibility = Visibility.Visible
        Me.Imports = New [Imports]
        Me.Declarations = New Declarations
        Me.References = New References

        Me.ImportsDataGrid.ItemsSource = Me.Imports
        Me.RefDataGrid.ItemsSource = Me.References
        Me.DeclarationsDataGrid.ItemsSource = Me.Declarations
    End Sub


    Private Sub AddSnippetsButton_Click(sender As Object, e As Windows.RoutedEventArgs)
        Dim dlg As New OpenFileDialog
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
            System.Windows.MessageBox.Show("The metadata information is incomplete. Fix errors before compiling.",
                                           "Code Snippet Studio",
                                           Windows.MessageBoxButton.OK,
                                           Windows.MessageBoxImage.Error)
            Exit Sub
        End If

        If theData.CodeSnippets.Any = False Then
            MessageBox.Show("The code snippet list is empty. Please add at least one before proceding.",
                            "Code Snippet Studio",
                            Windows.MessageBoxButton.OK,
                            Windows.MessageBoxImage.Error)
            Exit Sub
        End If

        Dim testLang = Me.theData.TestLanguageGroup()

        If testLang = False Then
            System.Windows.MessageBox.Show("You have added code snippets of different programming languages. " + Environment.NewLine + "VSIX packages offer the best customer experience possible with snippets of only one language." +
                                           "For this reason, leave snippets of only one language and remove others before building the package.", "Code Snippet Studio", Windows.MessageBoxButton.OK, Windows.MessageBoxImage.Warning)
            Exit Sub
        End If

        Try
            Dim dlg As New SaveFileDialog
            With dlg
                .Title = "Specify the .vsix name and location"
                .OverwritePrompt = True
                .Filter = "VSIX packages|*.vsix"
                If .ShowDialog = True Then
                    Me.theData.Build(.FileName)
                    Dim result = MessageBox.Show("Package " + IO.Path.GetFileName(.FileName) + " created. Would you like to install the package for testing now?", "Code Snippet Studio", Windows.MessageBoxButton.YesNo, Windows.MessageBoxImage.Question)
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
        Process.Start("https://github.com/AlessandroDelSole/CodeSnippetStudio")
    End Sub

    Private Sub VsiButton_Click(sender As Object, e As Windows.RoutedEventArgs)
        If theData.HasErrors Then
            MessageBox.Show("The package metadata contain errors that must be fixed before performing a conversion." & Environment.NewLine &
                            "Please go to the Package and Share tab and check values in the Metadata nested tab.")
            Exit Sub
        End If

        Dim dlg As New OpenFileDialog
        Dim inputFile As String, outputFile As String

        With dlg
            .Title = "Select .vsi file"
            .Filter = "Visual Studio Community Installer (*.vsi)|*.vsi|All files|*.*"
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

        VsiService.Vsi2Vsix(inputFile, outputFile, theData.SnippetFolderName,
                             theData.PackageAuthor, theData.ProductName,
                             theData.PackageDescription,
                             theData.IconPath, theData.PreviewImagePath,
                             theData.MoreInfoURL)
        MessageBox.Show($"Successfully converted {inputFile} into {outputFile}")
    End Sub

    Private Sub SignButton_Click(sender As Object, e As RoutedEventArgs)
        If PfxPassword.Password.Length = 0 Then
            MessageBox.Show("Please enter the password for the certificate file.")
            Exit Sub
        End If

        If PfxTextBox.Text = "" Or String.IsNullOrEmpty(PfxTextBox.Text) Then
            MessageBox.Show("Please specify a valid X.509 certificate file.")
            Exit Sub
        End If

        If Not IO.File.Exists(PfxTextBox.Text) Then
            MessageBox.Show("The specified certificate file does not exist.")
            Exit Sub
        End If

        Dim dlg As New OpenFileDialog
        With dlg
            .Title = "Select the .vsix you want to sign"
            .Filter = "VSIX packages (*.vsix)|*.pfx|All files|*.*"
            If .ShowDialog = True Then
                VSIXPackage.SignVsix(.FileName, PfxTextBox.Text, PfxPassword.Password)
                MessageBox.Show($"{ .FileName} signed successfully.")
            End If
        End With
    End Sub

    Private Sub PfxButton_Click(sender As Object, e As RoutedEventArgs)
        Dim dlg As New OpenFileDialog
        With dlg
            .Title = "Select certificate file"
            .Filter = "Certificate files (*.pfx)|*.pfx|All files|*.*"
            If .ShowDialog = True Then
                Me.PfxTextBox.Text = .FileName
            End If
        End With
    End Sub

    Private Sub LanguageCombo_SelectionChanged(sender As Object, e As System.Windows.Controls.SelectionChangedEventArgs) Handles LanguageCombo.SelectionChanged
        Dim cb = CType(sender, ComboBox)
        Select Case cb.SelectedIndex
            Case = 0
                Me.editControl1.DocumentLanguage = Languages.VisualBasic
                EnableDataGrids()
            Case = 1
                Me.editControl1.DocumentLanguage = Languages.CSharp
                DisableDataGrids()
            Case = 2
                Me.editControl1.DocumentLanguage = Languages.SQL
                DisableDataGrids()
            Case = 3
                Me.editControl1.DocumentLanguage = Languages.XML
                DisableDataGrids()
            Case = 4
                Me.editControl1.DocumentLanguage = Languages.Text
                DisableDataGrids()
            Case = 5
                Me.editControl1.DocumentLanguage = Languages.XML
                DisableDataGrids()
            Case = 6
                Me.editControl1.DocumentLanguage = Languages.XML
                DisableDataGrids()
        End Select
    End Sub

    Private Sub SaveSnippetButton_Click(sender As Object, e As RoutedEventArgs)
        Me.snippetProperties = CType(Me.SnippetPropertyGrid.SelectedObject, CodeSnippet)

        If snippetProperties.Author = "" Or String.IsNullOrEmpty(snippetProperties.Author) Then
            MessageBox.Show("Snippet author is missing.", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            Exit Sub
        End If

        If snippetProperties.Title = "" Or String.IsNullOrEmpty(snippetProperties.Title) Then
            MessageBox.Show("Snippet title is missing.", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            Exit Sub
        End If

        If snippetProperties.Description = "" Or String.IsNullOrEmpty(snippetProperties.Description) Then
            MessageBox.Show("Snippet description is missing.", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            Exit Sub
        End If

        Dim dlg2 As New SaveFileDialog
        With dlg2
            .OverwritePrompt = True
            .Title = "Output .snippet file"
            .Filter = ".snippet files (*.snippet)|*.snippet|All files|*.*"
            If Not .ShowDialog = True Then
                Exit Sub
            End If

            SaveSnippet(.FileName)
        End With
    End Sub

    Private Sub SaveSnippet(fileName As String)
        Dim currentLang = CType(LanguageCombo.SelectedItem, ComboBoxItem).Tag.ToString()
        Dim selectedSnippet = CType(SnippetPropertyGrid.SelectedObject, CodeSnippet)

        Dim keywords As IEnumerable(Of String) = selectedSnippet?.Keywords?.Split(","c).AsEnumerable

        SnippetService.SaveSnippet(fileName, selectedSnippet.Kind, currentLang, selectedSnippet.Title,
                       selectedSnippet.Description, selectedSnippet.HelpUrl,
                       selectedSnippet.Author, selectedSnippet.Shortcut, editControl1.Text,
                       Me.Imports, References, Declarations, keywords)

        MessageBox.Show($"{fileName} saved correctly.")

    End Sub

    Private Sub ExtractButton_Click(sender As Object, e As RoutedEventArgs)
        Dim dlg As New OpenFileDialog
        Dim inputFile As String, outputFolder As String

        With dlg
            .Title = "Select .vsix file"
            .Filter = "Visual Studio Extension (*.vsix)|*.vsix|All files|*.*"
            If Not .ShowDialog = True Then
                Exit Sub
            End If
            inputFile = .FileName
        End With

        Dim dlg2 As New CommonOpenFileDialog()
        dlg2.Title = "Destination folder"
        dlg2.IsFolderPicker = True
        dlg2.InitialDirectory = Environment.SpecialFolder.MyDocuments
        dlg2.DefaultDirectory = Environment.SpecialFolder.MyDocuments
        dlg2.EnsureFileExists = True
        dlg2.EnsurePathExists = True
        dlg2.EnsureValidNames = True
        dlg2.Multiselect = False
        dlg2.ShowPlacesList = True

        If Not dlg2.ShowDialog = CommonFileDialogResult.Ok Then
            Exit Sub
        End If

        outputFolder = dlg2.FileName

        VSIXPackage.ExtractVsix(inputFile, outputFolder, OnlySnippetsCheckBox.IsChecked)
        MessageBox.Show($"Successfully extracted {inputFile} into {outputFolder}")
    End Sub

    Private Sub DisableDataGrids()
        Me.ImportsDataGrid.IsEnabled = False
        Me.RefDataGrid.IsEnabled = False
    End Sub

    Private Sub EnableDataGrids()
        Me.ImportsDataGrid.IsEnabled = True
        Me.RefDataGrid.IsEnabled = True
    End Sub

    Private Shared Function ReturnSnippetKind(kind As SnippetTools.CodeSnippetKinds) As String
        Dim snippetKind As String
        Select Case kind
            Case SnippetTools.CodeSnippetKinds.MethodBody
                snippetKind = "method body"
            Case SnippetTools.CodeSnippetKinds.MethodDeclaration
                snippetKind = "method decl"
            Case SnippetTools.CodeSnippetKinds.File
                snippetKind = "file"
            Case SnippetTools.CodeSnippetKinds.TypeDeclaration
                snippetKind = "type decl"
            Case Else
                snippetKind = "any"
        End Select

        Return snippetKind
    End Function

    Private Sub DeclarationsDataGrid_SelectionChanged(sender As Object, e As GridSelectionChangedEventArgs) Handles DeclarationsDataGrid.SelectionChanged
        Try
            Dim item = CType(CType(sender, SfDataGrid).SelectedItem, Declaration)
            Dim fo As New FindOptions(editControl1)
            fo.FindText = item.Default
            'fo.IsSelectionSelected = True
            If editControl1.DocumentLanguage = Languages.VisualBasic Then
                fo.IsMatchCase = False
            Else
                fo.IsMatchCase = True
            End If
            editControl1.FindAllOccurences()
            For Each result In editControl1.SearchResults.FindAllResult
                editControl1.SelectLines(result.LineNumber, result.LineNumber, result.Index, result.Index)
            Next
            Dim res = editControl1.SearchResults
        Catch ex As Exception

        End Try

    End Sub

    Private Sub SaveCodeFileButton_Click(sender As Object, e As RoutedEventArgs)
        If editControl1.Text = "" Then
            MessageBox.Show("Write some code first!")
            Exit Sub
        End If

        Dim filter As String = "All files|*.*"
        Select Case LanguageCombo.SelectedIndex
            Case = 0
                filter = "Visual Basic code file (.vb)|*.vb|All files|*.*"
            Case = 1
                filter = "C# code file (.cs)|*.cs|All files|*.*"
            Case = 2
                filter = "SQL code file (.sql)|*.sql|All files|*.*"
            Case = 3
                filter = "XML file (.xml)|*.xml|All files|*.*"
            Case = 4
                filter = "C++ code file (.cpp)|*.cpp|All files|*.*"
            Case = 5
                filter = "HTML file (.htm)|*.htm|All files|*.*"
            Case = 6
                filter = "JavaScript code file (.js)|*.js|All files|*.*"
        End Select

        Dim dlg As New SaveFileDialog
        dlg.Title = "Specify code file name"
        dlg.Filter = filter
        dlg.OverwritePrompt = True

        If dlg.ShowDialog = True Then
            My.Computer.FileSystem.WriteAllText(dlg.FileName, editControl1.Text, False)
        End If
    End Sub

    Private Sub AddDecButton_Click(sender As Object, e As RoutedEventArgs)
        If editControl1.SelectedText = "" Then Exit Sub

        If editControl1.SelectedText.ToLower = "end" Or editControl1.SelectedText.ToLower = "select" Then
            MessageBox.Show("Declarations are not supported for Select and End words.", "Code Snippet Studio", MessageBoxButton.OK, MessageBoxImage.Error)
            Exit Sub
        End If

        Dim newDecl As New Declaration
        newDecl.Default = editControl1.SelectedText
        newDecl.ID = editControl1.SelectedText
        newDecl.ToolTip = "Replace with yours...."

        Dim query = From decl In Me.Declarations
                    Where decl.Default = newDecl.Default
                    Select decl

        If query.Any Then
            MessageBox.Show("A declaration already exists for the specified word", "Code Snippet Studio", MessageBoxButton.OK, MessageBoxImage.Error)
            Exit Sub
        End If

        Me.Declarations.Add(newDecl)
    End Sub

    Private Sub DeleteDecButton_Click(sender As Object, e As RoutedEventArgs) Handles DeleteDecButton.Click
        If Me.DeclarationsDataGrid.SelectedItem Is Nothing Then
            Exit Sub
        End If
        Me.DeclarationsDataGrid.View.Remove(Me.DeclarationsDataGrid.SelectedItem)
    End Sub
End Class

