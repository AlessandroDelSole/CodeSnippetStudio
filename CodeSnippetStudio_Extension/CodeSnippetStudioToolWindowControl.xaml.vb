Imports DelSole.VSIX
Imports Microsoft.Win32
Imports Syncfusion.Windows.Edit
Imports <xmlns="http://schemas.microsoft.com/VisualStudio/2005/CodeSnippet">
Imports Syncfusion.UI.Xaml.Grid
Imports System.ComponentModel
Imports DelSole.VSIX.VsiTools, DelSole.VSIX.SnippetTools
Imports System.Diagnostics
Imports EnvDTE
Imports System.Windows.Navigation
Imports System.Linq
Imports System
Imports System.Windows, System.Windows.Controls, Syncfusion.SfSkinManager
Imports System.Xml.Linq
Imports System.Collections.Generic
Imports Newtonsoft.Json
Imports System.Collections.ObjectModel
Imports System.Windows.Input
Imports Microsoft.CodeAnalysis
Imports System.Windows.Media

'''<summary>
''' Interaction logic for CodeSnippetStudioToolWindowControl.xaml
'''</summary>
Partial Public Class CodeSnippetStudioToolWindowControl
    Inherits System.Windows.Controls.UserControl

    Private Property vsixData As VsixPackage
    Private Property snippetData As CodeSnippet
    Private Property IntelliSenseReferences As ObservableCollection(Of Uri)
    Private Property snippetLib As SnippetLibrary

    Private LibraryName As String = IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "CodeSnippetStudioLibrary.xml")
    Private selectionBackground As Brush

    Private Sub ResetPkg()
        Me.vsixData = New VsixPackage

        Me.VsixGrid.DataContext = Me.vsixData
        Me.PackageTab.Focus()
    End Sub

    Private Sub Hyperlink_RequestNavigate(ByVal sender As Object, ByVal e As RequestNavigateEventArgs)
        System.Diagnostics.Process.Start(New ProcessStartInfo(e.Uri.AbsoluteUri))
        e.Handled = True
    End Sub

    Private Sub HidePropertiesFromPropertyGrid()
        'Properties that must be hidden from the PropertyGrid
        Me.SnippetPropertyGrid.HidePropertiesCollection.Add("Namespaces")
        Me.SnippetPropertyGrid.HidePropertiesCollection.Add("Declarations")
        Me.SnippetPropertyGrid.HidePropertiesCollection.Add("References")
        Me.SnippetPropertyGrid.HidePropertiesCollection.Add("Language")
        Me.SnippetPropertyGrid.HidePropertiesCollection.Add("Code")
        Me.SnippetPropertyGrid.HidePropertiesCollection.Add("Error")
        Me.SnippetPropertyGrid.HidePropertiesCollection.Add("HasErrors")
        Me.SnippetPropertyGrid.HidePropertiesCollection.Add("IsDirty")
        Me.SnippetPropertyGrid.HidePropertiesCollection.Add("FileName")
        Me.SnippetPropertyGrid.HidePropertiesCollection.Add("Diagnostics")
        SnippetPropertyGrid.RefreshPropertygrid()
    End Sub

    Private Sub LoadSnippetLibrary()
        Me.snippetLib = New SnippetLibrary
        Me.LibraryTreeview.ItemsSource = snippetLib.Folders

        My.Settings.LibraryName = LibraryName

        Me.EditorRoot.DataContext = Me.snippetData
        Try
            snippetLib.LoadLibrary(My.Settings.LibraryName)
        Catch ex As Exception
            'error loading library, ignore
        End Try
    End Sub

    Private Sub EditorSetup()
        'My.Settings.Reset()
        Me.RootTabControl.SelectedIndex = 0
        Me.editControl1.DocumentLanguage = LoadPreferredLanguage()

        If My.Settings.EditorForeColor IsNot Nothing Then
            editControl1.Foreground = My.Settings.EditorForeColor
            EditorForeColorPicker.Brush = My.Settings.EditorForeColor
        Else
            EditorForeColorPicker.Brush = editControl1.Foreground
        End If

        If My.Settings.EditorSelectionColor IsNot Nothing Then
            editControl1.SelectionForeground = My.Settings.EditorSelectionColor
            EditorSelectionColorPicker.Brush = My.Settings.EditorSelectionColor
        Else
            EditorSelectionColorPicker.Brush = editControl1.SelectionForeground
        End If
        Try
            SfSkinManager.SetVisualStyle(Me, My.Settings.PreferredTheme)
            SetPreferredThemeOption(My.Settings.PreferredTheme)
        Catch ex As Exception

        End Try

        Me.selectionBackground = editControl1.SelectionBackground

        Me.IntelliSenseReferences = New ObservableCollection(Of Uri)
        Me.editControl1.AssemblyReferences = IntelliSenseReferences
    End Sub

    Private Sub MyControl_Loaded(sender As Object, e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        If Me.snippetData Is Nothing Then
            GenerateNewSnippet()

            LoadSnippetLibrary()
            HidePropertiesFromPropertyGrid()

            ResetPkg()

            EditorSetup()

            Me.ImportsDataGrid.ItemsSource = snippetData.Namespaces
            Me.RefDataGrid.ItemsSource = snippetData.References
            Me.DeclarationsDataGrid.ItemsSource = snippetData.Declarations

            SfSkinManager.SetVisualStyle(Me, My.Settings.PreferredTheme)
        End If
    End Sub


    Private Sub AddSnippetsButton_Click(sender As Object, e As System.Windows.RoutedEventArgs)
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
                    Me.vsixData.CodeSnippets.Add(sninfo)
                Next
                Me.VSVsixTabControl.SelectedIndex = 1
            Else
                Exit Sub
            End If
        End With
    End Sub

    Private Sub BuildVsixButton_Click(sender As Object, e As System.Windows.RoutedEventArgs)
        If Me.vsixData.HasErrors Then
            System.Windows.MessageBox.Show("The metadata information is incomplete. Fix errors before compiling.",
                                           "Code Snippet Studio",
                                           System.Windows.MessageBoxButton.OK,
                                           System.Windows.MessageBoxImage.Error)
            Exit Sub
        End If

        If vsixData.CodeSnippets.Any = False Then
            MessageBox.Show("The code snippet list is empty. Please add at least one before proceding.",
                            "Code Snippet Studio",
                            System.Windows.MessageBoxButton.OK,
                            System.Windows.MessageBoxImage.Error)
            Exit Sub
        End If

        Dim testLang = Me.vsixData.TestLanguageGroup()

        If testLang = False Then
            System.Windows.MessageBox.Show("You have added code snippets of different programming languages. " + Environment.NewLine + "VSIX packages offer the best customer experience possible with snippets of only one language." +
                                           "For this reason, leave snippets of only one language and remove others before building the package.", "Code Snippet Studio", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning)
            Exit Sub
        End If

        Try
            Dim dlg As New SaveFileDialog
            With dlg
                .Title = "Specify the .vsix name and location"
                .OverwritePrompt = True
                .Filter = "VSIX packages|*.vsix"
                If .ShowDialog = True Then
                    Me.vsixData.Build(.FileName, IDEType.VisualStudio)
                    Dim result = MessageBox.Show("Package " + IO.Path.GetFileName(.FileName) + " created. Would you like to install the package for testing now?", "Code Snippet Studio", System.Windows.MessageBoxButton.YesNo, System.Windows.MessageBoxImage.Question)
                    If result = System.Windows.MessageBoxResult.No Then
                        Exit Sub
                    Else
                        System.Diagnostics.Process.Start(.FileName)
                        Exit Sub
                    End If
                End If
            End With
        Catch ex As Exception
            System.Windows.MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub GuideButton_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles GuideButton.Click
        Dim dlg As New OpenFileDialog
        With dlg
            .Title = "Select documentation file"
            .Filter = "All supported files (*.doc, *.docx, *.rtf, *.txt, *.htm, *.html)|*.doc;*.docx;*.rtf;*.htm;*.html;*.txt|All files|*.*"
            If .ShowDialog = True Then
                Me.vsixData.GettingStartedGuide = .FileName
            End If
        End With
    End Sub

    Private Sub IconButton_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles IconButton.Click
        Dim dlg As New OpenFileDialog
        With dlg
            .Title = "Select icon file"
            .Filter = "All supported files (*.jpg, *.png, *.ico, *.bmp, *.tif, *.tiff, *.gif)|*.jpg;*.png;*.ico;*.bmp;*.tiff;*.tif;*.gif|All files|*.*"
            If .ShowDialog = True Then
                Me.vsixData.IconPath = .FileName
            End If
        End With
    End Sub

    Private Sub ImageButton_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles ImageButton.Click
        Dim dlg As New OpenFileDialog
        With dlg
            .Title = "Select preview image"
            .Filter = "All supported files (*.jpg, *.png, *.ico, *.bmp, *.tif, *.tiff, *.gif)|*.jpg;*.png;*.ico;*.bmp;*.tiff;*.tif;*.gif|All files|*.*"
            If .ShowDialog = True Then
                Me.vsixData.PreviewImagePath = .FileName
            End If
        End With
    End Sub

    Private Sub LicenseButton_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles LicenseButton.Click
        Dim dlg As New OpenFileDialog
        With dlg
            .Title = "Select documentation file"
            .Filter = "All supported files (*.rtf, *.txt)|*.rtf;*.txt|All files|*.*"
            If .ShowDialog = True Then
                Me.vsixData.License = .FileName
            End If
        End With
    End Sub

    Private Sub RelNotesButton_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles RelNotesButton.Click
        Dim dlg As New OpenFileDialog
        With dlg
            .Title = "Select documentation file"
            .Filter = "All supported files (*.rtf, *.txt, *.htm, *.html)|*.rtf;*.htm;*.html;*.txt|All files|*.*"
            If .ShowDialog = True Then
                Me.vsixData.ReleaseNotes = .FileName
            End If
        End With
    End Sub

    Private Sub RemoveSnippetButton_Click(sender As Object, e As System.Windows.RoutedEventArgs)
        Try
            Dim selectedList = Me.CodeSnippetsDataGrid.SelectedItems.Cast(Of SnippetInfo).ToList
            If Not selectedList.Any Then Exit Sub

            Me.VSVsixTabControl.SelectedIndex = 1

            Dim result = MessageBox.Show("Are you sure you want to remove the selected snippet(s)?", "Confirmation", MessageBoxButton.YesNo, MessageBoxImage.Question)
            If result = MessageBoxResult.No Then Exit Sub

            For Each item In selectedList
                Me.vsixData.CodeSnippets.Remove(TryCast(item, SnippetInfo))
            Next
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ResetPkgButton_Click(sender As Object, e As System.Windows.RoutedEventArgs)
        Dim result = MessageBox.Show("Are you sure?", "Confirmation", MessageBoxButton.YesNo, MessageBoxImage.Warning)
        If result = MessageBoxResult.Yes Then ResetPkg()
    End Sub

    Private Sub AboutButton_Click(sender As Object, e As System.Windows.RoutedEventArgs)
        System.Diagnostics.Process.Start("https://github.com/AlessandroDelSole/CodeSnippetStudio")
    End Sub

    Private Sub VsiButton_Click(sender As Object, e As System.Windows.RoutedEventArgs)
        If vsixData.HasErrors Then
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

        VsiService.Vsi2Vsix(inputFile, outputFile, vsixData.SnippetFolderName,
                             vsixData.PackageAuthor, vsixData.ProductName,
                             vsixData.PackageDescription,
                             vsixData.IconPath, vsixData.PreviewImagePath,
                             vsixData.MoreInfoURL)
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
            .Filter = "VSIX packages (*.vsix)|*.vsix|All files|*.*"
            If .ShowDialog = True Then
                VsixPackage.SignVsix(.FileName, PfxTextBox.Text, PfxPassword.Password)
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
                Me.editControl1.DocumentLanguage = Syncfusion.Windows.Edit.Languages.VisualBasic
                Me.snippetData.Language = "VB"
                EnableDataGrids()
                AnalyzeCode()
            Case = 1
                Me.editControl1.DocumentLanguage = Syncfusion.Windows.Edit.Languages.CSharp
                Me.snippetData.Language = "CSharp"
                DisableDataGrids()
                AnalyzeCode()
            Case = 2
                Me.editControl1.DocumentLanguage = Syncfusion.Windows.Edit.Languages.SQL
                Me.snippetData.Language = "SQL"
                DisableDataGrids()
            Case = 3
                Me.editControl1.DocumentLanguage = Syncfusion.Windows.Edit.Languages.XML
                Me.snippetData.Language = "XML"
                DisableDataGrids()
            Case = 4
                Me.editControl1.DocumentLanguage = Syncfusion.Windows.Edit.Languages.XAML
                Me.snippetData.Language = "XAML"
                DisableDataGrids()
            Case = 5
                Me.editControl1.DocumentLanguage = Syncfusion.Windows.Edit.Languages.Text
                Me.snippetData.Language = "CPP"
                DisableDataGrids()
            Case = 6
                Me.editControl1.DocumentLanguage = Syncfusion.Windows.Edit.Languages.XML
                Me.snippetData.Language = "HTML"
                DisableDataGrids()
            Case = 7
                Me.editControl1.DocumentLanguage = Syncfusion.Windows.Edit.Languages.XML
                Me.snippetData.Language = "JavaScript"
                DisableDataGrids()
        End Select
    End Sub

    Private Sub SaveSnippetButton_Click(sender As Object, e As RoutedEventArgs)
        If snippetData.HasErrors Then
            MessageBox.Show("The current code snippet has errors that must be fixed before saving." _
                            & Environment.NewLine &
                            "Ensure that Author, Title, Description, and snippet language have been supplied properly.",
                            "Error", MessageBoxButton.OK, MessageBoxImage.Error)
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

            If snippetData.Language = "" Or String.IsNullOrEmpty(snippetData.Language) Then
                snippetData.Language = My.Settings.PreferredLanguage
            End If

            snippetData.SaveSnippet(.FileName, IDEType.VisualStudio)
            editControl1.SetValue(Syncfusion.Windows.Tools.Controls.DockingManager.HeaderProperty, .FileName)
            MessageBox.Show($"{ .FileName} saved correctly.")
        End With
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

        Dim dlg2 As New System.Windows.Forms.FolderBrowserDialog
        dlg2.Description = "Select destination folder"
        dlg2.ShowNewFolderButton = True

        If Not dlg2.ShowDialog = Forms.DialogResult.OK Then
            Exit Sub
        End If

        outputFolder = dlg2.SelectedPath

        VsixPackage.ExtractVsix(inputFile, outputFolder, OnlySnippetsCheckBox.IsChecked)
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

    Private Sub DeclarationsDataGrid_SelectionChanged(sender As Object, e As GridSelectionChangedEventArgs) Handles DeclarationsDataGrid.SelectionChanged
        Try
            Dim item = CType(CType(sender, SfDataGrid).SelectedItem, Declaration)
            Dim fo As New FindOptions(editControl1)
            fo.FindText = item.Default
            'fo.IsSelectionSelected = True
            If editControl1.DocumentLanguage = Syncfusion.Windows.Edit.Languages.VisualBasic Then
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
                filter = "XAML file (.xaml)|*.xaml|All files|*.*"
            Case = 5
                filter = "C++ code file (.cpp)|*.cpp|All files|*.*"
            Case = 6
                filter = "HTML file (.htm)|*.htm|All files|*.*"
            Case = 7
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
        DockingManager.ActivateWindow("DeclarationsDataGrid")
        Dim newDecl As New Declaration
        newDecl.Default = editControl1.SelectedText

        newDecl.ID = editControl1.SelectedText
        newDecl.ToolTip = "Replace with yours...."

        Dim query = From decl In snippetData.Declarations
                    Where decl.Default = newDecl.Default
                    Select decl

        If query.Any Then
            MessageBox.Show("A declaration already exists for the specified word",
                            "Code Snippet Studio", MessageBoxButton.OK, MessageBoxImage.Error)
            Exit Sub
        End If

        snippetData.Declarations.Add(newDecl)
    End Sub

    Private Sub DeleteDecButton_Click(sender As Object, e As RoutedEventArgs) Handles DeleteDecButton.Click
        If Me.DeclarationsDataGrid.SelectedItem Is Nothing Then
            Exit Sub
        End If
        Me.DeclarationsDataGrid.View.Remove(Me.DeclarationsDataGrid.SelectedItem)
    End Sub

    Private Sub OpenVsixButton_Click(sender As Object, e As RoutedEventArgs)
        Dim dlg As New OpenFileDialog

        With dlg
            .Title = "Select .vsix file"
            .Filter = "Visual Studio Extension (*.vsix)|*.vsix|All files|*.*"
            If Not .ShowDialog = True Then
                Exit Sub
            End If
            vsixData = VsixPackage.OpenVsix(.FileName)
            Me.VsixGrid.DataContext = vsixData
        End With
    End Sub

    Private Sub editControl1_TextChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs) Handles editControl1.TextChanged
        snippetData.Code = editControl1.Text
        AnalyzeCode()
    End Sub

    Private Sub LoadCodeFileButton_Click(sender As Object, e As RoutedEventArgs)
        If snippetData.IsDirty Then
            Dim result = MessageBox.Show("The current snippet has unsaved changes. Are you sure?", "Confirmation",
                                         MessageBoxButton.YesNo, MessageBoxImage.Question)
            If result = MessageBoxResult.No Then
                Exit Sub
            End If
        End If

        Dim dlg As New OpenFileDialog

        With dlg
            .Title = "Select code snippet file"
            .Filter = "Snippet files (*.snippet)|*.snippet;*.vbsnippet;*.vssnippet|Json snippets for VS Code|*.json|All files|*.*"
            If Not .ShowDialog = True Then
                Exit Sub
            End If
            Try
                Dim tempData = CodeSnippet.LoadSnippet(.FileName)
                If tempData IsNot Nothing Then
                    Me.snippetData = Nothing
                    Me.snippetData = tempData
                    Me.EditorRoot.DataContext = Me.snippetData
                    Me.snippetData.IsDirty = False
                    editControl1.SetValue(Syncfusion.Windows.Tools.Controls.DockingManager.HeaderProperty, .FileName)
                    If Not IO.Path.GetExtension(.FileName).ToLower = "json" Then SetCurrentLanguage(snippetData.Language)
                End If
            Catch ex As JsonReaderException
                MessageBox.Show("The .json snippet file is invalid", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
                Exit Sub
            Catch ex As Exception
                MessageBox.Show(ex.Message)
                Exit Sub
            End Try
        End With
    End Sub

    Private Sub ImportVsiButton_Click(sender As Object, e As RoutedEventArgs)
        Dim dlg As New OpenFileDialog

        With dlg
            .Title = "Select .vsi file"
            .Filter = "Visual Studio Content Installer (*.vsi)|*.vsi|All files|*.*"
            If Not .ShowDialog = True Then
                Exit Sub
            End If
            vsixData = VsiService.Vsi2Vsix(.FileName)
            Me.VsixGrid.DataContext = vsixData
        End With
    End Sub

    Private Sub SaveVSCodeSnippetButton_Click(sender As Object, e As RoutedEventArgs)
        If snippetData.HasErrors Then
            MessageBox.Show("The current code snippet has errors that must be fixed before saving." _
                            & Environment.NewLine &
                            "Ensure that Author, Title, Description, and snippet language have been supplied properly.",
                            "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            Exit Sub
        End If

        Dim dlg2 As New SaveFileDialog
        With dlg2
            .OverwritePrompt = True
            .Title = "Output .json file"
            .Filter = ".json files (*.json)|*.json|All files|*.*"
            If Not .ShowDialog = True Then
                Exit Sub
            End If

            snippetData.SaveSnippet(.FileName, IDEType.Code)
            editControl1.SetValue(Syncfusion.Windows.Tools.Controls.DockingManager.HeaderProperty, .FileName)
            MessageBox.Show($"{ .FileName} saved correctly. Please visit: " &
                            Environment.NewLine & "https://code.visualstudio.com/docs/customization/userdefinedsnippets" & Environment.NewLine &
                            "to learn how to consume custom snippets in Visual Studio Code", "Save info", MessageBoxButton.OK, MessageBoxImage.Information)
        End With
    End Sub

    Private Sub HelpButton_Click(sender As Object, e As RoutedEventArgs)
        System.Diagnostics.Process.Start("https://github.com/AlessandroDelSole/CodeSnippetStudio/blob/master/CodeSnippetStudio_StandAlone/Assets/Code_Snippet_Studio_User_Guide.pdf")
    End Sub

    Private Sub SaveSettings()
        My.Settings.PreferredLanguage = PrefLanguageCombo.SelectedItem.ToString
    End Sub

    Private Sub PrefLanguageCombo_SelectionChanged(sender As Object, e As System.Windows.Controls.SelectionChangedEventArgs)
        Dim cb = CType(sender, ComboBox)
        Select Case cb.SelectedIndex
            Case = 0
                My.Settings.PreferredLanguage = "VB"
                My.Settings.Save()
            Case = 1
                My.Settings.PreferredLanguage = "CSharp"
                My.Settings.Save()
            Case = 2
                My.Settings.PreferredLanguage = "SQL"
                My.Settings.Save()
            Case = 3
                My.Settings.PreferredLanguage = "XML"
                My.Settings.Save()
            Case = 4
                My.Settings.PreferredLanguage = "XAML"
                My.Settings.Save()
            Case = 5
                My.Settings.PreferredLanguage = "CPP"
                My.Settings.Save()
            Case = 6
                My.Settings.PreferredLanguage = "HTML"
                My.Settings.Save()
            Case = 7
                My.Settings.PreferredLanguage = "JavaScript"
                My.Settings.Save()
        End Select
    End Sub

    Private Sub SetPreferredThemeOption(style As VisualStyles)
        Select Case style
            Case VisualStyles.Metro
                ThemeCombo.SelectedIndex = 0
            Case VisualStyles.Blend
                ThemeCombo.SelectedIndex = 1
            Case VisualStyles.VisualStudio2015
                ThemeCombo.SelectedIndex = 2
            Case VisualStyles.Office2016Colorful
                ThemeCombo.SelectedIndex = 0
        End Select

    End Sub

    Private Sub ThemeCombo_SelectionChanged(sender As Object, e As System.Windows.Controls.SelectionChangedEventArgs)
        Dim cb = CType(sender, ComboBox)
        Select Case cb.SelectedIndex
            Case = 0
                SfSkinManager.SetVisualStyle(Me, VisualStyles.Metro)
                My.Settings.PreferredTheme = VisualStyles.Metro
                My.Settings.Save()
            Case = 1
                SfSkinManager.SetVisualStyle(Me, VisualStyles.Blend)
                My.Settings.PreferredTheme = VisualStyles.Blend
                My.Settings.Save()
            Case = 2
                SfSkinManager.SetVisualStyle(Me, VisualStyles.VisualStudio2015)
                My.Settings.PreferredTheme = VisualStyles.VisualStudio2015
                My.Settings.Save()
            Case = 3
                SfSkinManager.SetVisualStyle(Me, VisualStyles.Office2016Colorful)
                My.Settings.PreferredTheme = VisualStyles.Office2016Colorful
                My.Settings.Save()
        End Select
    End Sub

    Private Sub SetCurrentLanguage(snippetLanguage As String)
        Select Case snippetLanguage.ToUpper
            Case = "VB"
                Me.LanguageCombo.SelectedIndex = 0
            Case = "CSHARP"
                Me.LanguageCombo.SelectedIndex = 1
            Case = "SQL"
                Me.LanguageCombo.SelectedIndex = 2
            Case = "XML"
                Me.LanguageCombo.SelectedIndex = 3
            Case = "XAML"
                Me.LanguageCombo.SelectedIndex = 4
            Case = "CPP"
                Me.LanguageCombo.SelectedIndex = 5
            Case = "HTML"
                Me.LanguageCombo.SelectedIndex = 6
            Case = "JAVASCRIPT"
                Me.LanguageCombo.SelectedIndex = 7
            Case Else
                Me.LanguageCombo.SelectedIndex = 7
        End Select
    End Sub

    Private Function LoadPreferredLanguage() As Syncfusion.Windows.Edit.Languages
        Select Case My.Settings.PreferredLanguage
            Case = "VB"
                Me.LanguageCombo.SelectedIndex = 0
                Me.PrefLanguageCombo.SelectedIndex = 0
                Return Syncfusion.Windows.Edit.Languages.VisualBasic
            Case = "CSharp"
                Me.LanguageCombo.SelectedIndex = 1
                Me.PrefLanguageCombo.SelectedIndex = 1
                Return Syncfusion.Windows.Edit.Languages.CSharp
            Case = "SQL"
                Me.LanguageCombo.SelectedIndex = 2
                Me.PrefLanguageCombo.SelectedIndex = 2
                Return Syncfusion.Windows.Edit.Languages.SQL
            Case = "XML"
                Me.LanguageCombo.SelectedIndex = 3
                Me.PrefLanguageCombo.SelectedIndex = 3
                Return Syncfusion.Windows.Edit.Languages.XML
            Case = "XAML"
                Me.LanguageCombo.SelectedIndex = 4
                Me.PrefLanguageCombo.SelectedIndex = 4
                Return Syncfusion.Windows.Edit.Languages.XAML
            Case = "CPP"
                Me.LanguageCombo.SelectedIndex = 5
                Me.PrefLanguageCombo.SelectedIndex = 5
                Return Syncfusion.Windows.Edit.Languages.CSharp
            Case = "HTML"
                Me.LanguageCombo.SelectedIndex = 6
                Me.PrefLanguageCombo.SelectedIndex = 6
                Return Syncfusion.Windows.Edit.Languages.XML
            Case = "JavaScript"
                Me.LanguageCombo.SelectedIndex = 7
                Me.PrefLanguageCombo.SelectedIndex = 7
                Return Syncfusion.Windows.Edit.Languages.XML
            Case Else
                Return Syncfusion.Windows.Edit.Languages.Text
        End Select
    End Function

    Private Sub BrowseSnippetFolderButton_Click(sender As Object, e As RoutedEventArgs)
        Dim dlg As New OpenFileDialog
        With dlg
            .Title = "Select .json snippet"
            .Filter = ".json files (*.json)|*.json"
            If .ShowDialog = True Then
                Me.JsonSnippetFolderTextBox.Text = .FileName
            End If
        End With
    End Sub

    Private Sub BuildVsCodePackageButton_Click(sender As Object, e As RoutedEventArgs)
        If JsonSnippetFolderTextBox.Text = "" Or String.IsNullOrEmpty(JsonSnippetFolderTextBox.Text) Then
            MessageBox.Show("Please select a snippet first.")
            Exit Sub
        End If

        If SnippetLanguageTextBox.Text = "" Or String.IsNullOrEmpty(SnippetLanguageTextBox.Text) Then
            MessageBox.Show("Please specify the language first.")
            Exit Sub
        End If

        Dim sninfo As New SnippetInfo
        sninfo.SnippetFileName = IO.Path.GetFileName(JsonSnippetFolderTextBox.Text)
        sninfo.SnippetPath = IO.Path.GetDirectoryName(JsonSnippetFolderTextBox.Text)
        sninfo.SnippetDescription = $"{SnippetLanguageTextBox.Text} snippets"
        sninfo.SnippetLanguage = SnippetLanguageTextBox.Text

        Dim v As New VsixPackage
        v.PackageAuthor = CodeAuthorNameTextBox.Text
        v.PackageDescription = CodePackageDescriptionTextBox.Text
        v.ProductName = CodeProductNameTextBox.Text
        v.PackageVersion = CodePackageVersionTextBox.Text
        v.CodeSnippets.Add(sninfo)
        If v.HasErrors Then
            MessageBox.Show("Please make sure you have provided all the required information.")
            Exit Sub
        End If

        Dim snippetFolder = IO.Path.GetDirectoryName(JsonSnippetFolderTextBox.Text)

        v.Build(snippetFolder, IDEType.Code)
    End Sub

    Private Sub GenerateNewSnippet()
        Me.snippetData = New CodeSnippet
        Me.EditorRoot.DataContext = snippetData
        editControl1.SetValue(Syncfusion.Windows.Tools.Controls.DockingManager.HeaderProperty, "Untitled")
        Me.snippetData.IsDirty = False
    End Sub

    Private Sub NewSnippetButton_Click(sender As Object, e As RoutedEventArgs)
        If Me.snippetData.IsDirty Then
            Dim result = MessageBox.Show("There are unsaved changes. Are you sure?", "Confirmation",
                                         MessageBoxButton.YesNo, MessageBoxImage.Warning)
            If result = MessageBoxResult.No Then Exit Sub
        End If

        Me.snippetData = Nothing
        GenerateNewSnippet()
    End Sub

    Private Sub FontSizeTextBox_TextChanged(sender As Object, e As TextChangedEventArgs)
        Dim result As Double
        If Double.TryParse(Me.FontSizeTextBox.Text, result) = True Then
            My.Settings.EditorFontSize = result
            My.Settings.Save()
            Me.editControl1.FontSize = result
        Else
            MessageBox.Show("Invalid value", "", MessageBoxButton.OK, MessageBoxImage.Error)
            Exit Sub
        End If
    End Sub

    Private Sub AddRefButton_Click(sender As Object, e As RoutedEventArgs)
        Dim dlg As New OpenFileDialog
        With dlg
            .Title = "Select .NET assembly"
            .Filter = ".dll files (*.dll)|*.dll|All files|*.*"
            .Multiselect = True
            If .ShowDialog = True Then
                For Each fname In .FileNames
                    Me.IntelliSenseReferences.Add(New Uri(fname))
                    Dim ref As New Reference
                    ref.Assembly = IO.Path.GetFileName(fname)
                    snippetData.References.Add(ref)
                Next
            End If
        End With
    End Sub

    Private Sub DeleteRefButton_Click(sender As Object, e As RoutedEventArgs)
        Dim ref = TryCast(RefDataGrid.SelectedItem, Reference)
        If ref IsNot Nothing Then
            snippetData.References.Remove(ref)
        End If
    End Sub

    Private Sub LibraryTreeview_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs)
        Dim item = TryCast(LibraryTreeview.SelectedItem, CodeSnippet)
        If item IsNot Nothing Then
            Me.snippetData = Nothing
            Me.snippetData = item
            Me.EditorRoot.DataContext = snippetData
            editControl1.SetValue(Syncfusion.Windows.Tools.Controls.DockingManager.HeaderProperty, snippetData.FileName)
            If Not IO.Path.GetExtension(snippetData.FileName).ToLower = "json" Then SetCurrentLanguage(snippetData.Language)
        End If
    End Sub

    Private Sub AddLibFolderButton_Click(sender As Object, e As RoutedEventArgs)
        Dim dlg As New System.Windows.Forms.FolderBrowserDialog
        dlg.Description = "New library folder"
        dlg.ShowNewFolderButton = True

        If Not dlg.ShowDialog = Forms.DialogResult.OK Then
            Exit Sub
        End If

        Dim query = From fold In snippetLib.Folders
                    Where fold.FolderName.ToLower = dlg.SelectedPath.ToLower
                    Select fold

        If query.Any Then
            'already exist
            MessageBox.Show("Folder already exist in the library", "Not allowed", MessageBoxButton.OK, MessageBoxImage.Warning)
            Exit Sub
        End If

        Dim newFolder As New SnippetFolder With {.FolderName = dlg.SelectedPath}
        snippetLib.Folders.Add(newFolder)
        snippetLib.SaveLibrary(My.Settings.LibraryName)
    End Sub

    Private Sub DeleteLibFolderButton_Click(sender As Object, e As RoutedEventArgs)
        Dim item = TryCast(LibraryTreeview.SelectedItem, SnippetFolder)
        If item IsNot Nothing Then
            snippetLib.Folders.Remove(item)
            snippetLib.SaveLibrary(My.Settings.LibraryName)
        End If
    End Sub

    Private Sub FilterLibraryTextBox_KeyUp(sender As Object, e As KeyEventArgs) Handles FilterLibraryTextBox.KeyUp
        If e.Key = Key.Enter Then
            FilterSnippetList(Me.FilterLibraryTextBox.Text)
        End If
    End Sub

    Private Sub FilterSnippetList(criteria As String)
        If criteria = "" Then
            Me.LibraryTreeview.ItemsSource = Nothing
            Me.LibraryTreeview.ItemsSource = Me.snippetLib.Folders
            Exit Sub
        End If

        Try
            Dim query = Me.snippetLib.Folders.Where(Function(f) f?.SnippetFiles.Any(Function(s) s.FileName IsNot Nothing _
                        AndAlso s.FileName.ToLowerInvariant.Contains(criteria.ToLowerInvariant)))

            Me.LibraryTreeview.ItemsSource = New ObservableCollection(Of SnippetFolder)(query)
        Catch ex As Exception
            Me.LibraryTreeview.ItemsSource = Me.snippetLib.Folders
        End Try

    End Sub

    Private Sub FilterButton_Click(sender As Object, e As RoutedEventArgs)
        FilterSnippetList(Me.FilterLibraryTextBox.Text)
    End Sub

    Private Sub BackupLibButton_Click(sender As Object, e As RoutedEventArgs)
        If snippetLib.Folders.Any Then
            Dim dlg As New SaveFileDialog
            dlg.Title = "Specify a zip archive name"
            dlg.Filter = "Zip archives|*.zip|All files|*.*"
            dlg.OverwritePrompt = True

            If dlg.ShowDialog = True Then
                Try
                    snippetLib.BackupLibraryToZip(dlg.FileName)
                    MessageBox.Show($"{dlg.FileName} created successfully.")
                    Exit Sub
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
            End If
        End If
    End Sub

    Private Sub AddFromLibButton_Click(sender As Object, e As RoutedEventArgs)
        vsixData?.PopulateFromSnippetLibrary(snippetLib)
    End Sub

    Private Sub AnalyzeCode()
        Try
            If Not String.IsNullOrWhiteSpace(snippetData.Code) Then snippetData.AnalyzeCode()
            Me.ErrorList.ItemsSource = Me.snippetData.Diagnostics
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ErrorList_CurrentCellRequestNavigate(sender As Object, args As CurrentCellRequestNavigateEventArgs)
        Dim diag = CType(args.RowData, Diagnostic)
        If diag.Descriptor.HelpLinkUri <> "" Then
            System.Diagnostics.Process.Start(diag.Descriptor.HelpLinkUri)
        End If
    End Sub

    Private Sub ImportCodeFileButton_Click(sender As Object, e As RoutedEventArgs)
        If snippetData.IsDirty Then
            Dim result = MessageBox.Show("The current snippet has unsaved changes. Are you sure?", "Confirmation", MessageBoxButton.YesNo, MessageBoxImage.Question)
            If result = MessageBoxResult.No Then
                Exit Sub
            End If
        End If

        Dim dlg As New OpenFileDialog

        With dlg
            .Title = "Select code  file"
            .Filter = "Supported code files (.vb,.cs,.cpp,.js,.sql,.xml,.xaml)|*.cs;*.vb;*.cpp;*.sql;*.js;*.xml;*.xaml|All files|*.*"
            If Not .ShowDialog = True Then
                Exit Sub
            End If

            Try
                Dim tempData = CodeSnippet.ImportCodeFile(.FileName)
                If tempData IsNot Nothing Then
                    Me.snippetData = Nothing
                    Me.snippetData = tempData
                    Me.EditorRoot.DataContext = Me.snippetData
                    Me.snippetData.IsDirty = False
                    editControl1.SetValue(Syncfusion.Windows.Tools.Controls.DockingManager.HeaderProperty, "Untitled")
                    SetCurrentLanguage(snippetData.Language)
                End If
            Catch ex As UriFormatException
            Catch ex As Exception
                MessageBox.Show(ex.Message)
                Exit Sub
            End Try
        End With
    End Sub

    Private Sub ErrorList_SelectionChanged(sender As Object, e As GridSelectionChangedEventArgs) Handles ErrorList.SelectionChanged
        Dim diag = TryCast(ErrorList.SelectedItem, Diagnostic)

        If diag Is Nothing Then Exit Sub

        Dim span = diag.Location.GetLineSpan

        If diag.DefaultSeverity = DiagnosticSeverity.Error Then
            editControl1.SelectionBackground = New SolidColorBrush(Colors.Red)
        ElseIf diag.DefaultSeverity = DiagnosticSeverity.Warning Then
            editControl1.SelectionBackground = New SolidColorBrush(Colors.Yellow)
        Else
            editControl1.SelectionBackground = Me.selectionBackground
        End If

        editControl1.SelectLines(span.Span.Start.Line, span.Span.End.Line, span.Span.Start.Character, span.Span.End.Character)
    End Sub


    Private Sub editControl1_GotKeyboardFocus(sender As Object, e As KeyboardFocusChangedEventArgs) Handles editControl1.GotKeyboardFocus
        Me.editControl1.SelectionBackground = Me.selectionBackground
    End Sub

    Private Sub EditorForeColorPicker_SelectedBrushChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Me.editControl1.Foreground = EditorForeColorPicker.Brush
        My.Settings.EditorForeColor = EditorForeColorPicker.Brush
        My.Settings.Save()
    End Sub

    Private Sub EditorSelectionColorPicker_SelectedBrushChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Me.editControl1.SelectionForeground = EditorSelectionColorPicker.Brush
        My.Settings.EditorSelectionColor = EditorSelectionColorPicker.Brush
        My.Settings.Save()
    End Sub

    Private Sub ImportSumblimeButton_Click(sender As Object, e As RoutedEventArgs)
        If snippetData.IsDirty Then
            Dim result = MessageBox.Show("The current snippet has unsaved changes. Are you sure?", "Confirmation", MessageBoxButton.YesNo, MessageBoxImage.Question)
            If result = MessageBoxResult.No Then
                Exit Sub
            End If
        End If

        Dim dlg As New OpenFileDialog

        With dlg
            .Title = "Select code  file"
            .Filter = "All files|*.*"
            If Not .ShowDialog = True Then
                Exit Sub
            End If

            Try
                Dim tempData = CodeSnippet.ImportSublimeSnippet(.FileName)
                If tempData IsNot Nothing Then
                    Me.snippetData = Nothing
                    Me.snippetData = tempData
                    Me.EditorRoot.DataContext = Me.snippetData
                    Me.snippetData.IsDirty = False
                    editControl1.SetValue(Syncfusion.Windows.Tools.Controls.DockingManager.HeaderProperty, "Untitled")
                    SetCurrentLanguage(snippetData.Language)
                End If
            Catch ex As UriFormatException
            Catch ex As Exception
                MessageBox.Show(ex.Message)
                Exit Sub
            End Try
        End With
    End Sub
End Class