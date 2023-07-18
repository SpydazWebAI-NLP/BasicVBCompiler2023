
Namespace Basic_NLP
    Public Class VbCompiler

        Private Const Lang As String = "VisualBasic"
        Private Shared _results As CodeDom.Compiler.CompilerResults
        Private Shared Compiler As CodeDom.Compiler.CodeDomProvider = CodeDom.Compiler.CodeDomProvider.CreateProvider(Lang)
        Private Shared InternalRefferences As List(Of String)
        Private Shared Settings As New CodeDom.Compiler.CompilerParameters
        Private _compiled As Boolean
        Private _EmbededFiles As List(Of String)
        Private _Refferences As List(Of String)
        Private _source, _entClass, _entMethod As String

        Public Sub New(sourceCode As String, entClass As String, Optional entMethod As String = "Main",
                   Optional Assemblies As List(Of String) = Nothing,
                                     Optional ByVal EmbededFiles As List(Of String) = Nothing)
            _source = sourceCode
            _entClass = entClass
            _entMethod = entMethod
            _Refferences = Assemblies
            _EmbededFiles = EmbededFiles
        End Sub

        Public Shared ReadOnly Property Errors() As CodeDom.Compiler.CompilerErrorCollection
            Get
                Return _results.Errors
            End Get
        End Property

        Public Shared Sub AddHelperRefferences()

            AddInternalReferencesHelper()
            AddProjectSpydazWebAIRefferencesHelper()

        End Sub

        Public Shared Sub AddInternalReferencesHelper()
            InternalRefferences = New List(Of String)
            InternalRefferences.Add(Application.StartupPath & "\System.Web.Extensions.dll")
            InternalRefferences.Add(Application.StartupPath & "\System.Windows.Forms.dll")
            InternalRefferences.Add(Application.StartupPath & "\Microsoft.VisualBasic.dll")
            InternalRefferences.Add(Application.StartupPath & "\System.dll")
            InternalRefferences.Add(Application.StartupPath & "\System.Speech.dll")
            If InternalRefferences IsNot Nothing And InternalRefferences.Count > 0 Then
                For Each str As String In InternalRefferences
                    Settings.ReferencedAssemblies.Add(str)
                Next
            End If
        End Sub

        Public Shared Sub AddProjectSpydazWebAIRefferencesHelper()
            InternalRefferences = New List(Of String)
            InternalRefferences.Add(Application.StartupPath & "\SpydazWebAI_ControlLibrary.dll")
            InternalRefferences.Add(Application.StartupPath & "\AI_AGENT.dll")
            InternalRefferences.Add(Application.StartupPath & "\AI_SDK.dll")
            InternalRefferences.Add(Application.StartupPath & "\DirectShowLib-2005.dll")
            If InternalRefferences IsNot Nothing And InternalRefferences.Count > 0 Then
                For Each str As String In InternalRefferences
                    Settings.ReferencedAssemblies.Add(str)
                Next
            End If
        End Sub

        Public Shared Function GetFilenameDLL() As String
            GetFilenameDLL = ""
            Dim S As New SaveFileDialog
            With S

                .Filter = "Executable (*.Dll)|*.Dll"
                If (.ShowDialog() = DialogResult.OK) Then
                    Return .FileName
                End If
            End With
        End Function

        Public Shared Function GetFilenameEXE() As String
            GetFilenameEXE = ""
            Dim S As New SaveFileDialog
            With S

                .Filter = "Executable (*.exe)|*.exe"
                If (.ShowDialog() = DialogResult.OK) Then
                    Return .FileName
                End If
            End With
        End Function

        Public Shared Sub RunInteractive(ByRef CodeBlock As String, Optional ByVal iClassName As String = "MainClass",
                                      Optional iMethodName As String = "Execute", Optional Assemblies As List(Of String) = Nothing,
                                     Optional ByVal EmbededFiles As List(Of String) = Nothing)
            Try

                VB_CodeCompilerAsync(CodeBlock, "MEM", iClassName, iMethodName, Assemblies, EmbededFiles)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End Sub

        Public Shared Sub VB_CodeCompiler(CodeBlock As String,
                                      Optional ByVal CompileType As String = "MEM",
                                      Optional ByVal iClassName As String = "MainClass",
                                      Optional iMethodName As String = "Execute",
                                      Optional Assemblies As List(Of String) = Nothing,
                                     Optional ByVal EmbededFiles As List(Of String) = Nothing)

            Dim cmp As New VbCompiler(CodeBlock, iClassName, iMethodName)
            cmp.AddEmbeddedFiles(EmbededFiles)
            cmp.AddRefferences(Assemblies)
            Select Case CompileType
                Case "EXE"
                    cmp.Compile_EXE(GetFilenameDLL)
                Case "DLL"
                    cmp.Compile_DLL(GetFilenameDLL)
                Case "MEM"
                    cmp.Compile_MEM()
            End Select
        End Sub

        Public Shared Async Sub VB_CodeCompilerAsync(CodeBlock As String,
                                      Optional ByVal CompileType As String = "MEM",
                                      Optional ByVal iClassName As String = "MainClass",
                                      Optional iMethodName As String = "Execute",
                                      Optional Assemblies As List(Of String) = Nothing,
                                     Optional ByVal EmbededFiles As List(Of String) = Nothing)
            Try

                Await Task.Run(Sub() VB_CodeCompiler(CodeBlock, CompileType, iClassName, iMethodName, Assemblies, EmbededFiles))
            Catch ex As Exception
            End Try
        End Sub

        Public Function CompileDll() As Boolean
            AddEmbeddedFiles()
            AddRefferences()
            Compile_DLL(GetFilenameDLL)

            If getErrors.Count > 0 Then
                _compiled = False
                Return False
            Else
                _compiled = True
                Return True
            End If
        End Function

        Public Function CompileExE() As Boolean
            AddEmbeddedFiles()
            AddRefferences()
            Compile_EXE(GetFilenameEXE)
            If getErrors.Count > 0 Then
                _compiled = False
                Return False
            Else
                _compiled = True
                Return True
            End If
        End Function

        Public Function CompileMEM() As Boolean
            AddEmbeddedFiles()
            AddRefferences()
            Compile_MEM()
            If getErrors.Count > 0 Then
                _compiled = False
                Return False
            Else
                _compiled = True
                Return True
            End If
        End Function

        Public Function getErrors() As List(Of String)

#Region "Errors"

            Dim IntErrors As New List(Of String)
            'Determines if we have any errors when compiling if so loops through all of the CompileErrors in the Reults variable and displays their ErrorText property.
            If (_results.Errors.Count <> 0) Then

                '  MessageBox.Show("Exception Occured!", "Whoops!", MessageBoxButtons.OK, MessageBoxIcon.Information)
                For Each E As CodeDom.Compiler.CompilerError In _results.Errors

                    IntErrors.Add(E.ErrorText)

                Next
            Else
            End If

#End Region

            Return IntErrors
        End Function

        Private Sub AddEmbeddedFiles(Optional EmbededFiles As List(Of String) = Nothing)

            Try
                'handle Embedded Resources
                If EmbededFiles IsNot Nothing And EmbededFiles.Count > 0 Then
                    For Each str As String In EmbededFiles
                        Settings.EmbeddedResources.Add(str)
                    Next
                End If
                If _EmbededFiles IsNot Nothing And _EmbededFiles.Count > 0 Then
                    For Each str As String In _EmbededFiles
                        Settings.EmbeddedResources.Add(str)
                    Next
                End If
            Catch ex As Exception

            End Try
        End Sub

        Private Sub AddRefferences(Optional Refferences As List(Of String) = Nothing)
            'No References Held from Ctor
            Try
                If Refferences IsNot Nothing And Refferences.Count > 0 Then
                    For Each str As String In Refferences
                        Settings.ReferencedAssemblies.Add(str)
                    Next
                End If
                If _Refferences IsNot Nothing And _Refferences.Count > 0 Then
                    For Each str As String In _Refferences
                        Settings.ReferencedAssemblies.Add(str)
                    Next
                End If
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Compile_DLL(ExecuteableName As String)
            'Library type options : /target:library, /target:win, /target:winexe
            'Generates an executable instead of a class library.
            'Compiles in memory.
            Settings.GenerateInMemory = False
            Settings.GenerateExecutable = False
            Settings.CompilerOptions = " /target:library"
            'Set the assembly file name / path
            Settings.OutputAssembly = ExecuteableName
            'Read the documentation for the result again variable.
            'Calls the CompileAssemblyFromSource that will compile the specified source code using the parameters specified in the settings variable.
            _results = Compiler.CompileAssemblyFromSource(Settings, _source)
        End Sub

        Private Sub Compile_EXE(ExecuteableName As String)
            'Library type options : /target:library, /target:win, /target:winexe
            'Generates an executable instead of a class library.
            'Compiles in memory.
            Settings.GenerateInMemory = True
            Settings.GenerateExecutable = True
            Settings.CompilerOptions = " /target:winexe"
            'Set the assembly file name / path
            Settings.OutputAssembly = ExecuteableName
            'Read the documentation for the result again variable.
            'Calls the CompileAssemblyFromSource that will compile the specified source code using the parameters specified in the settings variable.
            _results = Compiler.CompileAssemblyFromSource(Settings, _source)
        End Sub

        Private Sub Compile_MEM()
            'Library type options : /target:library, /target:win, /target:winexe
            'Generates an executable instead of a class library.
            'Compiles in memory.
            Settings.GenerateInMemory = True
            'Read the documentation for the result again variable.
            'Calls the CompileAssemblyFromSource that will compile the specified source code using the parameters specified in the settings variable.
            _results = Compiler.CompileAssemblyFromSource(Settings, _source)
        End Sub

    End Class
End Namespace
