Imports System
Imports System.IO
Imports System.Resources
Imports System.Reflection
Imports System.Runtime.InteropServices

''' <summary>
''' Main Module for the GenerateLineMap utility
''' (c) 2008-2011 Darin Higgins All Rights Reserved
''' This is a command line utility, so the SUB MAIN is the main entry point.
''' </summary>
''' <remarks></remarks>
Module modMain

    <DllImport("kernel32.dll", CharSet:=CharSet.Auto)> _
    Public Sub LoadLibrary(ByVal lpFileName As String)
    End Sub


    ''' <summary>
    ''' Command line application entry point
    ''' </summary>
    ''' <remarks></remarks>
    Sub Main()
        Dim FileName As String = ""
        Dim bReport As Boolean = False
        Dim bFile As Boolean = False
        Dim bAPIResource As Boolean = True
        Dim bNETResource As Boolean = False
        Dim exitcode As Integer = 0
        Dim outfile As String = ""

        Try
            For Each s As String In My.Application.CommandLineArgs
                Dim bHandled = False
                If Len(s) = 2 And (s.Contains("?") Or s.ToLower.Contains("h")) Then
                    ShowHelp()
                    bHandled = True
                End If
                If String.Compare(s, "/report", True) = 0 Then
                    bReport = True
                    bHandled = True
                End If
                If s.StartsWith("/out:", StringComparison.InvariantCultureIgnoreCase) Then
                    bHandled = True
                    outfile = s.Substring(5).Trim
                    If outfile.StartsWith("""") Then
                        outfile = outfile.Substring(1)
                        If outfile.EndsWith("""") Then
                            outfile = outfile.Substring(0, outfile.Length - 1)
                        End If
                    End If
                End If
                If String.Compare(s, "/file", True) = 0 Then
                    '---- write the line map to a separate file
                    '     normally, it's written back into the EXE as a resource
                    bFile = True
                    bAPIResource = False
                    bNETResource = False
                    bHandled = True
                End If
                If String.Compare(s, "/apiresource", True) = 0 Then
                    '---- write the line map to a winAPI resource
                    '     normally, it's written back into the EXE as a .net resource
                    bAPIResource = True
                    bNETResource = False
                    bHandled = True
                End If
                If String.Compare(s, "/resource", True) = 0 Then
                    '---- write the line map to a .net resource
                    bAPIResource = False
                    bNETResource = True
                    bHandled = True
                End If

                If Not bHandled Then
                    If My.Computer.FileSystem.FileExists(s) Then
                        FileName = s
                    End If
                End If
            Next

            Console.WriteLine(My.Application.Info.Title & " v" & My.Application.Info.Version.ToString)
            Console.WriteLine(My.Application.Info.Description)
            Console.WriteLine(String.Format("   {0}", My.Application.Info.Copyright))
            Console.WriteLine()

            If Len(FileName) = 0 Then ShowHelp() : End

            '---- extract the necessary dbghelp.dll
            If ExtractDbgHelp() Then
                Console.WriteLine("Unable to extract dbghelp.dll to this folder.")
                End
            End If

            Dim lmb = New LineMapBuilder(FileName, outfile)

            If bReport Then
                '---- just set a flag to gen a report
                lmb.CreateMapReport = True
            End If

            If bFile Then
                Console.WriteLine(String.Format("Creating linemap file for file {0}...", FileName))
                lmb.CreateLineMapFile()
            End If

            If bAPIResource Then
                Console.WriteLine(String.Format("Adding linemap WIN resource in file {0}...", FileName))
                lmb.CreateLineMapAPIResource()
            End If

            If bNETResource Then
                Console.WriteLine(String.Format("Adding linemap .NET resource in file {0}...", FileName))
                lmb.CreateLineMapResource()
            End If

        Catch ex As Exception
            '---- let em know we had a failure
            Console.WriteLine("Unable to complete operation. Error: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace)

            exitcode = 1

        End Try

        '---- Return an exit code of 0 on success, 1 on failure
        Environment.Exit(exitcode)
    End Sub


    Public Sub ShowHelp()
        Console.WriteLine()
        Console.WriteLine("Usage:")
        Console.WriteLine(String.Format("   {0} FilenameOfExeOrDllFile [options]", My.Application.Info.Title))
        Console.WriteLine("where options are:")
        Console.WriteLine("   [/report] [[/file]|[/resource]|[/apiresource]]")
        Console.WriteLine()
        Console.WriteLine("/report        Generate report of contents of PDB file")
        Console.WriteLine("/file          Output a linemap file with the symbol and line num buffers")
        Console.WriteLine("/resource      (default) Create a linemap .NET resource in the target")
        Console.WriteLine("               EXE/DLL file")
        Console.WriteLine("/apiresource   Create a linemap windows resource in the target EXE/DLL file")
        Console.WriteLine()
        Console.WriteLine("The default is 'apiresource' which embeds the linemap into")
        Console.WriteLine("the target executable as a standard windows resource.")
        Console.WriteLine(".NET resource support is experimental at this point.")
		Console.WriteLine("The 'file' option is mainly for testing. The resulting *.lmp")
		Console.WriteLine("file will contain source names and line numbers but no other")
        Console.WriteLine("information commonly found in PDB files.")
        Console.WriteLine()
        Console.WriteLine("Returns an exitcode of 0 on success, 1 on failure")
    End Sub


    ''' <summary>
    ''' Extract the dbghelp.dll file from the exe
    ''' </summary>
    ''' <returns>true on failure</returns>
    ''' <remarks></remarks>
    Private Function ExtractDbgHelp() As Boolean
        Try
            '---- Get our assembly
            Dim executing_assembly = Assembly.GetExecutingAssembly()

            '---- Get our namespace
            '     Note that this is different from the Appname because we compile to a 
            '     file call *Raw so that we can run ILMerge and result in the final filename
            Dim my_namespace = executing_assembly.EntryPoint.DeclaringType.Namespace

            '---- write the file back out
            Dim dbghelp_stream As Stream

            dbghelp_stream = executing_assembly.GetManifestResourceStream(my_namespace & ".dbghelp.dll")
            If Not (dbghelp_stream Is Nothing) Then
                '---- write stream to file
                Dim AppPath = Path.GetDirectoryName(executing_assembly.Location)
                Dim outname = Path.Combine(AppPath, "dbghelp.dll")
                Dim bExtract = True
                If My.Computer.FileSystem.FileExists(outname) Then
                    '---- is it the right file (just check length for now)
                    If My.Computer.FileSystem.GetFileInfo(outname).Length = CInt(dbghelp_stream.Length) Then
                        bExtract = False
                    End If
                End If

                If bExtract Then
                    Dim reader As BinaryReader = New BinaryReader(dbghelp_stream)
                    Dim buffer = reader.ReadBytes(CInt(dbghelp_stream.Length))
                    Dim output = New FileStream(outname, FileMode.Create)
                    output.Write(buffer, 0, CInt(dbghelp_stream.Length))
                    output.Close()
                    reader.Close()
                End If
            End If
        Catch ex As Exception
            Console.WriteLine("")
            Console.WriteLine("!!! Unable to extract dbghelp.dll to current directory.")
            Return True
        End Try
        Return False
    End Function
End Module
