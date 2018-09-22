Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text


''' <summary>
''' Utility Class to generate a LineMap based on the
''' contents of the PDB for a specific EXE or DLL
''' (c) 2008-2017 Darin Higgins All Rights Reserved
''' </summary>
''' <remarks>
''' If some of the functions seem a little odd here, that's mostly because
''' I was experimenting with "Fluent" style apis during most of the 
''' development of this project.
''' </remarks>
Public Class LineMapBuilder

#Region " Structures"
    ''' <summary>
    ''' Structure used when Enumerating symbols from a PDB file
    ''' Note that this structure must be marshalled very specifically
    ''' to be compatible with the SYM* dbghelp dll calls
    ''' </summary>
    ''' <remarks></remarks>
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Ansi)> _
    Private Structure SYMBOL_INFO
        Public SizeOfStruct As Integer
        Public TypeIndex As Integer
        Public Reserved0 As Int64
        Public Reserved1 As Int64
        Public Index As Integer
        Public size As Integer
        Public ModBase As Int64
        Public Flags As CV_SymbolInfoFlags
        Public Value As Int64
        Public Address As Int64
        Public Register As Integer
        Public Scope As Integer
        Public Tag As Integer
        Public NameLen As Integer
        Public MaxNameLen As Integer
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=300)> Public Name As String
    End Structure


    Private Const MAX_PATH = 260


    ''' <summary>
    ''' Structure used when Enumerating line numbers from a PDB file
    ''' Note that this structure must be marshalled very specifically
    ''' to be compatible with the SYM* dbghelp dll calls
    ''' </summary>
    ''' <remarks></remarks>
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Ansi)> _
    Private Structure SRCCODEINFO
        Public SizeOfStruct As Integer
        Public Key As Integer
        Public ModBase As Int64
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=MAX_PATH + 1)> Public obj As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=MAX_PATH + 1)> Public FileName As String
        Public LineNumber As Integer
        Public Address As Int64
    End Structure


    Private Enum CV_SymbolInfoFlags
        IMAGEHLP_SYMBOL_INFO_VALUEPRESENT = &H1
        IMAGEHLP_SYMBOL_INFO_REGISTER = &H8
        IMAGEHLP_SYMBOL_INFO_REGRELATIVE = &H10
        IMAGEHLP_SYMBOL_INFO_FRAMERELATIVE = &H20
        IMAGEHLP_SYMBOL_INFO_PARAMETER = &H40
        IMAGEHLP_SYMBOL_INFO_LOCAL = &H80
        IMAGEHLP_SYMBOL_INFO_CONSTANT = &H100
        IMAGEHLP_SYMBOL_FUNCTION = &H800
        IMAGEHLP_SYMBOL_VIRTUAL = &H1000
        IMAGEHLP_SYMBOL_THUNK = &H2000
        IMAGEHLP_SYMBOL_TLSREL = &H4000
        IMAGEHLP_SYMBOL_SLOT = &H8000
        IMAGEHLP_SYMBOL_ILREL = &H10000
        IMAGEHLP_SYMBOL_METADATA = &H20000
        IMAGEHLP_SYMBOL_CLR_TOKEN = &H40000
    End Enum
#End Region


#Region " API Definitions"
    Private Declare Function SymInitialize Lib "dbghelp.dll" ( _
       ByVal hProcess As Integer, _
       ByVal UserSearchPath As String, _
       ByVal fInvadeProcess As Boolean) As Integer


    Private Declare Function SymLoadModuleEx Lib "dbghelp.dll" ( _
       ByVal hProcess As Integer, _
       ByVal hFile As Integer, _
       ByVal ImageName As String, _
       ByVal ModuleName As Integer, _
       ByVal BaseOfDll As Int64, _
       ByVal SizeOfDll As Integer, _
       ByVal pModData As Integer, _
       ByVal flags As Integer) As Int64

    '---- I believe this is deprecatedin dbghelp 6.0 or later
    'Private Declare Function SymLoadModule Lib "dbghelp.dll" ( _
    '   ByVal hProcess As Integer, _
    '   ByVal hFile As Integer, _
    '   ByVal ImageName As String, _
    '   ByVal ModuleName As Integer, _
    '   ByVal BaseOfDll As Integer, _
    '   ByVal SizeOfDll As Integer) As Integer


    Private Declare Function SymCleanup Lib "dbghelp.dll" (ByVal hProcess As Integer) As Integer


    Private Declare Function UndecorateSymbolName Lib "dbghelp.dll" Alias "UnDecorateSymbolName" ( _
       ByVal DecoratedName As String, _
       ByVal UnDecoratedName As String, _
       ByVal UndecoratedLength As Integer, _
       ByVal Flags As Integer) As Integer


    Private Declare Function SymEnumSymbols Lib "dbghelp.dll" ( _
       ByVal hProcess As Integer, _
       ByVal BaseOfDll As Int64, _
       ByVal Mask As Integer, _
       ByVal lpCallback As SymEnumSymbolsCallback, _
       ByVal UserContext As Integer) As Integer


    Private Declare Function SymEnumLines Lib "dbghelp.dll" ( _
       ByVal hProcess As Integer, _
       ByVal BaseOfDll As Int64, _
       ByVal obj As Integer, _
       ByVal File As Integer, _
       ByVal lpCallback As SymEnumLinesCallback, _
       ByVal UserContext As Integer) As Integer


    Private Declare Function SymEnumSourceLines Lib "dbghelp.dll" ( _
       ByVal hProcess As Integer, _
       ByVal BaseOfDll As Int64, _
       ByVal obj As Integer, _
       ByVal File As Integer, _
       ByVal Line As Integer, _
       ByVal Flags As Integer, _
       ByVal lpCallback As SymEnumLinesCallback, _
       ByVal UserContext As Integer) As Integer


    '---- I believe this is deprecated in dbghelp 6.0+
    'Private Declare Function SymUnloadModule Lib "dbghelp.dll" ( _
    '   ByVal hProcess As Integer, _
    '   ByVal BaseOfDll As Integer) As Integer

    Private Declare Function SymUnloadModule64 Lib "dbghelp.dll" ( _
       ByVal hProcess As Integer, _
       ByVal BaseOfDll As Int64) As Boolean

#End Region


#Region " Exceptions"
    Public Class UnableToEnumerateSymbolsException
        Inherits ApplicationException
    End Class


    Public Class UnableToEnumLinesException
        Inherits ApplicationException
    End Class
#End Region


    ''' <summary>
    ''' Constructor to setup this class to read a specific PDB
    ''' </summary>
    ''' <param name="FileName"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal FileName As String)
        Me.Filename = FileName
    End Sub


    Public Sub New(ByVal FileName As String, ByVal OutFileName As String)
        Me.Filename = FileName
        Me.OutFileName = OutFileName
    End Sub


    ''' <summary>
    ''' Name of EXE/DLL file to process
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Public Property Filename() As String
        Get
            Return rFilename
        End Get
        Set(ByVal value As String)
            If Not System.IO.File.Exists(value) Then
                Throw New FileNotFoundException("The file could not be found.", value)
            End If
            rFilename = value
        End Set
    End Property
    Private rFilename As String = ""


    ''' <summary>
    ''' Name of the output file
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Public Property OutFilename() As String
        Get
            If String.IsNullOrEmpty(rOutFilename) Then
                Return rFilename
            Else
                Return rOutFilename
            End If
        End Get
        Set(ByVal value As String)
            rOutFilename = value
        End Set
    End Property
    Private rOutFilename As String


    Private rbCreateMapReport As Boolean = False
    ''' <summary>
    ''' True to generate a Line map report
    ''' this is mainly for Debugging purposes
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Public Property CreateMapReport() As Boolean
        Get
            Return rbCreateMapReport
        End Get
        Set(ByVal value As Boolean)
            rbCreateMapReport = value
        End Set
    End Property


    ''' <summary>
    ''' Create a linemap file from the PDB for the given executable file
    ''' Theoretically, you could read the LineMap info from the linemap file
    ''' OR from a resource in the EXE/DLL, but I doubt you'd ever really want
    ''' to. This is mainly for testing all the streaming functions
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub CreateLineMapFile()
        '---- compress and encrypt the stream
        '     technically, I should probably CLOSE and DISPOSE the streams
        '     but this is just a quick and dirty tool
        Dim alm = pGetAssemblyLineMap(Me.Filename)

        Dim CompressedStream = pCompressStream(alm.ToStream)
        Dim EncryptedStream = pEncryptStream(CompressedStream)

		'---- swap out the below two lines to generate a linemap (lnm) file that is not compressed or encrypted
		pStreamToFile(Me.OutFilename & ".lmp", EncryptedStream)
		'pStreamToFile(Me.Filename & ".lmp", alm.ToStream)

		'---- write the report
		pCheckToWriteReport(alm)
    End Sub


    ''' <summary>
    ''' Inject a linemap resource into the given EXE/DLL file
    ''' from the PDB for that file
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub CreateLineMapAPIResource()
        '---- retrieve all the symbols
        Dim alm = pGetAssemblyLineMap(Me.Filename)

        '---- done one step at a time for debugging purposes
        Dim CompressedStream = pCompressStream(alm.ToStream)
        Dim EncryptedStream = pEncryptStream(CompressedStream)
        pStreamToAPIResource(Me.OutFilename, _
                          LineMap.LineMapKeys.ResTypeName, _
                          LineMap.LineMapKeys.ResName, _
                          LineMap.LineMapKeys.ResLang, _
                          EncryptedStream)

        '---- write the report
        pCheckToWriteReport(alm)
    End Sub


    ''' <summary>
    ''' Inject a linemap .net resource into the given EXE/DLL file
    ''' from the PDB for that file
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub CreateLineMapResource()
        '---- retrieve all the symbols
        Dim alm = pGetAssemblyLineMap(Me.Filename)

        '---- done one step at a time for debugging purposes
        Dim CompressedStream = pCompressStream(alm.ToStream)
        Dim EncryptedStream = pEncryptStream(CompressedStream)
        pStreamToResource(Me.OutFilename, _
                          LineMap.LineMapKeys.ResName, _
                          EncryptedStream)

        '---- write the report
        pCheckToWriteReport(alm)
    End Sub


    ''' <summary>
    ''' Internal function to write out a line map report if asked to
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub pCheckToWriteReport(ByVal AssemblyLineMap As LineMap.AssemblyLineMap)
        If Me.CreateMapReport Then
            Console.WriteLine("Creating symbol buffer report")
            Using tw = My.Computer.FileSystem.OpenTextFileWriter(Me.Filename & ".linemapreport", False)
                tw.Write(pCreatePDBReport(AssemblyLineMap).ToString)
                tw.Flush()
            End Using
        End If
    End Sub


    ''' <summary>
    ''' Create a linemap report buffer from the PDB for the given executable file
    ''' </summary>
    ''' <remarks></remarks>
    Private Function pCreatePDBReport(ByVal AssemblyLineMap As LineMap.AssemblyLineMap) As StringBuilder
        Dim sb = New StringBuilder

        sb.AppendLine("========")
        sb.AppendLine("SYMBOLS:")
        sb.AppendLine("========")
        sb.AppendLine(String.Format("   {0,-10}  {1,-10}  {2}  ", "Token", "Address", "Symbol"))
        For Each symbolEx In AssemblyLineMap.Symbols.Values
            sb.AppendLine(String.Format("   {0,-10:X}  {1,-10}  {2}", symbolEx.Token, symbolEx.Address, symbolEx.Name))
        Next
        sb.AppendLine("========")
        sb.AppendLine("LINE NUMBERS:")
        sb.AppendLine("========")
        sb.AppendLine(String.Format("   {0,-10}  {1,-11}  {2,-10}  {3}", "Address", "Line number", "Token", "Symbol/FileName"))
        Dim y As Integer = 0
        For Each lineex In AssemblyLineMap.AddressToLineMap
            '---- find the symbol for this line number
            Dim sym As LineMap.AssemblyLineMap.SymbolInfo = Nothing
            For Each symbolex In AssemblyLineMap.Symbols.Values
                If symbolex.Address = lineex.Address Then
                    '---- found the symbol for this line
                    sym = symbolex
                    Exit For
                End If
            Next
            Dim n = If(sym IsNot Nothing, lineex.ObjectName & "." & sym.Name, "") & " / " & lineex.SourceFile
            Dim t = If(sym IsNot Nothing, sym.Token, 0)
            sb.AppendLine(String.Format("   {0,-10}  {1,-11}  {2,-10:X}  {3}", lineex.Address, lineex.Line, t, n))
        Next

        sb.AppendLine("========")
        sb.AppendLine("NAMES:")
        sb.AppendLine("========")
        sb.AppendLine(String.Format("   {0,-10}  {1}", "Index", "Name"))
        For i = 0 To AssemblyLineMap.Names.Count - 1
            sb.AppendLine(String.Format("   {0,-10}  {1}", i, AssemblyLineMap.Names(i)))
        Next

        Return sb
    End Function


    ''' <summary>
    ''' Retrieve symbols and linenums and write them to a memory stream
    ''' </summary>
    ''' <param name="FileName"></param>
    ''' <returns></returns>
    Private Function pGetAssemblyLineMap(ByVal FileName As String) As LineMap.AssemblyLineMap
        '---- create a new map to capture symbols and line info with
        _alm = LineMap.AssemblyLineMaps.Add(FileName)

        If Not System.IO.File.Exists(FileName) Then
            Throw New FileNotFoundException("The file could not be found.", FileName)
        End If


        Dim hProcess As Integer = System.Diagnostics.Process.GetCurrentProcess.Id
        Dim dwModuleBase As Int64

        '---- clear the map
        _alm.Clear()

        Try
            If SymInitialize(hProcess, "", False) <> 0 Then
                dwModuleBase = SymLoadModuleEx(hProcess, 0, FileName, 0, 0, 0, 0, 0)
                If dwModuleBase <> 0 Then
                    '---- Enumerate all the symbol names 
                    Dim rEnumSymbolsDelegate As New SymEnumSymbolsCallback(AddressOf SymEnumSymbolsProc)
                    If SymEnumSymbols(hProcess, dwModuleBase, 0, rEnumSymbolsDelegate, 0) = 0 Then
                        '---- unable to retrieve the symbol list
                        Throw New UnableToEnumerateSymbolsException
                    End If

                    '---- now enum all the source lines and their respective addresses
                    Dim rEnumLinesDelegate As New SymEnumLinesCallback(AddressOf SymEnumLinesProc)
                    If SymEnumSourceLines(hProcess, dwModuleBase, 0, 0, 0, 0, rEnumLinesDelegate, 0) = 0 Then
                        '---- unable to retrieve the line number list
                        Throw New UnableToEnumLinesException
                    End If
                End If
            End If

        Catch ex As Exception
            '---- return return vars
            Console.WriteLine("Unable to retrieve symbols. Error: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace)
            _alm.Clear()
            '---- and rethrow
            Throw ex

        Finally
            Console.WriteLine("Retrieved {0} symbols", _alm.Symbols.Count)
            Console.WriteLine("Retrieved {0} lines", _alm.AddressToLineMap.Count)
            Console.WriteLine("Retrieved {0} strings", _alm.Names.Count)

            '---- release the module
            If dwModuleBase <> 0 Then SymUnloadModule64(hProcess, dwModuleBase)

            '---- can clean up the dbghelp system
            Call SymCleanup(hProcess)
        End Try

        Return _alm
    End Function
    Private _alm As LineMap.AssemblyLineMap


    ''' <summary>
    ''' Given a Filename and memorystream, write the stream to the file
    ''' </summary>
    ''' <param name="Filename"></param>
    ''' <param name="MemoryStream"></param>
    ''' <remarks></remarks>
    Private Sub pStreamToFile(ByVal Filename As String, ByVal MemoryStream As MemoryStream)

        Console.WriteLine("Writing symbol buffer to file")
        Using FileStream As System.IO.FileStream = New System.IO.FileStream(Filename, FileMode.Create)
            MemoryStream.WriteTo(FileStream)
            FileStream.Flush()
        End Using
    End Sub


    ''' <summary>
    ''' Write out a memory stream to a named resource in the given 
    ''' WIN32 format executable file (iether EXE or DLL)
    ''' </summary>
    ''' <param name="Filename"></param>
    ''' <param name="ResourceName"></param>
    ''' <param name="MemoryStream"></param>
    ''' <remarks></remarks>
    Private Sub pStreamToAPIResource(ByVal Filename As String, ByVal ResourceType As Object, ByVal ResourceName As Object, ByVal ResourceLanguage As Int16, ByVal MemoryStream As MemoryStream)
        Dim ResWriter = New ResourceWriter

        '---- the target file has to exist
        If Me.Filename <> Filename Then
            FileCopy(Me.Filename, Filename)
        End If
        '---- convert memorystream to byte array and write out to 
        '     the linemap resource
        Console.WriteLine("Writing symbol buffer to resource")
        ResWriter.FileName = Filename
        ResWriter.Update(ResourceType, ResourceName, ResourceLanguage, MemoryStream.ToArray())
    End Sub


    ''' <summary>
    ''' Write out a memory stream to a named resource in the given 
    ''' WIN32 format executable file (iether EXE or DLL)
    ''' </summary>
    ''' <param name="OutFilename"></param>
    ''' <param name="ResourceName"></param>
    ''' <param name="MemoryStream"></param>
    ''' <remarks></remarks>
    Private Sub pStreamToResource(ByVal OutFilename As String, ByVal ResourceName As String, ByVal MemoryStream As MemoryStream)
        Dim ResWriter = New ResourceWriterCecil

        '---- the target file has to exist
        If Me.Filename <> Filename Then
            FileCopy(Me.Filename, Filename)
        End If
        '---- convert memorystream to byte array and write out to 
        '     the linemap resource
        Console.WriteLine("Writing symbol buffer to resource")
        ResWriter.FileName = Me.Filename
        ResWriter.Add(ResourceName, MemoryStream.ToArray())
        ResWriter.Save(OutFilename)
    End Sub


    ''' <summary>
    ''' Given an input stream, compress it into an output stream
    ''' </summary>
    ''' <param name="UncompressedStream"></param>
    ''' <returns></returns>
    Private Function pCompressStream(ByVal UncompressedStream As MemoryStream) As MemoryStream
        Dim CompressedStream As New System.IO.MemoryStream()
        '---- note, the LeaveOpen parm MUST BE TRUE into order to read the memory stream afterwards!
        Console.WriteLine("Compressing symbol buffer")
        Using GZip As New System.IO.Compression.GZipStream(CompressedStream, System.IO.Compression.CompressionMode.Compress, True)
            GZip.Write(UncompressedStream.ToArray, 0, CInt(UncompressedStream.Length))
        End Using
        Return CompressedStream
    End Function


    ''' <summary>
    ''' Given a stream, encrypt it into an output stream
    ''' </summary>
    ''' <param name="CompressedStream"></param>
    ''' <returns></returns>
    Private Function pEncryptStream(ByVal CompressedStream As MemoryStream) As MemoryStream
        Dim Enc As New System.Security.Cryptography.RijndaelManaged

        Console.WriteLine("Encrypting symbol buffer")
        '---- setup our encryption key
        Enc.KeySize = 256
        '---- KEY is 32 byte array
        Enc.Key = LineMap.LineMapKeys.ENCKEY
        '---- IV is 16 byte array
        Enc.IV = LineMap.LineMapKeys.ENCIV

        Dim EncryptedStream = New System.IO.MemoryStream
        Dim cryptoStream = New Security.Cryptography.CryptoStream(EncryptedStream, Enc.CreateEncryptor, Security.Cryptography.CryptoStreamMode.Write)

        cryptoStream.Write(CompressedStream.ToArray, 0, CInt(CompressedStream.Length))
        cryptoStream.FlushFinalBlock()

        Return EncryptedStream
    End Function


#Region " Symbol Enumeration Delegates"
    ''' <summary>
    ''' Delegate to handle Symbol Enumeration Callback
    ''' </summary>
    ''' <param name="syminfo"></param>
    ''' <param name="Size"></param>
    ''' <param name="UserContext"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Delegate Function SymEnumSymbolsCallback(ByRef syminfo As SYMBOL_INFO, ByVal Size As Integer, ByVal UserContext As Integer) As Integer
    Private Function SymEnumSymbolsProc(ByRef syminfo As SYMBOL_INFO, ByVal Size As Integer, ByVal UserContext As Integer) As Integer
        If syminfo.Flags = (CV_SymbolInfoFlags.IMAGEHLP_SYMBOL_CLR_TOKEN Or CV_SymbolInfoFlags.IMAGEHLP_SYMBOL_METADATA) Then
            '---- we only really care about CLR metadata tokens
            '     anything else is basically a variable or internal
            '     info we wouldn't be worried about anyway.
            '     This might change and I get to know more about debugging
            '     .net!

            Dim Tokn = syminfo.Value
            Dim si As LineMap.AssemblyLineMap.SymbolInfo = New LineMap.AssemblyLineMap.SymbolInfo(Left(syminfo.Name, syminfo.NameLen), syminfo.Address, Tokn)
            _alm.Symbols.Add(Tokn, si)
        End If

        '---- return this to the call to let it keep enumerating
        SymEnumSymbolsProc = -1
    End Function


    ''' <summary>
    ''' Handle the callback to enumerate all the Line numbers in the given DLL/EXE
    ''' based on info from the PDB
    ''' </summary>
    ''' <param name="srcinfo"></param>
    ''' <param name="UserContext"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function SymEnumLinesProc(ByRef srcinfo As SRCCODEINFO, ByVal UserContext As Integer) As Integer
        If srcinfo.LineNumber = &HFEEFEE Then
            '---- skip these entries
            '     I believe they mark the end of a set of linenumbers associated
            '     with a particular method, but they don't appear to contain
            '     valid line number info in any case.
        Else
            Try
                '---- add the new line number and it's address
                '		NOTE, this address is an IL Offset, not a native code offset
                Dim FileName() = srcinfo.FileName.Split("\"c)
                Dim Name As String
                Dim i = UBound(FileName)
                If i > 2 Then
                    Name = "...\" & FileName(i - 2) & "\" & FileName(i - 1) & "\" & FileName(i)
                ElseIf i > 1 Then
                    Name = "...\" & FileName(i - 1) & "\" & FileName(i)
                Else
                    Name = Path.GetFileName(srcinfo.FileName)
                End If
                _alm.AddAddressToLine(srcinfo.LineNumber, srcinfo.Address, Name, srcinfo.obj)

            Catch ex As Exception
                '---- catch everything because we DO NOT
                '     want to throw an exception from here!
            End Try
        End If
        '---- Tell the caller we succeeded
        SymEnumLinesProc = -1
    End Function
    Private Delegate Function SymEnumLinesCallback(ByRef srcinfo As SRCCODEINFO, ByVal UserContext As Integer) As Integer


#End Region
End Class

