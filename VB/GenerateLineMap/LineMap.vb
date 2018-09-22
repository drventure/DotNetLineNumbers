Imports System.Reflection
Imports System.Runtime.Serialization
Imports System.Runtime.InteropServices
Imports System.IO
Imports System.Xml


''' <summary>
''' Track LineMap information
''' for persistence purposes
''' (c) 2008-2009 Darin Higgins All Rights Reserved
''' </summary>
''' <remarks></remarks>
''' <editHistory></editHistory>
Public Class LineMap
    Private Const LINEMAPNAMESPACE As String = "http://schemas.linemap.net"


    ''' <summary>
    ''' No need for a constructor on this class
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub New()
    End Sub


#Region " API Declarations for working with Resources"
    Private Declare Function FindResourceEx Lib "kernel32" Alias "FindResourceExA" (ByVal hModule As Int32, <MarshalAs(UnmanagedType.LPStr)> ByVal lpType As String, <MarshalAs(UnmanagedType.LPStr)> ByVal lpName As String, ByVal wLanguage As Int16) As IntPtr
    Private Declare Function LoadResource Lib "kernel32" (ByVal hInstance As IntPtr, ByVal hResInfo As IntPtr) As IntPtr
    Private Declare Function LockResource Lib "kernel32" (ByVal hResData As IntPtr) As IntPtr
    Private Declare Function FreeResource Lib "kernel32" (ByVal hResData As IntPtr) As IntPtr
    Private Declare Function SizeofResource Lib "kernel32" (ByVal hInstance As IntPtr, ByVal hResInfo As IntPtr) As Int32

    Private Declare Ansi Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" ( _
         ByRef pvDest As Byte, _
         ByVal pvSrc As IntPtr, ByVal cbCopy As Int32)
#End Region


    ''' <summary>
    ''' The class simply defines several values we need to share between 
    ''' the GenerateLineMap utility and functions in the Err handler object
    ''' </summary>
    ''' <remarks></remarks>
    ''' <editHistory></editHistory>
    Friend Class LineMapKeys
        Public Shared ENCKEY As Byte() = New Byte(31) {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32}
        Public Shared ENCIV As Byte() = New Byte(15) {65, 2, 68, 26, 7, 178, 200, 3, 65, 110, 68, 13, 69, 16, 200, 219}
		Public Shared ResTypeName As String = "LNM"
		Public Shared ResName As String = "LNMDATA"
		Public Shared ResLang As Short = 0
    End Class


    ''' <summary>
    ''' Tracks symbol cache entries for all assemblies that appear in a stack frame
    ''' </summary>
    ''' <remarks></remarks>
    ''' <editHistory></editHistory>
    <DataContract(Namespace:=LINEMAPNAMESPACE)> _
    Public Class AssemblyLineMap
        ''' <summary>
        ''' Track the assembly this map is for
        ''' no need to persist this information
        ''' </summary>
        ''' <remarks></remarks>
        Public FileName As String

        <DataContract(Namespace:=LINEMAPNAMESPACE)> _
        Public Class AddressToLine
            '---- these members must get serialized
            <DataMember()> _
            Public Address As Int64
            <DataMember()> _
            Public Line As Int32
            <DataMember()> _
            Public SourceFileIndex As Integer
            <DataMember()> _
            Public ObjectNameIndex As Integer

            '---- these members do not need to be serialized
            Public SourceFile As String
            Public ObjectName As String


            ''' <summary>
            ''' Parameterless constructor for serialization
            ''' </summary>
            ''' <remarks></remarks>
            Public Sub New()
            End Sub


            Public Sub New(ByVal Line As Int32, ByVal Address As Int64, ByVal SourceFile As String, ByVal SourceFileIndex As Integer, ByVal ObjectName As String, ByVal ObjectNameIndex As Integer)
                Me.Line = Line
                Me.Address = Address
                Me.SourceFile = SourceFile
                Me.SourceFileIndex = SourceFileIndex
                Me.ObjectName = ObjectName
                Me.ObjectNameIndex = ObjectNameIndex
            End Sub
        End Class
        ''' <summary>
        ''' Track the Line number list enumerated from the PDB
        ''' Note, the list is already sorted
        ''' </summary>
        ''' <remarks></remarks>
        <DataMember()> _
        Public AddressToLineMap As List(Of AddressToLine) = New List(Of AddressToLine)


        ''' <summary>
        ''' Private class to track Symbols read from the PDB file
        ''' </summary>
        ''' <remarks></remarks>
        ''' <editHistory></editHistory>
        <DataContract(Namespace:=LINEMAPNAMESPACE)> _
        Public Class SymbolInfo
            '---- these need to be persisted
            <DataMember()> _
            Public Address As Long
            <DataMember()> _
            Public Token As Long

            '---- these aren't persisted
            Public Name As String


            ''' <summary>
            ''' Parameterless constructor for serialization
            ''' </summary>
            ''' <remarks></remarks>
            Public Sub New()
            End Sub


            Public Sub New(ByVal Name As String, ByVal Address As Long, ByVal Token As Long)
                Me.Name = Name
                Me.Address = Address
                Me.Token = Token
            End Sub
        End Class
        ''' <summary>
        ''' Track the Symbols enumerated from the PDB keyed by their token
        ''' </summary>
        ''' <remarks></remarks>
        <DataMember()> _
        Public Symbols As Dictionary(Of Long, SymbolInfo) = New Dictionary(Of Long, SymbolInfo)


        ''' <summary>
        ''' Track a list of string values
        ''' </summary>
        ''' <remarks></remarks>
        ''' <editHistory></editHistory>
        Public Class NamesList
            Inherits List(Of String)


            ''' <summary>
            ''' When adding names, if the name already exists in the list
            ''' don't bother to add it again
            ''' </summary>
            ''' <param name="Name"></param>
            ''' <returns></returns>
            Public Overloads Function Add(ByVal Name As String) As Integer
                Name = Name.ToLower
                Dim i = Me.IndexOf(Name)
                If i >= 0 Then Return i

                '---- gotta add the name
                MyBase.Add(Name)
                Return Me.Count - 1
            End Function


            ''' <summary>
            ''' Override this prop so that requested an index that doesn't exist just
            ''' returns a blank string
            ''' </summary>
            ''' <param name="Index"></param>
            ''' <value></value>
            ''' <remarks></remarks>
            Public Shadows Property Item(ByVal Index As Integer) As String
                Get
                    If Index >= 0 And Index < Me.Count Then
                        Return MyBase.Item(Index)
                    Else
                        Return String.Empty
                    End If
                End Get
                Set(ByVal value As String)
                    MyBase.Item(Index) = value
                End Set
            End Property
        End Class
        ''' <summary>
        ''' Tracks various string values in a flat list that is indexed into
        ''' </summary>
        ''' <remarks></remarks>
        <DataMember()> _
        Public Names As NamesList = New NamesList


        ''' <summary>
        ''' Create a new map based on an assembly filename
        ''' </summary>
        ''' <param name="FileName"></param>
        ''' <remarks></remarks>
        Public Sub New(ByVal FileName As String)
            Me.FileName = FileName
            Load()
        End Sub


        ''' <summary>
        ''' Parameterless constructor for serialization
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New()

        End Sub


        ''' <summary>
        ''' Clear out all internal information
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Clear()
            Me.Symbols.Clear()
            Me.AddressToLineMap.Clear()
            Me.Names.Clear()
        End Sub


        ''' <summary>
        ''' Load a LINEMAP file
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub Load()
			Transfer(Depersist(DecompressStream(DecryptStream(FileToStream(Me.FileName & ".lmp")))))
		End Sub


        ''' <summary>
        ''' Create a new assemblylinemap based on an assembly
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New(ByVal Assembly As Assembly)
            'LoadLineMapResource(Assembly.GetExecutingAssembly(), Symbols, Lines)

            Me.FileName = Assembly.CodeBase.Replace("file:///", "")

            Try
                '---- Get the hInstance of the indicated exe/dll image
                Dim curmodule = Assembly.GetLoadedModules()(0)
                Dim hInst = System.Runtime.InteropServices.Marshal.GetHINSTANCE(curmodule)

				'---- retrieve a handle to the Linemap resource
				'     Since it's a standard Win32 resource, the nice .NET resource functions
				'     can't be used
				'
				'     Important Note: The FindResourceEx function appears to be case
				'     sensitive in that you really HAVE to pass in UPPER CASE search
				'     arguments
				Dim hres = FindResourceEx(hInst.ToInt32(), LineMapKeys.ResTypeName, LineMapKeys.ResName, LineMapKeys.ResLang)

				'---- Load the resource to get it into memory
				Dim hresdata = LoadResource(hInst, hres)

                Dim lpdata As IntPtr = LockResource(hresdata)
                Dim sz = SizeofResource(hInst, hres)

                Dim bytes() As Byte
                If lpdata <> IntPtr.Zero And sz > 0 Then
                    '---- able to lock it,
                    '     so copy the data into a byte array
                    ReDim bytes(sz - 1)
                    CopyMemory(bytes(0), lpdata, sz)
                    FreeResource(hresdata)

                    '---- deserialize the symbol map and line num list
                    Using MemStream As System.IO.MemoryStream = New MemoryStream(bytes)
                        '---- release the byte array to free up the memory
                        Erase bytes
                        '---- and depersist the object
                        Transfer(Depersist(DecompressStream(DecryptStream(MemStream))))
                    End Using
                End If

            Catch ex As Exception
                '---- yes, it's bad form to catch all exceptions like this
                '     but this is part of an exception handler
                '     so it really can't be allowed to fail with an exception!
            End Try

            Try
                If Me.Symbols.Count = 0 Then
					'---- weren't able to load resources, so try the LINEMAP (LNM) file
					Load()
                End If
            Catch ex As Exception

            End Try
        End Sub


        ''' <summary>
        ''' Transfer a given AssemblyLineMap's contents to this one
        ''' </summary>
        ''' <param name="alm"></param>
        ''' <remarks></remarks>
        Private Sub Transfer(ByVal alm As AssemblyLineMap)
            '---- transfer Internal variables over
            Me.AddressToLineMap = alm.AddressToLineMap
            Me.Symbols = alm.Symbols
            Me.Names = alm.Names

            '---- debugging
            Debug.Print(Me.Symbols.Count.ToString)
            Debug.Print(Me.AddressToLineMap.Count.ToString)
            Debug.Print(Me.Names.Count.ToString)
        End Sub


        ''' <summary>
        ''' Read an entire file into a memory stream
        ''' </summary>
        ''' <param name="Filename"></param>
        ''' <returns></returns>
        Private Function FileToStream(ByVal Filename As String) As MemoryStream
            If File.Exists(Filename) Then
                Using FileStream As System.IO.FileStream = New System.IO.FileStream(Filename, FileMode.Open)
                    If FileStream.Length > 0 Then
                        Dim Buffer() As Byte
                        ReDim Buffer(CInt(FileStream.Length - 1))
                        FileStream.Read(Buffer, 0, CInt(FileStream.Length))
                        Return New MemoryStream(Buffer)
                    End If
                End Using
            End If

            '---- just return an empty stream
            Return New MemoryStream
        End Function


        ''' <summary>
        ''' Decrypt a stream based on fixed internal keys
        ''' </summary>
        ''' <param name="EncryptedStream"></param>
        ''' <returns></returns>
        Private Function DecryptStream(ByRef EncryptedStream As Stream) As MemoryStream

            Try
                Dim Enc As New System.Security.Cryptography.RijndaelManaged
                Enc.KeySize = 256
                '---- KEY is 32 byte array
                Enc.Key = LineMapKeys.ENCKEY
                '---- IV is 16 byte array
                Enc.IV = LineMapKeys.ENCIV

                Dim cryptoStream = New Security.Cryptography.CryptoStream(EncryptedStream, Enc.CreateDecryptor, Security.Cryptography.CryptoStreamMode.Read)

                Dim buf() As Byte
                ReDim buf(1023)
                Dim DecryptedStream As New MemoryStream()
                Do While EncryptedStream.Length > 0
                    Dim l = cryptoStream.Read(buf, 0, 1024)
                    If l = 0 Then Exit Do
                    If l < 1024 Then ReDim Preserve buf(l - 1)
                    DecryptedStream.Write(buf, 0, UBound(buf) + 1)
                    If l < 1024 Then Exit Do
                Loop
                DecryptedStream.Position = 0
                Return DecryptedStream

            Catch ex As Exception
                '---- any problems, nothing much to do, so return an empty stream
                Return New MemoryStream
            End Try
        End Function


        ''' <summary>
        ''' Uncompress a memory stream
        ''' </summary>
        ''' <param name="CompressedStream"></param>
        ''' <returns></returns>
        Private Function DecompressStream(ByRef CompressedStream As MemoryStream) As MemoryStream

            Dim GZip As New System.IO.Compression.GZipStream(CompressedStream, System.IO.Compression.CompressionMode.Decompress)

            Dim buf() As Byte
            ReDim buf(1023)
            Dim UncompressedStream As New MemoryStream()
            Do While CompressedStream.Length > 0
                Dim l = GZip.Read(buf, 0, 1024)
                If l = 0 Then Exit Do
                If l < 1024 Then ReDim Preserve buf(l - 1)
                UncompressedStream.Write(buf, 0, UBound(buf) + 1)
                If l < 1024 Then Exit Do
            Loop
            UncompressedStream.Position = 0
            Return UncompressedStream
        End Function


        Private Function Depersist(ByVal MemoryStream As MemoryStream) As AssemblyLineMap
            MemoryStream.Position = 0

            If MemoryStream.Length <> 0 Then
                Dim binaryDictionaryreader = XmlDictionaryReader.CreateBinaryReader(MemoryStream, New XmlDictionaryReaderQuotas())
                Dim serializer = New DataContractSerializer(GetType(AssemblyLineMap))
                Return DirectCast(serializer.ReadObject(binaryDictionaryreader), AssemblyLineMap)
            Else
                '---- jsut return an empty object
                Return New AssemblyLineMap
            End If
        End Function


        ''' <summary>
        ''' Stream this object to a memory stream and return it
        ''' All other persistence is handled externally because none of that 
        ''' is necessary under normal usage, only when generating a linemap
        ''' </summary>
        ''' <returns></returns>
        Public Function ToStream() As MemoryStream
            Dim MemStream As System.IO.MemoryStream = New System.IO.MemoryStream()

            Dim binaryDictionaryWriter = XmlDictionaryWriter.CreateBinaryWriter(MemStream)
            Dim serializer = New DataContractSerializer(GetType(AssemblyLineMap))
            serializer.WriteObject(binaryDictionaryWriter, Me)
            binaryDictionaryWriter.Flush()

            Return MemStream
        End Function


        ''' <summary>
        ''' Helper function to add an address to line map entry
        ''' </summary>
        ''' <param name="Line"></param>
        ''' <param name="Address"></param>
        ''' <param name="SourceFile"></param>
        ''' <param name="ObjectName"></param>
        ''' <remarks></remarks>
        Public Sub AddAddressToLine(ByVal Line As Int32, ByVal Address As Int64, ByVal SourceFile As String, ByVal ObjectName As String)
            Dim SourceFileIndex = Me.Names.Add(SourceFile)
            Dim ObjectNameIndex = Me.Names.Add(ObjectName)
            Dim atl = New AssemblyLineMap.AddressToLine(Line, Address, SourceFile, SourceFileIndex, ObjectName, ObjectNameIndex)
            Me.AddressToLineMap.Add(atl)
        End Sub
    End Class


    ''' <summary>
    ''' Define a collection of assemblylinemaps (one for each assembly we
    ''' need to generate a stack trace through)
    ''' </summary>
    ''' <remarks></remarks>
    ''' <editHistory></editHistory>
    Public Class AssemblyLineMapCollection
        Inherits Dictionary(Of String, AssemblyLineMap)

        ''' <summary>
        ''' Load a linemap file given the NAME of an assembly
        ''' obviously, the Assembly must exist.
        ''' </summary>
        ''' <param name="FileName"></param>
        ''' <remarks></remarks>
        Public Overloads Function Add(ByVal FileName As String) As AssemblyLineMap
            If Not File.Exists(FileName) Then
                Throw New FileNotFoundException("The file could not be found.", FileName)
            End If

            If Me.ContainsKey(FileName) Then
                '---- no need, already loaded (should it reload?)
                Return Me.Item(FileName)
            Else
                Dim alm = New AssemblyLineMap(FileName)
                Me.Add(FileName, alm)
                Return alm
            End If
        End Function


        Public Overloads Function Add(ByVal Assembly As Assembly) As AssemblyLineMap
            Dim FileName = Assembly.CodeBase
            If Me.ContainsKey(FileName) Then
                '---- no need, already loaded (should it reload?)
                Return Me.Item(FileName)
            Else
                Dim alm = New AssemblyLineMap(Assembly)
                Me.Add(FileName, alm)
                Return alm
            End If
        End Function

    End Class
    Public Shared AssemblyLineMaps As AssemblyLineMapCollection = New AssemblyLineMapCollection
End Class
