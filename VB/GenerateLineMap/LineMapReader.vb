Imports System.IO
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Runtime.Serialization
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.Security.Cryptography


Public Class LineMapKeys
    Public Shared ENCKEY As Byte() = New Byte(31) {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32}
    Public Shared ENCIV As Byte() = New Byte(15) {65, 2, 68, 26, 7, 178, 200, 3, 65, 110, 68, 13, 69, 16, 200, 219}
    Public Shared ResTypeName As String = "LINEMAP"
    Public Shared ResName As String = "LINEMAPDATA"
    Public Shared ResLang As Integer = 0
End Class


Public Class LineMapReader
    Private Declare Function FindResourceEx Lib "kernel32" Alias "FindResourceExA" (ByVal hModule As Int32, <MarshalAs(UnmanagedType.LPStr)> ByVal lpType As String, <MarshalAs(UnmanagedType.LPStr)> ByVal lpName As String, ByVal wLanguage As Int16) As Int32
    'private Declare Function FindResource Lib "kernel32" Alias "FindResourceA" (ByVal hModule As Int32, <MarshalAs(UnmanagedType.LPStr)> ByVal lpName As String, <MarshalAs(UnmanagedType.LPStr)> ByVal lpType As String) As Int32
    'private Declare Function FindResource Lib "kernel32" Alias "FindResourceA" (ByVal hModule As Int32, <MarshalAs(UnmanagedType.LPStr)> ByVal lpName As String, ByVal lpType As Int32) As Int32
    Private Declare Function LoadResource Lib "kernel32" (ByVal hInstance As Int32, ByVal hResInfo As Int32) As Int32
    Private Declare Function LockResource Lib "kernel32" (ByVal hResData As Int32) As Int32
    Private Declare Function FreeResource Lib "kernel32" (ByVal hResData As Int32) As Int32
    Private Declare Function SizeofResource Lib "kernel32" (ByVal hInstance As Int32, ByVal hResInfo As Int32) As Int32

    Private Declare Ansi Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" ( _
       ByRef pvDest As Byte, _
       ByVal pvSrc As IntPtr, ByVal cbCopy As Int32)

    ''' <summary>
    ''' Load a linemap file into the Symbols and Lines lists
    ''' </summary>
    ''' <param name="FileName"></param>
    ''' <param name="Symbols"></param>
    ''' <param name="Lines"></param>
    ''' <remarks></remarks>
    Public Shared Sub LoadLineMapFile(ByVal FileName As String, ByRef Symbols As Dictionary(Of Long, Int64), ByRef Lines As List(Of KeyValuePair(Of Int64, Int32)))
        If Not System.IO.File.Exists(FileName) Then
            Throw New FileNotFoundException("The file could not be found.", FileName)
        End If

        '---- write the uncompressed stream out to a file (for debugging mainly)
        'pStreamToFile(FileName & ".linemap", pSymbolsToStream(Symbols, Lines))

        '---- compressed and encrypt the stream
        '     technically, I should probably CLOSE and DISPOSE the streams
        '     but this is just a quick and dirty tool
        '     TODO Is this "functional programming?!?"
        pStreamToSymbols(pDecompressStream(pDecryptStream(pFileToStream(FileName & ".linemap"))), Symbols, Lines)

        Debug.Print(Symbols.Count.ToString)
        Debug.Print(Lines.Count.ToString)
    End Sub


    ''' <summary>
    ''' Retrieves the linemap from the LINEMAP resource in the current 
    ''' assembly
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared Sub LoadLineMapResource(ByRef Symbols As Dictionary(Of Long, Int64), ByRef Lines As List(Of KeyValuePair(Of Int64, Int32)))

        '---- if symbols are already loaded, no need to reload them
        If Symbols IsNot Nothing Then
            If Symbols.Count > 0 Then Exit Sub
        End If

        Try
            '---- Get the hInstance of the current exe/dll image
            Dim curmodule = Assembly.GetExecutingAssembly().GetLoadedModules()(0)
            Dim hInst = System.Runtime.InteropServices.Marshal.GetHINSTANCE(curmodule)

            '---- retrieve a handle to the Linemap resource
            '     Since it's a standard Win32 resource, we can't use the nice
            '     .NET functions.
            '
            '     Important Note: The FindResourceEx function appears to be case
            '     sensitive in that you really HAVE to pass in UPPER CASE search
            '     arguments
            '
            '     The routine that 
            Dim hres = FindResourceEx(hInst.ToInt32, LineMapKeys.ResTypeName, LineMapKeys.ResName, CShort(LineMapKeys.ResLang))

            '---- Load the resource to get it into memory
            Dim hresdata = LoadResource(hInst.ToInt32, hres)

            Dim lpdata = New IntPtr(LockResource(hresdata))
            Dim sz = SizeofResource(hInst.ToInt32, hres)

            Dim bytes() As Byte
            If lpdata.ToInt32 <> 0 And sz > 0 Then
                '---- looks like we were able to lock it
                '     so copy the data into a byte array
                ReDim bytes(sz - 1)
                CopyMemory(bytes(0), lpdata, sz)
                FreeResource(hresdata)

                '---- deserialize the symbol map and line num list
                Using MemStream As System.IO.MemoryStream = New MemoryStream(bytes)
                    '---- release the byte array to free up the memory
                    Erase bytes
                    '---- and depersist the symbol and lines lists
                    pStreamToSymbols(pDecompressStream(pDecryptStream(MemStream)), Symbols, Lines)
                End Using
            End If

        Catch ex As Exception
            '---- yes, it's bad form to catch all exceptions like this
            '     but this is part of an exception handler
            '     so we can't really let it fail with an exception!
        End Try
    End Sub



    Private Shared Sub pStreamToSymbols(ByVal MemoryStream As MemoryStream, ByRef Symbols As Dictionary(Of Long, Int64), ByRef Lines As List(Of KeyValuePair(Of Int64, Int32)))
        Dim bf As System.Runtime.Serialization.Formatters.Binary.BinaryFormatter = New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter

        '---- write out the symbol dictionary
        MemoryStream.Position = 0
        Symbols = DirectCast(bf.Deserialize(MemoryStream), Dictionary(Of Long, Int64))

        '---- then the line num list
        Lines = DirectCast(bf.Deserialize(MemoryStream), List(Of KeyValuePair(Of Int64, Int32)))
    End Sub


    ''' <summary>
    ''' Read an entire file into a memory stream
    ''' </summary>
    ''' <param name="Filename"></param>
    ''' <returns></returns>
    Private Shared Function pFileToStream(ByVal Filename As String) As MemoryStream

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
    Private Shared Function pDecryptStream(ByRef EncryptedStream As Stream) As MemoryStream

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
        Do
            Dim l = cryptoStream.Read(buf, 0, 1024)
            If l = 0 Then Exit Do
            If l < 1024 Then ReDim Preserve buf(l - 1)
            DecryptedStream.Write(buf, 0, UBound(buf) + 1)
            If l < 1024 Then Exit Do
        Loop
        DecryptedStream.Position = 0
        Return DecryptedStream
    End Function


    ''' <summary>
    ''' Uncompress a memory stream
    ''' </summary>
    ''' <param name="CompressedStream"></param>
    ''' <returns></returns>
    Private Shared Function pDecompressStream(ByRef CompressedStream As MemoryStream) As MemoryStream

        Dim GZip As New System.IO.Compression.GZipStream(CompressedStream, System.IO.Compression.CompressionMode.Decompress)

        Dim buf() As Byte
        ReDim buf(1023)
        Dim UncompressedStream As New MemoryStream()
        Do
            Dim l = GZip.Read(buf, 0, 1024)
            If l = 0 Then Exit Do
            If l < 1024 Then ReDim Preserve buf(l - 1)
            UncompressedStream.Write(buf, 0, UBound(buf) + 1)
            If l < 1024 Then Exit Do
        Loop
        UncompressedStream.Position = 0
        Return UncompressedStream
    End Function

End Class
