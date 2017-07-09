Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Diagnostics.SymbolStore
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Reflection
Imports System.Text


' Use MDbg's managed wrappers over the corysm.idl (diasymreader.dll) COM APIs
' Must reference MDbgCore.dll from the .NET SDK or corapi.dll from the MDbg sample: 
' http://www.microsoft.com/downloads/details.aspx?familyid=38449a42-6b7a-4e28-80ce-c55645ab1310&displaylang=en
Imports Microsoft.Samples.Debugging.CorSymbolStore


''' <summary>
''' A class for reading the PDB of a compiled .NET DLL/EXE and enumerating
''' certain debug information from it.
''' This class uses the CORSYM idl functions instead of the PDB debug enum functions
''' because this way, we can get ILoffsets and metadata tokens for the symbols
''' instead of actual addresses and symbols names, thus it should work even on 
''' obfuscated files.
'''
''' This technique is based on code for performing stack trace dumps originally written by 
''' Rick Byers - http://blogs.msdn.com/rmbyers
''' 
''' and that code itself was originally adapted from code at
''' http://blogs.msdn.com/jmstall/pages/sample-pdb2xml.aspx
''' </summary>
''' <remarks>
''' </remarks>
''' <editHistory></editHistory>
Public Class PDBReader

    'We could easily add other APIs similar to those available on StackTrace (and StackFrame)

    Private m_metadataDispenser As IMetaDataDispenser
    Private m_metadataimport As IMetaDataImport
    Private m_symBinder As SymbolBinder
    Private m_searchPath As String
    Private m_searchPolicy As SymSearchPolicies
    Private rSymReader As ISymbolReader

    'Map from module path to symbol reader
    Private m_symReaders As Dictionary(Of String, ISymbolReader) = New Dictionary(Of String, ISymbolReader)

    Declare Function CoCreateInstance Lib "ole32.dll" (ByVal rclsid As String, _
                                                       <MarshalAs(UnmanagedType.IUnknown)> ByRef pUnkOuter As Object, _
                                                       ByVal dwClsContext As UInt32, _
                                                       ByVal riid As String) As Object




    ''' <summary>
    ''' Create a new instance and automatically setup some internal vars
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New(ByVal FileName As String)
        'Create a COM Metadata dispenser to use for all modules
        Dim dispenserClassID = New Guid("{0xE5CB7A31, 0x7512, 0x11D2, {0x89, 0xCE, 0x0, 0x80, 0xC7, 0x92, 0xE5, 0xD8}}") 'CLSID_CorMetaDataDispenser
        Dim dispenserIID = New Guid("{0x809C652E, 0x7396, 0x11D2, {0x97, 0x71, 0x0, 0xA0, 0xC9, 0xB4, 0xD5, 0xC}}") ' IID_IMetaDataDispenser

        Dim obj As Object = Activator.CreateInstance(Type.GetTypeFromCLSID(dispenserClassID))
        m_metadataDispenser = DirectCast(obj, IMetaDataDispenser)

        'm_metadataDispenser = CoCreateInstance(dispenserClassID, Nothing, 1, dispenserIID)
        'm_metadataDispenser = (IMetaDataDispenser)objDispenser;

        'm_metadataDispenser = CreateObject("CLRMetadata.CorMetadataDispenser")

        '---- we'll just use this object latebound
        ' = DirectCast(objDispenser, IMetaDataDispenser)

        'Create a binder from MDbg's wrappers over ISymUnmanagedBinder2
        m_symBinder = New SymbolBinder()

        rSymReader = pGetSymbolReaderForFile(FileName)
        If (rSymReader Is Nothing) Then
            Throw New FileNotFoundException("Could not load symbols for file.", FileName)
        End If
    End Sub


    ''' <summary>
    ''' Create a symbol reader object corresponding to the specified module (DLL/EXE)
    ''' </summary>
    ''' <param name="Filename">Full path to the module of interest</param>
    ''' <returns>A symbol reader object, or null if no matching PDB symbols can located</returns>
    Private Function pCreateSymbolReaderForFile(ByVal Filename As String) As ISymbolReader

        'First we need to get a metadata importer for the module to provide to the symbol reader
        'This is basically the same as MDbg's SymbolAccess.GetReaderForFile method, except that it
        'unfortunately does not have an overload that allows us to provide the searchPolicies
        Dim importerIID = New Guid("{0x7DAC8207, 0xD3AE, 0x4C75,{0x9B,0x67,0x92,0x80,0x1A,0x49,0x7D,0x44}}") ' IID_IMetaDataImport

        'Open an Importer on the given filename. We'll end up passing this importer straight
        'through to the Binder.
        Dim objImporter As Object = Nothing
        'CallByName(m_metadataDispenser, "OpenScope", CallType.Method, Filename, 0, importerIID.ToByteArray, objImporter)
        m_metadataDispenser.OpenScope(Filename, 0, importerIID, objImporter)
        m_metadataimport = DirectCast(objImporter, IMetaDataImport)

        'Call ISymUnmanagedBinder2.GetReaderForFile2 to load the PDB file (if any)
        'Note that ultimately how this PDB file is located is determined by
        'IDiaDataSource::loadDataForExe.  See the DIA SDK documentation for details.
        Dim reader = m_symBinder.GetReaderForFile(m_metadataimport, Filename, Path.GetDirectoryName(Filename), SymSearchPolicies.AllowOriginalPathAccess)
        Return reader
    End Function


    ''' <summary>
    ''' Get or create a symbol reader for the specified module (caching the result)
    ''' </summary>
    ''' <param name="modulePath">Full path to the module of interest</param>
    ''' <returns>A symbol reader for the specified module or null if none could be found</returns>
    Private Function pGetSymbolReaderForFile(ByVal modulePath As String) As ISymbolReader
        Dim reader As ISymbolReader

        If m_symReaders.ContainsKey(modulePath) Then
            Return m_symReaders(modulePath)
        Else
            reader = pCreateSymbolReaderForFile(modulePath)
            m_symReaders.Add(modulePath, reader)
        End If

        Return reader
    End Function


    ''' <summary>
    ''' Get a texual representing of the supplied stack trace including source file names
    ''' and line numbers, using the PDB lookup options supplied at construction.
    ''' </summary>
    ''' <param name="stackTrace">The stack trace to convert to text</param>
    ''' <returns>
    ''' A string in a format similar to StackTrace.ToString but whith file names and
    ''' line numbers even when they're not available to the built-in StackTrace class.
    ''' </returns>
    Public Function StackTraceToStringWithSourceInfo(ByVal stackTrace As StackTrace) As String
        Dim sb = New System.Text.StringBuilder()

        For Each stackFrame In stackTrace.GetFrames()

            Dim method = stackFrame.GetMethod()

            'Format the stack trace line similarily to how the built-in StackTrace class does.
            'Some differences (simplifications here): generics, nested types, argument names
            Dim methodString = method.ToString()   ' this is "RetType FuncName(args)
            Dim sig = String.Format("  at {0}.{1}", method.DeclaringType.FullName, methodString.Substring(methodString.IndexOf(" ") + 1))

            'Append source location information if we can find PDBs
            Dim sourceLoc = pGetSourceLoc(method, stackFrame.GetILOffset())
            If (sourceLoc IsNot Nothing) Then
                sig += " in " + sourceLoc
            End If

            sb.AppendLine(sig)
        Next
        Return sb.ToString()
    End Function


    ''' <summary>
    ''' Get a string representing the source location for the given IL offset and method
    ''' </summary>
    ''' <param name="method">The method of interest</param>
    ''' <param name="ilOffset">The offset into the IL</param>
    ''' <returns>
    ''' A string of the format [filepath]:[line] (eg. "C:\temp\foo.cs:123"), or null
    ''' if a matching PDB couldn't be found
    ''' </returns>
    Private Function pGetSourceLoc(ByVal Method As MethodBase, ByVal ilOffset As Integer) As String

        'Get the symbol reader corresponding to the module of the supplied method
        Dim modulePath = Method.Module.FullyQualifiedName
        Dim symReader = pGetSymbolReaderForFile(modulePath)
        If (symReader Is Nothing) Then
            Return Nothing   ' no matching PDB found
        End If

        Dim symMethod = symReader.GetMethod(New SymbolToken(Method.MetadataToken))

        'Get all the sequence points for the method
        Dim docs(symMethod.SequencePointCount) As ISymbolDocument
        Dim lineNumbers(symMethod.SequencePointCount) As Integer
        Dim ilOffsets(symMethod.SequencePointCount) As Integer
        symMethod.GetSequencePoints(ilOffsets, docs, lineNumbers, Nothing, Nothing, Nothing)

        'Find the closest sequence point to the requested offset
        'Sequence points are returned sorted by offset so we're looking for the last one with
        'an offset less than or equal to the requested offset. 
        'Note that this won't necessarily match the real source location exactly if 
        'the code was jit-compiled with optimizations.
        Dim i As Integer
        For i = 0 To symMethod.SequencePointCount - 1
            If (ilOffsets(i) > ilOffset) Then
                Exit For
            End If
        Next
        'Found the first mismatch, back up if it wasn't the first
        If (i > 0) Then
            i -= 1
        End If

        'Now return the source file and line number for this sequence point
        Return String.Format("{0}:{1}", docs(i).URL, lineNumbers(i))
    End Function


    Public Function GetMethodLines() As Dictionary(Of Integer, KeyValuePair(Of Integer, Integer))

        Dim enumhandle As UInteger = 0


        'Get the symbol reader corresponding to the module of the supplied method
        For Each doc In rSymReader.GetDocuments
            'Dim MaxLine = doc.FindClosestLine(3000)
            Dim lastmethod As ISymbolMethod = Nothing
            For l As Integer = 1 To CInt(2 ^ 30)
                Dim method As ISymbolMethod = Nothing
                Try
                    method = rSymReader.GetMethodFromDocumentPosition(doc, l, 1)
                Catch ex As Exception
                    Exit For
                End Try
                If Not method.Equals(lastmethod) Then
                    lastmethod = method

                    'Get all the sequence points for the method
                    Dim docs(method.SequencePointCount) As ISymbolDocument
                    Dim lineNumbers(method.SequencePointCount) As Integer
                    Dim ilOffsets(method.SequencePointCount) As Integer
                    method.GetSequencePoints(ilOffsets, docs, lineNumbers, Nothing, Nothing, Nothing)

                    'Find the closest sequence point to the requested offset
                    'Sequence points are returned sorted by offset so we're looking for the last one with
                    'an offset less than or equal to the requested offset. 
                    'Note that this won't necessarily match the real source location exactly if 
                    'the code was jit-compiled with optimizations.
                    Dim i As Integer
                    For i = 0 To method.SequencePointCount - 1
                        Debug.Print("line=" & lineNumbers(i).ToString & "   ILOFFSET=" & ilOffsets(i).ToString)
                    Next
                End If
            Next
        Next
        Return Nothing
    End Function


#Region " Metadata Imports"

    'Bare bones COM-interop definition of the IMetaDataDispenser API
    <Guid("809c652e-7396-11d2-9771-00a0c9b4d50c"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown), _
     ComVisible(True)> _
    Private Interface IMetaDataDispenser

        '      'We need to be able to call OpenScope, which is the 2nd vtable slot.
        '      'Thus we need this one placeholder here to occupy the first slot..
        Sub DefineScope_Placeholder()

        '      'STDMETHOD(OpenScope)(                   // Return code.
        '      'LPCWSTR     szScope,                // [in] The scope to open.
        '      '  DWORD       dwOpenFlags,            // [in] Open mode flags.
        '      '  REFIID      riid,                   // [in] The interface desired.
        '      '  IUnknown    **ppIUnk) PURE;         // [out] Return interface on success.
        Sub OpenScope(<MarshalAs(UnmanagedType.LPWStr)> ByVal szScope As String, ByVal dwOpenFlags As Int32, ByRef riid As Guid, <MarshalAs(UnmanagedType.IUnknown)> ByRef punk As Object)

        '      'Don't need any other methods.
    End Interface


    ''   'Since we're just blindly passing this interface through managed code to the Symbinder, we don't care about actually
    ''   'importing the specific methods.
    ''   'This needs to be public so that we can call Marshal.GetComInterfaceForObject() on it to get the
    ''   'underlying metadata pointer.
    '<Guid("7DAC8207-D3AE-4c75-9B67-92801A497D44"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown), _
    ' ComVisible(True)> _
    'Public Interface IMetadataImport

    '   'Just need a single placeholder method so that it doesn't complain about an empty interface.
    '   Sub Placeholder()
    'End Interface


    ''' <summary>
    ''' Original Version of the IMetadataImport interface in C# from
    ''' http://pzsolt.blogspot.com/2005/01/reading-types-from-assembly.html
    ''' Converted to VB via
    ''' http://labs.developerfusion.co.uk/convert/csharp-to-vb.aspx
    ''' </summary>
    ''' <remarks></remarks>
    ''' <editHistory></editHistory>
    <ComImport(), GuidAttribute("7DAC8207-D3AE-4c75-9B67-92801A497D44"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)> _
    Public Interface IMetaDataImport
        Sub CloseEnum(ByVal hEnum As UInteger)

        Function CountEnum(ByVal hEnum As UInteger, ByRef count As UInteger) As UInteger

        Function ResetEnum(ByVal hEnum As UInteger, ByVal ulPos As UInteger) As UInteger

        Function EnumTypeDefs(ByRef phEnum As UInteger, _
                              <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=2)> _
                              ByVal rTypeDefs As UInteger(), ByVal cMax As UInteger, ByRef pcTypeDefs As UInteger) As UInteger

        Function EnumInterfaceImpls(ByRef phEnum As UInteger, _
                                    ByVal td As UInteger, _
                                    <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=3)> _
                                    ByVal rImpls As UInteger(), ByVal cMax As UInteger, ByRef pcImpls As UInteger) As UInteger

        Function EnumTypeRefs(ByRef phEnum As UInteger, _
                              <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=2)> _
                              ByVal rTypeDefs As UInteger(), ByVal cMax As UInteger, ByRef pcTypeRefs As UInteger) As UInteger

        Function FindTypeDefByName(<MarshalAs(UnmanagedType.LPWStr)> _
                                   ByVal szTypeDef As String, ByVal tkEnclosingClass As UInteger, ByRef ptd As UInteger) As UInteger

        Function GetScopeProps(<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=0)> _
                               ByVal szName As Char(), ByVal cchName As UInteger, ByRef pchName As UInteger, ByRef pmvid As Guid) As UInteger

        Function GetModuleFromScope(ByRef pmd As UInteger) As UInteger

        Function GetTypeDefProps(ByVal td As UInteger, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=2)> _
                                 ByVal szTypeDef As Char(), ByVal cchTypeDef As UInteger, ByRef pchTypeDef As UInteger, ByRef pdwTypeDefFlags As UInteger, ByRef ptkExtends As UInteger) As UInteger

        Function GetInterfaceImplProps(ByVal iiImpl As UInteger, ByRef pClass As UInteger, ByRef ptkIface As UInteger) As UInteger

        Function GetTypeRefProps(ByVal tr As UInteger, ByRef ptkResolutionScope As UInteger, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=3)> _
                                 ByVal szName As Char(), ByVal cchName As UInteger, ByRef pchName As UInteger) As UInteger

        Function ResolveTypeRef(ByVal tr As UInteger, ByRef riid As Guid, <MarshalAs(UnmanagedType.[Interface])> _
                                ByRef ppIScope As Object, ByRef ptd As UInteger) As UInteger

        Function EnumMembers(ByRef phEnum As UInteger, ByVal cl As UInteger, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=3)> _
                             ByVal rMembers As UInteger(), ByVal cMax As UInteger, ByRef pcTokens As UInteger) As UInteger

        Function EnumMembersWithName(ByRef phEnum As UInteger, ByVal cl As UInteger, <MarshalAs(UnmanagedType.LPWStr)> _
                                     ByVal szName As String, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=4)> _
                                     ByVal rMembers As UInteger(), ByVal cMax As UInteger, ByRef pcTokens As UInteger) As UInteger

        Function EnumMethods(ByRef phEnum As UInteger, ByVal cl As UInteger, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=3)> _
                             ByVal rMethods As UInteger(), ByVal cMax As UInteger, ByRef pcTokens As UInteger) As UInteger

        Function EnumMethodsWithName(ByRef phEnum As UInteger, ByVal cl As UInteger, <MarshalAs(UnmanagedType.LPWStr)> _
                                     ByVal szName As String, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=4)> _
                                     ByVal rMethods As UInteger(), ByVal cMax As UInteger, ByRef pcTokens As UInteger) As UInteger

        Function EnumFields(ByRef phEnum As UInteger, ByVal cl As UInteger, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=3)> _
                            ByVal rFields As UInteger(), ByVal cMax As UInteger, ByRef pcTokens As UInteger) As UInteger

        Function EnumFieldsWithName(ByRef phEnum As UInteger, ByVal cl As UInteger, <MarshalAs(UnmanagedType.LPWStr)> _
                                    ByVal szName As String, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=4)> _
                                    ByVal rFields As UInteger(), ByVal cMax As UInteger, ByRef pcTokens As UInteger) As UInteger

        Function EnumParams(ByRef phEnum As UInteger, ByVal mb As UInteger, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=3)> _
                            ByVal rParams As UInteger(), ByVal cMax As UInteger, ByRef pcTokens As UInteger) As UInteger

        Function EnumMemberRefs(ByRef phEnum As UInteger, ByVal tkParent As UInteger, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=3)> _
                                ByVal rMemberRefs As UInteger(), ByVal cMax As UInteger, ByRef pcTokens As UInteger) As UInteger

        Function EnumMethodImpls(ByRef phEnum As UInteger, ByVal td As UInteger, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=2)> _
                                 ByVal rMethodBody As UInteger(), <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=3)> _
                                 ByVal rMethodDecl As UInteger(), ByVal cMax As UInteger, ByRef pcTokens As UInteger) As UInteger

        Function EnumPermissionSets(ByRef phEnum As UInteger, ByVal tk As UInteger, ByVal dwActions As UInteger, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=3)> _
                                    ByVal rPermission As UInteger(), ByVal cMax As UInteger, ByRef pcTokens As UInteger) As UInteger

        Function FindMember(ByVal td As UInteger, <MarshalAs(UnmanagedType.LPWStr)> _
                            ByVal szName As String, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=2)> _
                            ByVal pvSigBlob As Byte(), ByVal cbSigBlob As UInteger, ByRef pmb As UInteger) As UInteger

        Function FindMethod(ByVal td As UInteger, <MarshalAs(UnmanagedType.LPWStr)> _
                            ByVal szName As String, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=2)> _
                            ByVal pvSigBlob As Byte(), ByVal cbSigBlob As UInteger, ByRef pmb As UInteger) As UInteger

        Function FindField(ByVal td As UInteger, <MarshalAs(UnmanagedType.LPWStr)> _
                           ByVal szName As String, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=2)> _
                           ByVal pvSigBlob As Byte(), ByVal cbSigBlob As UInteger, ByRef pmb As UInteger) As UInteger

        Function FindMemberRef(ByVal td As UInteger, <MarshalAs(UnmanagedType.LPWStr)> _
                               ByVal szName As String, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=2)> _
                               ByVal pvSigBlob As Byte(), ByVal cbSigBlob As Integer, ByRef pmr As UInteger) As UInteger

        Function GetMethodProps(ByVal mb As UInteger, ByRef pClass As UInteger, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=2)> _
                                ByVal szMethod As Char(), ByVal cchMethod As UInteger, ByRef pchMethod As UInteger, ByRef pdwAttr As UInteger, _
                                ByRef ppvSigBlob As IntPtr, ByRef pcbSigBlob As UInteger, ByRef pulCodeRVA As UInteger, ByRef pdwImplFlags As UInteger) As UInteger

        Function GetMemberRefProps(ByVal mr As UInteger, ByRef ptk As UInteger, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=2)> _
                                   ByVal szMember As Char(), ByVal cchMember As UInteger, ByRef pchMember As UInteger, ByRef ppvSigBlob As IntPtr, _
                                   ByRef pbSigBlob As UInteger) As UInteger

        Function EnumProperties(ByRef phEnum As UInteger, ByVal td As UInteger, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=2)> _
                                ByVal rProperties As UInteger(), ByVal cMax As UInteger, ByRef pcProperties As UInteger) As UInteger

        Function EnumEvents(ByRef phEnum As UInteger, ByVal td As UInteger, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=2)> _
                            ByVal rEvents As UInteger(), ByVal cMax As UInteger, ByRef pcEvents As UInteger) As UInteger

        Function GetEventProps(ByVal ev As UInteger, ByRef pClass As UInteger, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=2)> _
                               ByVal szEvent As Char(), ByVal cchEvent As UInteger, ByRef pchEvent As UInteger, ByRef pdwEventFlags As UInteger, _
                               ByRef ptkEventType As UInteger, ByRef pmdAddOn As UInteger, ByRef pmdRemoveOn As UInteger, ByRef pmdFire As UInteger, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=10)> _
                               ByVal rmdOtherMethod As UInteger(), ByVal cMax As UInteger, _
                               ByRef pcOtherMethod As UInteger) As UInteger

        Function EnumMethodSemantics(ByRef phEnum As UInteger, ByVal mb As UInteger, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=2)> _
                                     ByVal rEventProp As UInteger(), ByVal cMax As UInteger, ByRef pcEventProp As UInteger) As UInteger

        Function GetMethodSemantics(ByVal mb As UInteger, ByVal tkEventProp As UInteger, ByRef pdwSemanticsFlags As UInteger) As UInteger

        Function GetClassLayout(ByVal td As UInteger, ByRef pdwPackSize As UInteger, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=2)> _
                                ByVal rFieldOffset As Long(), ByVal cMax As UInteger, ByRef pcFieldOffset As UInteger, ByRef pulClassSize As UInteger) As UInteger

        Function GetFieldMarshal(ByVal tk As UInteger, ByRef ppvNativeType As IntPtr, ByRef pcbNativeType As UInteger) As UInteger

        Function GetRVA(ByVal tk As UInteger, ByRef pulCodeRVA As UInteger, ByRef pdwImplFlags As UInteger) As UInteger

        Function GetPermissionSetProps(ByVal pm As UInteger, ByRef pdwAction As UInteger, ByRef ppvPermission As IntPtr, ByRef pcbPermission As UInteger) As UInteger

        Function GetSigFromToken(ByVal mdSig As UInteger, ByRef ppvSig As IntPtr, ByRef pcbSig As UInteger) As UInteger

        Function GetModuleRefProps(ByVal mur As UInteger, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> _
                                   ByVal szName As Char(), ByVal cchName As UInteger, ByRef pchName As UInteger) As UInteger

        Function EnumModuleRefs(ByRef phEnum As UInteger, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> _
                                ByVal rModuleRefs As UInteger(), ByVal cmax As UInteger, ByRef pcModuleRefs As UInteger) As UInteger

        Function GetTypeSpecFromToken(ByVal typespec As UInteger, ByRef ppvSig As IntPtr, ByRef pcbSig As UInteger) As UInteger

        Function GetNameFromToken(ByVal tk As UInteger, ByRef pszUtf8NamePtr As IntPtr) As UInteger

        Function EnumUnresolvedMethods(ByRef phEnum As UInteger, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> _
                                       ByVal rMethods As UInteger(), ByVal cMax As UInteger, ByRef pcTokens As UInteger) As UInteger

        Function GetUserString(ByVal stk As UInteger, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> _
                               ByVal szString As Char(), ByVal cchString As UInteger, ByRef pchString As UInteger) As UInteger

        Function GetPinvokeMap(ByVal tk As UInteger, ByRef pdwMappingFlags As UInteger, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=2)> _
                               ByVal szImportName As Char(), ByVal cchImportName As UInteger, ByRef pchImportName As UInteger, ByRef pmrImportDLL As UInteger) As UInteger

        Function EnumSignatures(ByRef phEnum As UInteger, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> _
                                ByVal rSignatures As UInteger(), ByVal cmax As UInteger, ByRef pcSignatures As UInteger) As UInteger

        Function EnumTypeSpecs(ByRef phEnum As UInteger, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> _
                               ByVal rTypeSpecs As UInteger(), ByVal cmax As UInteger, ByRef pcTypeSpecs As UInteger) As UInteger

        Function EnumUserStrings(ByRef phEnum As UInteger, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> _
                                 ByVal rStrings As UInteger(), ByVal cmax As UInteger, ByRef pcStrings As UInteger) As UInteger

        Function GetParamForMethodIndex(ByVal md As UInteger, ByVal ulParamSeq As UInteger, ByRef ppd As UInteger) As UInteger

        Function EnumCustomAttributes(ByRef phEnum As UInteger, ByVal tk As UInteger, ByVal tkType As UInteger, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=3)> _
                                      ByVal rCustomAttributes As UInteger(), ByVal cMax As UInteger, ByRef pcCustomAttributes As UInteger) As UInteger

        Function GetCustomAttributeProps(ByVal cv As UInteger, ByRef ptkObj As UInteger, ByRef ptkType As UInteger, ByRef ppBlob As IntPtr, ByRef pcbSize As UInteger) As UInteger

        Function FindTypeRef(ByVal tkResolutionScope As UInteger, <MarshalAs(UnmanagedType.LPWStr)> _
                             ByVal szName As String, ByRef ptr As UInteger) As UInteger

        Function GetMemberProps(ByVal mb As UInteger, ByRef pClass As UInteger, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=2)> _
                                ByVal szMember As Char(), ByVal cchMember As UInteger, ByRef pchMember As UInteger, ByRef pdwAttr As UInteger, _
                                ByRef ppvSigBlob As IntPtr, ByRef pcbSigBlob As UInteger, ByRef pulCodeRVA As UInteger, ByRef pdwImplFlags As UInteger, ByRef pdwCPlusTypeFlag As UInteger, ByRef ppValue As IntPtr, _
                                ByRef pcchValue As UInteger) As UInteger

        Function GetFieldProps(ByVal mb As UInteger, ByRef pClass As UInteger, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=2)> _
                               ByVal szField As Char(), ByVal cchField As UInteger, ByRef pchField As UInteger, ByRef pdwAttr As UInteger, _
                               ByRef ppvSigBlob As IntPtr, ByRef pcbSigBlob As UInteger, ByRef pdwCPlusTypeFlag As UInteger, ByRef ppValue As IntPtr, ByRef pcchValue As UInteger) As UInteger

        Function GetPropertyProps(ByVal prop As UInteger, ByRef pClass As UInteger, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=2)> _
                                  ByVal szProperty As Char(), ByVal cchProperty As UInteger, ByRef pchProperty As UInteger, ByRef pdwPropFlags As UInteger, _
                                  ByRef ppvSig As IntPtr, ByRef pbSig As UInteger, ByRef pdwCPlusTypeFlag As UInteger, ByRef ppDefaultValue As IntPtr, ByRef pcchDefaultValue As UInteger, ByRef pmdSetter As UInteger, _
                                  ByRef pmdGetter As UInteger, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=13)> _
                                  ByVal rmdOtherMethod As UInteger(), ByVal cMax As UInteger, ByRef pcOtherMethod As UInteger) As UInteger

        Function GetParamProps(ByVal tk As UInteger, ByRef pmd As UInteger, ByRef pulSequence As UInteger, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=3)> _
                               ByVal szName As Char(), ByVal cchName As UInteger, ByRef pchName As UInteger, _
                               ByRef pdwAttr As UInteger, ByRef pdwCPlusTypeFlag As UInteger, ByRef ppValue As IntPtr, ByRef pcchValue As UInteger) As UInteger

        Function GetCustomAttributeByName(ByVal tkObj As UInteger, <MarshalAs(UnmanagedType.LPWStr)> _
                                          ByVal szName As String, ByRef ppData As IntPtr, ByRef pcbData As UInteger) As UInteger

        Function IsValidToken(ByVal tk As UInteger) As Boolean

        Function GetNestedClassProps(ByVal tdNestedClass As UInteger, ByRef ptdEnclosingClass As UInteger) As UInteger

        Function GetNativeCallConvFromSig(ByVal pvSig As IntPtr, ByVal cbSig As UInteger, ByRef pCallConv As UInteger) As UInteger

        Function IsGlobal(ByVal pd As UInteger, ByRef pbGlobal As UInteger) As UInteger
    End Interface

#End Region

End Class
