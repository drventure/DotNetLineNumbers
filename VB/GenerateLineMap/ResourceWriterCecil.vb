Imports Mono.Cecil
Imports System.Resources


''' <summary>
''' 
''' </summary>
''' <remarks></remarks>
Public Class ResourceWriterCecil
    Private _Inited As Boolean
    Private _Asm As AssemblyDefinition
    Private _Resources As Mono.Collections.Generic.Collection(Of Resource)


    Public Sub New()

    End Sub


    Public Sub New(ByVal AssemblyFileName As String)
        Me.FileName = AssemblyFileName
    End Sub


    ''' <summary>
    ''' The filename of the assembly to modify
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property FileName() As String
        Get
            Return rFilename
        End Get
        Set(ByVal value As String)
            rFilename = value
        End Set
    End Property
    Private rFilename As String


    Private Sub InitAssembly()
        If _Inited = False And Len(Me.FileName) > 0 Then
            _Inited = True

            '---- load up the assembly to be modified
            _Asm = AssemblyDefinition.ReadAssembly(Me.FileName)

            'Gets all types which are declared in the Main Module of the asm to be modified
            _Resources = _Asm.MainModule.Resources
        End If
    End Sub


    ''' <summary>
    ''' Adds a given resourcename and data to the app resources in the target assembly
    ''' </summary>
    ''' <param name="ResourceName"></param>
    ''' <param name="ResourceData"></param>
    ''' <remarks></remarks>
    Public Sub Add(ByVal ResourceName As String, ByVal ResourceData() As Byte)
        '---- make sure the writer is initialized
        InitAssembly()

        '---- have to enumerate this way
        For x = 0 To _Resources.Count - 1
            Dim res = _Resources(x)
            If res.Name.Contains(".Resources.resources") Then
                '---- Have to assume this is the root application's .net resources.
                '     That might not be the case though.

                '---- cast as embeded resource to get at the data
                Dim EmbededResource = DirectCast(res, Mono.Cecil.EmbeddedResource)

                '---- a Resource reader is required to read the resource data
                Dim ResReader = New ResourceReader(New IO.MemoryStream(EmbededResource.GetResourceData))

                '---- Use this output stream to capture all the resource data from the 
                '     existing resource block, so we can add the new resource into it
                Dim MemStreamOut = New IO.MemoryStream
                Dim ResWriter = New System.Resources.ResourceWriter(MemStreamOut)
                Dim ResEnumerator = ResReader.GetEnumerator
                Dim resdata() As Byte = Nothing
                Do While ResEnumerator.MoveNext
                    Dim resname = DirectCast(ResEnumerator.Key, String)
                    Dim restype As String = ""
                    '---- if we come across a resource named the same as the one
                    '     we're about to add, skip it
                    If StrComp(resname, ResourceName, CompareMethod.Text) <> 0 Then
                        ResReader.GetResourceData(resname, restype, resdata)
                        ResWriter.AddResourceData(resname, restype, resdata)
                    End If
                Loop

                '---- add the new resource data here
                ResWriter.AddResourceData(ResourceName, "ResourceTypeCode.ByteArray", ResourceData)
                '---- gotta call this to render the memory stream
                ResWriter.Generate()

                '---- update the resource
                Dim buf() = MemStreamOut.ToArray
                Dim NewEmbedRes = New EmbeddedResource(res.Name, res.Attributes, buf)
                _Resources.Remove(res)
                _Resources.Add(NewEmbedRes)
                '---- gotta bail out, there can't be 2 embedded resource chunks, right?
                Exit For
            End If
        Next
    End Sub


    ''' <summary>
    ''' Save all changes now.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Save(Optional ByVal OutFileName As String = "")
        If OutFileName = "" Then OutFileName = Me.FileName
        _Asm.Write(OutFileName)
    End Sub
End Class
