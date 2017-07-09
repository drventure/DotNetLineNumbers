Imports System.Runtime.InteropServices


''' <summary>
''' Class to handle updating resources in a Windows DLL or EXE
''' (c) 2008-2010 Darin Higgins All Rights Reserved
''' </summary>
''' <remarks></remarks>
''' <editHistory></editHistory>
Class ResourceWriter
    Implements IDisposable

#Region " Enumerations"
    Public Enum ResTypes
        RT_ACCELERATOR = 9&
        RT_ANICURSOR = (21)
        RT_ANIICON = (22)
        RT_BITMAP = 2&
        RT_CURSOR = 1&
        RT_DIALOG = 5&
        RT_DLGINCLUDE = 17
        RT_FONT = 8&
        RT_FONTDIR = 7&
        RT_HTML = 23
        RT_ICON = 3&
        RT_MENU = 4&
        RT_MESSAGETABLE = 11
        RT_PLUGPLAY = 19
        RT_RCDATA = 10&
        RT_STRING = 6&
        RT_VERSION = 16
        RT_VXD = 20
        DIFFERENCE = 11
        RT_GROUP_CURSOR = (RT_CURSOR + DIFFERENCE)
        RT_GROUP_ICON = (RT_ICON + DIFFERENCE)
    End Enum
#End Region


#Region " Exceptions"
    Public Class ResWriteCantOpenException
        Inherits ApplicationException

        Public Sub New(ByVal Filename As String, ByVal ErrCode As Integer)
            MyBase.New("Unable to file " & Filename & "   Code: " & ErrCode.ToString)
        End Sub
    End Class


    Public Class ResWriteCantUpdateException
        Inherits ApplicationException
        Public Sub New(ByVal ErrCode As Integer)
            MyBase.New("Problem updating the resource. Code: " & ErrCode.ToString)
        End Sub
    End Class


    Public Class ResWriteCantDeleteException
        Inherits ApplicationException
        Public Sub New(ByVal ErrCode As Integer)
            MyBase.New("Problem updating the resource. Code: " & ErrCode.ToString)
        End Sub
    End Class


    Public Class ResWriteCantEndException
        Inherits ApplicationException
        Public Sub New(ByVal ErrCode As Integer)
            MyBase.New("Problem finalizing resource modifications. Code: " & ErrCode.ToString)
        End Sub
    End Class
#End Region


    Public Property FileName() As String
        Get
            Return rFilename
        End Get
        Set(ByVal value As String)
            rFilename = value
        End Set
    End Property
    Private rFilename As String


    ''' <summary>
    ''' Add and Update are really the same thing, they both
    ''' add and update resources, as appropriate
    ''' </summary>
    ''' <param name="ResType"></param>
    ''' <param name="ResName"></param>
    ''' <param name="wLanguage"></param>
    ''' <param name="NewValue"></param>
    ''' <remarks></remarks>
    Public Sub Add( _
          ByVal ResType As Object, _
          ByVal ResName As Object, _
          ByVal wLanguage As Short, _
          ByRef NewValue As Byte())

        Me.Update(ResType, ResName, wLanguage, NewValue)
    End Sub


    ''' <summary>
    ''' Update or insert a resource into a WinPE image.
    ''' This version accepts a string buffer
    ''' and converts it to a byte buffer
    ''' </summary>
    ''' <param name="ResType"></param>
    ''' <param name="ResName"></param>
    ''' <param name="NewValue"></param>
    ''' <param name="wLanguage"></param>
    ''' <remarks></remarks>
    Public Sub Update( _
          ByVal ResType As Object, _
          ByVal ResName As Object, _
          ByVal wLanguage As Short, _
          ByRef NewValue As String)

        '---- make sure there's something to do
        If Len(NewValue) = 0 Then Exit Sub

        '---- convert raw string to byte buffer
        Dim buf() As Byte
        ReDim buf(Len(NewValue) + 1)
        Dim i As Integer = 0
        For Each c As Char In NewValue.ToCharArray
            buf(i) = CByte(Asc(c))
            i += 1
        Next

        '---- and update the resource
        Me.Update(ResType, ResName, wLanguage, buf)
    End Sub


    ''' <summary>
    ''' Update a resource using a raw byte buffer
    ''' </summary>
    ''' <param name="ResType"></param>
    ''' <param name="ResName"></param>
    ''' <param name="wLanguage"></param>
    ''' <param name="NewValue"></param>
    ''' <remarks></remarks>
    Public Sub Update( _
          ByVal ResType As Object, _
          ByVal ResName As Object, _
          ByVal wLanguage As Short, _
          ByRef NewValue As Byte())

        Dim Result As Integer

        Dim hUpdate As IntPtr = BeginUpdateResource(Me.FileName, 0)
        If hUpdate.ToInt32 = 0 Then
            Throw New ResWriteCantOpenException(Me.FileName, Err.LastDllError)
        End If

        '---- guarantee data is aligned to 4 byte boundary
        '	  not sure if this is strictly necessary, but it's documented
        Dim l = NewValue.Length
        Dim [Mod] As Integer = l Mod 4
        If [Mod] > 0 Then
            l += 4 - [Mod]
        End If
        ReDim Preserve NewValue(l - 1)

        '---- important note, The Typename and Resourcename are
        '     written as UPPER CASE. I'm not sure why but this appears to 
        '     be required for the FindResource and FindResourceEx API calls to 
        '     be able to work properly. Otherwise, they'll never find the resource
        '     entries...
        If VarType(ResType) = VariantType.String Then
            '---- Restype is a resource type name
            If VarType(ResName) = VariantType.String Then
                Result = UpdateResource(hUpdate, ResType.ToString.ToUpper, ResName.ToString.ToUpper, wLanguage, NewValue, NewValue.Length)
            Else
                Result = UpdateResource(hUpdate, ResType.ToString.ToUpper, CInt(ResName), wLanguage, NewValue, NewValue.Length)
            End If
        Else
            '---- Restype is a numeric resource ID
            If VarType(ResName) = VariantType.String Then
                Result = UpdateResource(hUpdate, CInt(ResType), ResName.ToString.ToUpper, wLanguage, NewValue, NewValue.Length)
            Else
                Result = UpdateResource(hUpdate, CInt(ResType), CInt(ResName), wLanguage, NewValue, NewValue.Length)
            End If
        End If
        If Result = 0 Then
            Throw New ResWriteCantUpdateException(Err.LastDllError)
        End If

        Result = EndUpdateResource(hUpdate, 0)
        If Result = 0 Then
            Throw New ResWriteCantEndException(Err.LastDllError)
        End If
    End Sub


    ''' <summary>
    ''' Delete a resource from the WinPE image.
    ''' Important note: If the target resource was added via ADD or UPDATE
    ''' in this class (and via BeginUpdateResource in general), this process
    ''' fails with an error 87 (bad parameter).
    ''' I have no idea why at this point.
    ''' If the resource in question already existed within the file (defined at
    ''' compile time, for instance), then there usually is no problem deleting it.
    ''' </summary>
    ''' <param name="ResType"></param>
    ''' <param name="ResName"></param>
    ''' <param name="wLanguage"></param>
    ''' <remarks></remarks>
    Public Sub Delete( _
          ByVal ResType As Object, _
          ByVal ResName As Object, _
          Optional ByVal wLanguage As Short = 0)

        Dim Result As Integer

        Dim hUpdate As IntPtr = BeginUpdateResource(Me.FileName, 0)
        If hUpdate.ToInt32 = 0 Then
            Throw New ResWriteCantOpenException(Me.FileName, Err.LastDllError)
        End If

        If VarType(ResType) = VariantType.String Then
            '---- Restype is a resource name
            If VarType(ResName) = VariantType.String Then
                Result = UpdateResource(hUpdate, ResType.ToString.ToUpper, ResName.ToString.ToUpper, wLanguage, IntPtr.Zero, 0)
            Else
                Result = UpdateResource(hUpdate, ResType.ToString.ToUpper, CInt(ResName), wLanguage, IntPtr.Zero, 0)
            End If
        Else
            '---- Restype is a numeric resource ID
            If VarType(ResName) = VariantType.String Then
                Result = UpdateResource(hUpdate, CInt(ResType), ResName.ToString.ToUpper, wLanguage, IntPtr.Zero, 0)
            Else
                Result = UpdateResource(hUpdate, CInt(ResType), CInt(ResName), wLanguage, IntPtr.Zero, 0)
            End If
        End If
        If Result = 0 Then
            Throw New ResWriteCantDeleteException(Err.LastDllError)
        End If

        Result = EndUpdateResource(hUpdate, 0)
        If Result = 0 Then
            Throw New ResWriteCantEndException(Err.LastDllError)
        End If
    End Sub


    Public Sub Dispose() Implements IDisposable.Dispose
        '---- Nothing really to dispose of
    End Sub
End Class


#Region " Resource API Declarations"
''' <summary>
''' this is an internal module for Resource API functions I use above
''' to write the linemap into the resources of an exe/dll
''' </summary>
''' <remarks></remarks>
Public Module ResourceAPIs
    <DllImport("KERNEL32.DLL", EntryPoint:="BeginUpdateResourceW", SetLastError:=True, CharSet:=CharSet.Unicode)> _
    Public Function BeginUpdateResource( _
          <MarshalAs(UnmanagedType.LPWStr)> ByVal pFileName As String, _
          ByVal bDeleteExistingResources As Int32 _
          ) As IntPtr
    End Function


    <DllImport("KERNEL32.DLL", EntryPoint:="EndUpdateResourceW", SetLastError:=True, CharSet:=CharSet.Unicode)> _
    Public Function EndUpdateResource( _
       ByVal hUpdate As IntPtr, _
       ByVal fDiscard As Int32 _
       ) As Int32
    End Function


    <DllImport("KERNEL32.DLL", EntryPoint:="UpdateResourceW", SetLastError:=True, CharSet:=CharSet.Unicode)> _
    Public Function UpdateResource( _
          ByVal hUpdate As IntPtr, _
          ByVal lpType As Int32, _
          <MarshalAs(UnmanagedType.LPWStr)> ByVal lpName As String, _
          ByVal wLanguage As Int16, _
          ByVal lpData As Byte(), _
          ByVal cbData As Int32 _
          ) As Int32
    End Function


    <DllImport("KERNEL32.DLL", EntryPoint:="UpdateResourceW", SetLastError:=True, CharSet:=CharSet.Unicode)> _
    Public Function UpdateResource( _
          ByVal hUpdate As IntPtr, _
          <MarshalAs(UnmanagedType.LPWStr)> ByVal lpType As String, _
          <MarshalAs(UnmanagedType.LPWStr)> ByVal lpName As String, _
          ByVal wLanguage As Int16, _
          ByVal lpData As Byte(), _
          ByVal cbData As Int32 _
          ) As Int32
    End Function


    <DllImport("KERNEL32.DLL", EntryPoint:="UpdateResourceW", SetLastError:=True, CharSet:=CharSet.Unicode)> _
    Public Function UpdateResource( _
          ByVal hUpdate As IntPtr, _
          <MarshalAs(UnmanagedType.LPWStr)> ByVal lpType As String, _
          ByVal lpName As Int32, _
          ByVal wLanguage As Int16, _
          ByVal lpData As Byte(), _
          ByVal cbData As Int32 _
          ) As Int32
    End Function


    <DllImport("KERNEL32.DLL", EntryPoint:="UpdateResourceW", SetLastError:=True, CharSet:=CharSet.Unicode)> _
    Public Function UpdateResource( _
          ByVal hUpdate As IntPtr, _
          ByVal lpType As Int32, _
          ByVal lpName As Int32, _
          ByVal wLanguage As Int16, _
          ByVal lpData As Byte(), _
          ByVal cbData As Int32 _
          ) As Int32
    End Function


    <DllImport("KERNEL32.DLL", EntryPoint:="UpdateResourceW", SetLastError:=True, CharSet:=CharSet.Unicode)> _
    Public Function UpdateResource( _
          ByVal hUpdate As IntPtr, _
          ByVal lpType As Int32, _
          <MarshalAs(UnmanagedType.LPWStr)> ByVal lpName As String, _
          ByVal wLanguage As Int16, _
          ByVal lpData As IntPtr, _
          ByVal cbData As Int32 _
          ) As Int32
    End Function


    <DllImport("KERNEL32.DLL", EntryPoint:="UpdateResourceW", SetLastError:=True, CharSet:=CharSet.Unicode)> _
    Public Function UpdateResource( _
          ByVal hUpdate As IntPtr, _
          <MarshalAs(UnmanagedType.LPWStr)> ByVal lpType As String, _
          <MarshalAs(UnmanagedType.LPWStr)> ByVal lpName As String, _
          ByVal wLanguage As Int16, _
          ByVal lpData As IntPtr, _
          ByVal cbData As Int32 _
          ) As Int32
    End Function


    <DllImport("KERNEL32.DLL", EntryPoint:="UpdateResourceW", SetLastError:=True, CharSet:=CharSet.Unicode)> _
    Public Function UpdateResource( _
          ByVal hUpdate As IntPtr, _
          <MarshalAs(UnmanagedType.LPWStr)> ByVal lpType As String, _
          ByVal lpName As Int32, _
          ByVal wLanguage As Int16, _
          ByVal lpData As IntPtr, _
          ByVal cbData As Int32 _
          ) As Int32
    End Function


    <DllImport("KERNEL32.DLL", EntryPoint:="UpdateResourceW", SetLastError:=True, CharSet:=CharSet.Unicode)> _
    Public Function UpdateResource( _
          ByVal hUpdate As IntPtr, _
          ByVal lpType As Int32, _
          ByVal lpName As Int32, _
          ByVal wLanguage As Int16, _
          ByVal lpData As IntPtr, _
          ByVal cbData As Int32 _
          ) As Int32
    End Function
End Module
#End Region
