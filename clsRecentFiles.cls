VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRecent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'An object to maintain the most recent files opened by a program.

Option Explicit
Private mcolRecent As New Collection
Private mintNumber As Integer

Public Sub Add(strFile As String)
'Add a file to the list.

    If mcolRecent.count = 0 Then
        mcolRecent.Add strFile, UCase(strFile)
    Else
        On Error Resume Next
        mcolRecent.Remove UCase(strFile)
        mcolRecent.Add strFile, UCase(strFile), 1
        If mcolRecent.count > mintNumber Then
            mcolRecent.Remove mcolRecent.count
        End If
    End If

End Sub

Public Sub Clear()
'Clears all files from the list.

    Do While mcolRecent.count > 0
        mcolRecent.Remove 1
    Loop

End Sub

Public Property Get count() As Long
'Returns the number of files in the list.

    count = mcolRecent.count

End Property

Public Property Get Item(intN As Integer) As String
'Returns the nth item from the list.

    On Error GoTo ItemError
    Item = mcolRecent(intN)
    Exit Property

ItemError:
    Item = ""

End Property

Public Sub Load(ByVal strPath As String)
'Load the most recently used files from the registry.

    Dim aryTemp As Variant
    Dim strItem As String
    Dim intUpperBound As Integer
    Dim intIndex As Integer

    aryTemp = GetAllFromRegistry(HKEY_CURRENT_USER, strPath)
    
    intUpperBound = UBound(aryTemp)
    
    If intUpperBound > 0 Then
        Me.Clear
        strItem = Pc(aryTemp(intUpperBound), " = ", 2)
        mcolRecent.Add Item:=strItem, Key:=UCase(strItem)
        For intIndex = intUpperBound - 1 To 1 Step -1
            strItem = Pc(aryTemp(intIndex), " = ", 2)
            mcolRecent.Add strItem, UCase(strItem), 1
        Next intIndex
    End If

End Sub

Public Property Get Number() As Integer
'Gets the maximum size of the list.

    Number = mintNumber

End Property

Public Property Let Number(intN As Integer)
'Sets the maximum size of the list.

    mintNumber = intN

End Property

Public Sub Remove(strFile As String)
'Remove a file from the list.

   On Error Resume Next
   mcolRecent.Remove UCase(strFile)

End Sub

Public Sub Save(ByVal strPath As String)
'Save the most recently used files list to the registry.

    Dim intCount As Integer

    On Error Resume Next

    Call DeleteRegistryKey(HKEY_CURRENT_USER, strPath)
    
    For intCount = 1 To mcolRecent.count
        SaveToRegistry HKEY_CURRENT_USER, strPath, intCount, mcolRecent(intCount)
    Next intCount

End Sub

Public Sub Update(objForm As Form)
' Update the most recently used files menu item.
' Note:  The form must contain a menu control array
'        named mnuFileRecent that is at least as big
'        as Number.
    
    Dim intCount As Long
    
    On Error GoTo LoadItems
    
    For intCount = 1 To mintNumber
        objForm.mnuFileRecent(intCount).Visible = False
    Next intCount

LoadItems:
    On Error GoTo LoadEnd
    
    If mcolRecent.count > 0 Then
        objForm.mnuFileRecent(0).Visible = True
        For intCount = 1 To mcolRecent.count
            objForm.mnuFileRecent(intCount).Caption = "&" & intCount & " " & mcolRecent(intCount)
            objForm.mnuFileRecent(intCount).Visible = True
        Next intCount
        Do
            objForm.mnuFileRecent(intCount).Visible = False
            intCount = intCount + 1
        Loop
    Else
        intCount = 0
        Do
            objForm.mnuFileRecent(intCount).Visible = False
            intCount = intCount + 1
        Loop
    End If

LoadEnd:

End Sub

Private Sub Class_Initialize()

   mintNumber = 5

End Sub
