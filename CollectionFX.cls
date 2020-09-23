VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CollectionFX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type mCollection
    Item As Variant
    Key As Variant
End Type

Enum ExistsType
    ItemName = 1
    KeyName = 2
End Enum

Private m_Collection_Count As Long 'Hold collection count
Private m_CollectionType() As mCollection 'Collection

Sub Add(Item As Variant, Optional Key)
    'Add a new item and key to the collection
    m_Collection_Count = m_Collection_Count + Lower 'Keep a count of the items in the collection
    ReDim Preserve m_CollectionType(Lower To m_Collection_Count) As mCollection 'Resize the collection
    m_CollectionType(m_Collection_Count).Item = Item 'Assign the item value
    If IsMissing(Key) Then Key = vbNullString 'If key is missing fill with a nullchar
    m_CollectionType(m_Collection_Count).Key = Key 'assign the key value
End Sub

Public Function Count() As Long
    'Returns the number of items in the collection
    Count = m_Collection_Count
End Function

Public Property Get Item(Index As Long) As Variant
On Error Resume Next
    'Return a items value in the collection
    If Not RangeOk(Index) Then
        Err.Raise 9
        Exit Property
    Else
        Item = m_CollectionType(Index).Item
    End If
  End Property

Public Property Let Item(Index As Long, vNewValue As Variant)
    'Chnage a items value in the collection
    If Not RangeOk(Index) Then
        Err.Raise 9
        Exit Property
    Else
        m_CollectionType(Index).Item = vNewValue
    End If
End Property

Public Property Get Key(Index As Long) As Variant
    'Return the value of a key in the collection
    If Not RangeOk(Index) Then
        Err.Raise 9
        Exit Property
    Else
        Key = m_CollectionType(Index).Key
    End If
  End Property

Public Property Let Key(Index As Long, vNewValue As Variant)
    'Chnage a keys value in the collection
    If Not RangeOk(Index) Then
        Err.Raise 9
        Exit Property
    Else
        m_CollectionType(Index).Key = vNewValue
    End If
End Property

Sub Clear()
    'Resets our collection
    m_Collection_Count = 0
    Erase m_CollectionType
End Sub

Private Sub Class_Terminate()
   Call Clear 'Clear up
End Sub

Private Function RangeOk(Index As Long) As Boolean
    'This just checks the range of an index the user has entered
    RangeOk = CBool((Index = 0) Or (Index > m_Collection_Count)) <> True
End Function

Function ItemExists(Item) As Boolean
    'Return true if an item is found otherwise false is returned
    ItemExists = ItemIndex(Item) <> -1
End Function

Function KeyExists(Key) As Boolean
    'Return true if an key is found otherwise false is returned
    KeyExists = KeyIndex(Key) <> -1
End Function

Public Property Get ItemIndex(Item) As Long
Dim x As Long
    'Function used to return the index number of an item
    If m_Collection_Count = 0 Then Err.Raise 9: Exit Function 'exit if we have no items in the collection
    ItemIndex = -1 'Default errror index
    
    For x = 1 To m_Collection_Count 'Loop tho the collection
        If m_CollectionType(x).Item = Item Then 'Check if item is here
            ItemIndex = x 'Yes so return the index
            Exit For 'exit out
        End If
    Next
End Property

Public Sub SaveToFile(lpFileName As String)
Dim nFile As Long, iSize As Long
    iSize = UBound(m_CollectionType)
    'Used to save the collection contents to a file
    nFile = FreeFile
    Open lpFileName For Binary As #nFile
        Put #nFile, , iSize
        Put #nFile, , m_CollectionType
    Close #nFile
End Sub

Public Sub LoadFromFile(lpFileName As String)
Dim nFile As Long, iSize As Long
    'Used to load a file and fill the collection
    Call Clear
    nFile = FreeFile
    Open lpFileName For Binary As #nFile
        Get #nFile, , iSize
        ReDim m_CollectionType(1 To iSize)
        Get #nFile, , m_CollectionType
    Close #nFile
    
    m_Collection_Count = iSize
End Sub

Public Sub RemoveItem(Index As Long)
Dim TempArr() As mCollection, lngDestIndex As Long, x As Long
On Error Resume Next
    'Remove an item form the collection
    lngDestIndex = Lower
    
    If (m_Collection_Count = Lower) Then
        'Clear the collection if we hit the lower bound
        Call Clear
        Exit Sub
    End If
    
    ReDim TempArr(Lower To m_Collection_Count - Lower) 'Resize the temp array
    
    For x = Lower To m_Collection_Count
        If x <> Index Then
            'X not equal to delete index then make a copy of the collection
            TempArr(lngDestIndex).Item = m_CollectionType(x).Item
            TempArr(lngDestIndex).Key = m_CollectionType(x).Key
            lngDestIndex = lngDestIndex + 1
        End If
    Next
    
    m_CollectionType = TempArr 'Pass temp array over to our collection
    m_Collection_Count = UBound(m_CollectionType) 'Update our collection count
    Erase TempArr 'Destroy the temp array
    
End Sub

Function Lower()
    Lower = 1 'Our collections lower bound
End Function

