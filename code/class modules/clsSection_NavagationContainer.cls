VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSection_NavagationContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private activeChild As String
Private activeChildIdx As String
Private MainNavCollection As Collection


Private Sub class_initialize()
    Const mNode_Height As Long = 30
    Const cNode_Height As Long = 20
    Const lStart_Left As Long = 10
    Const lStart_Top As Long = 10
    
    Set MainNavCollection = New Collection
    Dim permissionsArray() As Variant
    'Dim parentObject As Variant
    Dim MainNavs() As Variant
    Dim i As Long

    'Set parentObject = UserForm1

    permissionsArray = Worksheets("Sheet2").Range("A1:B15").value

    MainNavs = GetMainNavSections(permissionsArray)

    For i = LBound(MainNavs) To UBound(MainNavs)
        Dim mNode As clsComp_NavNode_Main
        Dim ParentDisplayName As String
        Dim ParentRefName As String
        Dim childCount As Long
        Dim lChildTop As Long
        Dim mCurrTop As Long
        Dim j As Long
       
        Set mNav = New clsComp_NavNode_Main

        mCurrTop = lStart_Top + mNode_Height * (i - 1)
        ParentDisplayName = MainNavs(i)
        ParentRefName = CreateRefName("Main", ParentDisplayName)

        Call mNav.Build(UserForm1, ParentRefName, ParentDisplayName, lStart_Left, mCurrTop, True, True, i)
        'Debug.Print ("MainNav: [pRefName: " & ParentRefName & "][DName: " & ParentDispalyName & "] [L: " & lStart_Left & "][T: " & mCurrTop & "][idx: " & i & "]")
        childCount = 1
        lChildTop = mCurrTop + mNode_Height

        For j = LBound(permissionsArray) To UBound(permissionsArray)
            Dim cNode As clsComp_NavNode_Sub
            Dim ChildDisplayName As String
            Dim ChildRefName As String
            

            If MainNavs(i) = permissionsArray(j, 1) Then
                Set cNode = New clsComp_NavNode_Sub

                ChildDisplayName = permissionsArray(j, 2)
                ChildRefName = CreateRefName(ParentRefName, ChildDisplayName)
                
                Call cNode.Build(UserForm1, ChildRefName, ChildDisplayName, lStart_Left, lChildTop, False)
                'Debug.Print ("ChildNav: [cRefName: " & ChildRefName & "][DName: " & ChildDisplayName & "] [L: " & lStart_Left & "][T: " & lChildTop & "]")
       
                Call mNav.AddChildNode(cNode, ChildRefName)
                Set cNode.Set_Parent = mNav
                Call mNav.AddParent(Me)
                
                lChildTop = lChildTop + cNode_Height
                childCount = childCount + 1
            End If
        Next j
         MainNavCollection.Add mNav, ParentRefName
    Next i
    UserForm1.Show vbModeless
End Sub


Function CreateRefName(ByVal section As String, ByVal currentItemName As String) As String
    currentItemName = Replace(currentItemName, " ", "_")
    CreateRefName = section & "_" & currentItemName
End Function

Public Function GetMainNavSections(permissionsArray As Variant) As Variant
    Dim navSections() As Variant
    Dim i As Long
    Dim currentNavIdx As Long
    
    ReDim navSections(1 To UBound(permissionsArray))
    
    currentNavIdx = 1
    For i = LBound(permissionsArray) To UBound(permissionsArray)
        If GetIndex(navSections, permissionsArray(i, 1)) = 0 Then
            navSections(currentNavIdx) = permissionsArray(i, 1)
            currentNavIdx = currentNavIdx + 1
        End If
    Next i
    ReDim Preserve navSections(1 To currentNavIdx - 1)
    GetMainNavSections = navSections
End Function

Function GetIndex(ByRef InputArr As Variant, ByVal value As Variant, Optional ByVal indexPos As Integer) As Long
    Dim i As Long
    For i = LBound(InputArr) To UBound(InputArr)
        If indexPos > 0 Then
            If InputArr(i, indexPos) = value Then: GetIndex = i: Exit For
        Else
            If InputArr(i) = value Then: GetIndex = i: Exit For
        End If
     Next i
End Function


Public Sub MainNavActivated(MainNavName As String)
    Dim MainNavIdx As Long
    
    MainNavIdx = MainNavCollection(MainNavName).NavPositionIdx
    If activeChild = MainNavName Then
        Exit Sub
    ElseIf activeChild = "" Then
        'ChildrenCollection(activeChild).DectivateChildren
        activeChild = MainNavName
        activeChildIdx = MainNavIdx
        Call AdjustNavPositionsExpanded(MainNavIdx)
        Debug.Print activeChild
        Exit Sub
    Else
    
        MainNavCollection(activeChild).HideChildren
        Call AdjustNavPositionCollapse(MainNavCollection(MainNavName).NavPositionIdx)
        Call AdjustNavPositionsExpanded(MainNavCollection(MainNavName).NavPositionIdx)
        
        activeChild = MainNavName
        activeChildIdx = MainNavIdx
    End If
End Sub

Private Sub AdjustNavPositionsExpanded(idx As Long)
    Dim i As Long
    Dim NavCount As Long
    Dim currNewTop As Long
    
    NavCount = MainNavCollection.Count
    currNewTop = MainNavCollection(idx).NextParentHeightExpanded
    For i = idx To NavCount
        If i = NavCount Then
            'currNewTop = MainNavCollection(i).NextParentHeightExpanded
            'MainNavCollection(i + 1).NewTop = currNewTop
        Else
            
            'currNewTop = MainNavCollection(i).NextParentHeightExpanded
            MainNavCollection(i + 1).NewTop = currNewTop
            currNewTop = currNewTop + 30
        End If
    Next i
    
End Sub

Private Sub AdjustNavPositionCollapse(idx As Long)
    Dim i As Long
    Dim NavCount As Long
    
    NavCount = MainNavCollection.Count
    
    For i = 1 To NavCount
        Call MainNavCollection(i).SetOriginalTop
    Next i
End Sub

