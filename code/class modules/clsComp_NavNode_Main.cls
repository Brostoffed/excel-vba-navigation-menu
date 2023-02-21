VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsComp_NavNode_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
' ==========================================================================
' Module Name       : clsComp_NavNode
' Module Type       : Class
' --------------------------------------------------------------------------
' Description       : Still trying to determine if I should have all Nodes
'                     as one class or seperate them into Parent Nodes
'                     Child Nodes.
' ==========================================================================

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Consumed Events
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private WithEvents frmContainer As MSForms.Frame
Attribute frmContainer.VB_VarHelpID = -1
Private WithEvents lbl_Icon As MSForms.Label
Attribute lbl_Icon.VB_VarHelpID = -1
Private WithEvents lbl_Caption As MSForms.Label
Attribute lbl_Caption.VB_VarHelpID = -1
Private WithEvents lbl_Indicator As MSForms.Label
Attribute lbl_Indicator.VB_VarHelpID = -1
Private WithEvents frmActiveIndicator As MSForms.Frame
Attribute frmActiveIndicator.VB_VarHelpID = -1
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Custom Events
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Event MouseOver(NodeName As String)
Public Event ShowChildren(NodeName As String)

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Enumerations & Types
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Class Variables
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private MainNavElm As Boolean
Private InstanceName As String
Private NodeName As String
Private NavIndexPosition As Long

Private ParentElement As Object
Private ParentLogical As clsSection_NavagationContainer


Private ChildrenActive As Boolean
Private ChildrenCollection As Collection

' These Two variables are for the chevron on the right
' side of the nav_label
Private ElementTop As Long
Private ElementLeft As Long

Private Indicator_Normal As Double
Private Indicator_Expanded As Double

Private activeChild As String
Private Element_NotActive As Variant
Private Element_Active As Variant

Private ForeColor_NotActive As Variant
Private ForeColor_Active As Variant

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Class: Properties
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Property Get NextParentHeightExpanded() As Long
    NextParentHeightExpanded = GetChildrenTotalHeight + ElementTop + frmContainer.Height
End Property

Property Get NavPositionIdx() As Long
    NavPositionIdx = NavIndexPosition
End Property

Property Let NewTop(ltop As Long)
    frmContainer.Top = ltop
End Property
' =--------~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Class: Initialization & Termination
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub class_initialize()
    activeChild = ""
    Call SetElementAttributes
    ChildrenActive = False
    Set ChildrenCollection = New Collection
End Sub
' ------------------
Private Sub Class_Terminate()
    Set ChildrenCollection = Nothing
End Sub

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Class Events
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub frmContainer_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call MouseOver
End Sub
' ------------------
Private Sub lbl_Icon_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call MouseOver
End Sub
' ------------------
Private Sub lbl_Caption_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call MouseOver
End Sub
' ------------------
Private Sub lbl_Indicator_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call MouseOver
End Sub
' ------------------
Private Sub frmContainer_Click()
    Call ToggleActive
End Sub
' ------------------
Private Sub lbl_Icon_Click()
    Call ToggleActive
End Sub
' ------------------
Private Sub lbl_Caption_Click()
    Call ToggleActive
End Sub
' ------------------
Private Sub lbl_Indicator_Click()
    Call ToggleActive
End Sub
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Public Subroutines
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub Build(parentObject As Object, NodeName As String, sCaption As String, lLeft As Long, ltop As Long, bVisible As Boolean, bParent As Boolean, navIdx As Long)
    NodeName = NodeName
    InstanceName = NodeName
    ElementTop = ltop
    ElementLeft = lLeft
    Set ParentElement = parentObject
    NavIndexPosition = navIdx
    
    Call BuildContainerFrame(parentObject, NodeName, lLeft, ltop, bVisible)
    Call BuildIconLabel(frmContainer, NodeName & "_Icon_lbl")
    Call BuildDisplayLabel(frmContainer, NodeName & "_Display_lbl", sCaption)
    Call BuildIndicatorLabel(frmContainer, NodeName & "_Indicator_lbl")
End Sub
' ------------------
Public Sub AddChildNode(Node As Object, ChildIdxPos As String)
     ChildrenCollection.Add Item:=Node, Key:=CStr(ChildIdxPos)
End Sub
' ------------------
Public Sub AddParent(Node As clsSection_NavagationContainer)
    Set ParentLogical = Node
End Sub
' ------------------
Public Sub Highlight()
    frmContainer.BackColor = Element_Active
    frmContainer.ForeColor = ForeColor_Active
    
    lbl_Icon.BackColor = Element_Active
    lbl_Icon.ForeColor = ForeColor_Active
    
    lbl_Caption.BackColor = Element_Active
    lbl_Caption.ForeColor = ForeColor_Active
    
    lbl_Indicator.BackColor = Element_Active
    lbl_Indicator.ForeColor = ForeColor_Active
End Sub
' ------------------
Public Sub DeHighlight()
    frmContainer.BackColor = Element_NotActive
    frmContainer.ForeColor = ForeColor_NotActive

    lbl_Icon.BackColor = Element_NotActive
    lbl_Icon.ForeColor = ForeColor_NotActive
    
    lbl_Caption.BackColor = Element_NotActive
    lbl_Caption.ForeColor = ForeColor_NotActive
    
    lbl_Indicator.BackColor = Element_NotActive
    lbl_Indicator.ForeColor = ForeColor_NotActive
End Sub
' ------------------
Public Sub ActivateChildren()
    ' NEED to determine if I want to have a good amount of logic in here.
    lbl_Indicator.Caption = Application.WorksheetFunction.Unichar(Indicator_Expanded)
    ChildrenActive = True
End Sub
' ------------------
Public Sub DectivateChildren()
    ' NEED to determine if I want to have a good amount of logic in here.
    lbl_Indicator.Caption = Application.WorksheetFunction.Unichar(Indicator_Normal)
    ChildrenActive = False
End Sub
' ------------------
Public Sub ChildActivated(ChildName As String)
    If activeChild = "" Then
        activeChild = ChildName
    Else
        ChildrenCollection(activeChild).DectivateChildren
        activeChild = ChildName
    End If
        
End Sub
' ~~~~~~~~
Public Sub HideChildren()
    Dim Child As clsComp_NavNode_Sub

    frmActiveIndicator.Visible = False
    DectivateChildren
    DeHighlight
    
    For Each Child In ChildrenCollection
        Child.Hide
    Next Child
End Sub

Public Sub SetOriginalTop()
    frmContainer.Top = ElementTop
End Sub
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Public Functions
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Private Subroutines
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub ShowChildren()
    Dim Child As clsComp_NavNode_Sub
    
    For Each Child In ChildrenCollection
        Child.Show
    Next Child
End Sub
' ------------------
' ------------------
Private Sub SetElementAttributes()
    Indicator_Normal = 10759
    Indicator_Expanded = 10760
    
    Element_NotActive = RGB(48, 64, 78)
    Element_Active = RGB(41, 56, 68)
    
    ForeColor_NotActive = RGB(168, 177, 193)
    ForeColor_Active = RGB(255, 255, 255)
End Sub
' ------------------
Private Sub MouseOver()
    'Highlight
    RaiseEvent MouseOver(NodeName)
    ' May Need to add some internal routing for this sub.
End Sub
' ------------------
Private Sub ToggleActive()
    If frmActiveIndicator Is Nothing Then
        Call BuildSideActiveIndicator(ParentElement, "ActiveIndicator", 30 + GetChildrenTotalHeight, ElementLeft, ElementTop)
    End If
    
    If ChildrenActive Then

        'HideChildren

    Else
        ActivateChildren
        Highlight
        ShowChildren

        frmActiveIndicator.Visible = True
        Call ParentLogical.MainNavActivated(InstanceName)
    End If
    RaiseEvent ShowChildren(NodeName)
    ' May Need to add some internal routing for this sub.
End Sub
' ------------------
Private Sub BuildContainerFrame(parentObject As Object, name As String, lLeft As Long, ltop As Long, bVisible As Boolean)
    Set frmContainer = parentObject.Controls.Add("Forms.Frame.1", name)
    With frmContainer
        .BackColor = Element_NotActive
        '.BorderColor = &H80000012
        .BorderStyle = 0 'fmBorderStyleNone
        .Caption = ""
        .Enabled = True
        .ForeColor = ForeColor_NotActive
        .Height = 30
        .Width = 150
        .Top = ltop
        .Left = lLeft
        .ScrollBars = 0 ' fmScrollBarsNone
        .Visible = bVisible
        .SpecialEffect = 0 'fmSpecialEfectFlat
    End With
End Sub
' ------------------
Private Sub BuildIconLabel(parentObject As MSForms.Frame, name As String)
    Set lbl_Icon = parentObject.Controls.Add("Forms.Label.1", name)
    With lbl_Icon
        .AutoSize = False
        '.BackColor = rgb(48,64,78)
        .BackStyle = 0 ' fmBackStyleTransparent
        '.BorderColor = &H80000006
        .BorderStyle = 0 'fmBorderStyleNone
        .Caption = Application.WorksheetFunction.Unichar(9253)
        .Font.name = "Arial"
        .Font.Size = 16
        '.ForeColor = ForeColor_NotActive
        .Enabled = True
        .Height = 18
        .Width = 16.5
        .Top = 7
        .Left = 6
        .SpecialEffect = 0 'fmSpecialEfectFlat

        .TextAlign = 1 'fmTextAlignLeft
        .WordWrap = False
        .Visible = True
    End With
End Sub
' ------------------
Private Sub BuildDisplayLabel(parentObject As MSForms.Frame, name As String, sCaption As String)
    Set lbl_Caption = parentObject.Controls.Add("Forms.Label.1", name)
        With lbl_Caption
            .AutoSize = False
            '.BackColor = &H8000000F
            .BackStyle = 1 ' fmBackStyleTransparent
            '.BorderColor = &H80000006
            .BorderStyle = 0 'fmBorderStyleNone
            .Caption = sCaption
            .Font.name = "Century Gothic Bold" '"Segoe UI Semilight"
            .Font.Size = 13
            '.ForeColor = ForeColor_NotActive
            .Enabled = True
            .Height = 18
            .Width = 100
            .Top = 8
            .Left = 25
            .SpecialEffect = 0 'fmSpecialEfectFlat
          
            .TextAlign = 1 'fmTextAlignLeft
            .WordWrap = False
            .Visible = True
        End With
    End Sub
' ------------------
Private Sub BuildIndicatorLabel(parentObject As MSForms.Frame, name As String)
    Set lbl_Indicator = parentObject.Controls.Add("Forms.Label.1", name)
        With lbl_Indicator
            .AutoSize = False
            '.BackColor = &H8000000F
            .BackStyle = 0 ' fmBackStyleTransparent
            '.BorderColor = &H80000006
            .BorderStyle = 0 'fmBorderStyleNone
            .Caption = Application.WorksheetFunction.Unichar(Indicator_Normal)
            .Font.name = "Webdings"
            .Font.Size = 12
            .ForeColor = ForeColor_NotActive
            .Enabled = True
            .Height = 12
            .Width = 12.75
            .Top = 9
            .Left = 130
            .SpecialEffect = 0 'fmSpecialEfectFlat
            
            .TextAlign = 1 'fmTextAlignLeft
            .WordWrap = False
            .Visible = True
        End With
    End Sub
    
Private Sub BuildSideActiveIndicator(parentObject As Object, name As String, lHeight As Long, lLeft As Long, ltop As Long)
    Set frmActiveIndicator = parentObject.Controls.Add("Forms.Frame.1", name)
    With frmActiveIndicator
        .BackColor = RGB(26, 179, 148)
        .BorderStyle = 0 'fmBorderStyleNone
        .Caption = ""
        .Enabled = True
        .Height = lHeight
        .Width = 3
        .Top = ltop
        .Left = lLeft
        .ScrollBars = 0 ' fmScrollBarsNone
        .Visible = True
        .SpecialEffect = 0 'fmSpecialEfectFlat
    End With
End Sub
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Private Functions
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Function GetChildrenTotalHeight() As Long

    Dim ChildHeights As Long
    
    Dim Child As clsComp_NavNode_Sub
    
    For Each Child In ChildrenCollection
        ChildHeights = ChildHeights + Child.Height
    Next Child
    GetChildrenTotalHeight = ChildHeights
End Function



