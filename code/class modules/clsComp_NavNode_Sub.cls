VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsComp_NavNode_Sub"
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
Private WithEvents lbl_Caption As MSForms.Label
Attribute lbl_Caption.VB_VarHelpID = -1

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Custom Events
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Event MouseOver(ChildName As String)
Public Event ShowChildren(ChildName As String)
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Enumerations & Types
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Class Variables
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private MainNavElm As Boolean
Private ChildName As String


Private ParentElement As Object
Private ParentLogical As clsComp_NavNode_Main


Private ChildrenActive As Boolean
'Private children As

' These Two variables are for the chevron on the right
' side of the nav_label
Private Indicator_Normal As Double
Private Indicator_Expanded As Double


Private Element_NotActive As Variant
Private Element_Active As Variant

Private ForeColor_NotActive As Variant
Private ForeColor_Active As Variant

Private strDelegate_Activate As String
'Private Const HightlightColor As Variant = &H8000000D 'What is this...? what datatype
'Private Const StandardBackColor As Variant = &H80000012  'What is this...? what datatype


' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Class: Properties
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Property Get Height()
    Height = frmContainer.Height
End Property


Property Set Set_Parent(parent As clsComp_NavNode_Main)
    Set ParentLogical = parent
End Property
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Class: Initialization & Termination
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub class_initialize()
    Call SetElementAttributes
    ChildrenActive = False
End Sub
' ------------------
Private Sub Class_Terminate()
End Sub

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Class Events
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub frmContainer_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call MouseOver
End Sub
' ------------------
Private Sub lbl_Caption_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call MouseOver
End Sub
' ------------------
Private Sub frmContainer_Click()
    Call ToggleActive
End Sub
' ------------------
Private Sub lbl_Caption_Click()
    Call ToggleActive
End Sub

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Public Subroutines
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub Build(parentObject As Object, NodeName As String, sCaption As String, lLeft As Long, ltop As Long, bVisible As Boolean)
    ChildName = NodeName

    Call BuildContainerFrame(parentObject, NodeName, lLeft, ltop, bVisible)
    'Call BuildIconLabel(frmContainer, NodeName & "_Icon_lbl")
    Call BuildDisplayLabel(frmContainer, NodeName & "_Display_lbl", sCaption)
    'Call BuildIndicatorLabel(frmContainer, NodeName & "_Indicator_lbl")
End Sub
' ------------------
Public Sub Show()
    frmContainer.Visible = True
End Sub

Public Sub Hide()
    frmContainer.Visible = False
    DeHighlight
End Sub
' ------------------
Public Sub Highlight()
    frmContainer.BackColor = Element_Active
    frmContainer.ForeColor = ForeColor_Active

    lbl_Caption.BackColor = Element_Active
    lbl_Caption.ForeColor = ForeColor_Active
End Sub
' ------------------
Public Sub DeHighlight()
    'frmContainer.BackColor = Element_NotActive
    frmContainer.ForeColor = ForeColor_NotActive

    'lbl_Caption.BackColor = Element_NotActive
    lbl_Caption.ForeColor = ForeColor_NotActive
End Sub
' ------------------
Public Sub DectivateChildren()
    ' NEED to determine if I want to have a good amount of logic in here.
    DeHighlight
    ChildrenActive = False
    Debug.Print ("Body Unload: " & ChildName)
End Sub

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Public Functions
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Private Subroutines
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub ActivateChildren()
    ' NEED to determine if I want to have a good amount of logic in here.
    If Not ParentLogical Is Nothing Then
        Call ParentLogical.ChildActivated(ChildName)
    End If
    ChildrenActive = True
    Highlight
    Debug.Print ("Body Set: " & ChildName)

End Sub

Private Sub SetElementAttributes()
    Indicator_Normal = 10759
    Indicator_Expanded = 10760
    
    Element_NotActive = RGB(48, 64, 78)
    Element_Active = RGB(41, 56, 68)
    
    ForeColor_NotActive = RGB(168, 177, 193)
    ForeColor_Active = RGB(255, 255, 255)
End Sub

Private Sub MouseOver()

    RaiseEvent MouseOver(ChildName)
    ' May Need to add some internal routing for this sub.
End Sub
' ------------------
Private Sub ToggleActive()
    If ChildrenActive Then
        ' Doing nothing here since I don't want any actions
        ' to be preformed if the user reclicks the same nav button
        ' that he's currently on.
        'DectivateChildren
        
        
    Else
        ActivateChildren

    End If
    RaiseEvent ShowChildren(ChildName)
    ' May Need to add some internal routing for this sub.
End Sub
' ------------------
Private Sub BuildContainerFrame(parentObject As Object, name As String, lLeft As Long, ltop As Long, bVisible As Boolean)
    Set frmContainer = parentObject.Controls.Add("Forms.Frame.1", name)
    With frmContainer
        .BackColor = Element_Active
        '.BorderColor = &H80000012
        .BorderStyle = 0 'fmBorderStyleNone
        .Caption = ""
        .Enabled = True
        .ForeColor = ForeColor_NotActive
        .Height = 20
        .Width = 150
        .Top = ltop
        .Left = lLeft
        .ScrollBars = 0 ' fmScrollBarsNone
        .Visible = bVisible
        .SpecialEffect = 0 'fmSpecialEfectFlat
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
            .Width = 110
            .Top = 2
            .Left = 35
            .SpecialEffect = 0 'fmSpecialEfectFlat
          
            .TextAlign = 1 'fmTextAlignLeft
            .WordWrap = False
            .Visible = True
        End With
    End Sub

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Private Functions
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~





