Attribute VB_Name = "Accessibility"

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Definitions and Procedures relating to Accessibility, used by the Ribbon VBA  '
' Demonstration UserForm. The constants have been lifted from oleacc.h, and are '
' just a subset of those available.                                             '
'                                                                               '
'                                                    Tony Jollans, August 2008. '
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '

' http://stackoverflow.com/questions/615482/how-to-get-ribbon-custom-tabs-ids
' http://stackoverflow.com/questions/15113874/is-there-a-way-of-showing-a-customui-tab-in-word-when-the-document-is-opened
' http://www.wordarticles.com/Shorts/RibbonVBA/RibbonVBADemo.php


Option Explicit

Public Const CHILDID_SELF                  As Long = &H0&
Public Const STATE_SYSTEM_UNAVAILABLE      As Long = &H1&
Public Const STATE_SYSTEM_INVISIBLE        As Long = &H8000&
Public Const STATE_SYSTEM_SELECTED         As Long = &H2&

Public Enum RoleNumber
    ROLE_SYSTEM_CLIENT = &HA&
    ROLE_SYSTEM_PANE = &H10&
    ROLE_SYSTEM_GROUPING = &H14&
    ROLE_SYSTEM_TOOLBAR = &H16&
    ROLE_SYSTEM_PAGETAB = &H25&
    ROLE_SYSTEM_PROPERTYPAGE = &H26&
    ROLE_SYSTEM_GRAPHIC = &H28&
    ROLE_SYSTEM_STATICTEXT = &H29&
    ROLE_SYSTEM_TEXT = &H2A&
    ROLE_SYSTEM_BUTTONDROPDOWNGRID = &H3A&
    ROLE_SYSTEM_PAGETABLIST = &H3C&
End Enum

Private Enum NavigationDirection
    NAVDIR_FIRSTCHILD = &H7&
End Enum

Private Declare PtrSafe Function AccessibleChildren Lib "oleacc.dll" _
                    (ByVal paccContainer As Object, ByVal iChildStart As Long, ByVal cChildren As Long, _
                           rgvarChildren As Variant, pcObtained As Long) _
                As Long

Public Function GetAccessible _
                    (Element As IAccessible, _
                     RoleWanted As RoleNumber, _
                     NameWanted As String, _
                     Optional GetClient As Boolean) _
                As IAccessible

    ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
    ' This procedure recursively searches the accessibility hierarchy, starting '
    ' with the element given, for an object matching the given name and role.   '
    ' If requested, the Client object, assumed to be the first child, will be   '
    ' returned instead of its parent.                                           '
    '                                                                           '
    ' Called by: RibbonForm procedures to get parent objects as required        '
    '            Itself, recursively, to move down the hierarchy                '
    ' Calls: GetChildren to, well, get children.                                '
    '        Itself, recursively, to move down the hierarchy                    '
    ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '

    Dim ChildrenArray(), Child As IAccessible, ndxChild As Long, ReturnElement As IAccessible

    If Element.accRole(CHILDID_SELF) = RoleWanted And Element.accName(CHILDID_SELF) = NameWanted Then

        Set ReturnElement = Element

    Else ' not found yet
        ChildrenArray = GetChildren(Element)

        If (Not ChildrenArray) <> True Then
            For ndxChild = LBound(ChildrenArray) To UBound(ChildrenArray)
                If TypeOf ChildrenArray(ndxChild) Is IAccessible Then

                    Set Child = ChildrenArray(ndxChild)
                    Set ReturnElement = GetAccessible(Child, RoleWanted, NameWanted)
                    If Not ReturnElement Is Nothing Then Exit For

                End If                  ' Child is IAccessible
            Next ndxChild
        End If                          ' there are children
    End If                              ' still looking

    If GetClient Then
        Set ReturnElement = ReturnElement.accNavigate(NAVDIR_FIRSTCHILD, CHILDID_SELF)
    End If

    Set GetAccessible = ReturnElement

End Function

Private Function GetChildren(Element As IAccessible) As Variant()
    ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
    ' General purpose subroutine to get an array of children of an IAccessible  '
    ' object. The returned array is Variant because the elements may be either  '
    ' IAccessible objects or simple (Long) elements, and the caller must treat  '
    ' them appropriately.                                                       '
    '                                                                           '
    ' Called by: GetAccessible when searching for an Accessible element         '
    ' Calls: AccessibleChildren API                                             '
    ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
    Const FirstChild As Long = 0&
    Dim NumChildren As Long, NumReturned As Long, ChildrenArray()

    NumChildren = Element.accChildCount

    If NumChildren > 0 Then
        ReDim ChildrenArray(NumChildren - 1)
        AccessibleChildren Element, FirstChild, NumChildren, ChildrenArray(0), NumReturned
    End If

    GetChildren = ChildrenArray
End Function





Public Sub SwitchTab(TabName As String)
    Dim RibbonTab   As IAccessible

    'Get the Ribbon as an accessiblity object and the
    Set RibbonTab = GetAccessible(CommandBars("Ribbon"), ROLE_SYSTEM_PAGETAB, TabName)

    'If we've found the ribbon then we can loop through the tabs
    If Not RibbonTab Is Nothing Then
        'If the tab state is valid (not unavailable or invisible)
        If ((RibbonTab.accState(CHILDID_SELF) And (STATE_SYSTEM_UNAVAILABLE Or _
                     STATE_SYSTEM_INVISIBLE)) = 0) Then
            'Then we can change to that tab
            RibbonTab.accDoDefaultAction CHILDID_SELF
        End If
    End If

End Sub

