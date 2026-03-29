Attribute VB_Name = "KeyState"
' http://www.cpearson.com/excel/keytest.aspx
' https://macexcel.com/examples/setupinfo/detectkeypress/

Option Explicit
Option Compare Text
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' modKeyState
' By Chip Pearson, www.cpearson.com, chip@cpearson.com
' This code is at www.cpearson.com/Excel/KeyTest.aspx
' This module contains functions for testing the state of the SHIFT, ALT, and CTRL
' keys.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public KeysEnabled As Boolean

'''''''''''''''''''''''''''''''''''''''''''''''''''''
' Platform API declarations for key state detection.
''''''''''''''''''''''''''''''''''''''''''''''''''''
#If Mac Then
    'Private Declare PtrSafe Function AppleScriptTask Lib "AppleScriptTask" (ByVal ScriptFile As String, ByVal HandlerName As String, ByVal ParameterString As String) As String

    ' Script file + handler used by AppleScriptTask on macOS.
    Private Const MAC_SCRIPT_FILE As String = "BKTKeyState.scpt"
    Private Const MAC_SCRIPT_HANDLER_MODIFIER_FLAGS As String = "modifierFlags"

    ' NSEventModifierFlags bit masks.
    Private Const MAC_MOD_SHIFT As Long = 131072
    Private Const MAC_MOD_OPTION As Long = 524288
    Private Const MAC_MOD_COMMAND As Long = 1048576

    Private mMacModifierSnapshotActive As Boolean
    Private mMacModifierSnapshotValid As Boolean
    Private mMacModifierSnapshotFlags As Long
#Else
    Private Declare PtrSafe Function GetKeyState Lib "user32" (ByVal vKey As Long) As Integer

    ' This constant is used in a bit-wise AND operation with the result of
    ' GetKeyState to determine if the specified key is down.
    Private Const KEY_MASK As Integer = &HFF80 ' decimal -128

    ' KEY CONSTANTS. Values taken from VC++ 6.0 WinUser.h file.
    Private Const VK_LSHIFT = &HA0
    Private Const VK_RSHIFT = &HA1
    Private Const VK_LCONTROL = &HA2
    Private Const VK_RCONTROL = &HA3
    Private Const VK_LMENU = &HA4
    Private Const VK_RMENU = &HA5

    ' Familiar aliases.
    Private Const VK_LALT = VK_LMENU
    Private Const VK_RALT = VK_RMENU
    Private Const VK_LCTRL = VK_LCONTROL
    Private Const VK_RCTRL = VK_RCONTROL
#End If

''''''''''''''''''''''''''''''''''''''''''''
' The following constants are used to specify,
' when testing CTRL, ALT, or SHIFT, whether
' the Left key, the Right key, either the
' Left OR Right key, or BOTH the Left AND
' Right keys are down.
'
' By default, the key-test procedures make
' no distinction between the Left and Right
' keys and will return TRUE if either the
' Left or Right (or both) key is down.
''''''''''''''''''''''''''''''''''''''''''''
Public Const BothLeftAndRightKeys = 0
Public Const LeftKey = 1
Public Const RightKey = 2
Public Const LeftKeyOrRightKey = 3

#If Mac Then
Private Function TryGetMacModifierFlags(ByRef ModifierFlags As Long) As Boolean
    Dim resultText As String

    On Error GoTo ErrHandler

    resultText = AppleScriptTask(MAC_SCRIPT_FILE, MAC_SCRIPT_HANDLER_MODIFIER_FLAGS, "")
    ModifierFlags = CLng(Trim$(resultText))
    TryGetMacModifierFlags = True
    Exit Function

ErrHandler:
    ' Never break ribbon actions because key-state probing failed.
    ModifierFlags = 0
    TryGetMacModifierFlags = False
End Function

Private Function RefreshMacModifierSnapshot() As Boolean
    mMacModifierSnapshotValid = TryGetMacModifierFlags(mMacModifierSnapshotFlags)
    If Not mMacModifierSnapshotValid Then
        mMacModifierSnapshotFlags = 0
    End If
    RefreshMacModifierSnapshot = mMacModifierSnapshotValid
End Function

Public Sub BeginModifierKeySnapshot()
    If mMacModifierSnapshotActive Then Exit Sub

    mMacModifierSnapshotActive = True
    Call RefreshMacModifierSnapshot
End Sub

Public Sub EndModifierKeySnapshot()
    mMacModifierSnapshotActive = False
    mMacModifierSnapshotValid = False
    mMacModifierSnapshotFlags = 0
End Sub

Private Function IsMacModifierDown(ByVal ModifierMask As Long) As Boolean
    Dim flags As Long

    If mMacModifierSnapshotActive Then
        If mMacModifierSnapshotValid Then
            flags = mMacModifierSnapshotFlags
            IsMacModifierDown = ((flags And ModifierMask) <> 0)
        Else
            IsMacModifierDown = False
        End If
    ElseIf TryGetMacModifierFlags(flags) Then
        IsMacModifierDown = ((flags And ModifierMask) <> 0)
    Else
        IsMacModifierDown = False
    End If
End Function
#Else
Public Sub BeginModifierKeySnapshot()
End Sub

Public Sub EndModifierKeySnapshot()
End Sub
#End If

Public Function IsMacScriptFileAccessible() As Boolean
#If Mac Then
    Dim flags As Long
    IsMacScriptFileAccessible = TryGetMacModifierFlags(flags)
#Else
    IsMacScriptFileAccessible = True
#End If
End Function

Public Sub SetKeysEnabled(Optional enabled As Boolean = True)
    If enabled = False Then
        KeysEnabled = False
        Exit Sub
    End If
    If IsMacScriptFileAccessible() Then
        KeysEnabled = True
    Else
        KeysEnabled = False
#If Mac Then
        MsgBox "Please install BKTKeyState.scpt to enable keys on Mac", vbExclamation
#End If
    End If
End Sub



Public Function IsShiftKeyDown(Optional LeftOrRightKey As Long = LeftKeyOrRightKey) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''
' IsShiftKeyDown
' Returns TRUE or FALSE indicating whether the
' SHIFT key is down.
'
' If LeftOrRightKey is omitted or LeftKeyOrRightKey,
' the function return TRUE if either the left or the
' right SHIFT key is down. If LeftKeyOrRightKey is
' LeftKey, then only the Left SHIFT key is tested.
' If LeftKeyOrRightKey is RightKey, only the Right
' SHIFT key is tested. If LeftOrRightKey is
' BothLeftAndRightKeys, the codes tests whether
' both the Left and Right keys are down. The default
' is to test for either Left or Right, making no
' distiction between Left and Right.
''''''''''''''''''''''''''''''''''''''''''''''''
    Dim Res As Long
    
    If Not KeysEnabled Then
        IsShiftKeyDown = False
        Exit Function
    End If
    
    #If Mac Then
        IsShiftKeyDown = IsMacModifierDown(MAC_MOD_SHIFT)
    #Else
    
        Select Case LeftOrRightKey
            Case LeftKey
                Res = GetKeyState(VK_LSHIFT) And KEY_MASK
            Case RightKey
                Res = GetKeyState(VK_RSHIFT) And KEY_MASK
            Case BothLeftAndRightKeys
                Res = (GetKeyState(VK_LSHIFT) And GetKeyState(VK_RSHIFT) And KEY_MASK)
            Case Else
                Res = GetKeyState(vbKeyShift) And KEY_MASK
        End Select
        
        IsShiftKeyDown = CBool(Res)
    
    #End If
End Function

Public Function IsControlKeyDown(Optional LeftOrRightKey As Long = LeftKeyOrRightKey) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''
' IsControlKeyDown
' Returns TRUE or FALSE indicating whether the
' CTRL key is down.
'
' If LeftOrRightKey is omitted or LeftKeyOrRightKey,
' the function return TRUE if either the left or the
' right CTRL key is down. If LeftKeyOrRightKey is
' LeftKey, then only the Left CTRL key is tested.
' If LeftKeyOrRightKey is RightKey, only the Right
' CTRL key is tested. If LeftOrRightKey is
' BothLeftAndRightKeys, the codes tests whether
' both the Left and Right keys are down. The default
' is to test for either Left or Right, making no
' distiction between Left and Right.
''''''''''''''''''''''''''''''''''''''''''''''''
    Dim Res As Long
    
    If Not KeysEnabled Then
        IsControlKeyDown = False
        Exit Function
    End If
    
    #If Mac Then
        IsControlKeyDown = IsMacModifierDown(MAC_MOD_COMMAND)
    #Else
    
        Select Case LeftOrRightKey
            Case LeftKey
                Res = GetKeyState(VK_LCTRL) And KEY_MASK
            Case RightKey
                Res = GetKeyState(VK_RCTRL) And KEY_MASK
            Case BothLeftAndRightKeys
                Res = (GetKeyState(VK_LCTRL) And GetKeyState(VK_RCTRL) And KEY_MASK)
            Case Else
                Res = GetKeyState(vbKeyControl) And KEY_MASK
        End Select
        
        IsControlKeyDown = CBool(Res)
    
    #End If

End Function

Public Function IsAltKeyDown(Optional LeftOrRightKey As Long = LeftKeyOrRightKey) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''
' IsAltKeyDown
' Returns TRUE or FALSE indicating whether the
' ALT key is down.
'
' If LeftOrRightKey is omitted or LeftKeyOrRightKey,
' the function return TRUE if either the left or the
' right ALT key is down. If LeftKeyOrRightKey is
' LeftKey, then only the Left ALT key is tested.
' If LeftKeyOrRightKey is RightKey, only the Right
' ALT key is tested. If LeftOrRightKey is
' BothLeftAndRightKeys, the codes tests whether
' both the Left and Right keys are down. The default
' is to test for either Left or Right, making no
' distiction between Left and Right.
''''''''''''''''''''''''''''''''''''''''''''''''
    Dim Res As Long
    
    If Not KeysEnabled Then
        IsAltKeyDown = False
        Exit Function
    End If
    
    #If Mac Then
        IsAltKeyDown = IsMacModifierDown(MAC_MOD_OPTION)
    #Else
    
        Select Case LeftOrRightKey
            Case LeftKey
                Res = GetKeyState(VK_LALT) And KEY_MASK
            Case RightKey
                Res = GetKeyState(VK_RALT) And KEY_MASK
            Case BothLeftAndRightKeys
                Res = (GetKeyState(VK_LALT) And GetKeyState(VK_RALT) And KEY_MASK)
            Case Else
                Res = GetKeyState(vbKeyMenu) And KEY_MASK
        End Select
        
        IsAltKeyDown = CBool(Res)
    
    #End If

End Function

