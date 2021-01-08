Attribute VB_Name = "KeyState"
' http://www.cpearson.com/excel/keytest.aspx

Option Explicit
Option Compare Text
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' modKeyState
' By Chip Pearson, www.cpearson.com, chip@cpearson.com
' This code is at www.cpearson.com/Excel/KeyTest.aspx
' This module contains functions for testing the state of the SHIFT, ALT, and CTRL
' keys.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


''''''''''''''''''''''''''''''''''''''''''''''''''''
' Declaration of GetKeyState API function. This
' tests the state of a specified key.
''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Declare PtrSafe Function GetKeyState Lib "user32" (ByVal vKey As Long) As Integer
    
''''''''''''''''''''''''''''''''''''''''''
' This constant is used in a bit-wise AND
' operation with the result of GetKeyState
' to determine if the specified key is
' down.
''''''''''''''''''''''''''''''''''''''''''
Private Const KEY_MASK As Integer = &HFF80 ' decimal -128

'''''''''''''''''''''''''''''''''''''''''
' KEY CONSTANTS. Values taken
' from VC++ 6.0 WinUser.h file.
'''''''''''''''''''''''''''''''''''''''''
Private Const VK_LSHIFT = &HA0
Private Const VK_RSHIFT = &HA1
Private Const VK_LCONTROL = &HA2
Private Const VK_RCONTROL = &HA3
Private Const VK_LMENU = &HA4
Private Const VK_RMENU = &HA5
'''''''''''''''''''''''''''''''''''''''''
' The following four constants simply
' provide other names, CTRL and ALT,
' for CONTROL and MENU. "CTRL" and
' "ALT" are more familiar than
' "CONTROL" and "MENU". These constants
' provide no additional functionality.
' They simply provide more familiar
' names.
'''''''''''''''''''''''''''''''''''''''''
Private Const VK_LALT = VK_LMENU
Private Const VK_RALT = VK_RMENU
Private Const VK_LCTRL = VK_LCONTROL
Private Const VK_RCTRL = VK_RCONTROL

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
    
    #If Mac Then
        IsShiftKeyDown = False
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
    
    #If Mac Then
        IsControlKeyDown = False
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
    
    #If Mac Then
        IsAltKeyDown = False
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

