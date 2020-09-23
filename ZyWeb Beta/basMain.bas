Attribute VB_Name = "basMain"
Option Explicit

'-------------
'API Declares.
'=============

Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long

'--------------
'API Constants.
'==============

'Used for Hooking into the key board.
Public Const WH_KEYBOARD = 2

'Find a string.
Public Const CB_FINDSTRING = &H14C

'-----------------
'Public variables.
'=================

'Used for the windows hook.
Public hHook As Long

'Used to track the last key pressed.
Public LastKeyPressed As Long
'
'Used to handle keys pressed before windows gets to them.
'
Public Function KeyboardProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If nCode >= 0 Then
        KeyboardProc = 1
        LastKeyPressed = wParam
    End If
    KeyboardProc = CallNextHookEx(hHook, nCode, wParam, lParam)
End Function
Public Function AutoComplete(cbo As ComboBox) As Boolean
    
    Dim sText As String
    Dim lIndex As Long
    Static iLen As Integer
    
    With cbo
        
        'Check if backspace was pressed.
        If LastKeyPressed = 8 Then
        
            'Backspace pressed so
            'strip off the last letter
            'of the combo text.
            If iLen = 0 Then iLen = 1
            .Text = Left$(.Text, iLen - 1)
        
        End If
        
        'Check if delete was pressed.
        If LastKeyPressed = 46 Then
            
            'Just exit and let
            'the delete go ahead.
            Exit Function
            
        End If
        
        sText = .Text
        iLen = Len(.Text)
        
        'Used the SendMessage API for a fast search of the combo.
        lIndex = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal .Text)
    
        'Check if a match if found.
        If lIndex >= 0 Then
            
            'If so, select it.
            .ListIndex = lIndex
            
            'set the highlight text to
            'the auto completed section.
            .SelStart = iLen
            .SelLength = Len(.List(.ListIndex)) - iLen
            
            'Set the function to true.
            AutoComplete = True
    
        Else
            
            'Select nothing.
            .ListIndex = -1
            
            'Set the function to false.
            AutoComplete = False
    
        End If
        
    End With
    
End Function

