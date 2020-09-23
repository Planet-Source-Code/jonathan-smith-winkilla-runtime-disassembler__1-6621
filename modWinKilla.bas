Attribute VB_Name = "modWinKilla"
Private Declare Function EnumWindows Lib "user32" _
    (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" _
    (ByVal hwnd As Long) As Long
Private Declare Function GetParent Lib "user32" _
    (ByVal hwnd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" _
    Alias "GetWindowLongA" (ByVal hwnd As Long, _
    ByVal nIndex As Long) As Long
Private Declare Function GetWindowText Lib "user32" _
    Alias "GetWindowTextA" (ByVal hwnd As Long, _
    ByVal lpString As String, ByVal cch As Long) As Long

Private ListAddNum As Long

Private Const GWL_HWNDPARENT = (-8)

Private Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
    'A lot of the comments in here are from a task list
    'program I was writing that displays only visible
    'windows. WinKilla, however, displays ALL CWnd
    'derived windows.
    'WinKilla also filters itself out, so the user
    'cannot involuntarily destroy WinKilla while its
    'running
    
    Static WindowText As String
    Static nRet As Long
    '***** First filter *****
    If hwnd <> frmWinKilla.hwnd Then
    'If IsWindowVisible(hWnd) Then
    '    If GetParent(hWnd) = 0 Then
    '        If GetWindowLong(hWnd, GWL_HWNDPARENT) = 0 Then
                WindowText = Space$(256)
                WindowLong = GetWindowLong(hwnd, GWL_HWNDPARENT)
                nRet = GetWindowText(hwnd, WindowText, Len(WindowText))
                WindowText = Left$(WindowText, nRet)
                '***** Second filter *****
                If Not WindowText Like "*WinKilla*" Then
                    'nRet = SendMessage(lParam, LB_ADDSTRING, 0, ByVal WindowText)
                    frmWinKilla.lstWins.AddItem Hex(WindowLong) & " - " & WindowText & " (" & hwnd & ")"
                    'Call SendMessage(lParam, LB_SETITEMDATA, nRet, ByVal hWnd)
                    frmWinKilla.lstWins.ItemData(ListAddNum) = hwnd
                    ListAddNum = ListAddNum + 1
                End If
    '        End If
    '    End If
    'End If
    End If
    
    EnumWindowsProc = True
    
End Function

Public Function FillTaskListBox(lst As ListBox) As Long
    ' This function fills a list box with the currently
    ' running apps. It then returns the final count
    lst.Clear
    Call EnumWindows(AddressOf EnumWindowsProc, lst.hwnd)
    ListAddNum = 0
    For i = 1 To lst.ListCount
        lst.ListIndex = i - 1
    Next i
    
    FillTaskListBox = lst.ListCount
End Function
