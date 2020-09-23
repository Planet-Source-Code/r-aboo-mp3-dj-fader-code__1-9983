Attribute VB_Name = "modListViewEnh"


'=====
' needed for Enhancements
Private Const LVIS_STATEIMAGEMASK As Long = &HF000

Private Type LVITEM
    mask         As Long
    iItem        As Long
    iSubItem     As Long
    state        As Long
    stateMask    As Long
    pszText      As String
    cchTextMax   As Long
    iImage       As Long
    lParam       As Long
    iIndent      As Long
End Type

Const SWP_DRAWFRAME = &H20
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const SWP_NOZORDER = &H4

Private Const LVS_EX_FULLROWSELECT = &H20
Private Const LVS_EX_GRIDLINES = &H1
Private Const LVS_EX_CHECKBOXES As Long = &H4
Private Const LVS_EX_HEADERDRAGDROP = &H10
Private Const LVS_EX_TRACKSELECT = &H8
Private Const LVS_EX_ONECLICKACTIVATE = &H40
Private Const LVS_EX_TWOCLICKACTIVATE = &H80
Private Const LVS_EX_SUBITEMIMAGES = &H2

Private Const LVM_FIRST = &H1000
Private Const LVM_SETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 54
Private Const LVM_GETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 55
Private Const LVM_GETHEADER = (LVM_FIRST + 31)

Public Const LVIF_STATE = &H8
Public Const LVM_SETITEMSTATE = (LVM_FIRST + 43)
Public Const LVM_GETITEMSTATE As Long = (LVM_FIRST + 44)

Private Const HDS_BUTTONS = &H2
Private Const GWL_STYLE = (-16)

Private Const SWP_FLAGS = SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME

Public Declare Function SendMessageAny _
                        Lib "user32" _
                        Alias "SendMessageA" _
                        (ByVal hwnd As Long, _
                        ByVal Msg As Long, _
                        ByVal wParam As Long, _
                        lParam As Any) _
                        As Long

Private Declare Function SendMessageLong Lib _
                        "user32" Alias _
                        "SendMessageA" _
                        (ByVal hwnd As Long, _
                        ByVal Msg As Long, _
                        ByVal wParam As Long, _
                        ByVal lParam As Long) _
                        As Long
                        
Private Declare Function GetWindowLong _
                        Lib "user32" _
                        Alias "GetWindowLongA" _
                        (ByVal hwnd As Long, _
                        ByVal nIndex As Long) _
                        As Long
                        
Private Declare Function SetWindowLong _
                        Lib "user32" _
                        Alias "SetWindowLongA" _
                        (ByVal hwnd As Long, _
                        ByVal nIndex As Long, _
                        ByVal dwNewLong As Long) _
                        As Long
                        
Private Declare Function SetWindowPos _
                        Lib "user32" _
                        (ByVal hwnd As Long, _
                        ByVal hWndInsertAfter As Long, _
                        ByVal x As Long, _
                        ByVal Y As Long, _
                        ByVal cx As Long, _
                        ByVal cy As Long, _
                        ByVal wFlags As Long) _
                        As Long
'=====

'=====
Public LengthPerCharacter As Long
'=====








'=====
' Description: Enables SubItem Images in a ListView
'=====
Public Function EnhListView_Add_SubitemImages( _
                lstListViewName As ListView, _
                Optional bolShowErrors As Boolean) _
                As Boolean
    
    '
    ' initiate error handler
    On Error GoTo err_EnhListView_Add_SubitemImages
    
    '
    ' set function return to true
    EnhListView_Add_SubitemImages = True
    
    '
    ' setup variables
    Dim rStyle  As Long
    Dim r       As Long
    
    '
    ' get the current styles
    rStyle = SendMessageLong(lstListViewName.hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
    
    '
    ' add the selected style to the current styles
    rStyle = rStyle Or LVS_EX_SUBITEMIMAGES
    
    '
    ' update the listview styles
    SendMessageLong lstListViewName.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle
    
    '
    ' exit before error handler
    Exit Function
    
'
' deal with errors
err_EnhListView_Add_SubitemImages:
    
    '
    ' set function return to false
    EnhListView_Add_SubitemImages = False
    '
    ' if you want notification on an error
    If bolShowErrors = True Then

    End If
    
    '
    ' initiate debug
    Debug.Print Now & vbTab & "Error in function: EnhListView_Add_SubitemImages" _
                & vbCrLf & _
                Err.Number & vbTab & Err.Description
    Debug.Assert False
    
    '
    ' exit
    Exit Function
    
End Function
'=====



'=====
' Description: Checks all Items in the ListView
'=====
Public Function EnhLitView_CheckAllItems( _
                lstListViewName As ListView, _
                Optional bolShowErrors As Boolean) _
                As Boolean
    
    '
    ' initiate error handler
    On Error GoTo err_EnhLitView_CheckAllItems
    
    '
    ' set function return to true
    EnhLitView_CheckAllItems = True
    
    '
    ' setup variables
    Dim LV          As LVITEM
    Dim lvCount     As Long
    Dim lvIndex     As Long
    Dim lvState     As Long
    Dim r           As Long
    
    '
    lvState = IIf(True, &H2000, &H1000)
    lvCount = lstListViewName.ListItems.Count - 1
    Do
        With LV
            .mask = LVIF_STATE
            .state = lvState
            .stateMask = LVIS_STATEIMAGEMASK
        End With
        r = SendMessageAny(lstListViewName.hwnd, LVM_SETITEMSTATE, lvIndex, LV)
        lvIndex = lvIndex + 1
    Loop Until lvIndex > lvCount
    
    '
    ' exit before error handler
    Exit Function
    
'
' deal with errors
err_EnhLitView_CheckAllItems:
    
    '
    ' set function return to false
    EnhLitView_CheckAllItems = False
    '
    ' if you want notification on an error
    If bolShowErrors = True Then

    End If
    
    '
    ' initiate debug
    Debug.Print Now & vbTab & "Error in function: EnhLitView_CheckAllItems" _
                & vbCrLf & _
                Err.Number & vbTab & Err.Description
    Debug.Assert False
    
    '
    ' exit
    Exit Function
    
End Function
'=====

'=====
' Description: Unchecks all items in a ListView
'=====
Public Function EnhLitView_UnCheckAllItems( _
                lstListViewName As ListView, _
                Optional bolShowErrors As Boolean) _
                As Boolean
    
    '
    ' initiate error handler
    On Error GoTo err_EnhLitView_UnCheckAllItems
    
    '
    ' set function return to true
    EnhLitView_UnCheckAllItems = True
    
    '
    ' setup variables
    Dim LV          As LVITEM
    Dim lvCount     As Long
    Dim lvIndex     As Long
    Dim lvState     As Long
    Dim r           As Long
    
    '
    lvState = IIf(True, &H2000, &H1000)
    lvCount = lstListViewName.ListItems.Count - 1
    Do
        With LV
            .mask = LVIF_STATE
            .state = lvState
            .stateMask = LVIS_STATEIMAGEMASK
        End With
        r = SendMessageAny(lstListViewName.hwnd, LVM_SETITEMSTATE, lvIndex, LV)
        lvIndex = lvIndex + 1
    Loop Until lvIndex > lvCount
    
    '
    ' exit before error handler
    Exit Function
    
'
' deal with errors
err_EnhLitView_UnCheckAllItems:
    
    '
    ' set function return to false
    EnhLitView_UnCheckAllItems = False
    '
    ' if you want notification on an error
    If bolShowErrors = True Then

    End If
    
    '
    ' initiate debug
    Debug.Print Now & vbTab & "Error in function: EnhLitView_UnCheckAllItems" _
                & vbCrLf & _
                Err.Number & vbTab & Err.Description
    Debug.Assert False
    
    '
    ' exit
    Exit Function
    
End Function
'=====


'=====
' Description: Inverts all checked items in a ListView
'=====
Public Function EnhListView_InvertAllChecks( _
                lstListViewName As ListView, _
                Optional bolShowErrors As Boolean) _
                As Boolean
    
    '
    ' initiate error handler
    On Error GoTo err_EnhListView_InvertAllChecks
    
    '
    ' set function return to true
    EnhListView_InvertAllChecks = True
    
    '
    ' setup variables
    Dim LV          As LVITEM
    Dim r           As Long
    Dim lvCount     As Long
    Dim lvIndex     As Long
    
    '
    lvCount = lstListViewName.ListItems.Count - 1
    Do
        r = SendMessageLong(lstListViewName.hwnd, LVM_GETITEMSTATE, lvIndex, LVIS_STATEIMAGEMASK)
        With LV
            .mask = LVIF_STATE
            .stateMask = LVIS_STATEIMAGEMASK
            If r And &H2000& Then
                .state = &H1000
            Else
                .state = &H2000
            End If
        End With
        r = SendMessageAny(lstListViewName.hwnd, LVM_SETITEMSTATE, lvIndex, LV)
        lvIndex = lvIndex + 1
    Loop Until lvIndex > lvCount
    
    '
    ' exit before error handler
    Exit Function
    
'
' deal with errors
err_EnhListView_InvertAllChecks:
    
    '
    ' set function return to false
    EnhListView_InvertAllChecks = False
    '
    ' if you want notification on an error
    If bolShowErrors = True Then

    End If
    
    '
    ' initiate debug
    Debug.Print Now & vbTab & "Error in function: EnhListView_InvertAllChecks" _
                & vbCrLf & _
                Err.Number & vbTab & Err.Description
    Debug.Assert False
    
    '
    ' exit
    Exit Function
    
End Function
'=====

'=====
' Description: Toggles FlatColumnHeaders in a ListView
'=====
Public Function EnhListView_Toggle_FlatColumnHeaders( _
                frmFormName As Form, _
                lstListViewName As ListView, _
                Optional bolShowErrors As Boolean) _
                As Boolean
    
    '
    ' initiate error handler
    On Error GoTo err_EnhListView_Toggle_FlatColumnHeaders
    
    '
    ' set function return to true
    EnhListView_Toggle_FlatColumnHeaders = True
    
    '
    SetWindowLong SendMessageLong(lstListViewName.hwnd, _
                                 LVM_GETHEADER, _
                                 0, _
                                 ByVal 0&), _
                                 GWL_STYLE, _
                                 GetWindowLong(SendMessageLong(lstListViewName.hwnd, _
                                                               LVM_GETHEADER, _
                                                               0, _
                                                               ByVal _
                                                               0&), _
                                                               GWL_STYLE) _
                                                               Xor HDS_BUTTONS
    SetWindowPos lstListViewName.hwnd, _
                 frmFormName.hwnd, _
                 0, _
                 0, _
                 0, _
                 0, _
                 SWP_FLAGS
    
    '
    ' exit before error handler
    Exit Function
    
'
' deal with errors
err_EnhListView_Toggle_FlatColumnHeaders:
    
    '
    ' set function return to false
    EnhListView_Toggle_FlatColumnHeaders = False
    '

    If bolShowErrors = True Then

    End If
    
    Debug.Print Now & vbTab & "Error in function: EnhListView_Toggle_FlatColumnHeaders" _
                & vbCrLf & _
                Err.Number & vbTab & Err.Description
    Debug.Assert False
 
    Exit Function
    
End Function

