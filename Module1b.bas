Attribute VB_Name = "General2"


Public gListViewTotalSelected As Long
Public gListViewSelected() As Long
Public gListViewItemToInsertBefore As Long

Public Const LVM_FIRST = &H1000
Public Const LVM_GETSELECTEDCOUNT = (LVM_FIRST + 50)


Public Const LVFI_PARAM = &H1
Public Const LVFI_STRING = &H2
Public Const LVFI_PARTIAL = &H8
Public Const LVFI_WRAP = &H20
Public Const LVFI_NEARESTXY = &H40

Public Sub ListViewGetSelectedItems(ByVal FormToUse As Form, ByVal ListViewControl As Control)

  Dim Counter As Long
  Dim SelectedCount As Long
  

  
  SelectedCount = 1
  gListViewTotalSelected = SendMessage(ListViewControl.hwnd, LVM_GETSELECTEDCOUNT, 0, 0)

  If gListViewTotalSelected > 0 Then

    ReDim gListViewSelected(gListViewTotalSelected) As Long
  
    For Counter = 1 To ListViewControl.ListItems.Count
     
       On Error GoTo ext:
       If ListViewControl.ListItems(Counter).Selected = True Then
       
         gListViewSelected(SelectedCount) = Counter
        
     
       
      
 On Error Resume Next
       Set A = frmMain.LstPlay.ListItems.Add(SelectedCount, ListViewControl.ListItems(Counter).Key, ListViewControl.ListItems(Counter).Text, , 4)
       A.SubItems(1) = ListViewControl.ListItems(Counter).SubItems(1)
       
       frmMain.LstHistory.ListItems.Remove frmMain.LstHistory.ListItems(Counter).Key
       Counter = Counter - 1
       
    SelectedCount = SelectedCount + 1
       End If
     
    Next Counter
ext:
  End If

  
  
End Sub

Public Sub ListViewDelSelectedItems(ByVal FormToUse As Form, ByVal ListViewControl As Control)
Dim Counter As Long, SelectedCount As Long

SelectedCount = 1
  gListViewTotalSelected = SendMessage(ListViewControl.hwnd, LVM_GETSELECTEDCOUNT, 0, 0)

  If gListViewTotalSelected > 0 Then
 ReDim gListViewSelected(gListViewTotalSelected) As Long
   For Counter = 1 To ListViewControl.ListItems.Count
     On Error GoTo ext:
       If ListViewControl.ListItems(Counter).Selected = True Then
       
         gListViewSelected(SelectedCount) = Counter
        On Error Resume Next
      ListViewControl.ListItems.Remove ListViewControl.ListItems(Counter).Key
       Counter = Counter - 1
       SelectedCount = SelectedCount + 1
       End If
     Next Counter
ext:
  End If
  End Sub

Public Sub ListViewSetAsNext(ByVal FormToUse As Form, ByVal ListViewControl As Control)

  Dim Counter As Long, SelectedCount As Long
  Dim T As String, A As String

  
  SelectedCount = 1
  gListViewTotalSelected = SendMessage(ListViewControl.hwnd, LVM_GETSELECTEDCOUNT, 0, 0)

  If gListViewTotalSelected > 0 Then

    ReDim gListViewSelected(gListViewTotalSelected) As Long
  
    For Counter = 1 To ListViewControl.ListItems.Count
     
       On Error GoTo ext:
       If ListViewControl.ListItems(Counter).Selected = True Then
       
         gListViewSelected(SelectedCount) = Counter
        
     
       
      
 On Error Resume Next
 



     T = frmMain.LstPlay.ListItems(Counter).Text
     k = frmMain.LstPlay.ListItems(Counter).Key
     A = frmMain.LstPlay.ListItems(Counter).SubItems(1)
         
       
       ListViewControl.ListItems.Remove ListViewControl.ListItems(Counter).Key
      
       
  MsgBox T, A
       
     Set x = frmMain.LstPlay.ListItems.Add(SelectedCount, k, T, , 4)
    x.SubItems(1) = A
       
    SelectedCount = SelectedCount + 1
     Counter = Counter - 1
       End If
     
    Next Counter
ext:
  End If

  
  
End Sub

Public Sub SetListViewToWholeRowSelect(ByVal ListViewhWnd As Long)
    
  Dim lStyle As Long
  
  lStyle = SendMessage(ListViewhWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0)
  lStyle = lStyle Or LVS_EX_FULLROWSELECT
  
  Call SendMessage(ListViewhWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, ByVal lStyle)

End Sub




