VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFileDlgEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

 

'cFileDialogEvents
'Handles events in an IFileDialog dialog
'This class is used by creating a new instance and then
' using the IFileDialog.Advise method
'Default controls (IFileDialogEvents) and custom controls
' (IFileDialogControlEvents) must be handled here.

Implements IFileDialogEvents
Implements IFileDialogControlEvents

Public Event TypeChange(nIdx As Long)
Public Event OnOk()
' Private Declare Function SysReAllocString Lib "oleaut32.dll" (ByVal pBSTR As Long, Optional ByVal pszStrPtr As Long) As Long

' Private Function LPWSTRtoSTR(lpWStr As Long, Optional ByVal CleanupLPWStr As Boolean = True) As String
' SysReAllocString VarPtr(LPWSTRtoSTR), lpWStr
' If CleanupLPWStr Then CoTaskMemFree lpWStr
' End Function
Private Sub IFileDialogEvents_OnFileOk(ByVal pfd As IFileDialog)

RaiseEvent OnOk
End Sub
Public Sub IFileDialogEvents_OnFolderChange(ByVal pfd As IFileDialog)

End Sub
Public Sub IFileDialogEvents_OnFolderChanging(ByVal pfd As IFileDialog, ByVal psiFolder As IShellItem)

End Sub

Public Sub IFileDialogEvents_OnOverwrite(ByVal pfd As IFileDialog, ByVal psi As IShellItem, pResponse As FDE_OVERWRITE_RESPONSE)
End Sub

Public Sub IFileDialogEvents_OnSelectionChange(ByVal pfd As IFileDialog)

End Sub

Public Sub IFileDialogEvents_OnShareViolation(ByVal pfd As IFileDialog, ByVal psi As IShellItem, pResponse As FDE_SHAREVIOLATION_RESPONSE)
End Sub
Public Sub IFileDialogEvents_OnTypeChange(ByVal pfd As IFileDialog)
Dim nTyp As Long
pfd.GetFileTypeIndex nTyp
RaiseEvent TypeChange(nTyp)
End Sub

Public Sub IFileDialogControlEvents_OnButtonClicked(ByVal pfdc As IFileDialogCustomize, ByVal dwIDCtl As Long)

End Sub
Public Sub IFileDialogControlEvents_OnCheckButtonToggled(ByVal pfdc As IFileDialogCustomize, ByVal dwIDCtl As Long, ByVal bChecked As BOOL)

End Sub
Public Sub IFileDialogControlEvents_OnControlActivating(ByVal pfdc As IFileDialogCustomize, ByVal dwIDCtl As Long)
End Sub
Public Sub IFileDialogControlEvents_OnItemSelected(ByVal pfdc As IFileDialogCustomize, ByVal dwIDCtl As Long, ByVal dwIDItem As Long)

End Sub


