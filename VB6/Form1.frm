VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Windows Imaging Component (WIC) Demo"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9885
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   9885
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check2 
      Caption         =   "Center"
      Height          =   315
      Left            =   660
      TabIndex        =   19
      Top             =   1080
      Width           =   1050
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1530
      TabIndex        =   18
      Text            =   "0"
      Top             =   855
      Width           =   330
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   945
      TabIndex        =   16
      Text            =   "0"
      Top             =   840
      Width           =   345
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Keep ratio"
      Height          =   285
      Left            =   90
      TabIndex        =   14
      Top             =   2280
      Value           =   1  'Checked
      Width           =   1140
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Scale"
      Height          =   285
      Left            =   1350
      TabIndex        =   10
      Top             =   1980
      Width           =   945
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   780
      TabIndex        =   9
      Top             =   1980
      Width           =   525
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   90
      TabIndex        =   7
      Top             =   1980
      Width           =   525
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save As..."
      Height          =   360
      Left            =   45
      TabIndex        =   6
      Top             =   3045
      Width           =   990
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   7395
      Left            =   2355
      ScaleHeight     =   489
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   489
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   510
      Width           =   7395
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Open"
      Height          =   360
      Left            =   585
      TabIndex        =   2
      Top             =   465
      Width           =   990
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   315
      Left            =   3300
      TabIndex        =   1
      Top             =   60
      Width           =   405
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   45
      TabIndex        =   0
      Top             =   30
      Width           =   3225
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "y="
      Height          =   195
      Left            =   1305
      TabIndex        =   17
      Top             =   870
      Width           =   210
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Position: x="
      Height          =   195
      Left            =   75
      TabIndex        =   15
      Top             =   870
      Width           =   870
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Supported output: PNG, JPG"
      Height          =   195
      Left            =   105
      TabIndex        =   13
      Top             =   2820
      Width           =   2055
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Supported input: JPG, PNG, GIF, BMP, ICO, TIF"
      Height          =   195
      Left            =   3810
      TabIndex        =   12
      Top             =   90
      Width           =   3435
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   75
      TabIndex        =   11
      Top             =   2670
      Width           =   2160
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      Height          =   195
      Left            =   645
      TabIndex        =   8
      Top             =   2025
      Width           =   90
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   90
      TabIndex        =   4
      Top             =   1665
      Width           =   2040
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   1380
      Width           =   2205
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cWI As cWICImage
Private mFile As String
Private fsd As FileSaveDialog
Private fdc As IFileDialogCustomize
Private WithEvents cEvents As cFileDlgEvents
Attribute cEvents.VB_VarHelpID = -1
Private dwCk As Long
Private strImgQ As String
Private bAuto As Boolean

Private Declare Function SysReAllocString Lib "oleaut32.dll" (ByVal pBSTR As Long, Optional ByVal pszStrPtr As Long) As Long

Private Function LPWSTRtoSTR(lpWStr As Long, Optional ByVal CleanupLPWStr As Boolean = True) As String
SysReAllocString VarPtr(LPWSTRtoSTR), lpWStr
If CleanupLPWStr Then CoTaskMemFree lpWStr
End Function

Private Sub cEvents_OnOk()
Dim lp As Long, sz As Long
fdc.GetEditBoxText 3000, lp
sz = LPWSTRtoSTR(lp)
strImgQ = sz

End Sub

Private Sub cEvents_TypeChange(nIdx As Long)
Debug.Print "TypeChange " & nIdx
If (nIdx = 2) Or (nIdx = 3) Then
    fdc.SetControlState 2000, CDCS_INACTIVE
    fdc.SetControlState 3000, CDCS_INACTIVE
ElseIf nIdx = 1 Then
    fdc.SetControlState 2000, CDCS_VISIBLE Or CDCS_ENABLED
    fdc.SetControlState 3000, CDCS_VISIBLE Or CDCS_ENABLED
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = vbChecked Then
    Text4.Enabled = False
    Text5.Enabled = False
Else
    Text4.Enabled = True
    Text5.Enabled = False
End If
End Sub

Private Sub Command1_Click()
Dim fod As FileOpenDialog
Set fod = New FileOpenDialog

Dim FileFilter() As COMDLG_FILTERSPEC
ReDim FileFilter(1)

FileFilter(0).pszName = "Supported Image Files"
FileFilter(0).pszSpec = "*.jpg;*.png;*.ico;*.gif;*.bmp;*.tiff;*.raw;*.webp"

FileFilter(1).pszName = "All Files"
FileFilter(1).pszSpec = "*.*"

fod.SetTitle "Choose an image..."
fod.SetFileTypes 2, VarPtr(FileFilter(0).pszName)
On Error Resume Next
fod.Show Me.hWnd

Dim siRes As IShellItem

fod.GetResult siRes
If (siRes Is Nothing) = False Then
    Dim lpFile As Long
    siRes.GetDisplayName SIGDN_FILESYSPATH, lpFile
    mFile = LPWSTRtoSTR(lpFile)
    Text1.Text = mFile
End If


End Sub

Private Sub Command2_Click()
Set cWI = New cWICImage

Dim x As Long, y As Long

If Check2.Value = vbChecked Then
    x = -1
Else
    x = CLng(Text4.Text)
    y = CLng(Text5.Text)
End If
    
Picture1.Cls

cWI.OpenFile mFile, Picture1.hDC, x, y, , Picture1.hWnd

Label1.Caption = "Dimensions: " & cWI.ImageWidth & "x" & cWI.ImageHeight & " (" & Round(cWI.ImageWidth / cWI.ImageHeight, 2) & ":1)"
Label2.Caption = "Frame count: " & cWI.FrameCount


Picture1.Refresh
Debug.Print "PictureBox(sw=" & Picture1.ScaleWidth & ") reports cx=" & (Picture1.Image.Width / (1.5)) / Screen.TwipsPerPixelX
End Sub

Private Sub Command4_Click()
If (cWI Is Nothing) Then Exit Sub
Dim cx As Long, cy As Long
cx = CLng(Text2.Text)
cy = CLng(Text3.Text)
Dim x As Long, y As Long

If Check2.Value = vbChecked Then
    x = -1
Else
    x = CLng(Text4.Text)
    y = CLng(Text5.Text)
End If
Picture1.Cls
cWI.ScaleImage Picture1.hDC, x, y, cx, cy, Picture1.hWnd
Picture1.Refresh
End Sub

Private Sub Command3_Click()
Set fsd = New FileSaveDialog

Dim SaveFilter() As COMDLG_FILTERSPEC
ReDim SaveFilter(2)
SaveFilter(0).pszName = "JPEG Image (*.jpg)"
SaveFilter(0).pszSpec = "*.jpg"
SaveFilter(1).pszName = "PNG Image (*.png)"
SaveFilter(1).pszSpec = "*.png"
SaveFilter(2).pszName = "BMP Image (*.bmp)"
SaveFilter(2).pszSpec = "*.bmp"

fsd.SetTitle "Save image as..."
fsd.SetFileTypes UBound(SaveFilter) + 1, VarPtr(SaveFilter(0).pszName)
fsd.SetOptions FOS_STRICTFILETYPES
Set fdc = fsd

fdc.AddText 2000, "Image Quality (Percent)"
fdc.AddEditBox 3000, "100"
 

On Error Resume Next
Set cEvents = New cFileDlgEvents
fsd.Advise cEvents, dwCk
fsd.Show Me.hWnd
Dim siRes As IShellItem
fsd.GetResult siRes
If (siRes Is Nothing) = False Then
    Dim sSave As String, lpSave As Long
    Dim nFmt As Long
    siRes.GetDisplayName SIGDN_FILESYSPATH, lpSave
    sSave = LPWSTRtoSTR(lpSave)
    fsd.GetFileTypeIndex nFmt
    Debug.Print "Calling save(" & nFmt & ") " & sSave
    Dim sHR As Long
    Select Case nFmt
        Case 1
            If Right$(sSave, 4) <> ".jpg" Then sSave = sSave & ".jpg"
            sHR = cWI.SaveJPG(sSave, CSng(strImgQ) / 100)
        Case 2
            If Right$(sSave, 4) <> ".png" Then sSave = sSave & ".png"
            sHR = cWI.SavePNG(sSave)
        Case 3
            If Right$(sSave, 4) <> ".bmp" Then sSave = sSave & ".bmp"
            sHR = cWI.SaveBMP(sSave)
    End Select
    If sHR = S_OK Then
        Label4.Caption = "Saved."
    Else
        Label4.Caption = "Error: 0x" & Hex$(sHR)
    End If
Else
    Debug.Print "No item"
End If
End Sub

Private Sub Text2_Change()
If Check1.Value = vbChecked Then
    If bAuto = False Then
        Dim ratio As Single
        ratio = cWI.ImageWidth / cWI.ImageHeight
        bAuto = True
        Text3.Text = CLng(Round(CLng(Text2.Text) / ratio, 0))
        bAuto = False
    End If
End If
End Sub

Private Sub Text3_Change()
If Check1.Value = vbChecked Then
    If bAuto = False Then
        Dim ratio As Single
        ratio = cWI.ImageWidth / cWI.ImageHeight
        bAuto = True
        Text2.Text = CLng(Round(CLng(Text3.Text) / ratio, 0))
        bAuto = False
    End If
End If
End Sub
