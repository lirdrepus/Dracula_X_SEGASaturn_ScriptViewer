VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   11550
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   19875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   11550
   ScaleWidth      =   19875
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Language 
      Caption         =   "Language"
      Height          =   495
      Left            =   13920
      TabIndex        =   12
      Tag             =   "语言"
      Top             =   10920
      Width           =   1215
   End
   Begin VB.CommandButton Save 
      Caption         =   "保存列表"
      Height          =   495
      Left            =   12480
      TabIndex        =   11
      Tag             =   "Save"
      Top             =   10920
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   120
      Top             =   10320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton import 
      Caption         =   "导入文件"
      Height          =   495
      Left            =   11040
      TabIndex        =   10
      Tag             =   "Import"
      Top             =   10920
      Width           =   1215
   End
   Begin VB.ListBox AddList 
      Height          =   10500
      Left            =   16920
      TabIndex        =   9
      Top             =   720
      Width           =   2655
   End
   Begin VB.TextBox fontStart 
      Height          =   495
      Left            =   5280
      TabIndex        =   3
      Text            =   "0"
      Top             =   10920
      Width           =   1215
   End
   Begin VB.CommandButton Add 
      Caption         =   "增加"
      Height          =   495
      Left            =   8160
      TabIndex        =   6
      Tag             =   "Add"
      Top             =   10920
      Width           =   1215
   End
   Begin VB.CommandButton Del 
      Caption         =   "删除"
      Height          =   495
      Left            =   9600
      TabIndex        =   5
      Tag             =   "Del"
      Top             =   10920
      Width           =   1215
   End
   Begin VB.TextBox fSize 
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Text            =   "100"
      Top             =   10920
      Width           =   1215
   End
   Begin VB.CommandButton PrintOut 
      Caption         =   "生成文本"
      Height          =   495
      Left            =   6720
      TabIndex        =   4
      Tag             =   "Output"
      Top             =   10920
      Width           =   1215
   End
   Begin VB.TextBox txtStart 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Text            =   "0"
      Top             =   10920
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "字符开始:"
      Height          =   375
      Left            =   4200
      TabIndex        =   8
      Tag             =   "Font Addr:"
      Top             =   11040
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "尺寸:"
      Height          =   495
      Left            =   2280
      TabIndex        =   7
      Tag             =   "Size:"
      Top             =   11040
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "文本开始:"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Tag             =   "Text Addr:"
      Top             =   11040
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim frontFile As String        ' Store file prefix (path + filename without extension) for global access
Dim TXTadd() As String         ' Array storing text start addresses
Dim TILEadd() As String        ' Array storing tile start addresses
Dim flag As Integer            ' Language flag or general toggle flag


'===========================================================
' Add_Click
'   Adds a new entry (text address + tile address) to the list.
'===========================================================
Private Sub Add_Click()
    With AddList
        ' Add formatted text and tile addresses to the list
        .AddItem Format(Right("0000000" & txtStart.Text, 8), ">") & Space(2) & _
                 Format(Right("0000000" & fontStart.Text, 8), ">")

        ' Resize arrays to store new entry
        ReDim Preserve TXTadd(.ListCount) As String, TILEadd(.ListCount) As String

        ' Save raw values
        TXTadd(.ListCount) = txtStart.Text
        TILEadd(.ListCount) = fontStart.Text
    End With
End Sub


'===========================================================
' AddList_DblClick
'   Loads selected list entry back into the textboxes.
'===========================================================
Private Sub AddList_DblClick()
    Dim i As Integer

    i = AddList.ListIndex
    txtStart.Text = Format(TXTadd(i + 1), ">")
    fontStart.Text = Format(TILEadd(i + 1), ">")
End Sub


'===========================================================
' Del_Click
'   Deletes the selected entry from the list.
'===========================================================
Private Sub Del_Click()
    If AddList.ListIndex >= 0 Then
        AddList.RemoveItem AddList.ListIndex
    End If
End Sub


'===========================================================
' import_Click
'   Opens a CSV/PTN/CHR file and loads text/tile address pairs.
'===========================================================
Private Sub import_Click()
    On Error Resume Next

    With CD1
        .DialogTitle = "Open File"
        .Filter = "Text Index (*.csv)|*.csv|CHR Font (*.CHR)|*.CHR|Original Text (*.PTN)|*.PTN"
        .InitDir = App.Path
        .ShowOpen
        opnFile = .FileName
    End With

    ' User cancelled
    If Err.Number = cdlCancel Then
        Exit Sub

    ' Other error
    ElseIf Err.Number <> 0 Then
        MsgBox "Error Code: " & Format$(Err.Number) & vbCrLf & Err.Description

    Else
        ' Extract file prefix (path + filename without extension)
        spFile = Split(opnFile, ".")
        frontFile = spFile(0)

        ' Read CSV file to determine number of lines
        Open frontFile & ".csv" For Binary As #1
            l = Split(Input(LOF(1), 1), vbCrLf)
        Close #1

        LineNum = UBound(l)

        ' Resize arrays to match number of lines
        ReDim TXTadd(0 To LineNum) As String, TILEadd(0 To LineNum) As String

        ' Load CSV content
        Open frontFile & ".csv" For Input As #1
            Input #1, TXTadd(0), TILEadd(0)

            n = 1
            AddList.Clear

            Do Until EOF(1)
                Input #1, TXTadd(n), TILEadd(n)

                AddList.AddItem Format(Right("0000000" & TXTadd(n), 8), ">") & Space(2) & _
                                 Format(Right("0000000" & TILEadd(n), 8), ">")

                n = n + 1
            Loop
        Close #1
    End If

    On Error GoTo 0
End Sub


'===========================================================
' Language_Click
'   Calls the global language switcher.
'===========================================================
Private Sub Language_Click()
    SwitchLanguage Me
End Sub


'===========================================================
' PrintOut_Click
'   Reads PTN text data and prints tiles using fPrint().
'===========================================================
Private Sub PrintOut_Click()
    Dim Size As Integer
    Dim Horizon As Integer, Vertical As Integer
    Dim TXT((15 + 1) * 16 - 1) As Byte
    Dim txtS As Long
    Dim H1 As Integer, V1 As Integer
    Dim n As Integer

    txtS = CLng("&h" & txtStart.Text)
    Size = Val(fSize.Text)

    ' Read PTN text data
    Open frontFile & ".ptn" For Binary As #2
        Get #2, txtS + 1, TXT
    Close #2

    Form1.Cls

    '-------------------------
    ' First 128 characters
    '-------------------------
    Horizon = 0: Vertical = 0
    H1 = 0: V1 = 0

    For n = 1 To 127 Step 2
        Call fPrint(Horizon, Vertical, TXT(n), frontFile, TXT(n - 1))

        H1 = H1 + 8
        Horizon = H1 * Size

        If H1 Mod 160 = 0 Then
            V1 = V1 + 16
            Vertical = V1 * Size
            Horizon = 0
            H1 = 0
        End If
    Next n

    '-------------------------
    ' Next 128 characters
    '-------------------------
    Horizon = 0: Vertical = 8 * Size
    H1 = 0: V1 = 8

    For n = 129 To 255 Step 2
        Call fPrint(Horizon, Vertical, TXT(n), frontFile, TXT(n - 1))

        H1 = H1 + 8
        Horizon = H1 * Size

        If H1 Mod 160 = 0 Then
            V1 = V1 + 16
            Vertical = V1 * Size
            Horizon = 0
            H1 = 0
        End If
    Next n
End Sub


'===========================================================
' Save_Click
'   Saves TXTadd() and TILEadd() into a text file.
'===========================================================
Private Sub Save_Click()
    Dim n As Integer

    If AddList.ListCount >= 1 Then
        Open "test.txt" For Output As #2
            Print #2, "Text Address,Tile Address"

            For n = 1 To UBound(TXTadd)
                Print #2, TXTadd(n) & "," & TILEadd(n)
            Next n
        Close #2
    End If
End Sub


