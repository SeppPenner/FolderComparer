VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Folder Comparer"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   18345
   Icon            =   "FolderComparer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   18345
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check3 
      Caption         =   "Enable 3. folder"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13440
      TabIndex        =   10
      Top             =   2400
      Width           =   1815
   End
   Begin VB.DriveListBox Drive3 
      Height          =   315
      Left            =   8880
      TabIndex        =   8
      Top             =   960
      Width           =   3855
   End
   Begin VB.DirListBox Dir3 
      Height          =   4815
      Left            =   8880
      TabIndex        =   7
      Top             =   1560
      Width           =   3855
   End
   Begin VB.CommandButton Start 
      Caption         =   "Compare folders"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13200
      TabIndex        =   6
      Top             =   1560
      Width           =   2295
   End
   Begin VB.DirListBox Dir2 
      Height          =   4815
      Left            =   4560
      TabIndex        =   5
      Top             =   1560
      Width           =   3855
   End
   Begin VB.DriveListBox Drive2 
      Height          =   315
      Left            =   4560
      TabIndex        =   3
      Top             =   960
      Width           =   3855
   End
   Begin VB.DirListBox Dir1 
      Height          =   4815
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   3855
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   1920
      Left            =   16080
      Picture         =   "FolderComparer.frx":C84A
      Top             =   480
      Width           =   1920
   End
   Begin VB.Shape Shape15 
      Height          =   615
      Left            =   16920
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Shape Shape14 
      Height          =   615
      Left            =   15720
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Shape Shape13 
      Height          =   615
      Left            =   14520
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Shape Shape12 
      Height          =   615
      Left            =   13320
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Shape Shape11 
      Height          =   615
      Left            =   16920
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Shape Shape10 
      Height          =   615
      Left            =   15720
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Shape Shape9 
      Height          =   615
      Left            =   14520
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Shape Shape8 
      Height          =   615
      Left            =   13320
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Shape Shape0 
      Height          =   615
      Left            =   13320
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Shape Shape7 
      Height          =   615
      Left            =   16920
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Shape Shape6 
      Height          =   615
      Left            =   15720
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Shape Shape5 
      Height          =   615
      Left            =   14520
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Shape Shape4 
      Height          =   615
      Left            =   13320
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Shape Shape3 
      Height          =   615
      Left            =   16920
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      Height          =   615
      Left            =   15720
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      Height          =   615
      Left            =   14520
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label LabelSubfolders 
      Caption         =   " Subfolders:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13320
      TabIndex        =   25
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label LabelFiles 
      Caption         =   " Files:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13320
      TabIndex        =   24
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label LabelSize 
      Caption         =   " Size:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13320
      TabIndex        =   23
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label LabelSubfolder3 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   17040
      TabIndex        =   22
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label LabelSubfolder2 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15840
      TabIndex        =   21
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label LabelSubfolder1 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14640
      TabIndex        =   20
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label LabelFileCount3 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   17040
      TabIndex        =   19
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label LabelFileCount2 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15840
      TabIndex        =   18
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label LabelFileCount1 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14640
      TabIndex        =   17
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label LabelSize1 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14640
      TabIndex        =   16
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label LabelSize2 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15840
      TabIndex        =   15
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label LabelSize3 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   17040
      TabIndex        =   14
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label LabelFolder3 
      Caption         =   "Folder 3:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   17040
      TabIndex        =   13
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label LabelFolder2 
      Caption         =   "Folder 2:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15840
      TabIndex        =   12
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label LabelFolder1 
      Caption         =   "Folder 1:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14640
      TabIndex        =   11
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Select folder 3:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8880
      TabIndex        =   9
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Select folder 2:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Select folder 1:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" _
  Alias "ShellExecuteA" ( _
  ByVal hWnd As Long, _
  ByVal lpOperation As String, _
  ByVal lpFile As String, _
  ByVal lpParameters As String, _
  ByVal lpDirectory As String, _
  ByVal nShowCmd As Long) As Long

Private Sub Check3_Click()
    If Check3.Value = 1 Then
        Label3.Visible = True
        Dir3.Visible = True
        Drive3.Visible = True
        LabelFolder3.Visible = True
        LabelSize3.Visible = True
        LabelFileCount3.Visible = True
        LabelSubfolder3.Visible = True
        Shape3.Visible = True
        Shape7.Visible = True
        Shape11.Visible = True
        Shape15.Visible = True
    Else
        Label3.Visible = False
        Dir3.Visible = False
        Drive3.Visible = False
        LabelFolder3.Visible = False
        LabelSize3.Visible = False
        LabelFileCount3.Visible = False
        LabelSubfolder3.Visible = False
        Shape3.Visible = False
        Shape7.Visible = False
        Shape11.Visible = False
        Shape15.Visible = False
        LabelSize3.Caption = ""
        LabelFileCount3.Caption = ""
        LabelSubfolder3.Caption = ""
        LabelSize3.BackColor = vbWhite
        LabelFileCount3.BackColor = vbWhite
        LabelSubfolder3.BackColor = vbWhite
    End If
End Sub

Private Sub Drive1_Change()
    Dim d As Drive
    Set d = FSO.GetDrive(Left$(Drive1.Drive, 2))
    If d.IsReady() Then
        Dir1.Path = Left$(Drive1.Drive, 1) & ":\"
    Else
        Message = MsgBox("Drive not available", vbCritical + vbApplicationModal, "Multiple File Renamer")
        Drive1.Drive = FSO.GetDrive("C:")
    End If
End Sub

Private Sub Drive2_Change()
    Dim d As Drive
    Set d = FSO.GetDrive(Left$(Drive2.Drive, 2))
    If d.IsReady() Then
        Dir2.Path = Left$(Drive2.Drive, 1) & ":\"
    Else
        Message = MsgBox("Drive not available", vbCritical + vbApplicationModal, "Multiple File Renamer")
        Drive2.Drive = FSO.GetDrive("C:")
    End If
End Sub

Private Sub Drive3_Change()
    Dim d As Drive
    Set d = FSO.GetDrive(Left$(Drive3.Drive, 2))
    If d.IsReady() Then
        Dir3.Path = Left$(Drive3.Drive, 1) & ":\"
    Else
        Message = MsgBox("Drive not available", vbCritical + vbApplicationModal, "Multiple File Renamer")
        Drive3.Drive = FSO.GetDrive("C:")
    End If
End Sub

Private Sub Form_Load()
    Check3.Value = 1
    Label3.Visible = True
    Dir3.Visible = True
    Drive3.Visible = True
    LabelFolder3.Visible = True
End Sub

Private Sub Start_Click()
    'Reset the values:
    FileCount = 0
    FolderCount = 0
    SubFolderCount1 = 0
    SubFolderCount2 = 0
    SubFolderCount3 = 0
    LabelSubfolder1.Caption = ""
    LabelSubfolder2.Caption = ""
    LabelSubfolder3.Caption = ""
    FileCount1 = 0
    FileCount2 = 0
    FileCount3 = 0
    LabelFileCount1.Caption = ""
    LabelFileCount2.Caption = ""
    LabelFileCount3.Caption = ""
    LabelSize1.Caption = ""
    LabelSize2.Caption = ""
    LabelSize3.Caption = ""
    Bytes1 = 0
    Bytes2 = 0
    Bytes3 = 0
    Kilobytes1 = 0
    Kilobytes2 = 0
    Kilobytes3 = 0
    Megabytes1 = 0
    Megabytes2 = 0
    Megabytes3 = 0
    Gigabytes1 = 0
    Gigabytes2 = 0
    Gigabytes3 = 0
    
    Dim FSO As New FileSystemObject

    Set Folder1 = FSO.GetFolder(Dir1.List(Dir1.ListIndex))
    Set Folder2 = FSO.GetFolder(Dir2.List(Dir2.ListIndex))
    If Check3.Value = 1 Then
        Set Folder3 = FSO.GetFolder(Dir3.List(Dir3.ListIndex))
    End If
    
    On Error Resume Next
    Bytes1 = Folder1.Size
    If Err.Number = 70 Then Bytes1 = -1
    On Error GoTo 0
    On Error Resume Next
    Bytes2 = Folder2.Size
    If Err.Number = 70 Then Bytes2 = -1
    On Error GoTo 0
    If Check3.Value = 1 Then
        On Error Resume Next
        Bytes3 = Folder3.Size
        If Err.Number = 70 Then Bytes3 = -1
        On Error GoTo 0
    End If
    
    On Error Resume Next
    FileCount1 = FileCount1 + Folder1.Files.Count
    SubFolderCount1 = SubFolderCount1 + Folder1.SubFolders.Count
    Call CheckSubfolders(Folder1, 1)
    If Err.Number = 70 Then FileCount1 = -1
    On Error GoTo 0
    On Error Resume Next
    FileCount2 = FileCount2 + Folder2.Files.Count
    SubFolderCount2 = SubFolderCount2 + Folder2.SubFolders.Count
    Call CheckSubfolders(Folder2, 2)
    If Err.Number = 70 Then FileCount2 = -1
    On Error GoTo 0
    If Check3.Value = 1 Then
        On Error Resume Next
        FileCount3 = FileCount3 + Folder3.Files.Count
        SubFolderCount3 = SubFolderCount3 + Folder3.SubFolders.Count
        Call CheckSubfolders(Folder3, 3)
        If Err.Number = 70 Then FileCount3 = -1
        On Error GoTo 0
    End If
    
    If Bytes1 = -1 Then
        LabelSize1.Caption = "Error"
    Else
        Kilobytes1 = CDbl(Bytes1 / 1024)
        If Kilobytes1 > 1000 Then
            Megabytes1 = CDbl(Kilobytes1 / 1024)
            If Megabytes1 > 1000 Then
                Gigabytes1 = CDbl(Megabytes1 / 1024)
                LabelSize1.Caption = Format(Gigabytes1, "0.0") & " GB"
            Else
                LabelSize1.Caption = Format(Megabytes1, "0.0") & " MB"
            End If
        Else
            LabelSize1.Caption = Format(Kilobytes1, "0.0") & " KB"
        End If
    End If
    If Bytes2 = -1 Then
        LabelSize2.Caption = "Error"
    Else
        Kilobytes2 = CDbl(Bytes2 / 1024)
        If Kilobytes2 > 1000 Then
            Megabytes2 = CDbl(Kilobytes2 / 1024)
            If Megabytes2 > 1000 Then
                Gigabytes2 = CDbl(Megabytes2 / 1024)
                LabelSize2.Caption = Format(Gigabytes2, "0.0") & " GB"
            Else
                LabelSize2.Caption = Format(Megabytes2, "0.0") & " MB"
            End If
        Else
            LabelSize2.Caption = Format(Kilobytes2, "0.0") & " KB"
        End If
    End If
    If Check3.Value = 1 Then
        If Bytes3 = -1 Then
            LabelSize3.Caption = "Error"
        Else
            Kilobytes3 = CDbl(Bytes3 / 1024)
            If Kilobytes3 > 1000 Then
                Megabytes3 = CDbl(Kilobytes3 / 1024)
                If Megabytes3 > 1000 Then
                    Gigabytes3 = CDbl(Megabytes3 / 1024)
                    LabelSize3.Caption = Format(Gigabytes3, "0.0") & " GB"
                Else
                    LabelSize3.Caption = Format(Megabytes3, "0.0") & " MB"
                End If
            Else
                LabelSize3.Caption = Format(Kilobytes3, "0.0") & " KB"
            End If
        End If
    End If
    
    If FileCount1 = -1 Then
        LabelFileCount1.Caption = "Error"
    Else
        LabelFileCount1.Caption = FileCount1
    End If
    If FileCount2 = -1 Then
        LabelFileCount2.Caption = "Error"
    Else
        LabelFileCount2.Caption = FileCount2
    End If
    If Check3.Value = 1 Then
        If FileCount3 = -1 Then
            LabelFileCount3.Caption = "Error"
        Else
            LabelFileCount3.Caption = FileCount3
        End If
    End If
    
    If SubFolderCount1 = -1 Then
        LabelSubfolder1.Caption = "Error"
    Else
        LabelSubfolder1.Caption = SubFolderCount1
    End If
    If SubFolderCount2 = -1 Then
        LabelSubfolder2.Caption = "Error"
    Else
        LabelSubfolder2.Caption = SubFolderCount2
    End If
    If Check3.Value = 1 Then
        If SubFolderCount3 = -1 Then
            LabelSubfolder3.Caption = "Error"
        Else
            LabelSubfolder3.Caption = SubFolderCount3
        End If
    End If
    
    If Check3.Value = 1 Then
        If Bytes1 = Bytes2 And Bytes2 = Bytes3 Then
            LabelSize1.BackColor = vbGreen
            LabelSize2.BackColor = vbGreen
            LabelSize3.BackColor = vbGreen
        Else
            LabelSize1.BackColor = vbRed
            LabelSize2.BackColor = vbRed
            LabelSize3.BackColor = vbRed
        End If
        
        If FileCount1 = FileCount2 And FileCount2 = FileCount3 Then
            LabelFileCount1.BackColor = vbGreen
            LabelFileCount2.BackColor = vbGreen
            LabelFileCount3.BackColor = vbGreen
        Else
            LabelFileCount1.BackColor = vbRed
            LabelFileCount2.BackColor = vbRed
            LabelFileCount3.BackColor = vbRed
        End If
        
        If SubFolderCount1 = SubFolderCount2 And SubFolderCount2 = SubFolderCount3 Then
            LabelSubfolder1.BackColor = vbGreen
            LabelSubfolder2.BackColor = vbGreen
            LabelSubfolder3.BackColor = vbGreen
        Else
            LabelSubfolder1.BackColor = vbRed
            LabelSubfolder2.BackColor = vbRed
            LabelSubfolder3.BackColor = vbRed
        End If
    Else
        If Bytes1 = Bytes2 Then
            LabelSize1.BackColor = vbGreen
            LabelSize2.BackColor = vbGreen
        Else
            LabelSize1.BackColor = vbRed
            LabelSize2.BackColor = vbRed
        End If
        
        If FileCount1 = FileCount2 Then
            LabelFileCount1.BackColor = vbGreen
            LabelFileCount2.BackColor = vbGreen
        Else
            LabelFileCount1.BackColor = vbRed
            LabelFileCount2.BackColor = vbRed
        End If
        
        If SubFolderCount1 = SubFolderCount2 Then
            LabelSubfolder1.BackColor = vbGreen
            LabelSubfolder2.BackColor = vbGreen
        Else
            LabelSubfolder1.BackColor = vbRed
            LabelSubfolder2.BackColor = vbRed
        End If
    End If
End Sub

Public Sub URLGoTo(ByVal hWnd As Long, ByVal URL As String)
  ' hWnd: Das Fensterhandle des aufrufenden Formulars
  Screen.MousePointer = 11
  Call ShellExecute(hWnd, "Open", URL, "", "", 1)
  Screen.MousePointer = 0
End Sub

Private Sub Image1_Click()
    URLGoTo Me.hWnd, "http://franzhuber23.blogspot.de/2014/09/folder-comparer.html"
    On Error Resume Next
End Sub

Private Sub CheckSubfolders(ByVal folder As folder, ByVal Foldernumber As Integer)
    Select Case Foldernumber
        Case 1:
            For Each subfolder In folder.SubFolders
                FileCount1 = FileCount1 + subfolder.Files.Count
                SubFolderCount1 = SubFolderCount1 + subfolder.SubFolders.Count
                If subfolder.SubFolders.Count > 0 Then
                    Call CheckSubfolders(subfolder, 1)
                End If
            Next
        Case 2:
            For Each subfolder In folder.SubFolders
                FileCount2 = FileCount2 + subfolder.Files.Count
                SubFolderCount2 = SubFolderCount2 + subfolder.SubFolders.Count
                If subfolder.SubFolders.Count > 0 Then
                    Call CheckSubfolders(subfolder, 2)
                End If
            Next
        Case 3:
            For Each subfolder In folder.SubFolders
                FileCount3 = FileCount3 + subfolder.Files.Count
                SubFolderCount3 = SubFolderCount3 + subfolder.SubFolders.Count
                If subfolder.SubFolders.Count > 0 Then
                    Call CheckSubfolders(subfolder, 3)
                End If
            Next
    End Select
End Sub
