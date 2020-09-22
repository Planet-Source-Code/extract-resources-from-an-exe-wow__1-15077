VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File explorer: No Open File"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   11010
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   330
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Open module"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1440
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0454
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":08A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0CFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1150
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":15A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":19F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1E4C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   5415
      Left            =   5160
      ScaleHeight     =   5355
      ScaleWidth      =   5595
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   5655
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5295
         Left            =   120
         ScaleHeight     =   5295
         ScaleWidth      =   5655
         TabIndex        =   3
         Top             =   120
         Width           =   5655
      End
      Begin VB.Label Info1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         Height          =   1335
         Left            =   -360
         TabIndex        =   2
         Top             =   -960
         Width           =   2535
      End
   End
   Begin VB.PictureBox Picture8 
      BorderStyle     =   0  'None
      Height          =   5535
      Left            =   5160
      ScaleHeight     =   5535
      ScaleWidth      =   8655
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   8655
      Begin VB.PictureBox Picture9 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   5760
         ScaleHeight     =   375
         ScaleWidth      =   135
         TabIndex        =   11
         Top             =   0
         Width           =   135
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5205
         Left            =   0
         TabIndex        =   10
         Top             =   230
         Width           =   5760
      End
      Begin VB.Line Line1 
         X1              =   5745
         X2              =   5745
         Y1              =   1440
         Y2              =   0
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   0
         Picture         =   "Form1.frx":22A0
         Top             =   0
         Width           =   8280
      End
   End
   Begin VB.PictureBox Picture7 
      BorderStyle     =   0  'None
      Height          =   1080
      Left            =   10920
      Picture         =   "Form1.frx":8A62
      ScaleHeight     =   1080
      ScaleWidth      =   615
      TabIndex        =   8
      Top             =   5650
      Width           =   615
   End
   Begin VB.PictureBox Picture6 
      BorderStyle     =   0  'None
      Height          =   5525
      Left            =   4995
      Picture         =   "Form1.frx":8F04
      ScaleHeight     =   5520
      ScaleWidth      =   75
      TabIndex        =   6
      Top             =   0
      Width           =   75
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   0
      Picture         =   "Form1.frx":B886
      ScaleHeight     =   1575
      ScaleWidth      =   11655
      TabIndex        =   5
      Top             =   5520
      Width           =   11655
      Begin VB.TextBox General 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   955
         Left            =   310
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   7
         Text            =   "Form1.frx":42BB8
         Top             =   180
         Width           =   10595
      End
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   345
      Width           =   4725
      _ExtentX        =   8334
      _ExtentY        =   8916
      _Version        =   393217
      Indentation     =   0
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   5535
      Left            =   0
      Picture         =   "Form1.frx":42BD2
      ScaleHeight     =   5535
      ScaleWidth      =   4935
      TabIndex        =   4
      Top             =   0
      Width           =   4935
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Resource tree"
         Height          =   240
         Left            =   1800
         TabIndex        =   12
         Top             =   0
         Width           =   1095
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   2280
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'For extracting the icon info from a file...
Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
'For removing (closing) the icon from access with the program...
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
'For drawing the icon into the picture box...
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
'Also for drawing icons in to the picture box...
Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
'A flag for running the DrawIconEx API normally without any options...
Private Const DI_NORMAL = &H3

Dim FileName As String
Dim aFunc(1 To 1000) As String
Dim x() As Integer
Dim PeInfo As clsPEInfo
Dim Info(100) As String
Dim InfoSize(100)
Dim de As Node
Dim sFun(1 To 1000)
Dim aOrd(1 To 1000)
Dim aNames(1 To 1000) As Integer
Dim Indexed As Long
Private Sub LoadModuleA()
Dim PeInfo As clsPEInfo, Idx As Long
Dim aNames() As String
Dim lRet As Long, aOrd() As Integer
Show
Refresh
Me.Caption = "File explorer: " & FileName & " (PE)"
'Form2.Show
Form2.ProgressBar1.Value = 10
General.Text = "Loading module, " & Form2.ProgressBar1.Value & "% complete."

'Dim aNames() As String

Set PeInfo = New clsPEInfo
'CD.ShowOpen

  FileInfo (FileName)
   PeInfo.Load FileName
   
   
 TreeView1.Nodes.Clear
 
   
TreeView1.Nodes.Add , , "Root", FileName, 1, 2

TreeView1.Nodes.Add "Root", tvwChild, "Header", "Header", 1, 2
TreeView1.Nodes.Add "Root", tvwChild, "res", "Resources", 1, 2
TreeView1.Nodes.Add "Root", tvwChild, "ver", "Version information", 1, 2

TreeView1.Nodes.Add "ver", tvwChild, "VNum", "1 [" & FileInfo(FileName).LanguageID & "]", 7, 7

TreeView1.Nodes.Add "Header", tvwChild, "Section", "Sections", 1, 2
TreeView1.Nodes.Add "Header", tvwChild, "ImageInfo", "Image information", 1, 2
TreeView1.Nodes.Add "Header", tvwChild, "Headver", "Header version information", 1, 2
TreeView1.Nodes.Add "Header", tvwChild, "chsum", "Checksum information", 1, 2

TreeView1.Nodes.Add "ImageInfo", tvwChild, "Size", "Image size", 3, 3
TreeView1.Nodes.Add "ImageInfo", tvwChild, "Version", "Image version", 3, 3

TreeView1.Nodes.Add "Header", tvwChild, "Head", "Header information", 3, 3
TreeView1.Nodes.Add "chsum", tvwChild, "Checksum", "Header CheckSum", 3, 3
TreeView1.Nodes.Add "Headver", tvwChild, "Link", "Link version", 3, 3
TreeView1.Nodes.Add "Header", tvwChild, "Machine", "Machine", 3, 3
TreeView1.Nodes.Add "Headver", tvwChild, "OS", "OS version", 3, 3
TreeView1.Nodes.Add "Header", tvwChild, "PreferredBase", "Preferred base", 3, 3
TreeView1.Nodes.Add "chsum", tvwChild, "RCheck", "Real CheckSum", 3, 3
TreeView1.Nodes.Add "Section", tvwChild, "Sections", "Section count", 3, 3
TreeView1.Nodes.Add "Header", tvwChild, "SubSys", "Subsystem", 3, 3
TreeView1.Nodes.Add "Headver", tvwChild, "SubSysver", "Subsystem version", 3, 3
Form2.ProgressBar1.Value = 15
General.Text = "Loading module, " & Form2.ProgressBar1.Value & "% complete."

TreeView1.Nodes.Add "res", tvwChild, "M", "Module export/import information", 1, 2
TreeView1.Nodes.Add "res", tvwChild, "icons", "Icons", 1, 2
TreeView1.Nodes.Add "res", tvwChild, "bmp", "Bitmaps", 1, 2
TreeView1.Nodes.Add "res", tvwChild, "cur", "Cursors", 1, 2
TreeView1.Nodes.Add "res", tvwChild, "str", "Strings", 1, 2
TreeView1.Nodes.Add "res", tvwChild, "version1", "Version table", 1, 2

TreeView1.Nodes.Add "version1", tvwChild, "Num", "1 [" & FileInfo(FileName).LanguageID & "]", 7, 7

TreeView1.Nodes.Add "Num", tvwChild, "Cname1", "Company name", 3, 3
TreeView1.Nodes.Add "Num", tvwChild, "Fdes1", "File description", 3, 3
TreeView1.Nodes.Add "Num", tvwChild, "Fver1", "File version", 3, 3
TreeView1.Nodes.Add "Num", tvwChild, "InternalName1", "Internal name", 3, 3
TreeView1.Nodes.Add "Num", tvwChild, "Legal1", "Legal copyright", 3, 3
TreeView1.Nodes.Add "Num", tvwChild, "OriginalF1", "Original file name", 3, 3
TreeView1.Nodes.Add "Num", tvwChild, "Pro1", "Product name", 3, 3
TreeView1.Nodes.Add "Num", tvwChild, "ProVer1", "Product version", 3, 3


TreeView1.Nodes.Add "M", tvwChild, "Dependencies", "Dependencies", 1, 2
TreeView1.Nodes.Add "M", tvwChild, "import", "Imported library functions", 1, 2
TreeView1.Nodes.Add "M", tvwChild, "export", "Exported library functions", 1, 2
TreeView1.Nodes.Add "M", tvwChild, "importf", "Imported modules", 1, 2
Form2.ProgressBar1.Value = 20
General.Text = "Loading module, " & Form2.ProgressBar1.Value & "% complete."

TreeView1.Nodes.Add "VNum", tvwChild, "Cname", "Company name", 3, 3
TreeView1.Nodes.Add "VNum", tvwChild, "Fdes", "File description", 3, 3
TreeView1.Nodes.Add "VNum", tvwChild, "Fver", "File version", 3, 3
TreeView1.Nodes.Add "VNum", tvwChild, "InternalName", "Internal name", 3, 3
TreeView1.Nodes.Add "VNum", tvwChild, "Legal", "Legal copyright", 3, 3
TreeView1.Nodes.Add "VNum", tvwChild, "OriginalF", "Original file name", 3, 3
TreeView1.Nodes.Add "VNum", tvwChild, "Pro", "Product name", 3, 3
TreeView1.Nodes.Add "VNum", tvwChild, "ProVer", "Product version", 3, 3

For c = 1 To TreeView1.Nodes.Count
Next c


If IconCount <> 0 Then
For i = 1 To IconCount
Form2.Caption = "Processing icon: #" & i & "/" & IconCount
Form2.ProgressBar1.Value = 20 + (i / 4)
General.Text = "Loading module, " & Format(Form2.ProgressBar1.Value, "##") & "% complete."

TreeView1.Nodes.Add "icons", tvwChild, "I" & i, i & " [" & FileInfo(FileName).LanguageID & "]", 5, 5
Next i
End If
'TreeView1.Nodes.Add "Header", tvwChild, "Sction", "Sections", 1, 2
'TreeView1.Nodes.Add "Header", tvwChild, "Sction", "Sections", 1, 2
GetBitmapsA FileName
Form2.Caption = "Scanning for Bitmaps..."

Form2.ProgressBar1.Value = 50
General.Text = "Loading module, " & Form2.ProgressBar1.Value & "% complete."

GetCursorsA FileName
Form2.ProgressBar1.Value = 55
General.Text = "Loading module, " & Form2.ProgressBar1.Value & "% complete."
PeInfo.Load FileName
Form2.Caption = "Scanning for Cursors..."
PeInfo.EnumerateSections aNames()
      
      For Idx = 0 To UBound(aNames)
         With TreeView1.Nodes.Add("Sections", twvchild, "A" & Idx + 1, aNames(Idx, 0), 3, 3)
            Info(Idx + 1) = aNames(Idx, 1)
            InfoSize(Idx + 1) = aNames(Idx, 2)
         End With
      Next
Form2.Caption = "Loading data..."
Form2.ProgressBar1.Value = 60
General.Text = "Loading module, " & Form2.ProgressBar1.Value & "% complete."

LoadTree FileName

Form2.ProgressBar1.Value = 70
General.Text = "Loading module, " & Form2.ProgressBar1.Value & "% complete."

'PeInfo.EnumerateImportedModules aNames(), Idx
   
'      If Idx > 0 Then
      
'         For Idx = 0 To Idx - 1
'            TreeView1.Nodes.Add "importf", tvwChild, "", aNames(Idx)
'         Next
      
'      End If
   

  
      PeInfo.EnumerateExportedFunctions aNames(), Idx
      
      If Idx > 0 Then
      
         For Idx = 0 To Idx - 1
            'With TreeView1.Nodes.Add("import", tvwChild, aNames(Idx, 1), aNames(Idx, 1))
'               treeview1.Nodes.add SubItems(1) = aNames(Idx, 0)
               If aNames(Idx, 0) <> "" Then
                    TreeView1.Nodes.Add "export", tvwChild, , aNames(Idx, 0), 8, 8
                End If
            'End With
         Next
      
      End If
      
   Form2.ProgressBar1.Value = 80
General.Text = "Loading module, " & Form2.ProgressBar1.Value & "% complete."
   
   Form2.Caption = "Scaning for DLL references..."
      PeInfo.EnumerateImportedModules aNames(), Idx
        
        Form2.ProgressBar1.Value = 90
General.Text = "Loading module, " & Form2.ProgressBar1.Value & "% complete."
      
      If Idx > 0 Then
      Form2.Caption = "Complete..."
         
         For Idx = 0 To Idx - 1
            'PeInfo.EnumerateImportedFunctions aNames(Idx), sFun(), aOrd, Indexed
            TreeView1.Nodes.Add "import", tvwChild, , aNames(Idx), 8, 8
            TreeView1.Nodes.Add "importf", tvwChild, , aNames(Idx), 8, 8
            Module aNames(Idx)
            For i = 0 To lRet + 1
                'TreeView1.Nodes.Add aNames(Idx), tvwChild, , aFunc(i)
            Next
         Next
      
      End If
     Form2.ProgressBar1.Value = 100
General.Text = "Loading module, " & Form2.ProgressBar1.Value & "% complete."
  
  '''''''''''''''''''''''''''''
FillVer
      Form2.Hide
      End Sub

 
Private Sub Form_Load()
CD.DialogTitle = "Select module"
'CD.Filter = "All windows executables (*.DLL, *.EXE, *.OCX)|*.DLL|*.EXE|*.OCX"
CD.Filter = "Standard executable (*.EXE)|*.EXE|"
CD.Filter = CD.Filter & "Dynamic library (*.DLL)|*.DLL|"
CD.Filter = CD.Filter & "ActiveX control (*.OCX)|*.OCX"

CD.ShowOpen
If CD.CancelError = False Then
    If CD.FileName <> "" Then
        FileName = CD.FileName
        LoadModuleA ' CD.FileName
Exit Sub
    End If
End If
General.Text = "Please select module..."
TreeView1.Nodes.Clear
List1.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub General_Change()
Me.Refresh

End Sub

Private Sub Picture1_Click()
MsgBox "Picture1"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Form_Load

End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
Dim PeInfo As clsPEInfo

Set PeInfo = New clsPEInfo
PeInfo.Load FileName
'MsgBox FileInfo(FileName).vfi
Picture2.Cls
Picture2.Refresh
Picture8.Visible = False
Picture1.Visible = False
Picture1.Top = -999999

'Picture2.Visible = False

List1.Clear
'List1.Selected(0) = False
'List1.Selected(1) = False
'List1.Se
'lected(2) = False
'List1.Selected(3) = False
'List1.Selected(4) = False
'List1.Selected(5) = False
'List1.Selected(6) = False
'List1.Selected(7) = False

If Node.Key = "Num" Then
Picture8.Visible = True
FillVer
End If
If Node.Key = "VNum" Then
Picture8.Visible = True
FillVer
End If
'If Node.Parent = "VNum" Then Picture8.Visible = True

If Left(Node.Text, 3) <> "1 [" Then Node.ExpandedImage = 2
'If Node.Expanded = False Then Node.Image = 1
'Form2.Show
'Form2.Caption = "Processing..."
Select Case TreeView1.SelectedItem.Text

Case "Strings"
    Info1.Caption = "No strings detected."
  '  General.Text = General.Text & "Scanning for string..." & vbCrLf
  '  General.Text = General.Text & "No strings detected." & vbCrLf
    
Case "File version"
    Info1.Caption = FileInfo(FileName).FileVersion
FillVer
  
  '  General.Text = General.Text & "FileVersion: " & FileInfo(FileName).FileVersion & vbCrLf
List1.Selected(2) = True
 
Picture8.Visible = True
Case "Company name"
FillVer
    
Info1.Caption = FileInfo(FileName).CompanyName
List1.Selected(0) = True
 
'MsgBox Info1.Caption
  '  General.Text = General.Text & "CompanyName: " & FileInfo(FileName).CompanyName & vbCrLf
Picture8.Visible = True
'Picture1.Top = -99999
'Picture2.Visible = False
'MsgBox Picture8.Visible
Exit Sub
Case "File description"
FillVer
    
    Info1.Caption = FileInfo(FileName).FileDescription
 List1.Selected(1) = True
   
   ' General.Text = General.Text & "FileDescription: " & FileInfo(FileName).FileDescription & vbCrLf
Picture8.Visible = True

Case "Internal name"
FillVer
    
    Info1.Caption = FileInfo(FileName).InternalName
    'General.Text = General.Text & "InternalName: " & FileInfo(FileName).InternalName & vbCrLf
List1.Selected(3) = True
Picture8.Visible = True
'Picture1.Top = -99999
Exit Sub
Case "Legal copyright"
FillVer
    
    Info1.Caption = FileInfo(FileName).LegalCopyright
'    General.Text = General.Text & "LegalCopyright: " & FileInfo(FileName).LegalCopyright & vbCrLf
List1.Selected(4) = True
 
Picture8.Visible = True
Case "Original file name"
FillVer
    
    Info1.Caption = FileInfo(FileName).OrigionalFileName
 '   General.Text = General.Text & "OrigionalFilename: " & FileInfo(FileName).OrigionalFileName & vbCrLf
List1.Selected(5) = True
 
Picture8.Visible = True
Case "Product name"
FillVer
    
    Info1.Caption = FileInfo(FileName).ProductName
  '  General.Text = General.Text & "ProductName: " & FileInfo(FileName).ProductName & vbCrLf
List1.Selected(6) = True
 Picture8.Visible = True
Case "Product version"
FillVer
    
    Info1.Caption = FileInfo(FileName).ProductVersion
   ' General.Text = General.Text & "ProductVersion: " & FileInfo(FileName).ProductVersion & vbCrLf
List1.Selected(7) = True
 
Picture8.Visible = True
Case "Machine"
   
    Info1.Caption = PeInfo.MachineName & " (" & Hex(PeInfo.Machine) & ")"
    
    List1.AddItem "Machine name     " & Info1.Caption
    Picture8.Visible = True
    'General.Text = General.Text & Info1.Caption & vbCrLf
Case "Header CheckSum"
    Info1.Caption = Hex(PeInfo.CheckSum)
List1.AddItem "CheckSum         " & Info1.Caption
Picture8.Visible = True

Case "Section count"
Info1.Caption = PeInfo.Sections
List1.AddItem "Section count    " & Info1.Caption
Picture8.Visible = True


'General.Text = General.Text & Info1.Caption & vbCrLf
Case "Real CheckSum"
    Info1.Caption = Hex(PeInfo.RealCheckSum)
 List1.AddItem "Real CheckSum    " & Info1.Caption
Picture8.Visible = True
 
 '   General.Text = General.Text & Info1.Caption & vbCrLf
Case "Link version"
    Info1.Caption = PeInfo.LinkerVer
  List1.AddItem "Linker version   " & Info1.Caption
Picture8.Visible = True
  
  '  General.Text = General.Text & Info1.Caption & vbCrLf
Case "Image version"
        Info1.Caption = PeInfo.ImageVer
  List1.AddItem "Image version    " & Info1.Caption
Picture8.Visible = True
  
Case "Header information"
        'Info1.Caption = "Select catagory."
  List1.AddItem "Select category  "
Picture8.Visible = True
  
  '
  General.Text = General.Text & Info1.Caption & vbCrLf
Case "Image size"
    Info1.Caption = PeInfo.ImageSize
  List1.AddItem "Image size       " & Info1.Caption
Picture8.Visible = True
  
  '  General.Text = General.Text & Info1.Caption & vbCrLf
Case "OS version"
    Info1.Caption = PeInfo.OSVer
  List1.AddItem "OS version       " & Info1.Caption
Picture8.Visible = True
  
  '  General.Text = General.Text & Info1.Caption & vbCrLf
Case "Preferred base"
    Info1.Caption = PeInfo.PreferredBase
  
  List1.AddItem "Preferred base   " & Info1.Caption
Picture8.Visible = True
  
  '  General.Text = General.Text & Info1.Caption & vbCrLf
Case "Subsystem"
    Info1.Caption = PeInfo.SubSystem
   List1.AddItem "Subsystem        " & Info1.Caption
Picture8.Visible = True
   
   ' General.Text = General.Text & Info1.Caption & vbCrLf
Case "Subsystem version"
    Info1.Caption = PeInfo.SubSystemVer
 List1.AddItem "SubSys version   " & Info1.Caption
Picture8.Visible = True
  
  ' General.Text = General.Text & Info1.Caption & vbCrLf
Case Else
    Info1.Caption = Node.Text
'General.Text = General.Text & Info1.Caption & vbCrLf
End Select

If TreeView1.SelectedItem.Key = "A0" Then
Info1.Caption = Info(0)
List1.AddItem "Type description " & Info1.Caption
Picture8.Visible = True

End If

If TreeView1.SelectedItem.Key = "A1" Then
Info1.Caption = Info(1)
List1.AddItem "Section info     " & Info1.Caption
Picture8.Visible = True

End If

If TreeView1.SelectedItem.Key = "A2" Then
Info1.Caption = Info(2)
List1.AddItem "Section info     " & Info1.Caption
Picture8.Visible = True

End If

If TreeView1.SelectedItem.Key = "A3" Then
Info1.Caption = Info(3)
List1.AddItem "Section info     " & Info1.Caption
Picture8.Visible = True

End If

If TreeView1.SelectedItem.Key = "A4" Then
Info1.Caption = Info(4)
List1.AddItem "Section info     " & Info1.Caption
Picture8.Visible = True

End If


If TreeView1.SelectedItem.Key = "A5" Then
Info1.Caption = Info(5)
List1.AddItem "Section info     " & Info1.Caption
Picture8.Visible = True

End If

If TreeView1.SelectedItem.Key = "A6" Then
Info1.Caption = Info(6)
List1.AddItem "Section info     " & Info1.Caption
Picture8.Visible = True

End If

If TreeView1.SelectedItem.Key = "A7" Then
Info1.Caption = Info(7)
List1.AddItem "Section info     " & Info1.Caption
Picture8.Visible = True

End If

If TreeView1.SelectedItem.Key = "A8" Then
Info1.Caption = Info(8)
List1.AddItem "Section info     " & Info1.Caption
Picture8.Visible = True

End If

If TreeView1.SelectedItem.Key = "A9" Then
Info1.Caption = Info(9)
List1.AddItem "Section info     " & Info1.Caption
Picture8.Visible = True

End If

On Error Resume Next

Picture2.Cls
Picture2.Refresh
    
If Left(TreeView1.SelectedItem.Key, 1) = "I" Then
    Picture1.Cls
   Picture1.Visible = True
   Picture1.Top = 0
    'MsgBox "I"
   ' General.Text = General.Text & "Found icon: " & Node.Text & vbCrLf
    IconList = ExtractIcon(0, FileName, Mid(TreeView1.SelectedItem.Key, 2, 100) - 1)
    'Put the icon into the picture box using the DrawIcon API...
    DrawIcon Picture2.hdc, 0, 0, IconList
   ' General.Text = General.Text & "Icon draw successfuly..." & vbCrLf
End If

If Left(TreeView1.SelectedItem.Key, 1) = "B" Then
    Picture1.Cls
    Picture1.Refresh
    'MsgBox "B"
    'General.Text = General.Text & "Found bitmap: " & Node.Text & vbCrLf
    Picture1.Visible = True
    Picture1.Top = 0
    Picture2.BackColor = vbWhite
    Picture2.Cls
    Picture2.Refresh
    
    'MsgBox "FOund"
    LoadBitmapA FileName, "#" & Mid(TreeView1.SelectedItem.Key, 2, 100), Picture2
    'General.Text = General.Text & "Bitmap successfuly loaded..." & vbCrLf
    'MsgBox "Done."
End If

If Left(TreeView1.SelectedItem.Key, 1) = "C" Then
    Picture1.Cls
Picture1.Visible = True
    Picture1.Refresh
    Picture1.Top = 0
    'MsgBox "C"
    'General.Text = General.Text & "Found cursor: " & Node.Text & vbCrLf
    LoadCursorA FileName, "#" & Mid(TreeView1.SelectedItem.Key, 2, 100), Picture2
   ' General.Text = General.Text & "Cursor loaded..." & vbCrLf
    
End If
'Form2.Hide
End Sub

Public Sub LoadTree(ByVal FileName As String)
Dim PEI As clsPEInfo, NewNode As Node
Dim aModules() As String, lModCount As Long

   On Error Resume Next

   If ParentNode Is Nothing Then
      
      'Me.Caption = "Dependecies Tree for " & FileName
      
      Set ParentNode = TreeView1.Nodes.Add("Dependencies", tvwChild, , FileName, 1, 2)
      
   End If
   
   Set PEI = New clsPEInfo
   
   With PEI
   
      On Error Resume Next
      
      .Load FileName
      
      If Err.Number = 0 Then
         
         .EnumerateImportedModules aModules, lModCount
         .Unload
         
      Else
         
        TreeView1.Nodes.Add ParentNode, tvwChild, , aModules(lModCount) & " - NOT FOUND", 8, 8
         
         Exit Sub
         
      End If
      
   End With

   For lModCount = 0 To lModCount - 1
      
      With TreeView1.Nodes
              
         Set NewNode = .Add(ParentNode, tvwChild, "*" & aModules(lModCount), aModules(lModCount), 8, 8)
         
         If Not NewNode Is Nothing Then
            ' This is the first reference
            ' to the module
            LoadTree aModules(lModCount) ', 'NewNode
            
            'lstDep.AddItem aModules(lModCount)
            
         Else
            ' This module has been
            ' added in another branch
            .Add ParentNode, tvwChild, , aModules(lModCount), 8, 8
         End If
         
         Set NewNode = Nothing
         
      End With
      
   Next
   
   ParentNode.Sorted = True
   
End Sub

Function Load()

   'PeInfo.EnumerateExportedFunctions aNames(), Idx
      
      'If Idx > 0 Then
      
      '   For Idx = 0 To Idx - 1
      '      With .Add(, , aNames(Idx, 1))
      '         .SubItems(1) = aNames(Idx, 0)
      '         .SubItems(2) = aNames(Idx, 2)
      '      End With
      '   Next
      '
      'End If
      
   'End With
   
   'With lstImpModules
   
   '   .Clear
      
      'PeInfo.EnumerateImportedModules aNames(), Idx
   
      If Idx > 0 Then
      
         For Idx = 0 To Idx - 1
            TreeView1.Nodes.Add "import", tvwChild, "*" & aNames(Idx), aNames(Idx), 8, 8
         Next
      
      End If
   
   'End With
   
End Function

Function Module(ModuleName As String)
Dim aFunc() As String
Dim lRet As Long, aOrd() As Integer
Dim PeInfo As clsPEInfo

Set PeInfo = New clsPEInfo

PeInfo.Load FileName

PeInfo.EnumerateImportedFunctions ModuleName, aFunc, aOrd(), lRet
   
   'PEInfo.
   
      If lRet > 0 Then
      
         For lRet = 0 To lRet - 1
            If aFunc(lRet) = "" Then
              ' Debug.Print aOrd(lRet)
            Else
                  TreeView1.Nodes.Add "*" & ModuleName, tvwChild, , aFunc(lRet), 8, 8
              '    Debug.Print aOrd(lRet)
               
            End If
         Next
      
      End If
   
End Function


Function IconCount()
'The sub for when you select a file from the files list...
'Declare variables...
Dim Count, Path, Procedure
'Clear the list...
'Make the Path variale equal the full path name (ex. C:\Hey\Sup.ico)...
'Path = Folders.Path & "\" & Files.FileName
'Make the Count variable equal the amount of icons in the specified file
'using the ExtractIcon...
Count = ExtractIcon(App.hInstance, FileName, -1)
'If there are less than 1 (0) icons then...
If Count < 1 Then
    'Tell the user...
IconCount = 0
'Exit the sub...
    Exit Function
'End the If statement...
End If
'Make the Procedure variable equal the amount of icons in the file...
For Procedure = 0 To Count - 1
     'Add that file to a listbox...
'     IconSelect.AddItem Procedure
'Keep doing this until all icons are added...
Next Procedure
'If there was only 1 icon in the file t hen...
If Count = 1 Then
    'Tell the user...
inconcount = 1
'Exit the sub...
    Exit Function
'End the If statement...
End If
'Tell the user how many icons and in what file...
IconCount = Count
End Function

Sub FillVer()
List1.AddItem "CompanyName      " & FileInfo(FileName).CompanyName
    List1.AddItem "FileDescription  " & FileInfo.FileDescription
    List1.AddItem "FileVersion      " & FileInfo.FileVersion
    List1.AddItem "InternalName     " & FileInfo.InternalName
    List1.AddItem "LegalCopyright   " & FileInfo.LegalCopyright
    List1.AddItem "OriginalFilename " & FileInfo.OrigionalFileName
    List1.AddItem "ProductName      " & FileInfo.ProductName
    List1.AddItem "ProductVersion   " & FileInfo.ProductVersion

End Sub
