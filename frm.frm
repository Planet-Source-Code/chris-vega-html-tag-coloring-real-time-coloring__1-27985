VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_HTML_Editor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " HTML Editor (Tag Coloring)"
   ClientHeight    =   6510
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   10590
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "frm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   10590
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picTagsColor 
      BackColor       =   &H00C00000&
      Height          =   285
      Left            =   6435
      ScaleHeight     =   225
      ScaleWidth      =   450
      TabIndex        =   7
      Top             =   6120
      Width           =   510
   End
   Begin VB.PictureBox picCommentsColor 
      BackColor       =   &H000080FF&
      Height          =   285
      Left            =   4545
      ScaleHeight     =   225
      ScaleWidth      =   405
      TabIndex        =   6
      Top             =   6120
      Width           =   465
   End
   Begin MSComDlg.CommonDialog getX 
      Left            =   2385
      Top             =   6075
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox chkHTMLColoring 
      Caption         =   "Enable HTML Coloring"
      Height          =   240
      Left            =   90
      TabIndex        =   3
      Top             =   6165
      Value           =   1  'Checked
      Width           =   1995
   End
   Begin RichTextLib.RichTextBox rtTemp 
      Height          =   465
      Left            =   45
      TabIndex        =   2
      Top             =   6030
      Visible         =   0   'False
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   820
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frm.frx":000C
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   330
      Left            =   9585
      TabIndex        =   1
      Top             =   6120
      Width           =   960
   End
   Begin RichTextLib.RichTextBox txtBody 
      Height          =   5910
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   10425
      _Version        =   393217
      ScrollBars      =   3
      Appearance      =   0
      RightMargin     =   65535
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frm.frx":008E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lbls 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tags Color"
      Height          =   195
      Index           =   1
      Left            =   5490
      TabIndex        =   5
      Top             =   6165
      Width           =   765
   End
   Begin VB.Label lbls 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comments Color"
      Height          =   195
      Index           =   0
      Left            =   3285
      TabIndex        =   4
      Top             =   6165
      Width           =   1140
   End
   Begin VB.Shape shShadow 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   5910
      Index           =   1
      Left            =   90
      Top             =   90
      Width           =   10455
   End
End
Attribute VB_Name = "frm_HTML_Editor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private DisableTagColoring As Boolean
Private FirstLoad As Boolean

Private Const e_key = 128

Private Const HTML = _
"Ό΅­­    ªªªªªªªªªªªªªªªªªªªªªªªªªªªªªªªªªªªªªªªªªªªªªªª   ª       " & _
"                                      ª   ª   ΘΤΝΜ Ταη Γομοςιξη Υσιξη " & _
"ÒιγθΤεψτ Εδιτος   ª   ª                                             ª" & _
"   ª  Γςεατεδ βω Ί Γθςισ Φεηα Ϋηχαποΐνοδεμσ®γονέ ª   ªªªªªªªªªªªªªªªª" & _
"ªªªªªªªªªªªªªªªªªªªªªªªªªªªªªªª­­ΎΌθτνμΎΌθεαδΎΌστωμεΎΌ" & _
"΅­­   ®τεψτΝιξε ϋ       ζοξτ­ζνιμωΊαςιαμ»       ζοξτ­σιϊε½±²πψ» " & _
"      γομοςΊδαςλβμυε»   ύ――­­ΎΌ―στωμεΎΌ―θεαδΎΌβοδω οξμοαδ½Άχ" & _
"ιξδοχ®γμοσε¨©»ΆΎΌζοξτ γμασσ½τεψτΝιξεΎΤθαξλ Ωου ζος Δοχξμοαδιξη τ" & _
"θισ ΘΤΝΜ Ταη ΓομοςιξηΠμεασε ςατε τθισ γοδε­ βω Γθςισ Φεηα Ϋηχα" & _
"ποΐνοδεμσ®γονέΌ―ζοξτΎΌ―βοδωΎΌ―θτνμΎ"

Private Sub chkHTMLColoring_Click()
    DisableTagColoring = (chkHTMLColoring.Value = 0)

    If DisableTagColoring Then
        With txtBody
            sSt = .SelStart
            sLx = .SelLength
            ClearColors txtBody, 0
            .SelStart = sSt
            .SelLength = sLx
        End With
    Else
        txtBody_Change
    End If
End Sub

Private Sub cmdExit_Click()
    MsgBox "Copyright 2001 by Chris Vega [gwapo@models.com]", vbInformation
    End
End Sub

Private Sub Form_Load()
    HTML_Color = picTagsColor.BackColor
    Comment_Color = picCommentsColor.BackColor
    
    txtBody = etxt(HTML)
    
    DisableTagColoring = False
    FirstLoad = False
End Sub

Private Sub picCommentsColor_Click()
    getX.color = picCommentsColor.BackColor
    getX.Flags = cdlCCRGBInit Or cdlCCFullOpen
    getX.ShowColor
    
    picCommentsColor.BackColor = getX.color
    
    Comment_Color = getX.color

    FirstLoad = True
    txtBody_Change
    FirstLoad = False
End Sub

Private Sub picTagsColor_Click()
    getX.color = picTagsColor.BackColor
    getX.Flags = cdlCCRGBInit Or cdlCCFullOpen
    getX.ShowColor
    
    picTagsColor.BackColor = getX.color
    HTML_Color = getX.color
    
    FirstLoad = True
    txtBody_Change
    FirstLoad = False
End Sub

Private Sub txtBody_Change()
    If DisableTagColoring Then Exit Sub
        
    With txtBody
        SelCursor = .SelStart
        SelLength = .SelLength
            
        rtTemp = txtBody
            
            If FirstLoad Then _
                ColorTags rtTemp, 0 Else _
                ColorTags rtTemp, .SelStart
        
        On Error Resume Next
        txtBody = rtTemp
        txtBody.SetFocus
        
        .SelStart = SelCursor
        .SelLength = SelLength
    End With
End Sub

Public Function etxt(strx As String)
    etxt = ""
    For i = 1 To Len(strx)
        etxt = etxt & Chr(Asc(Mid(strx, i, 1)) Xor e_key)
    Next
End Function
