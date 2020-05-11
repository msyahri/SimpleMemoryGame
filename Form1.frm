VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF0000&
   Caption         =   "Tes Daya Ingat - Manusa Simple Game"
   ClientHeight    =   7260
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9495
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   7260
   ScaleWidth      =   9495
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   6840
      TabIndex        =   18
      Top             =   1680
      Width           =   2040
      Begin VB.ListBox ListSkor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0045BC95&
         Height          =   1380
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   1860
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00008000&
      Caption         =   "Skor"
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   7440
      TabIndex        =   16
      Top             =   3720
      Width           =   975
      Begin VB.Label LblSkor 
         AutoSize        =   -1  'True
         BackColor       =   &H00008000&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   570
         Left            =   360
         TabIndex        =   17
         Top             =   240
         Width           =   240
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cara main"
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
      Left            =   7200
      TabIndex        =   11
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton bkeluar 
      Caption         =   "Keluar"
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
      Left            =   7200
      TabIndex        =   10
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00008000&
      Caption         =   "Main"
      ForeColor       =   &H00FFFFFF&
      Height          =   5415
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   5775
      Begin VB.CommandButton CmdStart 
         Caption         =   "MULAI"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2040
         TabIndex        =   8
         Top             =   2280
         Width           =   1575
      End
      Begin VB.CommandButton CmdPress 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   3
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   720
         Width           =   1575
      End
      Begin VB.CommandButton CmdPress 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   2
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3240
         Width           =   1575
      End
      Begin VB.CommandButton CmdPress 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   1
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3240
         Width           =   1575
      End
      Begin VB.CommandButton CmdPress 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   0
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00008000&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3960
         TabIndex        =   15
         Top             =   4560
         Width           =   225
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00008000&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1680
         TabIndex        =   14
         Top             =   4560
         Width           =   225
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00008000&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   3960
         TabIndex        =   13
         Top             =   240
         Width           =   225
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00008000&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   1680
         TabIndex        =   12
         Top             =   240
         Width           =   225
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00008000&
      Caption         =   "Skor Tertinggi"
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   6720
      TabIndex        =   2
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "RPL ~ MANUSA"
      BeginProperty Font 
         Name            =   "Futura Md BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   7920
      TabIndex        =   20
      Top             =   6960
      Width           =   1170
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "ASAH DAYA INGAT"
      BeginProperty Font 
         Name            =   "Futura Md BT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3240
      TabIndex        =   9
      Top             =   360
      Width           =   3045
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   7200
      TabIndex        =   1
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   450
      Left            =   6735
      TabIndex        =   0
      Top             =   4245
      Width           =   420
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim arr(100) As Integer
Dim cnt As Integer
Dim langkah(100) As Integer
Dim cntLangkah As Integer
Dim TopSkor As Integer, BottomSkor As Integer

'untuk var agar form bisa dipindah
Dim xx As Single, yy As Single, pindah As Boolean

Private Sub bkeluar_Click()
pesan = MsgBox("Yakin Keluar ?", vbQuestion + vbYesNo, "Question")
If pesan = vbYes Then
End
Else
Form1.SetFocus
End If
End Sub

Private Sub CmdPress_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.CmdPress(Index).BackColor = vbGreen
    Me.CmdPress(Index).Caption = cntLangkah + 1
End Sub

Private Sub CmdPress_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.CmdPress(Index).BackColor = &H8000000F  'normal color
    Me.CmdPress(Index).Caption = ""
    
    cntLangkah = cntLangkah + 1
    langkah(cntLangkah) = Index
    If cntLangkah = cnt Then 'terakhir
        StopLangkah
        'cek sama
        Dim tidakSama As Boolean
        tidakSama = False
        For i = 1 To cnt
            If arr(i) <> langkah(i) Then
                tidakSama = True
            End If
        Next i
        If tidakSama = True Then
            MsgBox "SALAH" & vbCrLf & "Skor = " & cnt - 1, vbCritical, "^_^ Game Over"
            ''-menyimpan dalam list jika masuk 5 besar
            If cnt - 1 > BottomSkor Then
                Dim nama As String
                nama = InputBox("Selamat anda masuk kedalam 5 besar," & vbCrLf & "Masukkan nama anda :", "^_^ Nama 5 Besar")
                If nama <> "" Then
                    'menyimpan dalam tabel
                    With DataEnvironment1.rstabel_skor
                        .AddNew
                        .Fields("skor") = cnt - 1
                        .Fields("nama") = nama
                        .Update
                    End With
                    IsiSkor
                End If
            End If
            '--
        Else
            Me.CmdStart.Caption = "Next..."
            Tunggu
            TambahLangkah
            Simulasi
            BacaLangkah
        End If
    End If
End Sub

Private Sub CmdStart_Click()
    If Me.CmdStart.Caption = "START" Then
        cnt = 0
        TambahLangkah
        Simulasi
        BacaLangkah
        'Me.CmdStart.Enabled = False
    End If
End Sub

Sub TambahLangkah()
    Me.LblSkor.Caption = cnt
    cnt = cnt + 1
    arr(cnt) = Int(Rnd * 4)
End Sub

Sub Simulasi()
    Me.CmdStart.Caption = "Wait..."
    For i = 1 To cnt
        Me.CmdPress(arr(i)).BackColor = vbGreen
        Tunggu
        Me.CmdPress(arr(i)).BackColor = &H8000000F  'normal color
        Tunggu
    Next i
End Sub

Sub BacaLangkah()
    cntLangkah = 0
    For i = 0 To 3
        Me.CmdPress(i).Enabled = True
    Next i
    Me.CmdStart.Caption = "Ready"
End Sub
Sub StopLangkah()
    For i = 0 To 3
        Me.CmdPress(i).Enabled = False
    Next i
    Me.CmdStart.Caption = "START"
    Me.CmdStart.Enabled = True
End Sub

Private Sub Command2_Click()

MsgBox "Klik tombol mulai untuk memulai permainan. Perhatikan urutan tombol yang menyala dengan seksama. Lalu klik tombol - tombol (1-4) sesuai urutan", vbInformation, "Cara Bermain"

End Sub

Private Sub Form_Activate()
    Me.CmdStart.SetFocus
End Sub

Private Sub Form_Load()
    Randomize Timer
    StopLangkah
    IsiSkor
End Sub

Sub Tunggu()
    '--buat komputer menghitung untuk waktu tunggu
    '  sebelum menjalankan perintah berikutnya
    For t = 1 To 200000
        DoEvents
    Next t
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        pindah = True
        xx = X
        yy = Y
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If pindah = True Then
        Me.Left = Me.Left - (xx - X)
        Me.Top = Me.Top - (yy - Y)
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    pindah = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub Label1_Click()
    End
End Sub

Private Sub Label2_Click()
    MsgBox "Untuk memainkan permainan ini klik tombol START " & vbCrLf _
        & "Ikuti langkah yang ditunjukkan dengan mengklik tombol yang dtandai warna merah secara berurutan" & vbCrLf _
        & "Sampai sejauh manakah daya ingat anda untuk mengingat langkah yang diberikan komputer??", vbInformation, "^_^ Testing Daya Ingat"
End Sub

Sub IsiSkor()
    Me.ListSkor.Clear
    
    With DataEnvironment1.rstabel_skor
        If .State = 0 Then .Open
        If .RecordCount > 0 Then
            'mengurutkan yang tertinggi
            .Sort = "skor desc"
            .MoveFirst
            TopSkor = .Fields("skor")
            'mengisi list
            For i = 1 To .RecordCount
                Me.ListSkor.AddItem .Fields("skor") & vbTab & .Fields("nama")
                
                'mengisi bottom skor di lima besar
                If i = 5 Then
                    BottomSkor = .Fields("skor")
                    Exit For
                End If
                .MoveNext
            Next i
            
            .MoveLast
            
        End If
    End With
End Sub


