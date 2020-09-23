VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   Caption         =   "Bungee Baby"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6540
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   315
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   436
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picBaby 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   360
      Picture         =   "frmMain.frx":1042
      ScaleHeight     =   145
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   2030
      TabIndex        =   0
      Top             =   2400
      Visible         =   0   'False
      Width           =   30450
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

Private Const SND_APPLICATION = &H80        '  look for application specific association
Private Const SND_ALIAS = &H10000           '  name is a WIN.INI [sounds] entry
Private Const SND_ALIAS_ID = &H110000       '  name is a WIN.INI [sounds] entry identifier
Private Const SND_ASYNC = &H1               '  play asynchronously
Private Const SND_FILENAME = &H20000        '  name is a file name
Private Const SND_LOOP = &H8                '  loop the sound until next sndPlaySound
Private Const SND_MEMORY = &H4              '  lpszSoundName points to a memory file
Private Const SND_NODEFAULT = &H2           '  silence not default, if sound not found
Private Const SND_NOSTOP = &H10             '  don't stop any currently playing sound
Private Const SND_NOWAIT = &H2000           '  don't wait if the driver is busy
Private Const SND_PURGE = &H40              '  purge non-static events for task
Private Const SND_RESOURCE = &H40004        '  name is a resource name or atom
Private Const SND_SYNC = &H0                '  play synchronously (default)

Private cBaby As clsBaby
Private mbRunning As Boolean

' Some settings
Private Const FrameTime As Long = 100       ' Time in ms between each frame
Private Const MoveTime As Long = 20         ' Time in ms between each movement

Private Sub MainLoop()
    Dim lFrameTimer As Long
    Dim lFrame As Long
    Dim lTimer As Long
    
    mbRunning = True
    lTimer = GetTickCount()
    lFrameTimer = GetTickCount()
    lFrame = 0
    
    Call PlaySound(App.Path & "\baby.wav", ByVal 0&, SND_ASYNC Or SND_LOOP Or SND_FILENAME Or SND_NODEFAULT)
    
    ' Draw
    Do While mbRunning
        If GetTickCount() - lFrameTimer >= FrameTime Then
            ' Move frame
            lFrame = lFrame + 1
            If lFrame = 14 Then lFrame = 0
            cBaby.Frame = lFrame
            lFrameTimer = GetTickCount()
        End If
        
        If GetTickCount() - lTimer >= MoveTime Then
            ' Move baby
            cBaby.MoveBaby
            
            ' Draw baby
            cBaby.Blit
            
            lTimer = GetTickCount()
        End If
        
        DoEvents
    Loop
    
    Call PlaySound("", ByVal 0&, SND_ASYNC Or SND_NODEFAULT)
End Sub

Private Sub Form_Load()
    ' Create the baby
    Set cBaby = New clsBaby
    
    ' Let's draw on the form
    cBaby.FrameWidth = 145
    cBaby.Destination = frmMain
    cBaby.Source = picBaby
    
    Me.Show
    DoEvents
    
    ' Start drawing
    Call MainLoop
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cBaby.MoveCursor CLng(X), CLng(Y)
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    mbRunning = False
End Sub


