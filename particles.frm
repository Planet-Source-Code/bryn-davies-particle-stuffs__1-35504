VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Particles - By Bryn Davies"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10620
   LinkTopic       =   "Form1"
   ScaleHeight     =   9180
   ScaleWidth      =   10620
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chktrail 
      Caption         =   "Trails"
      Height          =   495
      Left            =   13560
      TabIndex        =   18
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H000000FF&
      Caption         =   "EXIT!"
      Height          =   495
      Left            =   13920
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   10200
      Width           =   1215
   End
   Begin VB.PictureBox piccol 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   13560
      ScaleHeight     =   113
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   15
      Top             =   3720
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.OptionButton optuser 
      Caption         =   "User Defined"
      Height          =   495
      Left            =   12720
      TabIndex        =   14
      Top             =   7440
      Width           =   1455
   End
   Begin VB.HScrollBar scrstrength 
      Height          =   255
      Left            =   11640
      Max             =   2000
      TabIndex        =   12
      Top             =   2760
      Value           =   300
      Width           =   2175
   End
   Begin VB.OptionButton optpuls 
      Caption         =   "Pulsating"
      Height          =   495
      Left            =   12720
      TabIndex        =   10
      Top             =   7080
      Width           =   1215
   End
   Begin VB.OptionButton Optfizzy 
      Caption         =   "Fizzy"
      Height          =   495
      Left            =   12720
      TabIndex        =   9
      Top             =   6720
      Width           =   1215
   End
   Begin VB.OptionButton optheat 
      Caption         =   "Heat"
      Height          =   495
      Left            =   11400
      TabIndex        =   7
      Top             =   7440
      Width           =   1215
   End
   Begin VB.OptionButton optgrey 
      Caption         =   "Greys"
      Height          =   495
      Left            =   11400
      TabIndex        =   6
      Top             =   7080
      Width           =   1215
   End
   Begin VB.OptionButton optcol 
      Caption         =   "Colourful"
      Height          =   495
      Left            =   11400
      TabIndex        =   5
      Top             =   6720
      Width           =   1215
   End
   Begin VB.HScrollBar scrgrav 
      Height          =   255
      Left            =   11640
      Max             =   600
      TabIndex        =   2
      Top             =   1920
      Value           =   125
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Controls:"
      Height          =   7695
      Left            =   10680
      TabIndex        =   3
      Top             =   720
      Width           =   4455
      Begin VB.Frame Frame2 
         Caption         =   "Colour Scheme:"
         Height          =   1815
         Left            =   480
         TabIndex        =   8
         Top             =   5640
         Width           =   3735
      End
      Begin VB.Label lblusercol 
         Caption         =   "Hold your left mouse button over this colour selector to select a colour:"
         Height          =   1815
         Left            =   840
         TabIndex        =   16
         Top             =   3000
         Width           =   2055
      End
      Begin VB.Label lblstrength 
         Height          =   375
         Left            =   960
         TabIndex        =   13
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label lblgrav 
         Height          =   255
         Left            =   960
         TabIndex        =   4
         Top             =   840
         Width           =   1335
      End
   End
   Begin VB.Timer tmrloop 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdrun 
      Caption         =   "Create!"
      Height          =   495
      Left            =   12360
      TabIndex        =   1
      Top             =   9000
      Width           =   1335
   End
   Begin VB.PictureBox picarea 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   8895
      Left            =   360
      ScaleHeight     =   591
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   663
      TabIndex        =   0
      Top             =   360
      Width           =   9975
   End
   Begin VB.Label lblinfo 
      Caption         =   $"particles.frx":0000
      Height          =   1575
      Left            =   480
      TabIndex        =   11
      Top             =   9360
      Width           =   9855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'With my last submissions I made the effort to comment EVERY LINE, but I had no
'feedback about whether it was entirely neccessary (<-- spelling?)
'So here I will comment sparingly, if you need to know how to do something just ask:D


'Bryn Davies

Dim w As Integer, h As Integer  'screen dimensions
Dim a(2000) As particle            'each particle
Dim mouse_x As Integer, mouse_y As Integer

'Colour variables ####
Dim pulse As Integer
Dim pulseb As Boolean
Dim usercolb As Boolean
Dim usercol As Double
Dim col_opt As Integer
' ####

' type containing information for each particle
Private Type particle
X As Double
Y As Double
velocity As Double
speed As Double
xoff As Double
colour As Integer
heat As Double
End Type
' ####

'environment quantities
Dim strength As Double          ' particle energy
Dim gravity As Double
' ####

Dim chki As Integer
Dim trails As Boolean
Const num = 1000 'number of particles
'*NB   I have put auto redraw on to stop flickering but this will slow down the draw
'with larger values

'I suppose you couls use setpixel to allow much larger amounts to be used. I
'used pset for simplicity.


Private Sub init()  'Initialize the program
Dim i As Integer

w = picarea.ScaleWidth
h = picarea.ScaleHeight

piccol.Picture = LoadPicture(App.Path & "\colourchart.JPG")

mouse_x = w / 2
mouse_y = h / 2

gravity = scrgrav.Value / 100
lblgrav.Caption = "Gravity = " & gravity & "."
   
strength = scrstrength.Value / 100
lblstrength = "Strength = " & strength & "."
optcol.Value = True
setprt
               

End Sub

Private Sub mve()       'Move all particles
Dim i As Integer

Randomize
    For i = 1 To num
    
        a(i).speed = (a(i).speed + a(i).velocity) - gravity
        a(i).Y = (a(i).Y - a(i).speed)
        a(i).X = a(i).X + a(i).xoff
        a(i).heat = a(i).heat + 3
            If a(i).Y > h Or a(i).Y < 0 Then
            
               a(i).Y = mouse_y
               a(i).X = mouse_x
               a(i).xoff = Rnd * -strength + Rnd * strength
               a(i).speed = Rnd * 15
               a(i).heat = 0
            End If

            If a(i).X < 0 Or a(i).X > w Then
            
                a(i).Y = mouse_y
                a(i).X = mouse_x
                a(i).xoff = Rnd * -strength + Rnd * strength
                a(i).speed = Rnd * 15
                a(i).heat = 0
                
            End If

    Next i
End Sub

Private Sub setprt()        'Create initial settings
Dim i As Integer

    For i = 1 To num
        a(i).Y = mouse_y
        a(i).X = mouse_x
        a(i).xoff = Rnd * -strength + Rnd * strength
        a(i).velocity = 0.5
        a(i).speed = Rnd * 15
        a(i).colour = Rnd * 255
    Next i

End Sub
Private Sub penit()                 'Draw each particle to screen
Dim i As Integer
If pulse > 255 Then pulseb = False
    If pulse < 1 Then pulseb = True
    If pulseb = True Then
    pulse = pulse + 5
    Else
    pulse = pulse - 5
    End If
    
    For i = 1 To num
    
    Select Case col_opt
    
    Case 1 'colourful
    picarea.PSet (a(i).X, a(i).Y), RGB(Rnd * a(i).colour, Rnd * a(i).colour, Rnd * a(i).colour)
    
    Case 2 'greys
    picarea.PSet (a(i).X, a(i).Y), RGB(a(i).colour, a(i).colour, a(i).colour)
    
    Case 3 'heat
    If a(i).heat > 255 Then a(i).heat = 255
    picarea.PSet (a(i).X, a(i).Y), RGB(255 - a(i).heat, 255 - a(i).heat, 0)
    
    Case 4 'fizzy
    picarea.PSet (a(i).X, a(i).Y), RGB(Rnd * a(i).colour, Rnd * a(i).colour, 0)
    
    Case 5 ' pulse
    
    picarea.PSet (a(i).X, a(i).Y), RGB(0, 0, pulse)
    
    Case 6 ' pulse
    picarea.PSet (a(i).X, a(i).Y), usercol
        
    End Select
    
    Next i
    
End Sub

Private Sub main_loop()     ' do this every milisecond
     
mve
If trails = False Then picarea.Cls
penit
        
End Sub

Private Sub chktrail_Click()
chki = chki + 1
If chki > 1 Then chki = 0
chktrail.Value = chki

trails = chki
End Sub

Private Sub cmdexit_Click() 'close program
MsgBox lblinfo.Caption, vbOKOnly
End
End Sub

Private Sub cmdrun_Click()          'begin

tmrloop.Enabled = True
End Sub

Private Sub Form_Activate()         'initialize on startup
init
End Sub


'####################### OPTION SETTINGS #####################
Private Sub optcol_Click()
col_opt = 1
optgrey.Value = False
optheat.Value = False
Optfizzy.Value = False
optpuls.Value = False
piccol.Visible = False
usercolb = False
lblusercol.Visible = False
End Sub

Private Sub Optfizzy_Click()
col_opt = 4
optcol.Value = False
optheat.Value = False
optgrey.Value = False
optpuls.Value = False
piccol.Visible = False
usercolb = False
lblusercol.Visible = False
End Sub

Private Sub optgrey_Click()
col_opt = 2
optcol.Value = False
optheat.Value = False
Optfizzy.Value = False
optpuls.Value = False
piccol.Visible = False
usercolb = False
lblusercol.Visible = False
End Sub

Private Sub optheat_Click()
col_opt = 3
optgrey.Value = False
optcol.Value = False
Optfizzy.Value = False
optpuls.Value = False
piccol.Visible = False
usercolb = False
lblusercol.Visible = False
End Sub

Private Sub optpuls_Click()
col_opt = 5
optcol.Value = False
optheat.Value = False
optgrey.Value = False
Optfizzy.Value = False
piccol.Visible = False
usercolb = False
lblusercol.Visible = False
End Sub

Private Sub optuser_Click()
col_opt = 6
optcol.Value = False
optheat.Value = False
optgrey.Value = False
Optfizzy.Value = False
optpuls.Value = False

piccol.Visible = True

lblusercol.Visible = True
usercolb = True
End Sub

'############################################################################


Private Sub picarea_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 Then  'move particle source
    mouse_x = X
    mouse_y = Y
End If
End Sub



Private Sub piccol_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then  'Choose colour
usercol = piccol.Point(X, Y)
piccol.Cls
piccol.Circle (X, Y), 5, vbRed      'Just a marker
End If
End Sub

Private Sub scrgrav_Scroll()    'change gravity
gravity = scrgrav.Value / 100
lblgrav.Caption = "Gravity = " & gravity & "."
End Sub

Private Sub scrstrength_Scroll()        'change strength
strength = scrstrength.Value / 100
lblstrength = "Strength = " & strength & "."
End Sub

Private Sub tmrloop_Timer()
main_loop
End Sub

