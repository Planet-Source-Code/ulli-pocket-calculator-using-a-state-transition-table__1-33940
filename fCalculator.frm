VERSION 5.00
Object = "{A034B639-50EC-11D4-B07A-FBBD7E43DB02}#10.0#0"; "Gradient.ocx"
Begin VB.Form fCalculator 
   BackColor       =   &H00800000&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Pocket Calculator"
   ClientHeight    =   6495
   ClientLeft      =   1140
   ClientTop       =   1935
   ClientWidth     =   4125
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   4125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.Timer tmr 
      Interval        =   500
      Left            =   3465
      Top             =   1200
   End
   Begin VB.CommandButton btButton 
      BackColor       =   &H0080C0FF&
      Caption         =   "•"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   22
      Left            =   1335
      Style           =   1  'Grafisch
      TabIndex        =   34
      TabStop         =   0   'False
      Tag             =   "."
      Top             =   5640
      Width           =   645
   End
   Begin VB.CommandButton btButton 
      BackColor       =   &H00FFFF80&
      Caption         =   "±"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   21
      Left            =   1335
      Style           =   1  'Grafisch
      TabIndex        =   33
      TabStop         =   0   'False
      Tag             =   "P"
      ToolTipText     =   "Toggle sign"
      Top             =   2385
      Width           =   645
   End
   Begin VB.CommandButton btButton 
      BackColor       =   &H00FF8080&
      Caption         =   "^"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   20
      Left            =   2145
      Style           =   1  'Grafisch
      TabIndex        =   31
      TabStop         =   0   'False
      Tag             =   "^"
      ToolTipText     =   "Exponentiation"
      Top             =   2385
      Width           =   645
   End
   Begin VB.CommandButton btButton 
      BackColor       =   &H000080FF&
      Caption         =   "Off"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   19
      Left            =   2940
      Style           =   1  'Grafisch
      TabIndex        =   30
      TabStop         =   0   'False
      Tag             =   "OFF"
      ToolTipText     =   "Quit"
      Top             =   2385
      Width           =   660
   End
   Begin VB.CommandButton btButton 
      BackColor       =   &H00FFFFFF&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   18
      Left            =   2145
      Style           =   1  'Grafisch
      TabIndex        =   29
      TabStop         =   0   'False
      Tag             =   "="
      ToolTipText     =   "Calculate"
      Top             =   5640
      Width           =   1440
   End
   Begin VB.CommandButton btButton 
      BackColor       =   &H0080FFFF&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   17
      Left            =   525
      Style           =   1  'Grafisch
      TabIndex        =   28
      TabStop         =   0   'False
      Tag             =   "C"
      ToolTipText     =   "Clear"
      Top             =   2385
      Width           =   645
   End
   Begin VB.CommandButton btButton 
      BackColor       =   &H00FFFF80&
      Caption         =   "Ex"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   16
      Left            =   1335
      Style           =   1  'Grafisch
      TabIndex        =   27
      TabStop         =   0   'False
      Tag             =   "X"
      ToolTipText     =   "Exchange"
      Top             =   3045
      Width           =   645
   End
   Begin VB.CommandButton btButton 
      BackColor       =   &H00FFFF80&
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   15
      Left            =   525
      Style           =   1  'Grafisch
      TabIndex        =   26
      TabStop         =   0   'False
      Tag             =   "M"
      ToolTipText     =   "Memory"
      Top             =   3045
      Width           =   645
   End
   Begin VB.CommandButton btButton 
      BackColor       =   &H00FF8080&
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   14
      Left            =   2145
      Style           =   1  'Grafisch
      TabIndex        =   25
      TabStop         =   0   'False
      Tag             =   "R"
      ToolTipText     =   "Root"
      Top             =   3045
      Width           =   645
   End
   Begin VB.CommandButton btButton 
      BackColor       =   &H00FF8080&
      Caption         =   "÷"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   13
      Left            =   2940
      Style           =   1  'Grafisch
      TabIndex        =   24
      TabStop         =   0   'False
      Tag             =   "/"
      ToolTipText     =   "Divide"
      Top             =   3045
      Width           =   645
   End
   Begin VB.CommandButton btButton 
      BackColor       =   &H00FF8080&
      Caption         =   "×"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   12
      Left            =   2940
      Style           =   1  'Grafisch
      TabIndex        =   23
      TabStop         =   0   'False
      Tag             =   "*"
      ToolTipText     =   "Multiply"
      Top             =   3705
      Width           =   645
   End
   Begin VB.CommandButton btButton 
      BackColor       =   &H00FF8080&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   11
      Left            =   2940
      Style           =   1  'Grafisch
      TabIndex        =   22
      TabStop         =   0   'False
      Tag             =   "-"
      ToolTipText     =   "Subtract"
      Top             =   4350
      Width           =   645
   End
   Begin VB.CommandButton btButton 
      BackColor       =   &H00FF8080&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   10
      Left            =   2940
      Style           =   1  'Grafisch
      TabIndex        =   21
      TabStop         =   0   'False
      Tag             =   "+"
      ToolTipText     =   "Add"
      Top             =   4995
      Width           =   645
   End
   Begin VB.CommandButton btButton 
      BackColor       =   &H0080C0FF&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   2145
      Style           =   1  'Grafisch
      TabIndex        =   20
      TabStop         =   0   'False
      Tag             =   "9"
      Top             =   3705
      Width           =   645
   End
   Begin VB.CommandButton btButton 
      BackColor       =   &H0080C0FF&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   1335
      Style           =   1  'Grafisch
      TabIndex        =   19
      TabStop         =   0   'False
      Tag             =   "8"
      Top             =   3705
      Width           =   645
   End
   Begin VB.CommandButton btButton 
      BackColor       =   &H0080C0FF&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   525
      Style           =   1  'Grafisch
      TabIndex        =   18
      TabStop         =   0   'False
      Tag             =   "7"
      Top             =   3705
      Width           =   645
   End
   Begin VB.CommandButton btButton 
      BackColor       =   &H0080C0FF&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   2145
      Style           =   1  'Grafisch
      TabIndex        =   17
      TabStop         =   0   'False
      Tag             =   "6"
      Top             =   4350
      Width           =   645
   End
   Begin VB.CommandButton btButton 
      BackColor       =   &H0080C0FF&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   1335
      Style           =   1  'Grafisch
      TabIndex        =   16
      TabStop         =   0   'False
      Tag             =   "5"
      Top             =   4350
      Width           =   645
   End
   Begin VB.CommandButton btButton 
      BackColor       =   &H0080C0FF&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   525
      Style           =   1  'Grafisch
      TabIndex        =   15
      TabStop         =   0   'False
      Tag             =   "4"
      Top             =   4350
      Width           =   645
   End
   Begin VB.CommandButton btButton 
      BackColor       =   &H0080C0FF&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   2145
      Style           =   1  'Grafisch
      TabIndex        =   14
      TabStop         =   0   'False
      Tag             =   "3"
      Top             =   4995
      Width           =   645
   End
   Begin VB.CommandButton btButton 
      BackColor       =   &H0080C0FF&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   1335
      Style           =   1  'Grafisch
      TabIndex        =   13
      TabStop         =   0   'False
      Tag             =   "2"
      Top             =   4995
      Width           =   645
   End
   Begin VB.CommandButton btButton 
      BackColor       =   &H0080C0FF&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   525
      Style           =   1  'Grafisch
      TabIndex        =   12
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   4995
      Width           =   645
   End
   Begin VB.CommandButton btButton 
      BackColor       =   &H0080C0FF&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   525
      Style           =   1  'Grafisch
      TabIndex        =   11
      TabStop         =   0   'False
      Tag             =   "0"
      Top             =   5640
      Width           =   645
   End
   Begin VB.PictureBox pic 
      BackColor       =   &H00000000&
      Height          =   1890
      Left            =   187
      ScaleHeight     =   1830
      ScaleWidth      =   3690
      TabIndex        =   0
      Top             =   240
      Width           =   3750
      Begin VB.Shape shp 
         BorderColor     =   &H0000FFFF&
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Ausgefüllt
         Height          =   75
         Index           =   3
         Left            =   3570
         Shape           =   3  'Kreis
         Top             =   1545
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Shape shp 
         BorderColor     =   &H0000FF00&
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Ausgefüllt
         Height          =   75
         Index           =   1
         Left            =   3570
         Shape           =   3  'Kreis
         Top             =   615
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Shape shp 
         BorderColor     =   &H0000FF00&
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Ausgefüllt
         Height          =   75
         Index           =   0
         Left            =   3570
         Shape           =   3  'Kreis
         Top             =   225
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lbMem 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Left            =   45
         TabIndex        =   35
         Top             =   1425
         Width           =   225
      End
      Begin VB.Label lblEqual 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   300
         Left            =   30
         TabIndex        =   10
         Top             =   885
         Width           =   225
      End
      Begin VB.Label lblOper 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   300
         Left            =   30
         TabIndex        =   9
         ToolTipText     =   "Operation"
         Top             =   495
         Width           =   225
      End
      Begin VB.Label lblVZ 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Index           =   3
         Left            =   3435
         TabIndex        =   8
         ToolTipText     =   "Sign"
         Top             =   1425
         Width           =   105
      End
      Begin VB.Label lblVZ 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   300
         Index           =   2
         Left            =   3435
         TabIndex        =   7
         ToolTipText     =   "Sign"
         Top             =   885
         Width           =   105
      End
      Begin VB.Label lblVZ 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   300
         Index           =   1
         Left            =   3435
         TabIndex        =   6
         ToolTipText     =   "Sign"
         Top             =   495
         Width           =   105
      End
      Begin VB.Label lblVZ 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   300
         Index           =   0
         Left            =   3435
         TabIndex        =   5
         ToolTipText     =   "Sign"
         Top             =   105
         Width           =   105
      End
      Begin VB.Label lblReg 
         Alignment       =   1  'Rechts
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Index           =   3
         Left            =   285
         TabIndex        =   4
         ToolTipText     =   "Memory"
         Top             =   1425
         Width           =   3120
      End
      Begin VB.Label lblReg 
         Alignment       =   1  'Rechts
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   300
         Index           =   2
         Left            =   285
         TabIndex        =   3
         ToolTipText     =   "Result"
         Top             =   885
         Width           =   3120
      End
      Begin VB.Label lblReg 
         Alignment       =   1  'Rechts
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   300
         Index           =   1
         Left            =   285
         TabIndex        =   2
         ToolTipText     =   "Second Operand"
         Top             =   495
         Width           =   3120
      End
      Begin VB.Label lblReg 
         Alignment       =   1  'Rechts
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   300
         Index           =   0
         Left            =   285
         TabIndex        =   1
         ToolTipText     =   "First Operand"
         Top             =   105
         Width           =   3120
      End
   End
   Begin GradientOCX.Gradient graRainbow 
      Left            =   540
      Top             =   6180
      _ExtentX        =   529
      _ExtentY        =   503
      FromColor       =   255
      ToColor         =   12583104
      ColorSequence   =   1
   End
   Begin VB.Label lblCpyRite 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "(c) UMG EDV"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   180
      Left            =   2655
      TabIndex        =   32
      ToolTipText     =   "UMGEDV@AOL.COM"
      Top             =   6255
      Width           =   870
   End
End
Attribute VB_Name = "fCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This pocket calculator is meant as an example for the State Transition Table
'technique.
'
'A few words about this technique:
'
'State Transition Tables are in some way an extension of Decision Tables.
'Whereas DTs store decision information that is static, ie only depends
'on input, STTs store decision information that is dynamic (well, not
'the stored information is dynamic, but the access to it is).
'
'STTs should be used when the rection of a device (program) not only depends
'on input (trigger) but also on circumstances (state). An example is the
'Reset key of a pocket calculator - when pressed initially it clears the whole
'calculator, however when pressed after an input it only clears that last input.

'Another example is human behaviour: You hear a good joke (input) for the
'first time and you burst out with laughter (reaction), however when you hear
'the same joke (identical input) a second time - ... (different reaction).
'
'To recognize an input it may be necessary to classify it to be able to select the
'proper reaction; sticking to the above (simplified) example:
'
'The story (input) is classified as funny - ah, it's a joke
'                                         - hear it for the 1st time --> laugh;
'                                                                        change state to knowing.
'                                         - knowing it already --> smile;
'                                                                  interrupt.
'                                   sad   - oops, this is serious
'                                         - hear it for the 1st time --> look serious;
'                                                                        express sympathy;
'                                                                        change state to knowing.
'                                         - knowing it already --> look serious;
'                                                                  nod your head.
'
'STT's are THE choice for game programming for example; a character whose behaviour
'is controlled by an STT can be made to react in a much more differentiated and
'humanly way than if he was controlled by simple IF's and ELSE's because his state(!)
'of mind now also influences his behaviour.
'
'The complete behaviour pattern is stored in a table from which the Reaction
'to the current input is selected as well as the next state. The advantages of
'storing this information in a table are:
'
'  1 Completeness          since every STT-cell MUST contain an entry there are no
'                          unconsidered cases by design
'
'  2 Simplicity            decision making and the repertoire of reactions are
'                          completely separated from one another
'
'  3 Ease of Modification  to change the behaviour pattern you simply alter the
'                          contents of the STT and possibly add a new Reaction
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'The State Transition Table
Private STT(1 To 9)       As Variant

'The current program State - used to address the STT
Private State             As Long

'The current Trigger - used to address the STT (derived from Key pressed)
Private Trigger           As Long

Private Type Reg
    Value                 As String
    DecimalPoint          As String
    Sign                  As String
End Type

'0 - Reg A - First Operand
'1 - Reg B - Second Operand
'2 - Reg C - Result
'3 - Reg M - Memory Register
'4 - Temp  - used for swapping
Private Regs(0 To 4)      As Reg

'Index for the Registers 0 thru 3
Private RegPtr            As Long

'Index for Active Register Indicator
Private ActPtr            As Long

'the operator + - * / etc
Private Operator          As String

'the equal sign
Private Equal             As String

'for various purposes
Private i                 As Long

'the key, orig and translated
Private OriginalKey       As String
Private TranslatedKey     As String
Private IgnoreKey         As Boolean

'translation strings
Private Const PermittedKeys    As String = "0123456789,.+-*/HE^R[=CMXP" & vbCr
Private Const TranslatedKeys   As String = "0123456789..+-×÷^^^[[=CMXP="

Private Const TriggerClasses   As String = "000000000000111111111234562"
'             are grouped as follows        0 = Numeric
'                                           1 = Operator
'                                           2 = Equals
'                                           3 = Clear
'                                           4 = Memory
'                                           5 = Exchange
'                                           6 = Toggle Sign

Private Sub Arith(OpA As Reg, OpB As Reg, Rslt As Reg, Operator As String)

  'Arithmetic routine

  Dim Operand1 As Double
  Dim Operand2 As Double
  Dim Result      As Double

    Operand1 = Val(OpA.Sign & OpA.Value)
    Operand2 = Val(OpB.Sign & OpB.Value)
    On Error Resume Next
      Select Case Operator
        Case "+"
          Result = Operand1 + Operand2
        Case "-"
          Result = Operand1 - Operand2
        Case "×"
          Result = Operand1 * Operand2
        Case "÷"
          Result = Operand1 / Operand2
        Case "^"
          Result = Operand1 ^ Operand2
        Case "["
          If Operand1 = Int(Operand1) And Operand1 And 1 Then 'odd integer root exponent
              Result = Abs(Operand2) ^ (1 / Operand1) * Sgn(Operand2)
            Else 'NOT OPERAND1...
              Result = Operand2 ^ (1 / Operand1)
          End If
      End Select
      If Err Then
          Rslt.Value = Err.Description
          Rslt.Sign = ""
        Else 'ERR = FALSE
          If Result < 0 Then
              Rslt.Sign = "-"
            Else 'NOT RESULT...
              Rslt.Sign = ""
          End If
          Rslt.Value = Abs(Result)
          Rslt.Value = Replace(Rslt.Value, ",", ".")
      End If
    On Error GoTo 0

End Sub

Private Sub btButton_Click(Index As Integer)

    If btButton(Index).Tag = "OFF" Then
        Unload Me
      Else 'NOT BTBUTTON(INDEX).TAG...
        'simulate Key_Press - the key is in the button's Tag
        pic.SetFocus
        DoEvents
        Form_KeyPress Asc(btButton(Index).Tag)
    End If

End Sub

Private Sub ClrBC()

  'Clear Regs B and C, Operator and Equal Sign

    For i = 1 To 2
        Regs(i).Value = ""
        Regs(i).DecimalPoint = ""
        Regs(i).Sign = ""
    Next i
    Operator = ""
    Equal = ""

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If Shift And vbCtrlMask And KeyCode = vbKeyC Then
        Clipboard.Clear
        Clipboard.SetText Regs(2).Sign & Regs(2).Value
        IgnoreKey = True
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If IgnoreKey Then
        IgnoreKey = False
      Else 'IGNOREKEY = FALSE
        OriginalKey = UCase$(Chr$(KeyAscii))
        i = InStr(PermittedKeys, OriginalKey) 'is it among the permitted triggers?
        If i = 0 Then 'not a legal key
            Beep
          Else 'NOT I...
            TranslatedKey = Mid$(TranslatedKeys, i, 1) 'for display

            '------------------------------------------------------------------------------------
            'this is where it is happening

            Trigger = Val(Mid$(TriggerClasses, i, 1)) 'classification
            'Debug.Print "State "; State; " Trigger "; Trigger; " Reaction "; STT(State)(Trigger) \ 100
            Select Case STT(State)(Trigger) \ 100  'Reaction Code - first three digits in STT cell
              Case 100
                Beep 'Wrong trigger for the situation
              Case 101
                Reaction01 'Clear all Regs exept M
                Reaction20 'Set current Reg = A
              Case 102
                Reaction02 'Clear current Reg
              Case 103
                Reaction03 'Insert Key into current Reg
              Case 104
                Reaction02 'Clear current Reg
                Reaction03 'Insert Key into current Reg
              Case 105
                Reaction05 'Save Operator; Set current Reg = B
              Case 106
                Reaction06 'Reg A [Operator] Reg B -> Reg C
              Case 107
                Reaction07 'Reg C -> Reg A; Clear Regs B and C
                Reaction05 'Save Operator; Set current Reg = B
              Case 108
                Reaction06 'Reg A [Operator] Reg B -> Reg C
                Reaction07 'Reg C -> Reg A; Clear Regs B and C
                Reaction05 'Save Operator; Set current Reg = B
              Case 109
                Reaction09 'Clear Operator
              Case 110
                Reaction10 'Set current Reg = M
              Case 111
                Reaction11 'Reg M -> Reg A; Clear Regs B and C
              Case 112
                Reaction12 'Reg M [Operator] Reg C -> Reg M
              Case 113
                Reaction13 'Clear Reg M
              Case 114
                Reaction14 'Reg M <--> Reg A; Clear Regs B and C
              Case 115
                Reaction15 'Toggle sign of current Reg
              Case 116
                Reaction16 'Reg A -> Reg B
                Reaction06 'Reg A [Operator] Reg B -> Reg C
              Case 117
                Reaction17 'Reg A <--> Reg B
              Case 118
                Reaction18 'Reg M -> Reg B
              Case 119
                Reaction01 'Clear all Regs exept M
                Reaction20 'Set current Reg = A
                Reaction03 'Insert Key into current Reg
              Case 120
                Reaction20 'Set current Reg = A
            End Select

            State = STT(State)(Trigger) Mod 100     'Next State - last two digits in STT cell
            'Debug.Print "New State "; State
            '------------------------------------------------------------------------------------

            'update display
            For i = 0 To 3
                lblReg(i) = Regs(i).Value
                lblVZ(i) = Regs(i).Sign
            Next i
            lblOper = Operator
            lblEqual = Equal
            lbMem = Left$("M", Len(Regs(3).Value)) 'show M(emory) if there's something in RegM
            shp(0).Visible = False
            shp(1).Visible = False
            shp(3).Visible = False
        End If
    End If

End Sub

Private Sub Form_Load()

  'Setup the State Transition Table

  'The first three digits represent the ReactionCode and the
  'last two digits indicate the next state.
  'The Triggers are numbered from 0 to 6 across the top and the States are
  'numbered from 1 to 9 down the side. The initial State is 1.

    State = 1 'Initial State

    'We have only 9 states - so one digit would be enough for the states, but maybe the
    'State Transition Table will grow as the need arises and then we have a reserve.
    'Also it is a good idea to group the Reaction codes by the first digit (in our case
    'we only have Reaction group 1xx).

    'Triggers       num    opr     =     clr    mem    xch    +/-
    '---------------------------------------------------------------

    STT(1) = Array(10402, 10001, 10001, 10101, 11008, 10001, 10001)
    STT(2) = Array(10302, 10503, 10002, 10201, 11008, 10002, 11502)
    STT(3) = Array(10404, 10003, 11605, 10906, 11804, 10003, 10003)
    STT(4) = Array(10304, 10803, 10605, 10207, 10004, 11704, 11504)
    STT(5) = Array(11902, 10703, 10005, 10101, 11008, 10005, 10005)
    STT(6) = Array(10006, 10503, 10006, 10101, 10006, 10006, 10006)
    STT(7) = Array(10404, 10007, 11605, 10101, 10007, 10007, 10007)
    STT(8) = Array(10008, 11205, 11109, 11305, 12005, 11409, 11505)
    STT(9) = Array(11902, 10507, 10009, 10101, 11008, 10009, 10009)

    'and here are the Reactions:
    '  100 Error - Beep
    '  101 Clear all Regs exept M
    '      Set current Reg = A
    '  102 Clear current Reg
    '  103 Insert Key into current Reg
    '  104 Clear current Reg
    '      Insert Key into current Reg
    '  105 Save Operator; Set current Reg = B
    '  106 Reg A [Operator] Reg B -> Reg C
    '  107 Reg C -> Reg A; Clear Regs B and C
    '      save Operator; Set current Reg = B
    '  108 Reg A [Operator] Reg B -> Reg C
    '      Reg C -> Reg A; Clear Regs B and C
    '      Save Operator; Set current Reg = B
    '  109 Clear Operator
    '  110 Set current Reg = M
    '  111 Reg M -> Reg A; Clear Regs B and C
    '  112 Reg M [Operator] Reg C -> Reg M
    '  113 Clear Reg M
    '  114 Reg M <--> Reg A; Clear Regs B and C
    '  115 Toggle sign of current Reg
    '  116 Reg A -> Reg B
    '      Reg A [Operator] Reg B -> Reg C
    '  117 Reg A <--> Reg B
    '  118 Reg M -> Reg B
    '  119 Clear all Regs exept M
    '      Set current Reg = A
    '      Insert Key into current Reg
    '  120 Set current Reg = A'

End Sub

Private Sub Form_Paint()

    graRainbow.Paint vbRed, vbMagenta

End Sub

Private Sub Reaction01()

  'Clear all Regs exept M

    RegPtr = 0
    Reaction02
    ClrBC

End Sub

Private Sub Reaction02()

  'Clear current Reg

    Regs(RegPtr).Value = ""
    Regs(RegPtr).DecimalPoint = ""
    Regs(RegPtr).Sign = ""

End Sub

Private Sub Reaction03()

  'Insert Key into current Reg

    If (Regs(RegPtr).DecimalPoint = "." And TranslatedKey = ".") Or (Len(Regs(RegPtr).Value) > 13) Then
        Beep
      Else 'NOT (REGS(RegPtr).DECIMALPOINT...
        Regs(RegPtr).Value = Regs(RegPtr).Value & TranslatedKey
        If TranslatedKey = "." Then
            Regs(RegPtr).DecimalPoint = "." 'to prevent 2nd decimal point
        End If
    End If

End Sub

Private Sub Reaction05()

  'Save Operator; Set current Reg = B

    Operator = TranslatedKey
    If Regs(0).Value = "" Then
        Regs(0).Value = "0"
    End If
    RegPtr = 1
    ActPtr = 1

End Sub

Private Sub Reaction06()

  'Reg A [Operator] Reg B -> Reg C

    Equal = "="
    Arith Regs(0), Regs(1), Regs(2), Operator
    ActPtr = 0

End Sub

Private Sub Reaction07()

  'Reg C -> Reg A; Clear Regs B and C

    Regs(0) = Regs(2)
    ClrBC

End Sub

Private Sub Reaction09()

  'Clear Operator

    Operator = ""

End Sub

Private Sub Reaction10()

  'Set current Reg = M

    RegPtr = 3
    ActPtr = 3

End Sub

Private Sub Reaction11()

  'Reg M -> Reg A; Clear Regs B and C

    Regs(0) = Regs(3)
    ClrBC
    ActPtr = 0

End Sub

Private Sub Reaction12()

  'Reg M [Operator] Reg C -> Reg M

    Arith Regs(3), Regs(2), Regs(3), TranslatedKey
    ActPtr = 0

End Sub

Private Sub Reaction13()

  'Clear Reg M

    Regs(3).Value = ""
    Regs(3).DecimalPoint = ""
    Regs(3).Sign = ""
    ActPtr = 0

End Sub

Private Sub Reaction14()

  'Reg M <--> Reg A; Clear Regs B and C

    Regs(4) = Regs(0)
    Regs(0) = Regs(3)
    Regs(3) = Regs(4)
    ClrBC
    ActPtr = 0

End Sub

Private Sub Reaction15()

  'Toggle sign of current Reg

    If Regs(RegPtr).Value = "" Then
        Beep
      Else 'NOT REGS(RegPtr).VALUE...
        Regs(RegPtr).Sign = IIf(Regs(RegPtr).Sign = "", "-", "")
    End If
    If RegPtr = 3 Then
        ActPtr = 0
    End If

End Sub

Private Sub Reaction16()

  'Reg A -> Reg B

    Regs(1) = Regs(0)

End Sub

Private Sub Reaction17()

  'Reg A <--> Reg B

    Regs(4) = Regs(0)
    Regs(0) = Regs(1)
    Regs(1) = Regs(4)

End Sub

Private Sub Reaction18()

  'Reg M -> Reg B

    Regs(1) = Regs(3)

End Sub

Private Sub Reaction20()

  'Set current Reg = A

    RegPtr = 0
    ActPtr = 0

End Sub

Private Sub tmr_Timer()

    shp(ActPtr).Visible = Not shp(ActPtr).Visible

End Sub

':) Ulli's VB Code Formatter V2.12.7 (19.04.2002 15:07:36) 111 + 424 = 535 Lines
