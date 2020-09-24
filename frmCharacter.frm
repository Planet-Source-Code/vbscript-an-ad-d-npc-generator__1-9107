VERSION 5.00
Begin VB.Form frmCharacter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NPC Character Sheet"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7935
   Icon            =   "frmCharacter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   7935
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "Close Character Sheet"
      Default         =   -1  'True
      Height          =   975
      Left            =   3360
      Picture         =   "frmCharacter.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox txtWeap2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   5040
      Width           =   1815
   End
   Begin VB.TextBox txtWeap1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   5040
      Width           =   1815
   End
   Begin VB.TextBox txtArmor 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   5400
      Width           =   1815
   End
   Begin VB.TextBox txtSpell 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox txtBreath 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox txtRod 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox txtPetr 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox txtPara 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox txtChri 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   4560
      Width           =   1335
   End
   Begin VB.TextBox txtCons 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox txtWisd 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox txtInte 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox txtDext 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox txtStre 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox txtThac0 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox txtAC 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox txtHtP 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox txtWgt 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox txtHgt 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox txtAge 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox txtLvl 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox txtAln 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtCls 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtRce 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblSave 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Saving Throws"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   26
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label lblAbilities 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Abilities"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   7800
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label lblWeap2 
      Caption         =   "Secondary Weapon"
      Height          =   255
      Left            =   4320
      TabIndex        =   24
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label lblWeap1 
      Caption         =   "Primary Weapon"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label lblArmor 
      Caption         =   "Armor"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Label lblSpell 
      Caption         =   "Spell"
      Height          =   255
      Left            =   5400
      TabIndex        =   21
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label lblBreath 
      Caption         =   "Breath"
      Height          =   255
      Left            =   5400
      TabIndex        =   20
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label lblRods 
      Caption         =   "Rods"
      Height          =   255
      Left            =   5400
      TabIndex        =   19
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label lblPara 
      Caption         =   "Paralization"
      Height          =   255
      Left            =   5400
      TabIndex        =   18
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label lblPetr 
      Caption         =   "Petrification"
      Height          =   255
      Left            =   5400
      TabIndex        =   17
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label lblCharisma 
      Caption         =   "Charisma"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label lblConstitution 
      Caption         =   "Constitution"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label lblWisdom 
      Caption         =   "Wisdom"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label lblIntelligence 
      Caption         =   "Intelligence"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label lblDexterity 
      Caption         =   "Dexterity"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label lblStrength 
      Caption         =   "Strength"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label lblAlignment 
      Caption         =   "Alignment"
      Height          =   255
      Left            =   5400
      TabIndex        =   10
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblThac0 
      Caption         =   "THAC0"
      Height          =   255
      Left            =   5400
      TabIndex        =   9
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lblAC 
      Caption         =   "Armor Class"
      Height          =   255
      Left            =   2760
      TabIndex        =   8
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lblWeight 
      Caption         =   "Weight"
      Height          =   255
      Left            =   5400
      TabIndex        =   7
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label lblHeight 
      Caption         =   "Height"
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label lblAge 
      Caption         =   "Age"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label lblHP 
      Caption         =   "Hip Points"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lblLevel 
      Caption         =   "Level"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblRace 
      Caption         =   "Race"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblClass 
      Caption         =   "Class"
      Height          =   255
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000D&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      X1              =   120
      X2              =   7800
      Y1              =   2160
      Y2              =   2160
   End
End
Attribute VB_Name = "frmCharacter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    'All this does is close the Character Sheet form and
    'open the List.
    frmCharacter.Hide
    frmNPCGen.Show
End Sub

