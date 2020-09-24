VERSION 5.00
Begin VB.Form frmNPCGen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NPC Character Generator"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   Icon            =   "NPCGen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   6615
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Quit Program"
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show Selected"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1920
      Width           =   1455
   End
   Begin VB.ListBox lstCharacters 
      Height          =   1620
      ItemData        =   "NPCGen.frx":030A
      Left            =   120
      List            =   "NPCGen.frx":030C
      TabIndex        =   0
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "frmNPCGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    'Declare the variables used in subroutine
    Dim chrFile As Variant
    Dim InputStr As String
    Dim sText() As String
    
    'Here we get the path of the data file.
    chrFile = App.Path & "\npcs.dat"
    'Here we open the file.
    Open chrFile For Input As #1
    
    'This loop reads in, line by line, the file and puts
    'each line of date into a list.
    Do While Not EOF(1) 'Tests for the End Of File.
        Line Input #1, InputStr 'Reads the line of data.
        lstCharacters.AddItem (InputStr) 'Adds the data to the list.
    Loop
    Close 1 'Close the file.
End Sub

Private Sub cmdGenerate_Click()
    'Declare the variables.
    Dim Cls, Clas, Lvl, HtP, Rce, Age, Hgt, Wgt As String
    Dim Aln, Stre, Inte, Wisd, Cons, Dext, Chri, Weap1 As String
    Dim Weap2, Armor, Para, Petr, Rod, Breath, Spell, Thac0, AC As String
    Dim ClsNum, RaceNum, AlnNum As Integer
    Dim Character, InputStr, sText() As String
    Dim chrFile As Variant
    
    Randomize 'Initialize the random number generator.
    Lvl = Int(Rnd * 5) + 1  'Generates the level of the character.
    HtP = ((Int(Rnd * 6) + 1) + (Int(Rnd * 6) + 1) + (Int(Rnd * 6) + 1)) * Lvl 'Generates the Hit Points (3d6)
    Stre = (Int(Rnd * 6) + 1) + (Int(Rnd * 6) + 1) + (Int(Rnd * 6) + 1) 'Generates the Strength (3d6)
    Inte = (Int(Rnd * 6) + 1) + (Int(Rnd * 6) + 1) + (Int(Rnd * 6) + 1) 'Generates the Intelligence (3d6)
    Wisd = (Int(Rnd * 6) + 1) + (Int(Rnd * 6) + 1) + (Int(Rnd * 6) + 1) 'Generates the Wisdom (3d6)
    Cons = (Int(Rnd * 6) + 1) + (Int(Rnd * 6) + 1) + (Int(Rnd * 6) + 1) 'Generates the Constitution (3d6)
    Dext = (Int(Rnd * 6) + 1) + (Int(Rnd * 6) + 1) + (Int(Rnd * 6) + 1) 'Generates the Dexterity (3d6)
    Chri = (Int(Rnd * 6) + 1) + (Int(Rnd * 6) + 1) + (Int(Rnd * 6) + 1) 'Generates the Charisma (3d6)
    Para = (Int(Rnd * 6) + 1) + (Int(Rnd * 6) + 1) + (Int(Rnd * 6) + 1) 'Generates the Paralyze (3d6)
    Petr = (Int(Rnd * 6) + 1) + (Int(Rnd * 6) + 1) + (Int(Rnd * 6) + 1) 'Generates the Petrification (3d6)
    Rod = (Int(Rnd * 6) + 1) + (Int(Rnd * 6) + 1) + (Int(Rnd * 6) + 1) 'Generates the Rod (3d6)
    Breath = (Int(Rnd * 6) + 1) + (Int(Rnd * 6) + 1) + (Int(Rnd * 6) + 1) 'Generates the Breath (3d6)
    Spell = (Int(Rnd * 6) + 1) + (Int(Rnd * 6) + 1) + (Int(Rnd * 6) + 1) 'Generates the Spell (3d6)
    Thac0 = Int(Rnd * 5) + 5 'Generate the To Hit AC of 0
    Weap2 = "Dagger" 'Sets the secondary weapon (for simplicity at this time)
    ClsNum = Int(Rnd * 7) + 1 'Generate the Random number for the Class
    RaceNum = Int(Rnd * 5) + 1 'Generate the Random number for the Race
    AlnNum = Int(Rnd * 9) + 1 'Generate the Random number for the Alignment
    
    'Set the Class, then armor, weapon and AC based on Class
    Select Case ClsNum
        Case 1
            Cls = "Ranger"
            Armor = "Leather"
            Weap1 = "Bow"
            AC = "8"
        Case 2
            Cls = "Thief"
            Armor = "Leather"
            Weap1 = "Shortsword"
            AC = "8"
        Case 3
            Cls = "Mage"
            Armor = "Cloth"
            Weap1 = "Quarterstaff"
            AC = "10"
        Case 4
            Cls = "Cleric"
            Armor = "Chain Mail"
            Weap1 = "Mace"
            AC = "5"
        Case 5
            Cls = "Palidin"
            Armor = "Plate Mail"
            Weap1 = "Bastard Sword"
            AC = "3"
            AlnNum = 1
            RaceNum = 4
        Case 6
            Cls = "Fighter"
            Armor = "Plate Mail"
            Weap1 = "Longsword"
            AC = "3"
        Case 7
            Cls = "Druid"
            Armor = "Leather"
            Weap1 = "Quarterstaff"
            AC = "8"
    End Select
    
    'Set the alignment
    Select Case AlnNum
        Case 1
            Aln = "Lawful Good"
        Case 2
            Aln = "Lawful Neutral"
        Case 3
            Aln = "Lawful Evil"
        Case 4
            Aln = "Neutral Good"
        Case 5
            Aln = "True Neutral"
        Case 6
            Aln = "Neutral Evil"
        Case 7
            Aln = "Chaotic Good"
        Case 8
            Aln = "Chaotic Neutral"
        Case 9
            Aln = "Chaotic Evil"
    End Select
    
    'Set the Race
    Select Case RaceNum
        Case 1
            Rce = "Elf"
        Case 2
            Rce = "Dwarf"
        Case 3
            Rce = "Half-Elf"
        Case 4
            Rce = "Human"
        Case 5
            Rce = "Halfling"
    End Select
    
    'Set the Age, Height and Weight based on Race
    Select Case Rce
        Case "Elf"
            Age = 100 + (Int(Rnd * 100) + 1)
            Hgt = 40 + (Int(Rnd * 30) + 1)
            Wgt = 90 + (Int(Rnd * 30) + 1)
        Case "Dwarf"
            Age = 100 + (Int(Rnd * 100) + 1)
            Hgt = 30 + (Int(Rnd * 25) + 1)
            Wgt = 90 + (Int(Rnd * 30) + 1)
        Case "Half-Elf"
            Age = 60 + (Int(Rnd * 100) + 1)
            Hgt = 50 + (Int(Rnd * 25) + 1)
            Wgt = 100 + (Int(Rnd * 50) + 1)
        Case "Halfling"
            Age = 100 + (Int(Rnd * 50) + 1)
            Hgt = 30 + (Int(Rnd * 25) + 1)
            Wgt = 70 + (Int(Rnd * 30) + 1)
        Case "Human"
            Age = 18 + (Int(Rnd * 10) + 1)
            Hgt = 50 + (Int(Rnd * 25) + 1)
            Wgt = 110 + (Int(Rnd * 50) + 1)
    End Select
    
    'Put all the data into a variable in the correct format.
    Character = Cls & "," & Lvl & "," & HtP & "," & Rce & "," & Age & _
        "," & Hgt & "," & Wgt & "," & Aln & "," & Stre & "," & Inte & _
        "," & Wisd & "," & Cons & "," & Dext & "," & Chri & "," & Weap1 & _
        "," & Weap2 & "," & Armor & "," & Para & "," & Petr & "," & Rod & _
        "," & Breath & "," & Spell & "," & Thac0 & "," & AC
    
    'Here we get the correct path and file
    chrFile = App.Path & "\npcs.dat"
    'Now we open the file to append data to the end with write access.
    Open chrFile For Append Access Write As #1
    Print #1, Character 'Write the data to the file
    Close #1  'Close the file
    cmdRefresh_Click 'Call the subroutine to refresh the list.
End Sub

Private Sub lstCharacters_DblClick()
    'Declare all variables.
    Dim intLoopIndex
    Dim sText() As String
    
    'The For...Next loop counts the number of list items and gives them a number.
    'The If...Then statement tests to see what is selected, and then sends the data
    'to the module to show the character sheet.
    For intLoopIndex = 0 To lstCharacters.ListCount - 1
        If lstCharacters.Selected(intLoopIndex) Then
            GetSheet (lstCharacters.List(intLoopIndex))
        End If
    Next intLoopIndex
End Sub

Private Sub cmdShow_Click()
    lstCharacters_DblClick 'Calls the subroutine to open the selected item.
End Sub

Private Sub cmdRefresh_Click()
    lstCharacters.Clear  'Clears the list.
    Form_Load            'Re-reads the file and generates the list from fresh data
End Sub

Private Sub cmdExit_Click()
    End 'Close the program
End Sub
