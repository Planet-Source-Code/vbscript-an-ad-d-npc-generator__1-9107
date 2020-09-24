Attribute VB_Name = "Module"
Option Explicit

Function GetSheet(CharText As String)
    'Declare all variables used in the function.
    Dim Cls, Lvl, HtP, Rce, Age, Hgt, Wgt, Aln, Stre
    Dim Inte, Wisd, Cons, Dext, Chri, Weap1, Weap2, Armor
    Dim Para, Petr, Rod, Breath, Spell, Thac0, AC
    Dim sText() As String
    
    'Here I split the string with the comma as the
    'delimiter and put each field into the array sText()
    sText() = Split(CharText, ",")
    
    'Force the elements of the array into simple variables.
    Cls = sText(0)
    Lvl = sText(1)
    HtP = sText(2)
    Rce = sText(3)
    Age = sText(4)
    Hgt = sText(5)
    Wgt = sText(6)
    Aln = sText(7)
    Stre = sText(8)
    Inte = sText(9)
    Wisd = sText(10)
    Cons = sText(11)
    Dext = sText(12)
    Chri = sText(13)
    Weap1 = sText(14)
    Weap2 = sText(15)
    Armor = sText(16)
    Para = sText(17)
    Petr = sText(18)
    Rod = sText(19)
    Breath = sText(20)
    Spell = sText(21)
    Thac0 = sText(22)
    AC = sText(23)
        
    'Fill out the character sheet with the appropriate values.
    frmCharacter.txtCls.Text = Cls
    frmCharacter.txtLvl.Text = Lvl
    frmCharacter.txtHtP.Text = HtP
    frmCharacter.txtRce.Text = Rce
    frmCharacter.txtAge.Text = Age
    frmCharacter.txtHgt.Text = Hgt
    frmCharacter.txtWgt.Text = Wgt
    frmCharacter.txtAln.Text = Aln
    frmCharacter.txtStre.Text = Stre
    frmCharacter.txtInte.Text = Inte
    frmCharacter.txtWisd.Text = Wisd
    frmCharacter.txtCons.Text = Cons
    frmCharacter.txtDext.Text = Dext
    frmCharacter.txtChri.Text = Chri
    frmCharacter.txtWeap1.Text = Weap1
    frmCharacter.txtWeap2.Text = Weap2
    frmCharacter.txtArmor.Text = Armor
    frmCharacter.txtPara.Text = Para
    frmCharacter.txtPetr.Text = Petr
    frmCharacter.txtRod.Text = Rod
    frmCharacter.txtBreath.Text = Breath
    frmCharacter.txtSpell.Text = Spell
    frmCharacter.txtThac0.Text = Thac0
    frmCharacter.txtAC.Text = AC
    
    'Show the Character Sheet form, filled out.
    frmCharacter.Show
    
End Function
