VERSION 5.00
Begin VB.UserControl GradientHScroll 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1980
   PropertyPages   =   "GradientHScroll.ctx":0000
   ScaleHeight     =   900
   ScaleWidth      =   1980
   ToolboxBitmap   =   "GradientHScroll.ctx":0021
End
Attribute VB_Name = "GradientHScroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'*********************************
'*DA'                            *
'* +-+    +-+   +-+  +-+    +-+  *
'* |  )  |   | |   | |  |  |   | *
'* +-+   |___| |___| |  |  +---+ *
'* |  )  |  *| |  *| |  |  |   | *
'* +-+    +-+   +-+  +-+   |   | *
'*                               *
'*********************************
'
' Horizontal Gradiant ScrollBar V 1.1
' Nov 2000
' Well here it is, It's crap, but I hope you learn
' something from it, feel free to use these controls,
' but give credit where credit is due.
' I know why would I want credit for this
' POS(Piece of S#$T) Control.
' I designed this for a game I was making
' because I think that regular scrollbars
' look just gay and they ruined the effect
' I was going for.
'
' Public your Events, but Private your Variables
'
Public Event Change() 'the only event that the user needs
'These are the starting colors
Private Sr As Integer, Sg As Integer, Sb As Integer
'Whether are not the cell is outlined or not
Private CO As Boolean
'These are the ending colors
Private Er As Integer, Eg As Integer, Eb As Integer
'These tell the Max and Min of the scroll bar
Private Mx As Integer
Private Mn As Integer
'These are the colors for the out line
Private Lr As Integer, Lg As Integer, Lb As Integer
'These are the colors for the background
Private Br As Integer, Bg As Integer, Bb As Integer
'These are dummys for finding color inc.(eg..sr to er)
Private Pr As Single, Pg As Single, Pb As Single
'These are the colors for the barscroll
Private Rr As Integer, Rg As Integer, Rb As Integer
'These are the options for the barscroll
Private BV As Boolean, BS As Boolean
'This is cell size
Private St As Integer
'Current value of the scroll
Private V As Integer
'This tells if the control is changeable or just display
'eg...percentage bars
Private ScE As Boolean
'This toggle is to keep it from drawing a 100 times during
'initialization and reading
Private MTog As Integer
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ScE = True Then
Dim Dx As Integer 'Dummy x
Dx = X / Screen.TwipsPerPixelX 'Find the mouse x and convert to pixels
Dx = Dx / St 'divide by cell size
V = Dx + Mn 'add new x to mn and thats your value
Value = V 'make the property so
RaiseEvent Change 'tell the user app that the value has changed
End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Same as above, this emulates you scrolling a vb scroll bar
If ScE = True And Button = 1 Then
Dim Dx As Integer
Dx = X / Screen.TwipsPerPixelX
Dx = Dx / St
V = Dx + Mn
If V > Mx Then V = Mx
If V < Mn Then V = Mn
Value = V
RaiseEvent Change
End If
End Sub

Private Sub UserControl_Resize()
DrawME 'any time the user resizes, it redraws the scroll to fit
End Sub

Private Sub UserControl_Initialize()
'this sub is the first thing to happen
'when the control is refreshed
'so if you want to put a sound, or something fancy
'this is the place to do it
'but beware, your variables haven't been read yet
'so make dummys for this event
End Sub

Private Sub UserControl_InitProperties()
'Setup Properties, does this event once, when the user creates
'the control
MTog = 1
'Sets the default property values
StartRed = 0: StartGreen = 0: StartBlue = 0
EndRed = 255: EndGreen = 255: EndBlue = 255
Max = 100: Min = 0
BoxRed = 255: BoxGreen = 255: BoxBlue = 255
BackRed = 0: BackGreen = 0: BackBlue = 0
BarRed = 255: BarGreen = 255: BarBlue = 255
BarVisible = True: BarSolid = False
Step = 1: MTog = 0: Value = 100
ScrollEnabled = True
CellOutlined = True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'whenever a control is refreshed it reads the property
'values out of a prop bag, it compares them to your default
'and if the same, doesn't bother, saves memory
MTog = 1 'Keeps the bar from being redraw over and over
StartRed = PropBag.ReadProperty("StartRed", 0)
'these are all the same, so will breakdown just one
'your property = whatever is stored in the bag
'the name in captions is a data name
'where the data is stored
'the number at the end is a default value
'set this to whatever your start default is.
StartGreen = PropBag.ReadProperty("StartGreen", 0)
StartBlue = PropBag.ReadProperty("StartBlue", 0)
EndRed = PropBag.ReadProperty("EndRed", 255)
EndGreen = PropBag.ReadProperty("EndGreen", 255)
EndBlue = PropBag.ReadProperty("EndBlue", 255)
CellOutlined = PropBag.ReadProperty("CellOutLined", True)
Max = PropBag.ReadProperty("Max", 100)
Min = PropBag.ReadProperty("Min", 0)
BoxRed = PropBag.ReadProperty("BoxRed", 255)
BoxGreen = PropBag.ReadProperty("BoxGreen", 255)
BoxBlue = PropBag.ReadProperty("BoxBlue", 255)
BackRed = PropBag.ReadProperty("BackRed", 0)
BackGreen = PropBag.ReadProperty("BackGreen", 0)
BackBlue = PropBag.ReadProperty("BackBlue", 0)
BarRed = PropBag.ReadProperty("BarRed", 255)
BarGreen = PropBag.ReadProperty("BarGreen", 255)
BarBlue = PropBag.ReadProperty("BarBlue", 255)
BarVisible = PropBag.ReadProperty("BarVisible", True)
BarSolid = PropBag.ReadProperty("BarSolid", False)
Step = PropBag.ReadProperty("Step", 1)
ScrollEnabled = PropBag.ReadProperty("ScrollEnabled", True)
MTog = 0 'change tog back before last value to enable DrawMe
Value = PropBag.ReadProperty("Value", 100)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'Look at this this way, whenever the control
'becomes invisible, albeit minimizing or closing the user
'form. it writes all the property values here,
'to be read when the control is restored, or becomes
'seen.  This is a good vb extra, keeps the memory
'from being zapped by a hundred little unseen controls.
PropBag.WriteProperty "StartRed", Sr, 0
'Same thing as above, except you plug in the value
'of the current property after data name.
'also make sure your data name corresponds
'with the read data name.
'will save you hours of debug time, trust me, I have
'no hair left on my head after that incident.
PropBag.WriteProperty "StartGreen", Sg, 0
PropBag.WriteProperty "StartBlue", Sb, 0
PropBag.WriteProperty "EndRed", Er, 255
PropBag.WriteProperty "EndGreen", Eg, 255
PropBag.WriteProperty "EndBlue", Eb, 255
PropBag.WriteProperty "CellOutLined", CO, True
PropBag.WriteProperty "Max", Mx, 100
PropBag.WriteProperty "Min", Mn, 0
PropBag.WriteProperty "BoxRed", Lr, 255
PropBag.WriteProperty "BoxGreen", Lg, 255
PropBag.WriteProperty "BoxBlue", Lb, 255
PropBag.WriteProperty "BackRed", Br, 0
PropBag.WriteProperty "BackGreen", Bg, 0
PropBag.WriteProperty "BackBlue", Bb, 0
PropBag.WriteProperty "BarRed", Rr, 255
PropBag.WriteProperty "BarGreen", Rg, 255
PropBag.WriteProperty "BarBlue", Rb, 255
PropBag.WriteProperty "BarVisible", BV, True
PropBag.WriteProperty "BarSolid", BS, False
PropBag.WriteProperty "Step", St, 1
PropBag.WriteProperty "ScrollEnabled", ScE, True
PropBag.WriteProperty "Value", V, 100
End Sub

Private Sub UserControl_Terminate()
'this event executes whenever the control becomes
'unaccessable by the user
'good place to add a closing effect, (eg..Sound,message)
End Sub
Private Sub DrawME()
'Imagine with out active x I would have to cut and paste
'this code several times, but with it I just have to
'write it once.  which saves your program from getting
'cluttered with a bunch of repetitive code.
'This is the Sub where it all happens.
'
'This takes the control a resizes it to the mx and mn values
'this is neccessary because value is based on a pixel amount
UserControl.Width = (((Mx - Mn + 1) * St) + 2) * Screen.TwipsPerPixelX
'changes the backcolor
UserControl.BackColor = RGB(Br, Bg, Bb)
'clears original, if you didn't know this go back to qbasic
UserControl.Cls
'dummy x and y values
Dim X As Single, Y As Single
'find out how wide the cell is going to be
Y = UserControl.Height
Y = Y - 2 * Screen.TwipsPerPixelY
'd will be the length of the bar
Dim d As Integer
d = Mx - Mn
If d = 0 Then d = 1 'had to do this to keep division by zero from happening
Pr = (Er - Sr) / d 'these three find the color inc to add each cell
Pg = (Eg - Sg) / d
Pb = (Eb - Sb) / d
'simple dummys
Dim a As Single, b As Single
For a = 0 To V - Mn 'go from the begining to the end of bar
b = a * St 'calculate cell size
'Draw the cell
UserControl.Line ((1 + b) * Screen.TwipsPerPixelX, 1 * Screen.TwipsPerPixelY)-((1 + b + St - 1) * Screen.TwipsPerPixelX, Y), RGB(Sr + (a * Pr), Sg + (a * Pg), Sb + (a * Pb)), BF
'if cell is bigger than one the outline with backcolor
If CO = True And St > 2 Then
UserControl.Line ((1 + b) * Screen.TwipsPerPixelX, 1 * Screen.TwipsPerPixelY)-((1 + b + St - 1) * Screen.TwipsPerPixelX, Y), RGB(Br, Bg, Bb), B
End If
Next a
'this sets up the barscroll at the current value
b = (V - Mn) * St
If BarVisible = True Then
If BarSolid = False Then 'draw just outline
UserControl.Line ((1 + b) * Screen.TwipsPerPixelX, 1 * Screen.TwipsPerPixelY)-((1 + b + St - 1) * Screen.TwipsPerPixelX, Y), RGB(Rr, Rg, Rb), B
Else: 'draw solid box
UserControl.Line ((1 + b) * Screen.TwipsPerPixelX, 1 * Screen.TwipsPerPixelY)-((1 + b + St - 1) * Screen.TwipsPerPixelX, Y), RGB(Rr, Rg, Rb), BF
End If
End If
'Draw out line
UserControl.Line (0, 0)-(UserControl.Width - 1 * Screen.TwipsPerPixelX, UserControl.Height - 1 * Screen.TwipsPerPixelX), RGB(Lr, Lg, Lb), B
End Sub
'This last bit of code is all your propertys
'whenever you want a public property you have to put these
'two statements in, get and let, make sure the as is
'congruent with the as of the temp variable.
'and warning don't put the propertyname in the let
'if will give you a out of stack space error
'because it keeps running another get without closing the
'one before it.
'I will breakdown the first two
Public Property Get StartRed() As Integer
Attribute StartRed.VB_ProcData.VB_Invoke_Property = "HGColor"
'whenever the user asks for the property startred
'it comes to this sub and returns the temp value for it
StartRed = Sr
End Property

Public Property Let StartRed(ByVal vNewValue As Integer)
'make sure you make the temp = vnewvalue, or
'youll get the dreaded out of stack space error.
Sr = vNewValue
If Sr > 255 Then Sr = 255
If Sr < 0 Then Sr = 0
PropertyChanged "StartRed" 'whenever you change a property
                            'you have to call this method
                            'it changes it for the user
If MTog = 0 Then DrawME 'this again is to keep it from redrawing
                        'unless changed independently
End Property
Public Property Get StartGreen() As Integer
Attribute StartGreen.VB_ProcData.VB_Invoke_Property = "HGColor"
StartGreen = Sg
End Property

Public Property Let StartGreen(ByVal vNewValue As Integer)
Sg = vNewValue
If Sg > 255 Then Sg = 255
If Sg < 0 Then Sg = 0
PropertyChanged "StartGreen"
If MTog = 0 Then DrawME
End Property
Public Property Get StartBlue() As Integer
Attribute StartBlue.VB_ProcData.VB_Invoke_Property = "HGColor"
StartBlue = Sb
End Property

Public Property Let StartBlue(ByVal vNewValue As Integer)
Sb = vNewValue
If Sb > 255 Then Sb = 255
If Sb < 0 Then Sb = 0
PropertyChanged "StartBlue"
If MTog = 0 Then DrawME
End Property
Public Property Get EndRed() As Integer
Attribute EndRed.VB_ProcData.VB_Invoke_Property = "HGColor"
EndRed = Er
End Property

Public Property Let EndRed(ByVal vNewValue As Integer)
Er = vNewValue
If Er > 255 Then Er = 255
If Er < 0 Then Er = 0
PropertyChanged "EndRed"
If MTog = 0 Then DrawME
End Property
Public Property Get EndGreen() As Integer
Attribute EndGreen.VB_ProcData.VB_Invoke_Property = "HGColor"
EndGreen = Eg
End Property

Public Property Let EndGreen(ByVal vNewValue As Integer)
Eg = vNewValue
If Eg > 255 Then Eg = 255
If Eg < 0 Then Eg = 0
PropertyChanged "EndGreen"
If MTog = 0 Then DrawME
End Property
Public Property Get EndBlue() As Integer
Attribute EndBlue.VB_ProcData.VB_Invoke_Property = "HGColor"
EndBlue = Eb
End Property

Public Property Let EndBlue(ByVal vNewValue As Integer)
Eb = vNewValue
If Eb > 255 Then Eb = 255
If Eb < 0 Then Eb = 0
PropertyChanged EndBlue
If MTog = 0 Then DrawME
End Property
Public Property Get Max() As Integer
Attribute Max.VB_ProcData.VB_Invoke_Property = "HGSize"
Max = Mx
End Property

Public Property Let Max(ByVal vNewValue As Integer)
Mx = vNewValue
If Mx < Mn Then Mx = Mn
If Mx > 999 Then Mx = 999
If V > Mx Then V = Mx
PropertyChanged "Max"
If MTog = 0 Then DrawME
End Property
Public Property Get Min() As Integer
Attribute Min.VB_ProcData.VB_Invoke_Property = "HGSize"
Min = Mn
End Property

Public Property Let Min(ByVal vNewValue As Integer)
Mn = vNewValue
If Mn > Mx Then Mn = Mx
If Mn < 0 Then Mn = 0
If V < Mn Then V = Mn
PropertyChanged "Min"
If MTog = 0 Then DrawME
End Property
Public Property Get BoxRed() As Integer
Attribute BoxRed.VB_ProcData.VB_Invoke_Property = "HGColor"
BoxRed = Lr
End Property

Public Property Let BoxRed(ByVal vNewValue As Integer)
Lr = vNewValue
If Lr > 255 Then Lr = 255
If Lr < 0 Then Lr = 0
PropertyChanged "BoxRed"
If MTog = 0 Then DrawME
End Property
Public Property Get BoxGreen() As Integer
Attribute BoxGreen.VB_ProcData.VB_Invoke_Property = "HGColor"
BoxGreen = Lg
End Property

Public Property Let BoxGreen(ByVal vNewValue As Integer)
Lg = vNewValue
If Lg > 255 Then Lg = 255
If Lg < 0 Then Lg = 0
PropertyChanged "BoxGreen"
If MTog = 0 Then DrawME
End Property
Public Property Get BoxBlue() As Integer
Attribute BoxBlue.VB_ProcData.VB_Invoke_Property = "HGColor"
BoxBlue = Lb
End Property

Public Property Let BoxBlue(ByVal vNewValue As Integer)
Lb = vNewValue
If Lb > 255 Then Lb = 255
If Lb < 0 Then Lb = 0
PropertyChanged "BoxBlue"
If MTog = 0 Then DrawME
End Property
Public Property Get BackRed() As Integer
Attribute BackRed.VB_ProcData.VB_Invoke_Property = "HGColor"
BackRed = Br
End Property

Public Property Let BackRed(ByVal vNewValue As Integer)
Br = vNewValue
If Br > 255 Then Br = 255
If Br < 0 Then Br = 0
PropertyChanged "BackRed"
If MTog = 0 Then DrawME
End Property
Public Property Get BackGreen() As Integer
Attribute BackGreen.VB_ProcData.VB_Invoke_Property = "HGColor"
BackGreen = Bg
End Property

Public Property Let BackGreen(ByVal vNewValue As Integer)
Bg = vNewValue
If Bg > 255 Then Bg = 255
If Bg < 0 Then Bg = 0
PropertyChanged "BackGreen"
If MTog = 0 Then DrawME
End Property
Public Property Get BackBlue() As Integer
Attribute BackBlue.VB_ProcData.VB_Invoke_Property = "HGColor"
BackBlue = Bb
End Property

Public Property Let BackBlue(ByVal vNewValue As Integer)
Bb = vNewValue
If Bb > 255 Then Bb = 255
If Bb < 0 Then Bb = 0
PropertyChanged "BackBlue"
If MTog = 0 Then DrawME
End Property
Public Property Get BarRed() As Integer
Attribute BarRed.VB_ProcData.VB_Invoke_Property = "HGColor"
BarRed = Rr
End Property

Public Property Let BarRed(ByVal vNewValue As Integer)
Rr = vNewValue
If Rr > 255 Then Rr = 255
If Rr < 0 Then Rr = 0
PropertyChanged "BarRed"
If MTog = 0 Then DrawME
End Property
Public Property Get BarGreen() As Integer
Attribute BarGreen.VB_ProcData.VB_Invoke_Property = "HGColor"
BarGreen = Rg
End Property

Public Property Let BarGreen(ByVal vNewValue As Integer)
Rg = vNewValue
If Rg > 255 Then Rg = 255
If Rg < 0 Then Rg = 0
PropertyChanged "BarGreen"
If MTog = 0 Then DrawME
End Property
Public Property Get BarBlue() As Integer
Attribute BarBlue.VB_ProcData.VB_Invoke_Property = "HGColor"
BarBlue = Rb
End Property

Public Property Let BarBlue(ByVal vNewValue As Integer)
Rb = vNewValue
If Rb > 255 Then Rb = 255
If Rb < 0 Then Rb = 0
PropertyChanged "BarBlue"
If MTog = 0 Then DrawME
End Property
Public Property Get BarVisible() As Boolean
Attribute BarVisible.VB_ProcData.VB_Invoke_Property = "HGSize"
BarVisible = BV
End Property

Public Property Let BarVisible(ByVal vNewValue As Boolean)
BV = vNewValue
PropertyChanged "BarVisible"
If MTog = 0 Then DrawME
End Property
Public Property Get BarSolid() As Boolean
Attribute BarSolid.VB_ProcData.VB_Invoke_Property = "HGSize"
BarSolid = BS
End Property

Public Property Let BarSolid(ByVal vNewValue As Boolean)
BS = vNewValue
PropertyChanged "BarSolid"
If MTog = 0 Then DrawME
End Property

Public Property Get Step() As Integer
Attribute Step.VB_ProcData.VB_Invoke_Property = "HGSize"
Step = St
End Property

Public Property Let Step(ByVal vNewValue As Integer)
St = vNewValue
If St > 999 Then St = 999
If St < 1 Then St = 1
PropertyChanged "Step"
If MTog = 0 Then DrawME
End Property

Public Property Get Value() As Integer
Attribute Value.VB_ProcData.VB_Invoke_Property = "HGSize"
Value = V
End Property

Public Property Let Value(ByVal vNewValue As Integer)
V = vNewValue
If V > Mx Then V = Mx
If V < Mn Then V = Mn
PropertyChanged "Value"
If MTog = 0 Then DrawME
End Property


Public Property Get ScrollEnabled() As Boolean
Attribute ScrollEnabled.VB_ProcData.VB_Invoke_Property = "HGSize"
ScrollEnabled = ScE
End Property

Public Property Let ScrollEnabled(ByVal vNewValue As Boolean)
ScE = vNewValue
PropertyChanged "ScrollEnabled"
If MTog = 0 Then DrawME
End Property

Public Property Get CellOutlined() As Boolean
CellOutlined = CO
End Property

Public Property Let CellOutlined(ByVal vNewValue As Boolean)
CO = vNewValue
PropertyChanged "CellOutlined"
If MTog = 0 Then DrawME
End Property
