VERSION 5.00
Begin VB.Form frmShip 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Ha!"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   ScaleHeight     =   380
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   413
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmShip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'The gettickcount API delcaration
Private Declare Function GetTickCount Lib "kernel32" () As Long

Const PI = 3.14159
Const ACCEL = 0.1
Const ROTATION_RATE = 0.05
Dim ZOOM_LEVEL As Single
Const MS_DELAY = 25

Dim mlngTimer As Long       'Holds system time since last frame was displayed
Dim msngFacing As Single    'Angle the ship is facing (ok, ok, it's a triangle, not a ship! Shut it!)
Dim msngHeading As Single   'Current direction in which ship is moving
Dim msngSpeed As Single     'Current speed with which ship is moving
Dim msngX As Single         'Current X coordinate of ship within form
Dim msngY As Single         'Current Y coordinate of ship within form
Dim mblnRunning As Boolean  'Is the render loop running?
Dim mblnLeftKey As Boolean  'Is the left arrow-key depressed?
Dim mblnRightKey As Boolean 'Is the right arrow-key depressed?
Dim mblnUpKey As Boolean    'Is the up arrow-key depressed?
Dim mblnDownKey As Boolean    'Is the up arrow-key depressed?

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    'Check for keypresses
    If KeyCode = vbKeyLeft And Not mblnRightKey Then mblnLeftKey = True
    If KeyCode = vbKeyRight And Not mblnLeftKey Then mblnRightKey = True
    If KeyCode = vbKeyUp Then mblnUpKey = True
    If KeyCode = vbKeyDown Then mblnDownKey = True

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    'Check for keyreleases
    If KeyCode = vbKeyLeft Then mblnLeftKey = False
    If KeyCode = vbKeyRight Then mblnRightKey = False
    If KeyCode = vbKeyUp Then mblnUpKey = False
    If KeyCode = vbKeyDown Then mblnDownKey = False

End Sub

Private Sub Form_Load()

ZOOM_LEVEL = 75
AutoRedraw = True
    'Initialize the variables
    msngX = ScaleWidth / 2 - ZOOM_LEVEL / 2
    msngY = ScaleHeight / 2 - ZOOM_LEVEL / 2
    msngFacing = 0
    msngSpeed = 0
    msngHeading = 0
    mlngTimer = GetTickCount()
    mblnRunning = True
    
    'Display the form
    DrawShip
    frmShip.Show
    
    'Start the render loop
    Do While mblnRunning
        'Check if we've waited for the appropriate number of milliseconds
        If mlngTimer + MS_DELAY <= GetTickCount() Then
            mlngTimer = GetTickCount()  'Reset the timer variable
            Cls
            DrawShip                    'Draw the ship
            Physics
        End If
        'Allow other events to occur
        DoEvents
    Loop

End Sub

Private Sub Physics()
If mblnRightKey Then
    msngFacing = msngFacing + ROTATION_RATE
End If

If mblnLeftKey Then
    msngFacing = msngFacing - ROTATION_RATE
End If
End Sub

Private Sub DrawShip()

Dim intX1 As Integer    'Coordinates of the 4 rectangle verticies
Dim intY1 As Integer
Dim intX2 As Integer
Dim intY2 As Integer
Dim intX3 As Integer
Dim intY3 As Integer
Dim intX4 As Integer
Dim intY4 As Integer
Dim intX5 As Integer
Dim intY5 As Integer

'Do some zooming if needed
If mblnUpKey Then
    ZOOM_LEVEL = ZOOM_LEVEL + 1
End If
    
If mblnDownKey Then
    ZOOM_LEVEL = ZOOM_LEVEL - 1
End If

'Calculate new verticies
    intX1 = msngX + ZOOM_LEVEL * Sin(msngFacing)
    intY1 = msngY - ZOOM_LEVEL * Cos(msngFacing)
    intX2 = msngX + ZOOM_LEVEL * Sin(msngFacing + 2 * PI / 3)
    intY2 = msngY - ZOOM_LEVEL * Cos(msngFacing + 2 * PI / 3)
    intX3 = msngX + ZOOM_LEVEL * Sin(msngFacing + 2 * PI / 3)
    intY3 = msngY - ZOOM_LEVEL * Cos(msngFacing + 2 * PI / 3)
    intX4 = msngX - ZOOM_LEVEL * Sin(msngFacing)
    intY4 = msngY + ZOOM_LEVEL * Cos(msngFacing)
    intX5 = msngX - ZOOM_LEVEL * Sin(msngFacing + 2 * PI / 3)
    intY5 = msngY + ZOOM_LEVEL * Cos(msngFacing + 2 * PI / 3)

'Draw the rectangle
    Line (intX1, intY1)-(intX2, intY2), vbWhite
    Line (intX3, intY3)-(intX4, intY4), vbWhite
    Line (intX4, intY4)-(intX5, intY5), vbWhite
    Line (intX1, intY1)-(intX5, intY5), vbWhite
    
'Draw the center
    PSet (msngX, msngY), vbWhite

'Draw instructions
CurrentX = 10
CurrentY = 10
ForeColor = vbBlue
Print "Left and right to rotate, up and down to zoom"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    msngX = X
    msngY = Y
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    msngX = X
    msngY = Y
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    msngX = X
    msngY = Y
End If
End Sub

Private Sub Form_Resize()
    msngX = ScaleWidth / 2 - ZOOM_LEVEL / 2
    msngY = ScaleHeight / 2 - ZOOM_LEVEL / 2

End Sub

Private Sub Form_Unload(Cancel As Integer)

    'Terminate the render loop
    mblnRunning = False

End Sub
