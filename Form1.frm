VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   825
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2325
   LinkTopic       =   "Form1"
   ScaleHeight     =   825
   ScaleWidth      =   2325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   960
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
      ' DrawEdge.frm - Demonstrates a simple method of using the
      '                DrawEdge function.
      '********************************************************************
      Option Explicit

      '********************************************************************
      ' Prepares the form and Picture1 for use with the DrawEdge function.
      '********************************************************************

      Private Sub Form_Load()
      '--------------------------------------------------------------------
      ' Always set the ScaleMode to pixels when using API drawing
      ' functions.
      '--------------------------------------------------------------------
         ScaleMode = vbPixels
         With Picture1
            '--------------------------------------------------------------
            ' The next line is not required if you put your drawing code
            ' in the Paint event.
            '--------------------------------------------------------------
            .AutoRedraw = True

            '--------------------------------------------------------------
            ' Set the Backcolor, set the Borderstyle to none, and size
            ' the picture box to a more realistic button size.
            '--------------------------------------------------------------
            .BackColor = vb3DFace
            .BorderStyle = 0
            .Move 60, 10, 90, 30

            '--------------------------------------------------------------
            ' Make sure the picture box uses the pixel ScaleMode, and
            ' set the tag of the control to a caption for later use by
            ' the DrawControl function.
            '--------------------------------------------------------------
            .ScaleMode = vbPixels
            .Tag = "DrawEdge Test"
         End With

         '-----------------------------------------------------------------
         ' Draw the initial button.
         '-----------------------------------------------------------------
         DrawControl Picture1, Picture1.Tag, EDGE_RAISED

      End Sub

      '********************************************************************
      ' When the picture box gets a click event, draw a etched box on the
      ' upper-left corner of the form.
      '********************************************************************

      Private Sub Picture1_Click()
         Dim r As RECT   ' Used by DrawEdge to determine where to draw.

         '-----------------------------------------------------------------
         ' Location of the etched box.
         '-----------------------------------------------------------------
         With r
            .Left = 10
            .Top = 10
            .Right = 50
            .Bottom = 50
         End With

         '-----------------------------------------------------------------
         ' Draw it.
         '-----------------------------------------------------------------
         DrawEdge hdc, r, EDGE_ETCHED, BF_RECT

      End Sub

      '********************************************************************
      ' When the user presses the mouse down on the picture box,
      ' draw a sunken edge to simulate a depressed button.
      '********************************************************************
      Private Sub Picture1_MouseDown(Button%, Shift%, X!, Y!)
         DrawControl Picture1, Picture1.Tag, EDGE_SUNKEN
      End Sub

      '********************************************************************
      ' When the user releases the mouse over the picture box, draw a
      ' standard button.
      '********************************************************************
      Private Sub Picture1_MouseUp(Button%, Shift%, X!, Y!)
         DrawControl Picture1, Picture1.Tag, EDGE_RAISED
      End Sub

      '********************************************************************
      ' The DrawControl helper function is designed to make it easier to
      ' draw a button on a picture box.
      '********************************************************************
      Private Sub DrawControl(picControl As PictureBox, _
                  strCaption As String, Optional vntEdge)

      Dim r As RECT    ' Holds the location of the DrawEdge rectangle.
      Dim intOffset%   ' Used to shift the caption when the button is
                       ' pressed.
      '--------------------------------------------------------------------
      ' If the user does not provide a Edge flag, then use a default value.
      '--------------------------------------------------------------------
      vntEdge = IIf(IsMissing(vntEdge), EDGE_RAISED, vntEdge)

      '-------------------------------------------------------------------
      ' Clear the picture control and determine where to draw the new
      ' rectangle and caption.
      '-------------------------------------------------------------------
      With picControl
         .Cls
         r.Left = .ScaleLeft
         r.Top = .ScaleTop
         r.Right = .ScaleWidth
         r.Bottom = .ScaleHeight

         If vntEdge = EDGE_SUNKEN Then intOffset = 2

         .CurrentX = (.ScaleWidth - .TextWidth(strCaption) _
                     + intOffset) / 2
         .CurrentY = (.ScaleHeight - .TextHeight(strCaption) _
                     + intOffset) / 2

      End With

      '-------------------------------------------------------------------
      ' Draw the caption, then draw the rectangle.
      '-------------------------------------------------------------------
      Picture1.Print strCaption
      DrawEdge picControl.hdc, r, CLng(vntEdge), BF_RECT

      '-------------------------------------------------------------------
      ' If AutoRedraw is True then any drawing done by an API call is not
      ' seen until the picture box gets refreshed.
      '-------------------------------------------------------------------
      If picControl.AutoRedraw Then picControl.Refresh
      End Sub


