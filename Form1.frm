VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2175
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   3836
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":0000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   2520
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const AnInch As Long = 1440   '1440 twips per inch
Private Const QuarterInch As Long = 360

Private Sub Form_Load()
   Dim PrintableWidth As Long
   Dim PrintableHeight As Long
   Dim x As Single

   ' Initialize Form and Command button
   Me.Caption = "Rich Text Box WYSIWYG Printing Example"
   Command1.Move 10, 10, 600, 380
   Command1.Caption = "&Print"

   ' Set the font of the RTF to a TrueType font for best results
   RichTextBox1.SelFontName = "Arial"
   RichTextBox1.SelFontSize = 10
   
   'initialize the printer object
   x = Printer.TwipsPerPixelX
   Printer.Orientation = vbPRORPortrait  'vbPRORLandscape

   ' Tell the RTF to base it's display off of the printer
   Call WYSIWYG_RTF(RichTextBox1, QuarterInch, QuarterInch, QuarterInch, QuarterInch, PrintableWidth, PrintableHeight) '1440 Twips=1 Inch

   ' Set the form width to match the line width
   Me.Width = PrintableWidth + 200
   Me.Height = PrintableHeight + 800
End Sub

Private Sub Form_Resize()
   ' Position the RTF on form
   RichTextBox1.Move 100, 500, Me.ScaleWidth - 200, Me.ScaleHeight - 600
End Sub

Private Sub Command1_Click()
   ' Print the contents of the RichTextBox with a one inch margin
   PrintRTF RichTextBox1, AnInch, AnInch, AnInch, AnInch
End Sub

