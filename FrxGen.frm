VERSION 5.00
Begin VB.Form FrmFrxGen 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Frx Generator"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5595
   Icon            =   "FrxGen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   237
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   373
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox PicShow 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   2775
      Left            =   2280
      MousePointer    =   2  'Kreuz
      ScaleHeight     =   2775
      ScaleWidth      =   3135
      TabIndex        =   6
      Top             =   120
      Width           =   3135
   End
   Begin VB.PictureBox PicBinary 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   255
      Left            =   1320
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   7
      Top             =   4800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox ChkKeep 
      Caption         =   "Keep Picture as Background"
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   3240
      Value           =   1  'Aktiviert
      Width           =   3015
   End
   Begin VB.CommandButton SaveForm 
      Caption         =   "Save Form"
      Enabled         =   0   'False
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton PrevForm 
      Caption         =   "Preview Form"
      Enabled         =   0   'False
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton SelectColor 
      Caption         =   "Transparent Color"
      Enabled         =   0   'False
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton LoadPic 
      Caption         =   "Load Picture"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.PictureBox PicTmp 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   3015
      Left            =   2400
      ScaleHeight     =   3015
      ScaleWidth      =   4455
      TabIndex        =   0
      Top             =   3120
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Or click on the Picture"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Shape ShapeTrans 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   255
      Left            =   1680
      Top             =   1080
      Width           =   255
   End
End
Attribute VB_Name = "FrmFrxGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Frx Generator
'Create Shaped Forms fast and Easy
'© ScytheVB 2005
'www.scythe-tools.de

'Thx to Robert Gainor
'he made a nice Transparent Form Maker
'I took a little code from his version´s
'http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=50405&lngWId=1

'Thx to Robert Rayment who found some errors

'Thx to LaVolpe who inspired me to make the calculations routines faster.
'He also made a great Shaped form Code find it on PSC
'http://www.Planet-Source-Code.com/vb/default.asp?lngCId=54017&lngWId=1
'His creation routine is twice as fast than mine but
'He has to calculate the thing every time (No Saving for the Form and the data)
'and it seems to make some problems

'Thx to all others on Planet Source Code


'New in Version 1.0.1
'Removed some errors in Cmdialog.bas
'Changed one error in TransRout (calculation was from wrong Picturebox)
'Speeded up CalcPic routine (The new needs only about 20% of the time my old needed)

Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As RasterOpConstants) As Long

'Fast binary Data
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Private PicInfo As BITMAP

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type


Private PictureName As String 'Name of the Background Image
Private PicSize As Long 'Size for the last Picture we did

'Load a new Picture and Show it
'using its aspect ratio
Private Sub LoadNewPic()

    PicTmp.Picture = LoadPicture(PictureName) 'Load Picture
    FrmTest.Width = PicTmp.Width * Screen.TwipsPerPixelX 'Resize Testframe
    FrmTest.Height = PicTmp.Height * Screen.TwipsPerPixelY
    
'Resize the visible Picturebox
    If PicTmp.Width > PicTmp.Height Then
        PicShow.Height = 200 * (PicTmp.Height / PicTmp.Width)
        PicShow.Width = 200
    Else
        PicShow.Height = 200
        PicShow.Width = 200 * (PicTmp.Width / PicTmp.Height)
    End If
'Copy the Picture to the Visible Picturebox
    StretchBlt PicShow.hDC, 0, 0, PicShow.Width, PicShow.Height, PicTmp.hDC, 0, 0, PicTmp.Width, PicTmp.Height, vbSrcCopy
    PicShow.Refresh

End Sub


'Load a Picture
Private Sub LoadPic_Click()

    PictureName = OpenDialog(Me.hWnd, App.Path & "\", "All Files" & vbNullChar & "*.*" & vbNullChar & vbNullChar, False)
    If PictureName <> "" Then LoadNewPic
    SelectColor.Enabled = True
    PrevForm.Enabled = True
    SaveForm.Enabled = True
    CalculationDone = False

End Sub

'Select a Transparent Color
Private Sub PicShow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    TransColor = PicShow.Point(X, Y)
    ShapeTrans.BackColor = TransColor
    CalculationDone = False

End Sub
Private Sub SelectColor_Click()

Dim Col As Long
    
    Col = ColorDialog(Me.hWnd, TransColor)

    If Col <> -1 Then
        ShapeTrans.BackColor = Col
    End If
    CalculationDone = False

End Sub

'Show the Preview
Private Sub PrevForm_Click()

'Load Backgroundimage if wanted
    If ChkKeep.Value Then
     Set FrmTest.Picture = PicTmp.Image
    Else
        FrmTest.Picture = LoadPicture()
    End If
'Calculate the Region to Shape the form
    If CalculationDone = False Then CalcPic
'Show the Form
    FrmTest.Show

End Sub

'Hide Preview if vissible
Private Sub Form_Activate()

    FrmTest.Hide

End Sub

'Remove all end exit programm
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Unload FrmTest
    Unload Me
    End

End Sub

'Create a new Formfile
'Include a Frx file so we still have the picture
Private Sub SaveForm_Click()

Dim SaveName As String
Dim Fname As String
Dim BackPic As Long

'Not Calculate the Region to Shape the form so do it
    If CalculationDone = False Then CalcPic
    
'Show Save dialog
    SaveName = SaveDialog(Me.hWnd, App.Path & "\", "*.*")
    If SaveName = "" Then Exit Sub

'Get the Filename in a way we could use it
    SaveName = CutExt(SaveName, ".")
    Fname = CutAfter(SaveName, "\")


'Create a Picture to store our RegionData
'So the code in the form is short an clean
'and we need less data
    PicBinary.Width = UBound(RgnData) / 4 + 1
    PicBinary.Height = 1
    GetObject PicBinary.Image, Len(PicInfo), PicInfo
    SetBitmapBits PicBinary.Image, UBound(RgnData), RgnData(0)
    SavePicture PicBinary.Image, App.Path & "\tempBinary.dat"

'Now Create the FrxFile
    Open SaveName & ".frx" For Binary Access Write As #1
'Store the Background Picture on the new Form
        If ChkKeep.Value = 1 Then WriteFrx PictureName: BackPic = PicSize
'Store the Binary data (Temporary)
        WriteFrx App.Path & "\tempBinary.dat"
    Close
    
    Open SaveName & ".frm" For Output As #1
        Print #1, "VERSION 5.00"
        Print #1, "Begin VB.Form " & Fname
        Print #1, "   BorderStyle     =   0  'Kein"
        Print #1, "   ClientHeight    =   " & Trim(Str(FrmTest.Height))
        Print #1, "   ClientLeft      =   0"
        Print #1, "   ClientTop       =   0"
        Print #1, "   ClientWidth     =   " & Trim(Str(FrmTest.Width))
        Print #1, "   ControlBox      =   0   'False"
        Print #1, "   LinkTopic       =   """ & Fname & """"
        Print #1, "   MaxButton       =   0   'False"
        Print #1, "   MinButton       =   0   'False"
'You dont want a Background Picture
'so we dont load it in the form
        If BackPic > 0 Then
            Print #1, "   Picture         =   """ & Fname; ".frx"":0000"
        End If
        Print #1, "   ScaleHeight     =   " & Trim(Str(FrmTest.Height))
        Print #1, "   ScaleWidth      =   " & Trim(Str(FrmTest.Width))
        Print #1, "   ShowInTaskbar   =   0   'False"
        Print #1, "   StartUpPosition =   3  'Windows-Standard"
        Print #1, "   Begin VB.PictureBox PicHiddenData "
        Print #1, "      AutoRedraw      =   -1  'True"
        Print #1, "      BorderStyle     =   0  'Kein"
        Print #1, "      Height          =   " & Trim(Str(PicBinary.Height * Screen.TwipsPerPixelY))
        Print #1, "      Left            =   0"
'Write the correct Position
'for the binary data picture
        If BackPic > 0 Then
            Print #1, "      Picture         =   """ & Fname; ".frx"":" & Hex$(BackPic + 12)
        Else
            Print #1, "      Picture         =   """ & Fname; ".frx"":0000"
        End If
        Print #1, "      ScaleHeight     =   " & Trim(Str(PicBinary.Height * Screen.TwipsPerPixelY))
        Print #1, "      ScaleWidth      =   " & Trim(Str(PicBinary.Width * Screen.TwipsPerPixelX))
        Print #1, "      TabIndex        =   0"
        Print #1, "      Top             =   -2000"
        Print #1, "      Visible         =   0   'False"
        Print #1, "      Width           =   " & Trim(Str(PicBinary.Width * Screen.TwipsPerPixelX))
        Print #1, "   End"
        Print #1, "End"
        Print #1, "Attribute VB_Name = """ & Fname & """"
        Print #1, "Attribute VB_GlobalNameSpace = False"
        Print #1, "Attribute VB_Creatable = False"
        Print #1, "Attribute VB_PredeclaredId = True"
        Print #1, "Attribute VB_Exposed = False"
        Print #1, "Option Explicit"
        Print #1, ""
        Print #1, "Private Declare Function ExtCreateRegion Lib ""gdi32"" (lpXform As Any, ByVal nCount As Long, lpRgnData As Any) As Long"
        Print #1, "Private Declare Function SetWindowRgn Lib ""user32"" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long"
        Print #1, ""
        Print #1, "'Fast binary Data"
        Print #1, "Private Declare Function GetBitmapBits Lib ""gdi32"" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long"
        Print #1, "Private Declare Function GetObject Lib ""gdi32"" Alias ""GetObjectA"" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long"
        Print #1, ""
        Print #1, "Dim PicInfo As BITMAP"
        Print #1, ""
        Print #1, "Private Type BITMAP"
        Print #1, " bmType As Long"
        Print #1, " bmWidth As Long"
        Print #1, " bmHeight As Long"
        Print #1, " bmWidthBytes As Long"
        Print #1, " bmPlanes As Integer"
        Print #1, " bmBitsPixel As Integer"
        Print #1, " bmBits As Long"
        Print #1, "End Type"
        Print #1, ""
        Print #1, "Private Sub Form_Load()"
        Print #1, "Dim Region As Long"
        Print #1, "Dim ByteCtr As Long"
        Print #1, "Dim ByteData(" & Str(ByteCtr - 1) & ") As Byte"
        Print #1, ""
        Print #1, "ByteCtr =" & Str(ByteCtr)
        Print #1, ""
        Print #1, "'Get the Data"
        Print #1, "GetObject PicHiddenData.Image, Len(PicInfo), PicInfo"
        Print #1, "GetBitmapBits PicHiddenData.Image, ByteCtr, ByteData(0) "
        Print #1, ""
        Print #1, "'Shape The Form"
        Print #1, "Region = ExtCreateRegion(ByVal 0&, ByteCtr, ByteData(0))"
        Print #1, "SetWindowRgn Me.hWnd, Region, True"
        Print #1, "End Sub"
    Close
'Remove the Temporary file we produced
    Kill App.Path & "\tempBinary.dat"
End Sub
Private Function CutExt(ByVal StringToCut As String, ByVal CutString As String)

Dim Cutlenght As Long
    Cutlenght = Len(StringToCut) - Len(CutString) + 1

    Do Until Mid$(StringToCut, Cutlenght, Len(CutString)) = CutString
        Cutlenght = Cutlenght - 1
    Loop
    CutExt = Left$(StringToCut, Cutlenght - 1)

End Function
Private Function CutAfter(ByVal StringToCut As String, ByVal CutString As String)

Dim Cutlenght As Long
    Cutlenght = Len(StringToCut) - Len(CutString) + 1

    Do Until Mid$(StringToCut, Cutlenght, Len(CutString)) = CutString
        Cutlenght = Cutlenght - 1
    Loop
    CutAfter = Right$(StringToCut, Len(StringToCut) - Cutlenght)

End Function

'Write an FRX file to store data we need
'like Background Image and binary data
'without extra file loading
Private Sub WriteFrx(PicFileName As String)

'First of all create an FRX
'to store the data we need
Dim Sum(8)      As Byte    'The 8 unknow bytes for the header
Dim S           As Byte    'Counter
Dim I           As Byte    'Counter #2
Dim bytes()     As Byte    'Needed for a Fast reading/writing of the BMP Bytes
Dim b As String
Dim e As String
Dim FPos As Long
    FPos = Seek(1) - 1

    S = 1               'Sets the Counter to 1
    PicSize = FileLen(PicFileName)      'Getting Size of the Picture
    Open PicFileName For Binary Access Read As #2 'Open Picture
        b = Hex$(PicSize + 12)    'Filesize +12
        e = Hex$(PicSize)         'Filesize
        If Int(Len(e) / 2) <> Len(e) / 2 Then e = "0" & e 'If u cant divide with 2 make it longer
        If Int(Len(b) / 2) <> Len(b) / 2 Then b = "0" & b
        For I = 1 To Len(b) Step 2 'Reverse The numbers
            Sum(S) = Val("&H" & Mid$(b, Len(b) - I, 2)) 'write the numers to Sum(?)
            S = S + 1 'increase the counter
        Next I
        S = 5
        For I = 1 To Len(e) Step 2
            Sum(S) = Val("&H" & Mid$(e, Len(e) - I, 2))
            S = S + 1
        Next I
        For I = 1 To 4 'Write the 1. part of the header (4byte)
            Put #1, I + FPos, Sum(I)
        Next I
        Put #1, 5 + FPos, &H6C 'Write the 2. part of the header (2byte)
        Put #1, 6 + FPos, &H74
        For I = 5 To 8 'Write the 3. part of the header (4byte) starting with byte #9
            Put #1, I + 4 + FPos, Sum(I)
        Next I
        ReDim bytes(PicSize - 1) 'Size of the pic
        Get #2, , bytes() 'Read Bytes
        Put #1, 13 + FPos, bytes() 'Write Bytes
    Close #2

End Sub
