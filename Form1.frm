VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3075
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   3075
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picChecked 
      Height          =   285
      Left            =   2730
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   2
      Top             =   210
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picUnchecked 
      Height          =   285
      Left            =   2730
      Picture         =   "Form1.frx":0342
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   1
      Top             =   570
      Visible         =   0   'False
      Width           =   285
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2865
      Left            =   150
      TabIndex        =   0
      Top             =   180
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   5054
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strChecked As String

Private Sub Form_Load()
Dim i As Variant, ms_rows As Integer
    ' Start building the Grid

ms_rows = 20 ' This is the number of rows to print out
With MSFlexGrid1
        .Row = 0
        .Col = 0
        .Rows = ms_rows + 1  'We add 1 to ensure we get all the rows
        .Cols = 2
        .ColWidth(0) = 250 ' CheckBox column
        .ColWidth(1) = 1440 ' Index column
End With

' Now build the Grid

For i = 0 To 20 'm_rows - 1
     With MSFlexGrid1
          .Row = i: .Col = 0: .CellPictureAlignment = 4 ' Align the checkbox
          Set .CellPicture = picUnchecked.Picture  ' Set the default checkbox picture to the empty box
          .TextMatrix(i, 1) = i
     End With
Next
End Sub

Private Sub MSFlexGrid1_Click()
Dim oldx, oldy, cell2text As String, strTextCheck As String

' Check or uncheck the grid checkbox
With MSFlexGrid1
    oldx = .Col
    oldy = .Row
        If MSFlexGrid1.Col = 0 Then
            If MSFlexGrid1.CellPicture = picChecked Then
                Set MSFlexGrid1.CellPicture = picUnchecked
                .Col = .Col + 1  ' I use data that is in column #1, usually an Index or ID #
                strTextCheck = .Text
                ' When you de-select a CheckBox, we need to strip out the #
                strChecked = Replace(strChecked, strTextCheck & ",", "")
                ' Don't forget to strip off the trailing , before passing the string
                Debug.Print strChecked
            Else
                Set MSFlexGrid1.CellPicture = picChecked
                .Col = .Col + 1
                strTextCheck = .Text
                strChecked = strChecked & strTextCheck & ","
                Debug.Print strChecked
            End If
        End If
    .Col = oldx
    .Row = oldy
End With
End Sub
