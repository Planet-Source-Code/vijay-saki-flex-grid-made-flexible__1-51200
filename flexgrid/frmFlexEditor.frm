VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmFlexEditor 
   Caption         =   "Flex Grid Made Flexible"
   ClientHeight    =   1275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1275
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtEditor 
      Height          =   495
      Left            =   3015
      TabIndex        =   1
      Top             =   3450
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexEditor 
      Height          =   1245
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   4860
      _ExtentX        =   8573
      _ExtentY        =   2196
      _Version        =   393216
      Rows            =   5
      Cols            =   5
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   -2147483632
      ForeColor       =   -2147483643
      Appearance      =   0
   End
End
Attribute VB_Name = "frmFlexEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    ' Setting The Height and Width of The TextBox to The Size of the Cells of the Flex Grid
    txtEditor.Height = MSFlexEditor.CellHeight
    txtEditor.Width = MSFlexEditor.CellWidth
    
    'Function for Setting the Left and Top of the TextBox to Tune of the Col and Row of the FlexGrid
    Set_Flex_TextBox_Pos
End Sub

Private Sub MSFlexEditor_Click()
    'Function for Setting the Left and Top of the TextBox to Tune of the Col and Row of the FlexGrid
    Set_Flex_TextBox_Pos
    
    txtEditor.SetFocus
End Sub

Private Sub txtEditor_Change()
    MSFlexEditor.Text = txtEditor.Text 'Setting the Text to the Active Cell
End Sub

Private Sub txtEditor_KeyDown(KeyCode As Integer, Shift As Integer)
    'Moving of the TextBox to the Tune of the Arrow Keys
    If KeyCode = 37 Then
        If Not MSFlexEditor.Col = 0 Then
            MSFlexEditor.Col = MSFlexEditor.Col - 1
            Set_Flex_TextBox_Pos
        End If
    End If
    If KeyCode = 38 Then
        If Not MSFlexEditor.Row = 0 Then
            MSFlexEditor.Row = MSFlexEditor.Row - 1
            Set_Flex_TextBox_Pos
        End If
    End If
    If KeyCode = 39 Then
        If Not MSFlexEditor.Col = MSFlexEditor.Cols - 1 Then
            MSFlexEditor.Col = MSFlexEditor.Col + 1
            Set_Flex_TextBox_Pos
        End If
    End If
    If KeyCode = 40 Then
        If Not MSFlexEditor.Row = MSFlexEditor.Rows - 1 Then
            MSFlexEditor.Row = MSFlexEditor.Row + 1
            Set_Flex_TextBox_Pos
        End If
    End If
    txtEditor.SetFocus
End Sub
Private Sub Set_Flex_TextBox_Pos()
    'Setting Text Positions
    txtEditor.Left = MSFlexEditor.CellLeft
    txtEditor.Top = MSFlexEditor.CellTop
    txtEditor.Text = MSFlexEditor.Text
End Sub
