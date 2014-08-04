VERSION 5.00
Begin VB.UserControl ctlAuxBarra 
   ClientHeight    =   4080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5100
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   4080
   ScaleWidth      =   5100
   Begin VB.Shape shpMain 
      Height          =   90
      Index           =   0
      Left            =   1920
      Top             =   1320
      Visible         =   0   'False
      Width           =   90
   End
End
Attribute VB_Name = "ctlAuxBarra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private Sub UserControl_Click()
    Desenhar
End Sub

Private Sub UserControl_Initialize()

End Sub
Public Sub Desenhar(ByRef mEtapas As Etapas)

    Dim x As Integer
    Dim y As Integer
    Dim a As Integer
    
        For a = 1 To shpMain.UBound
            If Not shpMain(a) Is Nothing Then
                Unload shpMain(a)
            End If
        Next
    
        a = 1
    
        
    
        For y = 0 To UserControl.Height Step 105
            For x = 0 To UserControl.Width Step 105
                Load shpMain(a)
                shpMain(a).Left = x
                shpMain(a).Top = y
                shpMain(a).Visible = True
                a = a + 1
            Next x
        Next y


End Sub

Private Sub UserControl_Resize()
End Sub
