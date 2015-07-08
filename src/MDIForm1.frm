VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H80000004&
   Caption         =   "Ronda por Enrique Somolinos Pérez"
   ClientHeight    =   8190
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11280
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   NegotiateToolbars=   0   'False
   ScrollBars      =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu m 
      Caption         =   "Menu"
      Begin VB.Menu jn 
         Caption         =   "Juego nuevo"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu op 
      Caption         =   "Opciones"
      Begin VB.Menu qe 
         Caption         =   "Quién empieza"
         Begin VB.Menu ord 
            Caption         =   "Ordenador"
         End
         Begin VB.Menu usu 
            Caption         =   "Usuario"
            Checked         =   -1  'True
         End
      End
   End
   Begin VB.Menu ay 
      Caption         =   "Ayuda"
      Begin VB.Menu ad 
         Caption         =   "Acerca de"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ad_Click()
Load frmAbout
frmAbout.Show
End Sub

Private Sub jn_Click()
Dim c As Class1
Set c = New Class1
nueva = 1
Call c.iniciar
'Unload Form1
'If usu.Checked Then turno = 0 Else turno = 1'

'Load Form1
'Form1.Top = 0
'Form1.Left = 0

End Sub

Private Sub MDIForm_Load()
App.Title = "Ronda"

If App.PrevInstance = True Then End

Load Form1
End Sub


Private Sub MDIForm_Terminate()
Unload Form1

End Sub

Private Sub ord_Click()
ord.Checked = True
usu.Checked = False
End Sub

Private Sub salir_Click()
End
End Sub

Private Sub usu_Click()
ord.Checked = False
usu.Checked = True

End Sub
