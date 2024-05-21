VERSION 5.00
Begin VB.MDIForm frmIntegrador 
   BackColor       =   &H8000000C&
   Caption         =   "Piasentin y Soto - Integrador"
   ClientHeight    =   7995
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   14565
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mRegInfoCV 
      Caption         =   "Regimen Información Compras y Ventas"
      Begin VB.Menu mRegInfoCVSH 
         Caption         =   "Piasentin y Soto - SH"
      End
      Begin VB.Menu mRegInfoCVSRL 
         Caption         =   "Piasentin y Soto - SRL"
      End
   End
   Begin VB.Menu mFin 
      Caption         =   "Fin"
   End
End
Attribute VB_Name = "frmIntegrador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
Dim ctlCon As New clsCtlConnect

    ctlCon.configure_dbmdzsh
    ctlCon.configure_dbmdzsrl
    ctlCon.configure_dbuaqsrl
    
End Sub

Private Sub MDIForm_Terminate()

    mFin_Click
    
End Sub

Private Sub mFin_Click()

    dbmdzsh.closeDB
    dbmdzsrl.closeDB
    dbuaqsrl.closeDB
    
    End
    
End Sub

Private Sub mRegInfoCVSH_Click()

    frmRegInfoCVSH.Show
    
End Sub

Private Sub mRegInfoCVSRL_Click()

    frmRegInfoCVSRL.Show
    
End Sub
