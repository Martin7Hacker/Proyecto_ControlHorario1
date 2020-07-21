VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPrograma 
   Caption         =   "Control Horario"
   ClientHeight    =   2715
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6240
   Icon            =   "frmPrograma.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2715
   ScaleWidth      =   6240
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdcalcular 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Calcular"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CommandButton cmdcalculoX 
      BackColor       =   &H00C0E0FF&
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1320
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker dtp1 
      Height          =   540
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   953
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   16777215
      Format          =   3014658
      UpDown          =   -1  'True
      CurrentDate     =   0.805555555555556
   End
   Begin MSComCtl2.DTPicker dtp2 
      Height          =   540
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   953
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   16777215
      Format          =   3014658
      UpDown          =   -1  'True
      CurrentDate     =   0.805555555555556
   End
   Begin VB.Label lblnegat 
      AutoSize        =   -1  'True
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3480
      TabIndex        =   6
      Top             =   600
      Width           =   90
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   5400
      X2              =   3480
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label lblres 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "---:---:---"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3960
      TabIndex        =   5
      Top             =   840
      Width           =   990
   End
   Begin VB.Label lblhe 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "---:---:---"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3960
      TabIndex        =   4
      Top             =   480
      Width           =   990
   End
   Begin VB.Label lbl2 
      BackStyle       =   0  'Transparent
      Caption         =   "Hora de Salida del Personal:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   2565
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hora de Entrada del Personal:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   2580
   End
End
Attribute VB_Name = "frmPrograma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' programa de control de horario de cantidad de horas realizadas
' autor: martin grasso Castrillo 10 / 07 / 2017
' código finalizado a las 3:17:07 hs
'
'// constantes de filtro de armado de formato basico.
Const doblecero As String = "00"
Const doblepunto As String = ":"
Const fhora As String = " hs."


' variables de Vectores numericos
Dim h(1) As Byte
Dim m(1) As Byte
Dim s(1) As Byte

'variable auxiliar en el caso de que cualquiera de los valores de _
 los vectores sea corespondiente al 0 o a la variable del vector,
Dim Auxiliar(7) As String

Private Sub ingresar_digitos()
Dim rx As Byte
  For rx = 0 To 1 '///////////////////////////////*
      h(rx) = 0
      m(rx) = 0
      s(rx) = 0
      'define los caracteres de separación de h:m:s
  Next rx
   ' escribir en blanco los datos de suma */
   
  lblhe.Caption = doblecero & doblepunto & doblecero _
                  & doblepunto & doblecero
                  
  ' se iguala a la resta para no repetir caracteres /*
  lblres.Caption = lblhe.Caption
  
End Sub

Private Sub cmdcalcular_Click()
dtp1_Change: dtp2_Change
'! se produce una resta de vectores
On Error GoTo nose
 'cmdcalculoX.Caption = h(1) - h(0)
 cmdcalculoX.Caption = Format(TimeValue(dtp1.Value) - TimeValue(dtp2.Value), "hh:mm:ss") & fhora
nose:
End Sub

Private Sub cmdcalculoX_Click()
MsgBox "cantidad de horas trabajadas: " & cmdcalculoX.Caption, vbInformation, "Control Horario"
End Sub

'// se aplica en el cambio de registro de cada control del Evento Cambio /*
Private Sub dtp1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift _
  As Integer, ByVal CallbackField As String, CallbackDate As Date)
  dtp1_Change
End Sub

Private Sub dtp1_Change()
Ingresar_hora_minutos_segundos
End Sub

'// se aplica en el cambio de registro de cada control del Evento Cambio /*
Private Sub dtp2_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift _
As Integer, ByVal CallbackField As String, CallbackDate As Date)
dtp2_Change
End Sub

Private Sub dtp2_Change()
Ingresar_hora_minutos_segundos
End Sub

Private Sub Form_Load()
ingresar_digitos
'hora de entrada
dtp1.Value = Time

'hms a 00:00:00
With dtp2
     .Hour = 0
     .Minute = 0
     .Second = 0
End With

End Sub

Private Sub Ingresar_hora_minutos_segundos()
'/ en el control hora de entrada /
With dtp1
  h(0) = .Hour
  m(0) = .Minute
  s(0) = .Second
 End With
'/ en el control hora de salida /
  With dtp2
  h(1) = .Hour
  m(1) = .Minute
  s(1) = .Second
 End With
          'si existe posibilidad de 0.
    If h(0) = "00" Then
       Auxiliar(0) = doblecero
       Else
       Auxiliar(0) = h(0)
    End If
    If h(1) = "00" Then
       Auxiliar(1) = doblecero
       Else
       Auxiliar(1) = h(1)
    End If
    If m(0) = "00" Then
       Auxiliar(2) = doblecero
       Else
       Auxiliar(2) = m(0)
    End If
    If m(1) = "00" Then
       Auxiliar(3) = doblecero
       Else
       Auxiliar(3) = m(1)
    End If
    If s(0) = "00" Then
       Auxiliar(4) = doblecero
       Else
       Auxiliar(4) = s(0)
    End If
    If s(1) = "00" Then
       Auxiliar(5) = doblecero
       Else
       Auxiliar(5) = s(1)
    End If
 lblhe.Caption = (Auxiliar(0) & doblepunto & Auxiliar(2) & doblepunto & Auxiliar(4))
 lblres.Caption = (Auxiliar(1) & doblepunto & Auxiliar(3) & doblepunto & Auxiliar(5))
 
End Sub
