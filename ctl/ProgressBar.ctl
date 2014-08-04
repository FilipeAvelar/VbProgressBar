VERSION 5.00
Begin VB.UserControl ProgressBar 
   BackColor       =   &H00FFFFFF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   1995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3015
   HitBehavior     =   2  'Use Paint
   ScaleHeight     =   1995
   ScaleWidth      =   3015
   Windowless      =   -1  'True
   Begin VB.Label lblHora 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   360
      Width           =   1000
   End
   Begin VB.Label lblCaption2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1020
      TabIndex        =   3
      Top             =   240
      Width           =   6075
   End
   Begin VB.Label lblRestante 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Calculando tempo restante..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   1000
   End
   Begin VB.Label lblPercentual 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   -60
      Width           =   1000
   End
   Begin VB.Label lblMain 
      BackStyle       =   0  'Transparent
      Caption         =   "Aguarde ..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1020
      TabIndex        =   0
      Top             =   0
      Width           =   6075
   End
   Begin VB.Shape shpMain 
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   0
      Left            =   1020
      Top             =   480
      Visible         =   0   'False
      Width           =   90
   End
End
Attribute VB_Name = "NGBarraProgresso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim mEtapas As New Etapas
'Default Property Values:
Const m_def_TempoRestante = -1
Const m_def_TempoTotal = -1
Const m_def_Value = 0
'Property Variables:
Dim m_TempoRestante As Boolean
Dim m_TempoTotal As Boolean
Dim m_Value As Long
Dim MQuadrados As Integer
Dim mMaxValue As Long
Dim mEtapaAtual As Integer
Dim Cores(10, 1) As Long
Dim mInicio As Date
Dim mLastTempo As Integer
Dim mLastPercentual As Integer
'Event Declarations:
'Event Cancelar()
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Public Sub RefreshCaption()
    
End Sub
Public Property Let Caption2(ByVal Caption As String)
    lblCaption2.Caption = Caption
End Property


Public Property Let Etapa(ByVal EtapaAtual As Variant)
    Dim i As Integer
    Dim Valor As Long
        
    For i = 1 To Etapas.Count
        If IsNumeric(EtapaAtual) Then
            If i = EtapaAtual Then Exit For
        Else
            If Etapas(EtapaAtual).Key = Etapas(i).Key Then Exit For
        End If
        Valor = Valor + Etapas(i).Tamanho
    Next
    
    Value = Valor + 1
    
End Property


Public Property Set Etapas(vEtapas As Etapas)
    Set mEtapas = vEtapas
    ctlAuxBarra1.Desenhar mEtapas
    DefineCores
End Property
Public Sub DefineCores()
    Dim i As Integer
    
    Cores(1, 0) = 12648447
    Cores(1, 1) = &HFFFF&

    Cores(2, 0) = 12640511
    Cores(2, 1) = 33023

    Cores(3, 0) = 12648384
    Cores(3, 1) = 65280
    
    Cores(4, 0) = 16777152
    Cores(4, 1) = 16776960
    
    Cores(5, 0) = 16761024
    Cores(5, 1) = 16711680
    
    Cores(6, 0) = 16761087
    Cores(6, 1) = 16711935
    
    Cores(7, 0) = 12648384
    Cores(7, 1) = 65280
    
    Cores(8, 0) = 12648384
    Cores(8, 1) = 65280
    
    Cores(9, 0) = 12648384
    Cores(9, 1) = 65280
    
    Cores(10, 0) = 12648384
    Cores(10, 1) = 65280
    
    For i = 1 To mEtapas.Count
        If mEtapas(i).CorNormal = 0 Then mEtapas(i).CorNormal = Cores(i, 0)
        If mEtapas(i).CorCompleto = 0 Then mEtapas(i).CorCompleto = Cores(i, 1)
    Next
End Sub


Public Property Get Etapas() As Etapas
    Set Etapas = mEtapas
End Property

Private Sub ctlAuxBarra1_GotFocus()

End Sub
'
'Private Sub Button1_Click()
'    RaiseEvent Cancelar
'End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
    Desenhar
    
End Sub

Public Sub Desenhar()

    Dim x As Integer
    Dim y As Integer
    Dim a As Integer
    
        For a = 1 To shpMain.UBound
            If Not shpMain(a) Is Nothing Then
                Unload shpMain(a)
            End If
        Next
    
        a = 1
    
           
    
        For y = 480 To UserControl.Height - 105 Step 105
            For x = 1020 To UserControl.Width - 105 Step 105
                Load shpMain(a)
                shpMain(a).Left = x
                shpMain(a).Top = y
                shpMain(a).Visible = True
                a = a + 1
            Next x
        Next y

        MQuadrados = shpMain.UBound

        Colorir

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get Value() As Integer
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Integer)
    If m_Value = 0 Then
        mInicio = Time
    End If
    m_Value = New_Value
    Colorir
    PropertyChanged "Value"
End Property
Private Sub Colorir()
Dim i As Integer
Dim a As Double
Dim Etapa As Integer
Dim Max As Double
Dim Valor As Double

        
    AtualizaEtapas
    
    Etapa = 1
    Max = mEtapas(1).Tamanho
    a = mMaxValue / MQuadrados
    Valor = 0
    For i = 1 To MQuadrados
        If i * a <= m_Value Then
            shpMain(i).FillColor = mEtapas(Etapa).CorCompleto
            shpMain(i).BorderColor = mEtapas(Etapa).CorCompleto
        Else
            shpMain(i).BorderColor = mEtapas(Etapa).CorNormal
            shpMain(i).FillColor = mEtapas(Etapa).CorNormal
        End If
        Valor = Valor + a
        If Valor >= Max Then
            Etapa = Etapa + 1
            If Etapa > mEtapas.Count Then Etapa = mEtapas.Count
            Max = mEtapas(Etapa).Tamanho
            Valor = 0
        End If
    Next

End Sub
Private Sub AtualizaEtapas()
Dim i As Integer
Dim a As Long
Dim Tempo As Long


    mEtapaAtual = 0
    mMaxValue = 0
    For i = 1 To mEtapas.Count
        mMaxValue = mMaxValue + mEtapas(i).Tamanho
        If m_Value <= mMaxValue And mEtapaAtual = 0 Then mEtapaAtual = i
    Next
        
        
    
    
    If mInicio <> #12:00:00 AM# Then
        a = m_Value * 100
        If mLastPercentual <> (Int(a / mMaxValue)) Then
            
            lblPercentual = Int(a / mMaxValue) & "%"
            
            If lblMain <> mEtapas(mEtapaAtual).Caption And m_Value > 0 Then
                lblMain = mEtapas(mEtapaAtual).Caption
            End If
            Tempo = DateDiff("s", mInicio, Time)
            If m_TempoTotal Then
                lblHora = Right("0" & Int(Tempo / 60), 2) & ":" & Right("0" & (Tempo Mod 60), 2)
            End If
            If m_TempoRestante Then
                Tempo = Int(Tempo * (mMaxValue - m_Value) / m_Value)
                If Tempo > 60 Then
                    Tempo = Int(Tempo / 60)
                    If Tempo < mLastTempo Or Tempo > mLastTempo + 1 Then
                        lblRestante = Tempo & " minutos restantes"
                        mLastTempo = Tempo
                    End If
                Else
                    If Tempo < mLastTempo Or Tempo > mLastTempo + 5 Then
                        lblRestante = Tempo & " segundos restantes"
                        mLastTempo = Tempo
                    End If
                End If
            End If
        End If
    Else
        lblRestante = ""
    End If
End Sub


Private Function QualEtapaheQuadrado(ByVal Index As Integer)
Dim a As Integer
Dim i As Integer


    a = Index * (mMaxValue / MQuadrados)
    
    For i = mEtapas.Count To 1 Step -1
        
        If a <= mEtapas(i).Tamanho Then QualEtapaheQuadrado = i
    
    Next
    
End Function
Private Function QualEtapaheValor(ByVal Value As Integer)
Dim i As Integer


    For i = 1 To mEtapas.Count
        If m_Value > mEtapas(i).Tamanho Then QualEtapa = i
    Next


End Function

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Value = m_def_Value
    m_TempoRestante = m_def_TempoRestante
    m_TempoTotal = m_def_TempoTotal
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 0)
    lblMain.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    lblRestante.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    lblCaption2.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    lblPercentual.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    m_TempoRestante = PropBag.ReadProperty("TempoRestante", m_def_TempoRestante)
    m_TempoTotal = PropBag.ReadProperty("TempoTotal", m_def_TempoTotal)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &HFFFFFF)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 0)
    Call PropBag.WriteProperty("ForeColor", lblMain.ForeColor, &H80000012)
    Call PropBag.WriteProperty("TempoRestante", m_TempoRestante, m_def_TempoRestante)
    Call PropBag.WriteProperty("TempoTotal", m_TempoTotal, m_def_TempoTotal)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblMain,lblMain,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = lblMain.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    lblMain.ForeColor() = New_ForeColor
    lblCaption2.ForeColor() = New_ForeColor
    lblRestante.ForeColor() = New_ForeColor
    lblPercentual.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,-1
Public Property Get TempoRestante() As Boolean
    TempoRestante = m_TempoRestante
End Property

Public Property Let TempoRestante(ByVal New_TempoRestante As Boolean)
    m_TempoRestante = New_TempoRestante
    PropertyChanged "TempoRestante"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,-1
Public Property Get TempoTotal() As Boolean
    TempoTotal = m_TempoTotal
End Property

Public Property Let TempoTotal(ByVal New_TempoTotal As Boolean)
    m_TempoTotal = New_TempoTotal
    PropertyChanged "TempoTotal"
End Property


