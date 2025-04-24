VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "TOP G Vspeed Calculator"
   ClientHeight    =   5385
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9960
   Icon            =   "Vspeed.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   9960
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox Picture1 
      Height          =   1215
      Left            =   4920
      Picture         =   "Vspeed.frx":2F50D2
      ScaleHeight     =   1155
      ScaleWidth      =   4875
      TabIndex        =   23
      Top             =   550
      Width           =   4935
   End
   Begin VB.ComboBox rnw_cond 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   3770
      Width           =   1215
   End
   Begin VB.CommandButton calculate 
      Caption         =   "CALCULATE"
      BeginProperty Font 
         Name            =   "AirbusDisp"
         Size            =   17.25
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   120
      TabIndex        =   21
      Top             =   4400
      Width           =   4575
   End
   Begin VB.TextBox output 
      BeginProperty Font 
         Name            =   "AirbusMCDUa"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   4920
      MultiLine       =   -1  'True
      TabIndex        =   20
      Top             =   1920
      Width           =   4935
   End
   Begin VB.ComboBox flap_select 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox weight_box 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2040
      TabIndex        =   14
      Top             =   2550
      Width           =   1215
   End
   Begin VB.TextBox elev_box 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2040
      TabIndex        =   12
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox press_box 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1620
      TabIndex        =   10
      Top             =   1080
      Width           =   1355
   End
   Begin VB.OptionButton inhg 
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   1080
      Width           =   495
   End
   Begin VB.OptionButton hpa 
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   1080
      Width           =   375
   End
   Begin VB.ComboBox temp_format 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "Vspeed.frx":2FE289
      Left            =   3000
      List            =   "Vspeed.frx":2FE28B
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox temp_box 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1620
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.ComboBox aircraft_selection 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2760
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   0
      Width           =   7215
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   4800
      X2              =   4800
      Y1              =   4680
      Y2              =   5400
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   4800
      X2              =   9960
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   4800
      X2              =   4800
      Y1              =   480
      Y2              =   4680
   End
   Begin VB.Label Label11 
      Caption         =   "Runway Condition"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label Label10 
      Caption         =   "Flap Setting"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   140
      TabIndex        =   18
      Top             =   3170
      Width           =   1575
   End
   Begin VB.Label Label9 
      Caption         =   "Kg"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   16
      Top             =   2580
      Width           =   495
   End
   Begin VB.Label Label8 
      Caption         =   "Takeoff Weight"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   15
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label7 
      Caption         =   "ft"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   13
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "Baro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   550
      TabIndex        =   11
      Top             =   1155
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Airport Elevation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1990
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "inhg"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "hpa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Temperature"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   3
      Top             =   680
      Width           =   1350
   End
   Begin VB.Label Label1 
      Caption         =   "SELECT AIRCRAFT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim temp As Integer
Dim pressure_alt As Double
Dim V1 As Integer
Dim V2 As Integer
Dim Vr As Integer
Dim flap As Integer
Dim weight As Long
Dim wet_rnw As Integer
Dim no_speed As Integer
Dim pressure As Double
Dim elevation As Integer
Dim press_format_output As String

Function FahrenheitToCelsius(ByVal fahrenheit As Double) As Double
FahrenheitToCelsius = (fahrenheit - 32) * (5 / 9)
End Function
Function b738()
If wet_rnw = 0 Then
    Select Case flap
        Case 1
        If weight >= 40000 And weight <= 42500 Then
        V1 = 105
        Vr = 106
        V2 = 125
        ElseIf weight > 42500 And weight <= 47500 Then
        V1 = 113
        Vr = 114
        V2 = 131
        ElseIf weight > 47500 And weight <= 52500 Then
        V1 = 121
        Vr = 122
        V2 = 137
        ElseIf weight > 52500 And weight <= 57500 Then
        V1 = 128
        Vr = 129
        V2 = 143
        ElseIf weight > 57500 And weight <= 62500 Then
        V1 = 135
        Vr = 136
        V2 = 148
        ElseIf weight > 62500 And weight <= 67500 Then
        V1 = 141
        Vr = 143
        V2 = 153
        ElseIf weight > 67500 And weight <= 72500 Then
        V1 = 147
        Vr = 149
        V2 = 158
        ElseIf weight > 72500 And weight <= 77500 Then
        V1 = 153
        Vr = 155
        V2 = 162
        ElseIf weight > 77500 And weight <= 82500 Then
        V1 = 158
        Vr = 160
        V2 = 167
        ElseIf weight > 82500 And weight <= 87500 Then
        V1 = 163
        Vr = 166
        V2 = 171
        ElseIf weight > 87500 And weight <= 90000 Then
        V1 = 169
        Vr = 171
        V2 = 175
        Else
        no_speed = 1
        End If
        
        '-------------------------------------------------------------------------------------------
        
        Case 5
        If weight >= 40000 And weight <= 42500 Then
        V1 = 101
        Vr = 102
        V2 = 120
        ElseIf weight > 42500 And weight <= 47500 Then
        V1 = 109
        Vr = 110
        V2 = 126
        ElseIf weight > 47500 And weight <= 52500 Then
        V1 = 116
        Vr = 117
        V2 = 132
        ElseIf weight > 52500 And weight <= 57500 Then
        V1 = 123
        Vr = 124
        V2 = 137
        ElseIf weight > 57500 And weight <= 62500 Then
        V1 = 129
        Vr = 131
        V2 = 143
        ElseIf weight > 62500 And weight <= 67500 Then
        V1 = 135
        Vr = 137
        V2 = 147
        ElseIf weight > 67500 And weight <= 72500 Then
        V1 = 141
        Vr = 143
        V2 = 152
        ElseIf weight > 72500 And weight <= 77500 Then
        V1 = 147
        Vr = 148
        V2 = 156
        ElseIf weight > 77500 And weight <= 82500 Then
        V1 = 152
        Vr = 154
        V2 = 160
        ElseIf weight > 82500 And weight <= 87500 Then
        V1 = 157
        Vr = 159
        V2 = 164
        ElseIf weight > 87500 And weight <= 90000 Then
        V1 = 161
        Vr = 163
        V2 = 168
        Else
        no_speed = 1
        End If
        
        '---------------------------------------------------------------------------------------------------
        
        Case 10
        If weight >= 40000 And weight <= 42500 Then
        V1 = 100
        Vr = 101
        V2 = 119
        ElseIf weight > 42500 And weight <= 47500 Then
        V1 = 108
        Vr = 108
        V2 = 125
        ElseIf weight > 47500 And weight <= 52500 Then
        V1 = 115
        Vr = 116
        V2 = 130
        ElseIf weight > 52500 And weight <= 57500 Then
        V1 = 122
        Vr = 123
        V2 = 136
        ElseIf weight > 57500 And weight <= 62500 Then
        V1 = 128
        Vr = 129
        V2 = 141
        ElseIf weight > 62500 And weight <= 67500 Then
        V1 = 134
        Vr = 136
        V2 = 146
        ElseIf weight > 67500 And weight <= 72500 Then
        V1 = 140
        Vr = 141
        V2 = 150
        ElseIf weight > 72500 And weight <= 77500 Then
        V1 = 146
        Vr = 147
        V2 = 154
        ElseIf weight > 77500 And weight <= 82500 Then
        V1 = 151
        Vr = 152
        V2 = 158
        ElseIf weight > 82500 And weight <= 87500 Then
        V1 = 156
        Vr = 157
        V2 = 162
        Else
        no_speed = 1
        End If
        
        '---------------------------------------------------------------------------------------------------
        
        Case 15
        If weight >= 40000 And weight <= 42500 Then
        V1 = 98
        Vr = 99
        V2 = 117
        ElseIf weight > 42500 And weight <= 47500 Then
        V1 = 105
        Vr = 106
        V2 = 122
        ElseIf weight > 47500 And weight <= 52500 Then
        V1 = 112
        Vr = 113
        V2 = 128
        ElseIf weight > 52500 And weight <= 57500 Then
        V1 = 119
        Vr = 120
        V2 = 133
        ElseIf weight > 57500 And weight <= 62500 Then
        V1 = 125
        Vr = 126
        V2 = 138
        ElseIf weight > 62500 And weight <= 67500 Then
        V1 = 131
        Vr = 133
        V2 = 143
        ElseIf weight > 67500 And weight <= 72500 Then
        V1 = 137
        Vr = 138
        V2 = 147
        ElseIf weight > 72500 And weight <= 77500 Then
        V1 = 142
        Vr = 144
        V2 = 151
        ElseIf weight > 77500 And weight <= 82500 Then
        V1 = 148
        Vr = 149
        V2 = 155
        Else
        no_speed = 1
        End If
        
        '-------------------------------------------------------------------------------------------------
        Case 25
        If weight >= 40000 And weight <= 42500 Then
        V1 = 96
        Vr = 97
        V2 = 115
        ElseIf weight > 42500 And weight <= 47500 Then
        V1 = 103
        Vr = 104
        V2 = 120
        ElseIf weight > 47500 And weight <= 52500 Then
        V1 = 110
        Vr = 111
        V2 = 126
        ElseIf weight > 52500 And weight <= 57500 Then
        V1 = 117
        Vr = 118
        V2 = 131
        ElseIf weight > 57500 And weight <= 62500 Then
        V1 = 123
        Vr = 124
        V2 = 136
        ElseIf weight > 62500 And weight <= 67500 Then
        V1 = 129
        Vr = 130
        V2 = 140
        ElseIf weight > 67500 And weight <= 72500 Then
        V1 = 135
        Vr = 136
        V2 = 145
        ElseIf weight > 72500 And weight <= 77500 Then
        V1 = 140
        Vr = 141
        V2 = 149
        ElseIf weight > 77500 And weight <= 82500 Then
        V1 = 145
        Vr = 146
        V2 = 153
        Else
        no_speed = 1
        End If
        Case Else
        MsgBox ("INVALID FLAP SETTING")
        no_speed = 1
    End Select
    If no_speed = 1 Then
    MsgBox ("NO SPEEDS")
    Else
    dry_rnw_correction
    End If
    '--------------------------------------------------------
    
    
ElseIf wet_rnw = 1 Then
    Select Case flap
        Case 1
        If weight >= 40000 And weight <= 42500 Then
        V1 = 96
        Vr = 106
        V2 = 125
        ElseIf weight > 42500 And weight <= 47500 Then
        V1 = 104
        Vr = 114
        V2 = 131
        ElseIf weight > 47500 And weight <= 52500 Then
        V1 = 111
        Vr = 122
        V2 = 137
        ElseIf weight > 52500 And weight <= 57500 Then
        V1 = 119
        Vr = 129
        V2 = 143
        ElseIf weight > 57500 And weight <= 62500 Then
        V1 = 126
        Vr = 136
        V2 = 148
        ElseIf weight > 62500 And weight <= 67500 Then
        V1 = 133
        Vr = 143
        V2 = 153
        ElseIf weight > 67500 And weight <= 72500 Then
        V1 = 139
        Vr = 149
        V2 = 158
        ElseIf weight > 72500 And weight <= 77500 Then
        V1 = 145
        Vr = 155
        V2 = 162
        ElseIf weight > 77500 And weight <= 82500 Then
        V1 = 151
        Vr = 160
        V2 = 167
        ElseIf weight > 82500 And weight <= 87500 Then
        V1 = 157
        Vr = 166
        V2 = 171
        ElseIf weight > 87500 And weight <= 90000 Then
        V1 = 164
        Vr = 171
        V2 = 175
        Else
        no_speed = 1
        End If
        
        '-------------------------------------------------------------------------------------------
        
        Case 5
        If weight >= 40000 And weight <= 42500 Then
        V1 = 92
        Vr = 102
        V2 = 120
        ElseIf weight > 42500 And weight <= 47500 Then
        V1 = 99
        Vr = 110
        V2 = 126
        ElseIf weight > 47500 And weight <= 52500 Then
        V1 = 107
        Vr = 117
        V2 = 132
        ElseIf weight > 52500 And weight <= 57500 Then
        V1 = 114
        Vr = 124
        V2 = 137
        ElseIf weight > 57500 And weight <= 62500 Then
        V1 = 121
        Vr = 131
        V2 = 143
        ElseIf weight > 62500 And weight <= 67500 Then
        V1 = 127
        Vr = 137
        V2 = 148
        ElseIf weight > 67500 And weight <= 72500 Then
        V1 = 133
        Vr = 143
        V2 = 152
        ElseIf weight > 72500 And weight <= 77500 Then
        V1 = 139
        Vr = 148
        V2 = 156
        ElseIf weight > 77500 And weight <= 82500 Then
        V1 = 145
        Vr = 154
        V2 = 160
        ElseIf weight > 82500 And weight <= 87500 Then
        V1 = 150
        Vr = 159
        V2 = 164
        ElseIf weight > 87500 And weight <= 90000 Then
        V1 = 156
        Vr = 164
        V2 = 168
        Else
        no_speed = 1
        End If
        
        '---------------------------------------------------------------------------------------------------
        
        Case 10
        If weight >= 40000 And weight <= 42500 Then
        V1 = 91
        Vr = 101
        V2 = 119
        ElseIf weight > 42500 And weight <= 47500 Then
        V1 = 99
        Vr = 108
        V2 = 125
        ElseIf weight > 47500 And weight <= 52500 Then
        V1 = 106
        Vr = 116
        V2 = 130
        ElseIf weight > 52500 And weight <= 57500 Then
        V1 = 113
        Vr = 123
        V2 = 136
        ElseIf weight > 57500 And weight <= 62500 Then
        V1 = 120
        Vr = 129
        V2 = 141
        ElseIf weight > 62500 And weight <= 67500 Then
        V1 = 127
        Vr = 136
        V2 = 146
        ElseIf weight > 67500 And weight <= 72500 Then
        V1 = 133
        Vr = 141
        V2 = 150
        ElseIf weight > 72500 And weight <= 77500 Then
        V1 = 139
        Vr = 147
        V2 = 154
        ElseIf weight > 77500 And weight <= 82500 Then
        V1 = 145
        Vr = 152
        V2 = 158
        ElseIf weight > 82500 And weight <= 87500 Then
        V1 = 151
        Vr = 157
        V2 = 162
        Else
        no_speed = 1
        End If
        
        '---------------------------------------------------------------------------------------------------
        
        Case 15
        If weight >= 40000 And weight <= 42500 Then
        V1 = 89
        Vr = 99
        V2 = 117
        ElseIf weight > 42500 And weight <= 47500 Then
        V1 = 96
        Vr = 106
        V2 = 122
        ElseIf weight > 47500 And weight <= 52500 Then
        V1 = 104
        Vr = 113
        V2 = 128
        ElseIf weight > 52500 And weight <= 57500 Then
        V1 = 111
        Vr = 120
        V2 = 133
        ElseIf weight > 57500 And weight <= 62500 Then
        V1 = 117
        Vr = 126
        V2 = 138
        ElseIf weight > 62500 And weight <= 67500 Then
        V1 = 124
        Vr = 133
        V2 = 143
        ElseIf weight > 67500 And weight <= 72500 Then
        V1 = 130
        Vr = 138
        V2 = 147
        ElseIf weight > 72500 And weight <= 77500 Then
        V1 = 136
        Vr = 144
        V2 = 151
        ElseIf weight > 77500 And weight <= 82500 Then
        V1 = 141
        Vr = 149
        V2 = 155
        Else
        no_speed = 1
        End If
        
        '-------------------------------------------------------------------------------------------------
        Case 25
        If weight >= 40000 And weight <= 42500 Then
        V1 = 87
        Vr = 97
        V2 = 115
        ElseIf weight > 42500 And weight <= 47500 Then
        V1 = 95
        Vr = 104
        V2 = 120
        ElseIf weight > 47500 And weight <= 52500 Then
        V1 = 102
        Vr = 111
        V2 = 126
        ElseIf weight > 52500 And weight <= 57500 Then
        V1 = 109
        Vr = 118
        V2 = 131
        ElseIf weight > 57500 And weight <= 62500 Then
        V1 = 115
        Vr = 124
        V2 = 136
        ElseIf weight > 62500 And weight <= 67500 Then
        V1 = 122
        Vr = 130
        V2 = 140
        ElseIf weight > 67500 And weight <= 72500 Then
        V1 = 128
        Vr = 136
        V2 = 145
        ElseIf weight > 72500 And weight <= 77500 Then
        V1 = 134
        Vr = 141
        V2 = 149
        ElseIf weight > 77500 And weight <= 82500 Then
        V1 = 140
        Vr = 146
        V2 = 153
        Else
        no_speed = 1
        End If
        Case Else
        MsgBox ("INVALID FLAP SETTING")
        no_speed = 1
    End Select
    If no_speed = 1 Then
    MsgBox ("NO SPEED")
    Else
    wet_rnw_correction
    End If
Else
MsgBox ("INVALID RUNWAY CONDITION")
End If
End Function

Function dry_rnw_correction()
If temp <= 25 Then
    If pressure_alt >= 1000 And pressure_alt <= 3000 Then
    V1 = V1 + 1
    Vr = Vr + 1
    V2 = V2 - 1
    ElseIf pressure_alt > 3000 And pressure_alt <= 5000 Then
    V1 = V1 + 2
    Vr = Vr + 2
    V2 = V2 - 1
    ElseIf pressure_alt > 5000 And pressure_alt <= 7000 Then
    V1 = V1 + 3
    Vr = Vr + 3
    V2 = V2 - 2
    ElseIf pressure_alt > 7000 And pressure_alt <= 9000 Then
    V1 = V1 + 4
    Vr = Vr + 4
    V2 = V2 - 3
    ElseIf pressure_alt > 9000 Then
    V1 = V1 + 5
    Vr = Vr + 5
    V2 = V2 - 3
    End If
    
'------------------------------------------------------------

ElseIf temp > 25 And temp <= 35 Then
    If pressure_alt >= 1000 And pressure_alt <= 3000 Then
    V1 = V1 + 1
    Vr = Vr + 1
    V2 = V2 - 1
    ElseIf pressure_alt > 3000 And pressure_alt <= 5000 Then
    V1 = V1 + 3
    Vr = Vr + 3
    V2 = V2 - 2
    ElseIf pressure_alt > 5000 And pressure_alt <= 7000 Then
    V1 = V1 + 4
    Vr = Vr + 4
    V2 = V2 - 2
    ElseIf pressure_alt > 7000 And pressure_alt <= 9000 Then
    V1 = V1 + 6
    Vr = Vr + 5
    V2 = V2 - 3
    ElseIf pressure_alt > 9000 Then
    V1 = V1 + 7
    Vr = Vr + 6
    V2 = V2 - 4
    End If
'------------------------------------------------------------

ElseIf temp > 35 And temp <= 45 Then
    If pressure_alt >= (-2000) And pressure_alt <= (-1000) Then
    V1 = V1 + 1
    Vr = Vr + 1
    V2 = V2 - 1
    ElseIf pressure_alt > (-1000) And pressure_alt <= 1000 Then
    V1 = V1 + 1
    Vr = Vr + 1
    V2 = V2 - 1
    ElseIf pressure_alt > 1000 And pressure_alt <= 3000 Then
    V1 = V1 + 3
    Vr = Vr + 3
    V2 = V2 - 2
    ElseIf pressure_alt > 3000 And pressure_alt <= 5000 Then
    V1 = V1 + 4
    Vr = Vr + 4
    V2 = V2 - 2
    ElseIf pressure_alt > 5000 And pressure_alt <= 7000 Then
    V1 = V1 + 6
    Vr = Vr + 5
    V2 = V2 - 3
    ElseIf pressure_alt > 7000 And pressure_alt <= 9000 Then
    V1 = V1 + 7
    Vr = Vr + 6
    V2 = V2 - 4
    ElseIf pressure_alt > 9000 Then
    V1 = V1 + 9
    Vr = Vr + 7
    V2 = V2 - 5
    End If
'---------------------------------------------------------------

ElseIf temp > 45 And temp <= 55 Then
    If pressure_alt >= (-2000) And pressure_alt <= (-1000) Then
    V1 = V1 + 2
    Vr = Vr + 2
    V2 = V2 - 2
    ElseIf pressure_alt > (-1000) And pressure_alt <= 1000 Then
    V1 = V1 + 3
    Vr = Vr + 3
    V2 = V2 - 2
    ElseIf pressure_alt > 1000 And pressure_alt <= 3000 Then
    V1 = V1 + 4
    Vr = Vr + 4
    V2 = V2 - 3
    ElseIf pressure_alt > 3000 And pressure_alt <= 5000 Then
    V1 = V1 + 5
    Vr = Vr + 5
    V2 = V2 - 3
    ElseIf pressure_alt > 5000 And pressure_alt <= 7000 Then
    V1 = V1 + 6
    Vr = Vr + 6
    V2 = V2 - 4
    ElseIf pressure_alt > 7000 And pressure_alt <= 9000 Then
    V1 = V1 + 7
    Vr = Vr + 7
    V2 = V2 - 5
    ElseIf pressure_alt > 9000 Then
    V1 = V1 + 9
    Vr = Vr + 8
    V2 = V2 - 6
    End If
    
'-------------------------------------------------------------

ElseIf temp > 55 And temp <= 65 Then
    If pressure_alt >= (-2000) And pressure_alt <= (-1000) Then
    V1 = V1 + 4
    Vr = Vr + 3
    V2 = V2 - 2
    ElseIf pressure_alt > (-1000) And pressure_alt <= 1000 Then
    V1 = V1 + 5
    Vr = Vr + 4
    V2 = V2 - 3
    ElseIf pressure_alt > 1000 And pressure_alt <= 3000 Then
    V1 = V1 + 6
    Vr = Vr + 5
    V2 = V2 - 3
    ElseIf pressure_alt > 3000 And pressure_alt <= 5000 Then
    V1 = V1 + 7
    Vr = Vr + 6
    V2 = V2 - 4
    End If
End If
End Function


Function pressure_altitude(pressure_param As Double, elevation_param As Integer) As Double
Dim pressure_alt As Double
pressure_alt_param = (29.92 - pressure_param) * 1000 + elevation_param
pressure_altitude = pressure_alt_param
End Function

Function pressure_convert(X As Double) As Double
pressure_convert = X / 33.86
End Function



Function round(var As Double) As Double
Dim value As Double
value = CInt(var * 100 + 0.5)
round = CDbl(value / 100)
End Function

Function wet_rnw_correction()
If temp <= 25 Then
    If pressure_alt >= 1000 And pressure_alt <= 3000 Then
    V1 = V1 + 1
    Vr = Vr + 1
    V2 = V2 - 1
    ElseIf pressure_alt > 3000 And pressure_alt <= 5000 Then
    V1 = V1 + 2
    Vr = Vr + 2
    V2 = V2 - 1
    ElseIf pressure_alt > 5000 And pressure_alt <= 7000 Then
    V1 = V1 + 4
    Vr = Vr + 3
    V2 = V2 - 2
    ElseIf pressure_alt > 7000 And pressure_alt <= 9000 Then
    V1 = V1 + 5
    Vr = Vr + 4
    V2 = V2 - 2
    ElseIf pressure_alt > 9000 Then
    V1 = V1 + 6
    Vr = Vr + 5
    V2 = V2 - 3
    End If

'------------------------------------------------------------

ElseIf temp > 25 And temp <= 35 Then
    If pressure_alt >= 1000 And pressure_alt <= 3000 Then
    V1 = V1 + 1
    Vr = Vr + 1
    V2 = V2 - 1
    ElseIf pressure_alt > 3000 And pressure_alt <= 5000 Then
    V1 = V1 + 3
    Vr = Vr + 3
    V2 = V2 - 2
    ElseIf pressure_alt > 5000 And pressure_alt <= 7000 Then
    V1 = V1 + 4
    Vr = Vr + 4
    V2 = V2 - 2
    ElseIf pressure_alt > 7000 And pressure_alt <= 9000 Then
    V1 = V1 + 6
    Vr = Vr + 5
    V2 = V2 - 3
    ElseIf pressure_alt > 9000 Then
    V1 = V1 + 7
    Vr = Vr + 6
    V2 = V2 - 4
    End If

'------------------------------------------------------------

ElseIf temp > 35 And temp <= 45 Then
    If pressure_alt >= (-2000) And pressure_alt <= (-1000) Then
    V1 = V1 + 1
    Vr = Vr + 1
    V2 = V2 - 1
    ElseIf pressure_alt > (-1000) And pressure_alt <= 1000 Then
    V1 = V1 + 2
    Vr = Vr + 1
    V2 = V2 - 1
    ElseIf pressure_alt > 1000 And pressure_alt <= 3000 Then
    V1 = V1 + 3
    Vr = Vr + 3
    V2 = V2 - 2
    ElseIf pressure_alt > 3000 And pressure_alt <= 5000 Then
    V1 = V1 + 4
    Vr = Vr + 4
    V2 = V2 - 2
    ElseIf pressure_alt > 5000 And pressure_alt <= 7000 Then
    V1 = V1 + 6
    Vr = Vr + 5
    V2 = V2 - 3
    ElseIf pressure_alt > 7000 And pressure_alt <= 9000 Then
    V1 = V1 + 7
    Vr = Vr + 6
    V2 = V2 - 4
    ElseIf pressure_alt > 9000 Then
    V1 = V1 + 9
    Vr = Vr + 7
    V2 = V2 - 5
    End If
    
'------------------------------------------------------------

ElseIf temp > 45 And temp <= 55 Then
    If pressure_alt >= (-2000) And pressure_alt <= (-1000) Then
    V1 = V1 + 3
    Vr = Vr + 2
    V2 = V2 - 2
    ElseIf pressure_alt > (-1000) And pressure_alt <= 1000 Then
    V1 = V1 + 4
    Vr = Vr + 3
    V2 = V2 - 2
    ElseIf pressure_alt > 1000 And pressure_alt <= 3000 Then
    V1 = V1 + 5
    Vr = Vr + 4
    V2 = V2 - 3
    ElseIf pressure_alt > 3000 And pressure_alt <= 5000 Then
    V1 = V1 + 6
    Vr = Vr + 5
    V2 = V2 - 3
    ElseIf pressure_alt > 5000 And pressure_alt <= 7000 Then
    V1 = V1 + 8
    Vr = Vr + 6
    V2 = V2 - 4
    ElseIf pressure_alt > 7000 And pressure_alt <= 9000 Then
    V1 = V1 + 9
    Vr = Vr + 7
    V2 = V2 - 5
    ElseIf pressure_alt > 9000 Then
    V1 = V1 + 12
    Vr = Vr + 8
    V2 = V2 - 6
    End If

'-------------------------------------------------------------

ElseIf temp > 55 And temp <= 65 Then
    If pressure_alt >= (-2000) And pressure_alt <= (-1000) Then
    V1 = V1 + 5
    Vr = Vr + 3
    V2 = V2 - 2
    ElseIf pressure_alt > (-1000) And pressure_alt <= 1000 Then
    V1 = V1 + 6
    Vr = Vr + 4
    V2 = V2 - 3
    ElseIf pressure_alt > 1000 And pressure_alt <= 3000 Then
    V1 = V1 + 7
    Vr = Vr + 5
    V2 = V2 - 3
    ElseIf pressure_alt > 3000 And pressure_alt <= 5000 Then
    V1 = V1 + 9
    Vr = Vr + 6
    V2 = V2 - 4
    End If
End If
End Function


'------------------------------------------------------------


Private Sub aircraft_selection_Change()

End Sub



Private Sub calculate_Click()
no_speed = 0
output.Text = ""
press_format_output = ""
If IsNumeric(temp_box.Text) = False Then
    If temp_box.Text = "" Then
    MsgBox ("Input Temperature")
    End If
    Exit Sub
Else: temp = CInt(temp_box.Text)
End If

If temp_format.ListIndex = 0 Then 'format is C

Else: temp = FahrenheitToCelsius(temp)
End If

If IsNumeric(press_box.Text) = False Then
    If press_box.Text = "" Then
    MsgBox ("Input Pressure")
    End If
    Exit Sub
Else: pressure = CDbl(press_box.Text)
End If

If hpa.value = True Then
    pressure = pressure_convert(pressure)
    pressure = round(pressure)
ElseIf inhg.value = True Then
    pressure = pressure
Else
    MsgBox ("Select Pressure format")
    Exit Sub
End If


If hpa.value = True Then
press_format_output = "HPA"
Else: press_format_output = "INHG"
End If

If IsNumeric(elev_box.Text) = False Then
    If elev_box.Text = "" Then
    MsgBox ("Input Airport Elevation")
    Exit Sub
    End If
Else: elevation = CInt(elev_box.Text)
End If

If IsNumeric(weight_box.Text) = False Then
    If weight_box.Text = "" Then
    MsgBox ("Input Takeoff Weight")
    Exit Sub
    End If
Else: weight = weight_box.Text
End If

If flap_select.Text = "" Then
MsgBox ("Select Flap setting")
Exit Sub
Else: flap = flap_select.Text
End If

If rnw_cond.Text = "Wet" Then
wet_rnw = 1
ElseIf rnw_cond = "Dry" Then
wet_rnw = 0
Else
MsgBox ("Select Runway Condition")
Exit Sub
End If

'//////////////Calculating pressure alt////////////
pressure_alt = pressure_altitude(pressure, elevation)

'-------------------Calculation-----------------
If aircraft_selection.ListIndex = 0 Then
b738
If no_speed <> 1 Then
output.Text = "-------BOEING 737-800--------" & vbNewLine & vbNewLine & temp & temp_format.Text & "     " & press_box.Text & " " & press_format_output & "     " & rnw_cond.Text & vbNewLine & vbNewLine & "TOW " & weight_box.Text & vbNewLine & vbNewLine & "THR TO" & vbNewLine & vbNewLine & "FLAPS " & flap_select.Text & vbNewLine & vbNewLine & "V1 " & V1 & vbNewLine & "Vr " & Vr & vbNewLine & "V2 " & V2
End If
End If
End Sub




Private Sub elev_box_Change()
elev_box.MaxLength = 5
End Sub

Private Sub flap_select_GotFocus()
If aircraft_selection.ListIndex = 0 Then
    flap_select.Clear
    flap_select.AddItem "1"
    flap_select.AddItem "5"
    flap_select.AddItem "10"
    flap_select.AddItem "15"
    flap_select.AddItem "25"
ElseIf aircraft_selection.ListIndex = 1 Then
    flap_select.Clear
    flap_select.AddItem "15"
    flap_select.AddItem "28"
ElseIf aircraft_selection.ListIndex = 2 Then
    flap_select.Clear
    flap_select.AddItem "10"
    flap_select.AddItem "11"
    flap_select.AddItem "12"
    flap_select.AddItem "13"
    flap_select.AddItem "14"
    flap_select.AddItem "15"
    flap_select.AddItem "16"
    flap_select.AddItem "17"
    flap_select.AddItem "18"
    flap_select.AddItem "19"
    flap_select.AddItem "20"
    flap_select.AddItem "21"
    flap_select.AddItem "22"
    flap_select.AddItem "23"
    flap_select.AddItem "24"
    flap_select.AddItem "25"
    flap_select.AddItem "28"
ElseIf aircraft_selection.ListIndex = 3 Then
    flap_select.Clear
    flap_select.AddItem "10"
    flap_select.AddItem "18"
End If
End Sub


Private Sub Form_Load()
temp_format.AddItem "C"
temp_format.AddItem "F"
temp_format.ListIndex = 0
aircraft_selection.AddItem "Boeing 737-800"
'aircraft_selection.AddItem "Tupolev TU-154"
'aircraft_selection.AddItem "Mcdonnell Douglas MD-11"
'aircraft_selection.AddItem "Lockheed Martin L1011"
no_speed = 0
rnw_cond.AddItem "Dry"
rnw_cond.AddItem "Wet"
output.Locked = True
V1 = 0
Vr = 0
V2 = 0
End Sub



Private Sub hpa_Click()
inhg.value = False
press_box.MaxLength = 4
End Sub

Private Sub inhg_Click()
hpa.value = False
press_box.MaxLength = 5
End Sub


Private Sub Option1_Click()

End Sub

Private Sub press_box_Change()
If hpa.value = True Then
press_box.MaxLength = 4
Else
press_box.MaxLength = 5
End If
End Sub

Private Sub temp_box_Change()
temp_box.MaxLength = 2

End Sub


Private Sub weight_box_Change()
weight_box.MaxLength = 6
End Sub


Private Sub wet_Click()
dry.value = False
End Sub


