VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   ScaleHeight     =   2760
   ScaleWidth      =   5310
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdData 
      Caption         =   "Data"
      Height          =   285
      Left            =   90
      TabIndex        =   2
      Top             =   990
      Width           =   1230
   End
   Begin VB.CommandButton cmdOption2 
      Caption         =   "Option 2"
      Height          =   330
      Left            =   90
      TabIndex        =   1
      Top             =   495
      Width           =   1230
   End
   Begin VB.CommandButton cmdOption1 
      Caption         =   "Option 1"
      Height          =   330
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   1230
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' PLEASE start by reading comments in modStartup.
'
'


Private MySerial As String  'Local variable to hold
                            'a serial string.










Public Property Let Serial(ByVal vData As String)
    ' Write-only property for the serial string.
    ' There is nothing against reading the string,
    ' but we have no use for that in this example.
    MySerial = vData
End Property





Private Sub cmdData_Click()
    ' Show the contents of the serial string.  Let's you
    ' see what exactly gets stored and how.
    MsgBox MySerial
End Sub




Private Sub cmdOption1_Click()

    ' This is almost identical to cmdOption2_click.  I
    ' could have used a control array, but I wanted to
    ' demonstrate the operation a little better.
    '
    ' Here we pick apart the Section and KeysValues.
    '
    
    ' Locals, mind you.
    Dim MySection As New zabArray
    Dim MyKeyValue As New zabArray
    
    ' Pump the serial into the zabArray.  The zabArray
    ' should automatically clear and accept the new
    ' data.  Provided it is valid data...
    MySection.SetSerial MySerial
    
    With MyKeyValue
        'Two part statement.  We need to get the Option1
        'string.  Since we know that this string is a
        'Serial, we simply feed it into MyKeyValue
        .SetSerial MySection.GetValue("Option1")
    
        'Provided everything worked, we can now pull
        'data from the zabArray.  Notice that we aren't
        'case-sensitive, but you can force that if you
        'want with CaseSensitiveKeys.
        Me.Top = .GetValue("top")               'Long
        Me.Left = .GetValue("left")             'Long
        Me.Width = .GetValue("width")           'Long
        Me.Height = .GetValue("height")         'Long
        Me.BackColor = .GetValue("Backcolor")   'String
        Me.Caption = .GetValue("caption")       'Long
    End With
    
End Sub




Private Sub cmdOption2_Click()
    
    ' This is almost identical to cmdOption1_click.  I
    ' could have used a control array, but I wanted to
    ' demonstrate the operation a little better.
    '
    ' Here we pick apart the Section and KeysValues.
    '
    
    ' Locals, mind you.
    Dim MySection As New zabArray
    Dim MyKeyValue As New zabArray
    
    ' Pump the serial into the zabArray.  The zabArray
    ' should automatically clear and accept the new
    ' data.  Provided it is valid data...
    MySection.SetSerial MySerial
    
    With MyKeyValue
        'Two part statement.  We need to get the Option1
        'string.  Since we know that this string is a
        'Serial, we simply feed it into MyKeyValue
        .SetSerial MySection.GetValue("Option2")
    
        'Provided everything worked, we can now pull
        'data from the zabArray.  Notice that we aren't
        'case-sensitive, but you can force that if you
        'want with CaseSensitiveKeys.
        Me.Top = .GetValue("top")               'Long
        Me.Left = .GetValue("left")             'Long
        Me.Width = .GetValue("width")           'Long
        Me.Height = .GetValue("height")         'Long
        Me.BackColor = .GetValue("Backcolor")   'String
        Me.Caption = .GetValue("caption")       'Long
    End With
    
End Sub

