Attribute VB_Name = "modStartup"
'
' Look mom!  No Globals!
'
'



Sub main()
    '
    ' Please examine this demonstration carefully!
    '
        
    ' We want to make a group of key/value pairs.  This
    ' is as easy as creating a Section zabArray, and
    ' storing Key/Value zabArrays in it.  A zabArray
    ' doesn't allow object variants, but you can get
    ' the Serial string from KeyValue and store it as
    ' a string in Section.  :-)
    
    Dim Section As New zabArray
    Dim KeyValue As New zabArray
    
    
    With KeyValue
        ' Set some form options in KeyValue
        .SetValue "Top", 1000
        .SetValue "Left", 1000
        .SetValue "Height", 3000
        .SetValue "Width", 4000
        .SetValue "BackColor", RGB(40, 0, 80)
        .SetValue "Caption", "zabArray Test"
    End With
    
    'The data above will become Option1.
    Section.SetValue "Option1", KeyValue.GetSerial
      
            
    
    With KeyValue
        'Replace the old values with new ones.
        .SetValue "Top", 2000
        .SetValue "Left", 2000
        .SetValue "Height", 4000
        .SetValue "Width", 5000
        .SetValue "BackColor", RGB(0, 80, 40)
        .SetValue "Caption", "Testing zabArray"
    End With
    
    'Save the new data as Option2
    Section.SetValue "Option2", KeyValue.GetSerial
    
    
    
    
    ' We set up a public property in Form1 to recieve
    ' a string.  This is where we transmit the serial
    ' data to the form.
    
    Dim frm As New Form1
    
    Load frm
    
    frm.Serial = Section.GetSerial  'Just like that.
    
    frm.Show
    
    
End Sub
