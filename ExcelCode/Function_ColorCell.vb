Public Function ColorCell(oCell As Range, sColor As String, Optional bSet_Text_Color_NONE As Boolean = False)
Dim intColor As Integer

'Select the color you want

Select Case UCase(sColor)

Case "BLACK"

    intColor = BLACK
    
Case "BLUE"

    intColor = BLUE
    
Case "BLUEGREEN"

    intColor = TURQUOISE
    
Case "ORANGE"

    intColor = ORANGE
   
Case "GREY"

    intColor = GREY40
   
Case "LIGHTGREY"

    intColor = GREY25
   
Case "ICEBLUE"

    intColor = ICEBLUE
   
Case "NONE"

    intColor = NONE
   
Case "LIGHTGREEN"

   intColor = LIGHTGREEN
   
Case "YELLOWGREEN"

   intColor = BRIGHTGREEN
       
Case "LIGHTBLUE"

    intColor = LIGHTBLUE
   
Case "DARKGREEN"
 
     intColor = SEAGREEN
   
Case "LIGHTYELLOW"

    intColor = LIGHTYELLOW
   
Case "LIGHTORANGE"
   
    intColor = LIGHTORANGE

Case "RED"

    intColor = RED

Case "YELLOW"

    intColor = YELLOW
   
Case "WHITE"

    intColor = WHITE
   
Case "ICEBLUE"

    intColor = ICEBLUE
    
 Case "PINK"

  intColor = PALEPINK
  
Case ""

  intColor = NONE
   
Case Else

    intColor = NONE
   
End Select

'Set Fill color and turn Text color off "to white" or leave text color default color of black
'NOTE: This assumes that the default text color of the workbook is black

'Set with default text color.
 If bSet_Text_Color_NONE = False Then
 
 
       'Set Fill color with Text color (or white text color)
        If Not intColor = NONE Then
        
            With oCell.Interior
                .ColorIndex = intColor
                .Pattern = xlSolid
            End With
           
        Else
        
            With oCell.Interior
                .ColorIndex = xlNone
                .Pattern = xlNone
            End With
        
        End If
        
 End If

'Set Text Color
  If bSet_Text_Color_NONE = True Then
  
       'Set Fill color with NO text color (or white text color)
        If Not intColor = NONE Then
        
          'Set fill color
            With oCell.Interior
                .ColorIndex = intColor
                .Pattern = xlSolid
            End With
        
          'Set WITHOUT text color
            With oCell.Font
                .ColorIndex = WHITE
            End With

           
        Else
        
           'Set without fill color with Text color (default black) fill color and black (default text
            With oCell.Interior
                .ColorIndex = xlNone
                .Pattern = xlNone
            End With
            
            'Set to default text color
             With oCell.Font
                .ColorIndex = BLACK
            End With
        

        
        End If
        
  End If
    
    
End Function
