Attribute VB_Name = "DateValidation"
'######################################################
'# This function returns a date-value, which is 1 day #
'# later than the input-date. It checks for month and #
'# year changes and handles leap-years.               #
'# This is usefull, if you plan meetings and don't    #
'# want to quit the meeting before it begins by a     #
'# mistake.                                           #
'# I use DateTimePicker at the form. Those are not    #
'# neccessary for the function and is NOT included.   #
'# You'll find it in the Microsoft Windows Common     #
'# Controls-2 6.0 (MSCOMCT2.OCX).                     #
'#                                                    #
'#                                                    #
'# You can use any date-object or variable you like.  #
'#                                                    #
'# Feel free to optimize and debug this :-)           #
'# And please tell me your opinion!                   #
'# LordOklar@gmx.net                                  #
'######################################################
Option Explicit

Public Function later(GetDate As Date)
                        'Date to stay over
  Dim dd As Integer     ' Day
  Dim mm As Integer     ' Month
  Dim yyyy As Integer   ' Year

  
  'This brings the date-value to the variables,
  'handling them like strings (converted to integer)
  dd = CInt(Left$(GetDate, 2))
  mm = CInt(Right$(Left$(GetDate, 5), 2))
  yyyy = CInt(Right$(GetDate, 2))
  
  'Counts up a day
  dd = dd + 1
  
  
  'Validates if the month got 28, 29, 30 or 31 days
  Select Case mm
         Case 1
              If dd >= 32 Then
                 dd = 1
                 mm = mm + 1
              End If
         
         Case 2
              'This prevents zero-division
              yyyy = yyyy + 4
              'Checks if the year is dividable through 4 (leap-years)
              If (yyyy / Int(yyyy / 4)) = 4 Then
                 If dd >= 30 Then
                    dd = 1
                    mm = mm + 1
                 End If
              Else
                If dd >= 29 Then
                   dd = 1
                   mm = mm + 1
                End If
              End If
              'To get the real year back
              yyyy = yyyy - 4
         
         Case 3
              If dd >= 32 Then
                 dd = 1
                 mm = mm + 1
              End If
         
         Case 4
              If dd >= 31 Then
                 dd = 1
                 mm = mm + 1
              End If
         
         Case 5
              If dd = 32 Then
                 dd = 1
                 mm = mm + 1
              End If
         
         Case 6
              If dd = 31 Then
                 dd = 1
                 mm = mm + 1
              End If
         
         Case 7
              If dd = 32 Then
                 dd = 1
                 mm = mm + 1
              End If
         
         Case 8
              If dd = 32 Then
                 dd = 1
                 mm = mm + 1
              End If
         
         Case 9
              If dd = 31 Then
                 dd = 1
                 mm = mm + 1
              End If
         
         Case 10
              If dd = 32 Then
                 dd = 1
                 mm = mm + 1
              End If
         
         Case 11
              If dd = 31 Then
                 dd = 1
                 mm = mm + 1
              End If
         
         Case 12
              If dd = 32 Then
                 dd = 1
                 mm = mm + 1
              End If
   
  End Select
  
  'Validates month and count up year
  If mm = 13 Then
     mm = 1
     yyyy = yyyy + 1
  End If
  
  'Returns the date-string with 1 day later
  later = dd & "." & mm & "." & yyyy
  
End Function

