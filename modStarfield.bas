Attribute VB_Name = "modStarfield"
Option Explicit

Type Star
    X As Integer
    Y As Integer
    Color As Long
End Type

'these are the speeds the stars will travel in certain planes
Public Const Plane1Velocity = 1
Public Const Plane2Velocity = 2
Public Const Plane3Velocity = 3

'this is an array used throughout the program
Dim stars(1 To 3, 1 To 100) As Star

Sub InitStars()
'this will be the planes
Dim i As Integer
'this will be the stars within the planes
Dim j As Integer

    For i = 1 To 3
        For j = 1 To 100

          'this randomizes the stars and sets their color
          Randomize
          stars(i, j).X = Int((frmMain.ScaleWidth * Rnd) + 1)
          stars(i, j).Y = Int((frmMain.ScaleHeight * Rnd) + 1)
          
          Select Case i
          Case 1
          stars(i, j).Color = RGB(50, 50, 50)
          Case 2
          stars(i, j).Color = RGB(100, 100, 100)
          Case 3
          stars(i, j).Color = RGB(255, 255, 255)
          End Select

        Next j
    Next i
End Sub

Sub DrawStars()
Dim i As Integer
Dim j As Integer
    
    For i = 1 To 3
        For j = 1 To 100
        
            'this erases the last set of stars
            frmMain.picField.PSet (stars(i, j).X, stars(i, j).Y), RGB(0, 0, 0)
            
            'this moves the stars, uncomment the ones below for different effects
            Select Case i
            Case 1
            stars(i, j).X = stars(i, j).X + Plane1Velocity
            'stars(i, j).Y = stars(i, j).Y + Plane1Velocity
            Case 2
            stars(i, j).X = stars(i, j).X + Plane2Velocity
            'stars(i, j).Y = stars(i, j).Y + Plane2Velocity
            Case 3
            stars(i, j).X = stars(i, j).X + Plane3Velocity
            'stars(i, j).Y = stars(i, j).Y + Plane3Velocity
            End Select
            
            If stars(i, j).X > frmMain.picField.Width Then
                stars(i, j).X = frmMain.picField.Left
            End If
            
            'If stars(i, j).Y > frmMain.picField.Height Then
            '    stars(i, j).Y = frmMain.picField.Top
            'End If
            
            'this draws the stars
            frmMain.picField.PSet (stars(i, j).X, stars(i, j).Y), stars(i, j).Color
            
        Next j
    Next i
End Sub
