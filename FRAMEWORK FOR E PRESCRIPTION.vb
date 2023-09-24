'This code works as a timer and makes the introduction page visible'
Private Sub Timer1_Timer ()
Form1.Visible = True
End Sub
'This code removes the introduction page if it's not already hidden'
Private Sub Timer2_Timer ()
Form1.Visible = False
'Checks if Introduction page is hidden'
If Form1.Visible = False, Then
'If introduction page is hidden, these codes turn off the timer'
Timer1.Enabled = False
Timer2.Enabled = False
'Visualizes the Login page'
Form2.Visible = True
End If
End Sub
SECOND PHASE
'This code executes the login function'
Private Sub Command1_Click ()
Form2.Visible = False
Form3.Visible = True
End Sub
'Executes the close button at the top left side page'
Private Sub mnuCLOSE_Click()
End
End Sub
'Makes the project description/title landing invisible'
Private Sub Timer1_Timer()
Label1.Visible = False
End Sub
'Timer effect, makes the project description/title landing visible'
Private Sub Timer2_Timer()
Label1.Visible = True
End Sub
THIRD PHASE
'This executes when login button is clicked, supersedes the login page and introduces the homepage/GP interface'
Private Sub Command1_Click()
Form3.Visible = False
Form6.Visible = True
End Sub
'This code updates and checks the login details'
Private Sub Command2_Click()
Adodc1.Recordset.Update
End Sub
'This brings back login page if CANCEL button is clicked'
Private Sub Command3_Click()
Form2.Visible = True
Form3.Visible = False
End Sub
'This code saves and checks the login details'
Private Sub Command4_Click()
Adodc1.Recordset.Save
End Sub


'This code executes when CLEAR button is clicked, and the boxes are cleared of input'
Private Sub Command5_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
End Sub
FOURTH phase
'This code executes the Login function'
Private Sub B_Click()
Form4.Visible = False
Form3.Visible = True
End Sub
Private Sub Command1_Click()
Form3.Visible = False
Form6.Visible = True
End Sub
'Shows the login page'
Private Sub Command3_Click()
loginfrm.Visible = True
Form3.Visible = False
End Sub
Private Sub Command4_Click()
Adodc1.Recordset.Save
End Sub
'This code executes when CLEAR button is clicked, and the boxes are cleared of input'
Private Sub Command5_Click()
Text1.Text = ""
Text2.Text = ""
Combo1.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
End Sub
Private Sub Command2_Click()
Form4.Visible = False
Form6.Visible = True
End Sub
'This code executes when CLEAR button is clicked, and the boxes are cleared of input'
Private Sub Command3_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
End Sub
Private Sub Timer1_Timer()
Picture5.Visible = True
End Sub
Private Sub Timer2_Timer()
Picture4.Visible = False
End Sub
FIFTH  PHASE
Private Sub Command1_Click()
Form6.Visible = False
Form2.Visible = True
End Sub
Private Sub Command2_Click()
Form6.Visible = False
Form4.Visible = True
End Sub
Private Sub Command3_Click()
Form6.Visible = False
t.Visible = True
End Sub
'This code executes when CLEAR button is clicked, and the boxes are cleared of input'
SIXTH PHASE
Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
End Sub
Private Sub Command3_Click()
t.Visible = False
Form2.Visible = True
End Sub
Private Sub Command4_Click()
Print Form
End Sub

