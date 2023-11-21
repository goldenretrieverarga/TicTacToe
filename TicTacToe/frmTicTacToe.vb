'***
' Action
'   - Creating the game TicTacToe. This is form as startpoint of the exercise
' Created
'   - CopyPaste – 20070226 – VVDW
' Changed
'   - Organisation – yyyymmdd – Initials of programmer – What changed
' Tested
'   - CopyPaste – 20070226 – VVDW
' Proposal (To Do)
'   - List of actions that can be added to the functionality
'***

Option Explicit On 
Option Strict On

Imports System.Drawing

Public Class frmTicTacToe
  Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "
  'Required by the Windows Form Designer
  Private components As System.ComponentModel.Container

  'NOTE: The following procedure is required by the Windows Form Designer
  'It can be modified using the Windows Form Designer.  
  'Do not modify it using the code editor.
  Friend WithEvents lblTitle As System.Windows.Forms.Label
  Friend WithEvents cmd09 As System.Windows.Forms.Button
  Friend WithEvents cmd08 As System.Windows.Forms.Button
  Friend WithEvents cmd07 As System.Windows.Forms.Button
  Friend WithEvents cmd06 As System.Windows.Forms.Button
  Friend WithEvents cmd05 As System.Windows.Forms.Button
  Friend WithEvents cmd04 As System.Windows.Forms.Button
  Friend WithEvents cmd03 As System.Windows.Forms.Button
  Friend WithEvents cmd02 As System.Windows.Forms.Button
  Friend WithEvents cmd01 As System.Windows.Forms.Button
  Friend WithEvents panHorizontal02 As System.Windows.Forms.Panel
  Friend WithEvents panHorizontal01 As System.Windows.Forms.Panel
  Friend WithEvents panVertical01 As System.Windows.Forms.Panel
  Friend WithEvents panVertical02 As System.Windows.Forms.Panel
  Friend WithEvents cmdRestart As System.Windows.Forms.Button
  Friend WithEvents cmdQuit As System.Windows.Forms.Button
  Friend WithEvents lblPlayer As System.Windows.Forms.Label
  <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmTicTacToe))
        Me.lblTitle = New System.Windows.Forms.Label()
        Me.cmd09 = New System.Windows.Forms.Button()
        Me.cmd08 = New System.Windows.Forms.Button()
        Me.cmd07 = New System.Windows.Forms.Button()
        Me.cmd06 = New System.Windows.Forms.Button()
        Me.cmd05 = New System.Windows.Forms.Button()
        Me.cmd04 = New System.Windows.Forms.Button()
        Me.cmd03 = New System.Windows.Forms.Button()
        Me.cmd02 = New System.Windows.Forms.Button()
        Me.cmd01 = New System.Windows.Forms.Button()
        Me.panHorizontal02 = New System.Windows.Forms.Panel()
        Me.panHorizontal01 = New System.Windows.Forms.Panel()
        Me.panVertical01 = New System.Windows.Forms.Panel()
        Me.panVertical02 = New System.Windows.Forms.Panel()
        Me.cmdRestart = New System.Windows.Forms.Button()
        Me.cmdQuit = New System.Windows.Forms.Button()
        Me.lblPlayer = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'lblTitle
        '
        Me.lblTitle.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTitle.Font = New System.Drawing.Font("Times New Roman", 36.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTitle.ForeColor = System.Drawing.Color.Red
        Me.lblTitle.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.lblTitle.Location = New System.Drawing.Point(6, 0)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(492, 48)
        Me.lblTitle.TabIndex = 1
        Me.lblTitle.Text = "Tic-Tac-Toe"
        Me.lblTitle.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cmd09
        '
        Me.cmd09.BackColor = System.Drawing.Color.LimeGreen
        Me.cmd09.Font = New System.Drawing.Font("Arial Black", 72.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmd09.Location = New System.Drawing.Point(348, 342)
        Me.cmd09.Name = "cmd09"
        Me.cmd09.Size = New System.Drawing.Size(90, 90)
        Me.cmd09.TabIndex = 0
        Me.cmd09.UseVisualStyleBackColor = False
        '
        'cmd08
        '
        Me.cmd08.BackColor = System.Drawing.Color.LimeGreen
        Me.cmd08.Font = New System.Drawing.Font("Arial Black", 72.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmd08.Location = New System.Drawing.Point(216, 342)
        Me.cmd08.Name = "cmd08"
        Me.cmd08.Size = New System.Drawing.Size(90, 90)
        Me.cmd08.TabIndex = 0
        Me.cmd08.UseVisualStyleBackColor = False
        '
        'cmd07
        '
        Me.cmd07.BackColor = System.Drawing.Color.LimeGreen
        Me.cmd07.Font = New System.Drawing.Font("Arial Black", 72.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmd07.Location = New System.Drawing.Point(84, 342)
        Me.cmd07.Name = "cmd07"
        Me.cmd07.Size = New System.Drawing.Size(90, 90)
        Me.cmd07.TabIndex = 0
        Me.cmd07.UseVisualStyleBackColor = False
        '
        'cmd06
        '
        Me.cmd06.BackColor = System.Drawing.Color.LimeGreen
        Me.cmd06.Font = New System.Drawing.Font("Arial Black", 72.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmd06.Location = New System.Drawing.Point(348, 210)
        Me.cmd06.Name = "cmd06"
        Me.cmd06.Size = New System.Drawing.Size(90, 90)
        Me.cmd06.TabIndex = 0
        Me.cmd06.UseVisualStyleBackColor = False
        '
        'cmd05
        '
        Me.cmd05.BackColor = System.Drawing.Color.LimeGreen
        Me.cmd05.Font = New System.Drawing.Font("Arial Black", 72.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmd05.Location = New System.Drawing.Point(216, 210)
        Me.cmd05.Name = "cmd05"
        Me.cmd05.Size = New System.Drawing.Size(90, 90)
        Me.cmd05.TabIndex = 0
        Me.cmd05.UseVisualStyleBackColor = False
        '
        'cmd04
        '
        Me.cmd04.BackColor = System.Drawing.Color.LimeGreen
        Me.cmd04.Font = New System.Drawing.Font("Arial Black", 72.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmd04.Location = New System.Drawing.Point(84, 210)
        Me.cmd04.Name = "cmd04"
        Me.cmd04.Size = New System.Drawing.Size(90, 90)
        Me.cmd04.TabIndex = 0
        Me.cmd04.UseVisualStyleBackColor = False
        '
        'cmd03
        '
        Me.cmd03.BackColor = System.Drawing.Color.LimeGreen
        Me.cmd03.Font = New System.Drawing.Font("Arial Black", 72.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmd03.Location = New System.Drawing.Point(348, 78)
        Me.cmd03.Name = "cmd03"
        Me.cmd03.Size = New System.Drawing.Size(90, 90)
        Me.cmd03.TabIndex = 0
        Me.cmd03.UseVisualStyleBackColor = False
        '
        'cmd02
        '
        Me.cmd02.BackColor = System.Drawing.Color.LimeGreen
        Me.cmd02.Font = New System.Drawing.Font("Arial Black", 72.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmd02.Location = New System.Drawing.Point(216, 78)
        Me.cmd02.Name = "cmd02"
        Me.cmd02.Size = New System.Drawing.Size(90, 90)
        Me.cmd02.TabIndex = 0
        Me.cmd02.UseVisualStyleBackColor = False
        '
        'cmd01
        '
        Me.cmd01.BackColor = System.Drawing.Color.LimeGreen
        Me.cmd01.Font = New System.Drawing.Font("Arial Black", 72.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmd01.Location = New System.Drawing.Point(84, 78)
        Me.cmd01.Name = "cmd01"
        Me.cmd01.Size = New System.Drawing.Size(90, 90)
        Me.cmd01.TabIndex = 0
        Me.cmd01.UseVisualStyleBackColor = False
        '
        'panHorizontal02
        '
        Me.panHorizontal02.BackColor = System.Drawing.Color.Blue
        Me.panHorizontal02.Location = New System.Drawing.Point(54, 312)
        Me.panHorizontal02.Name = "panHorizontal02"
        Me.panHorizontal02.Size = New System.Drawing.Size(400, 18)
        Me.panHorizontal02.TabIndex = 2
        '
        'panHorizontal01
        '
        Me.panHorizontal01.BackColor = System.Drawing.Color.Blue
        Me.panHorizontal01.Location = New System.Drawing.Point(48, 180)
        Me.panHorizontal01.Name = "panHorizontal01"
        Me.panHorizontal01.Size = New System.Drawing.Size(400, 18)
        Me.panHorizontal01.TabIndex = 2
        '
        'panVertical01
        '
        Me.panVertical01.BackColor = System.Drawing.Color.Blue
        Me.panVertical01.Location = New System.Drawing.Point(186, 60)
        Me.panVertical01.Name = "panVertical01"
        Me.panVertical01.Size = New System.Drawing.Size(18, 400)
        Me.panVertical01.TabIndex = 2
        '
        'panVertical02
        '
        Me.panVertical02.BackColor = System.Drawing.Color.Blue
        Me.panVertical02.Location = New System.Drawing.Point(318, 60)
        Me.panVertical02.Name = "panVertical02"
        Me.panVertical02.Size = New System.Drawing.Size(18, 400)
        Me.panVertical02.TabIndex = 2
        '
        'cmdRestart
        '
        Me.cmdRestart.Location = New System.Drawing.Point(348, 492)
        Me.cmdRestart.Name = "cmdRestart"
        Me.cmdRestart.Size = New System.Drawing.Size(72, 24)
        Me.cmdRestart.TabIndex = 4
        Me.cmdRestart.Text = "Start"
        '
        'cmdQuit
        '
        Me.cmdQuit.Location = New System.Drawing.Point(426, 492)
        Me.cmdQuit.Name = "cmdQuit"
        Me.cmdQuit.Size = New System.Drawing.Size(72, 24)
        Me.cmdQuit.TabIndex = 4
        Me.cmdQuit.Text = "Stop"
        '
        'lblPlayer
        '
        Me.lblPlayer.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblPlayer.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPlayer.Location = New System.Drawing.Point(6, 492)
        Me.lblPlayer.Name = "lblPlayer"
        Me.lblPlayer.Size = New System.Drawing.Size(336, 24)
        Me.lblPlayer.TabIndex = 3
        Me.lblPlayer.Text = "Speler"
        Me.lblPlayer.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'frmTicTacToe
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(509, 525)
        Me.Controls.Add(Me.cmdRestart)
        Me.Controls.Add(Me.cmdQuit)
        Me.Controls.Add(Me.lblPlayer)
        Me.Controls.Add(Me.cmd09)
        Me.Controls.Add(Me.cmd08)
        Me.Controls.Add(Me.cmd07)
        Me.Controls.Add(Me.cmd06)
        Me.Controls.Add(Me.cmd05)
        Me.Controls.Add(Me.cmd04)
        Me.Controls.Add(Me.cmd03)
        Me.Controls.Add(Me.cmd02)
        Me.Controls.Add(Me.cmd01)
        Me.Controls.Add(Me.panHorizontal02)
        Me.Controls.Add(Me.panHorizontal01)
        Me.Controls.Add(Me.panVertical01)
        Me.Controls.Add(Me.lblTitle)
        Me.Controls.Add(Me.panVertical02)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmTicTacToe"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Tic-Tac-Toe"
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Constructors / Destructors"

    Protected Overloads Overrides Sub Dispose(ByVal blnDisposing As Boolean)
    '***
    ' Action
    '   - Cleanup after closing the form
    ' Called by
    '   - User action
    ' Calls
    '   - 
    ' Created
    '   - CopyPaste – 20070226 – VVDW
    ' Changed
    '   - Organisation – yyyymmdd – Initials of programmer – What changed
    ' Tested
    '   - CopyPaste – 20070226 – VVDW
    ' Keyboard key
    '   - 
    ' Proposal (To Do)
    '   - List of actions that can be added to the functionality
    '***

    If blnDisposing Then

      If components Is Nothing Then
      Else
        ' Not components Is Nothing
        components.Dispose()
      End If
      ' components Is Nothing

    Else
      ' Not blnDisposing
    End If
    ' blnDisposing

    MyBase.Dispose(blnDisposing)
  End Sub
  ' Dispose(Boolean)

  Public Sub New()
    '***
    ' Action
    '   - Creating an instance of the form
    '   - Initialize the components of that form
    ' Called by
    '   - User action
    ' Calls
    '   - 
    ' Created
    '   - CopyPaste – 20070226 – VVDW
    ' Changed
    '   - Organisation – yyyymmdd – Initials of programmer – What changed
    ' Tested
    '   - CopyPaste – 20070226 – VVDW
    ' Keyboard key
    '   - 
    ' Proposal (To Do)
    '   - List of actions that can be added to the functionality
    '***

    MyBase.New()
    InitializeComponent()
  End Sub
  ' New()

#End Region

  '#Region "Designer"
  '#End Region

  '#Region "Structures"
  '#End Region

#Region "Fields"
  Dim mchrToken As Char
  Dim mlngPlayer As Int32
#End Region

  '#Region "Properties"
  '#End Region

#Region "Methods"

  '#Region "Overrides"
  '#End Region

#Region "Controls"

  Private Sub cmd01_Click(ByVal theSender As System.Object, ByVal theEventArguments As System.EventArgs) Handles cmd01.Click
    '***
    ' Action
    '   - Set the text of the button to a character (mchrToken)
    '   - Disable the button
    ' Called by
    '   - User action (Clicking a button)
    ' Calls
    '   - CheckWinner()
    ' Created
    '   - CopyPaste – 20070226 – VVDW
    ' Changed
    '   - Organisation – yyyymmdd – Initials of programmer – What changed
    ' Tested
    '   - CopyPaste – 20070226 – VVDW
    ' Keyboard key
    '   - 
    ' Proposal (To Do)
    '   - List of actions that can be added to the functionality
    '***

    cmd01.Text = mchrToken
    cmd01.Enabled = False
    CheckWinner()
  End Sub
  ' cmd01_Click(System.Object, System.EventArgs) Handles cmd01.Click

  Private Sub cmd02_Click(ByVal theSender As System.Object, ByVal theEventArguments As System.EventArgs) Handles cmd02.Click
    '***
    ' Action
    '   - Set the text of the button to a character (mchrToken)
    '   - Disable the button
    ' Called by
    '   - User action (Clicking a button)
    ' Calls
    '   - CheckWinner()
    ' Created
    '   - CopyPaste – 20070226 – VVDW
    ' Changed
    '   - Organisation – yyyymmdd – Initials of programmer – What changed
    ' Tested
    '   - CopyPaste – 20070226 – VVDW
    ' Keyboard key
    '   - 
    ' Proposal (To Do)
    '   - List of actions that can be added to the functionality
    '***

    cmd02.Text = mchrToken
    cmd02.Enabled = False
    CheckWinner()
  End Sub
  ' cmd02_Click(System.Object, System.EventArgs) Handles cmd02.Click

  Private Sub cmd03_Click(ByVal theSender As System.Object, ByVal theEventArguments As System.EventArgs) Handles cmd03.Click
    '***
    ' Action
    '   - Set the text of the button to a character (mchrToken)
    '   - Disable the button
    ' Called by
    '   - User action (Clicking a button)
    ' Calls
    '   - CheckWinner()
    ' Created
    '   - CopyPaste – 20070226 – VVDW
    ' Changed
    '   - Organisation – yyyymmdd – Initials of programmer – What changed
    ' Tested
    '   - CopyPaste – 20070226 – VVDW
    ' Keyboard key
    '   - 
    ' Proposal (To Do)
    '   - List of actions that can be added to the functionality
    '***

    cmd03.Text = mchrToken
    cmd03.Enabled = False
    CheckWinner()
  End Sub
  ' cmd03_Click(System.Object, System.EventArgs) Handles cmd03.Click

  Private Sub cmd04_Click(ByVal theSender As System.Object, ByVal theEventArguments As System.EventArgs) Handles cmd04.Click
    '***
    ' Action
    '   - Set the text of the button to a character (mchrToken)
    '   - Disable the button
    ' Called by
    '   - User action (Clicking a button)
    ' Calls
    '   - CheckWinner()
    ' Created
    '   - CopyPaste – 20070226 – VVDW
    ' Changed
    '   - Organisation – yyyymmdd – Initials of programmer – What changed
    ' Tested
    '   - CopyPaste – 20070226 – VVDW
    ' Keyboard key
    '   - 
    ' Proposal (To Do)
    '   - List of actions that can be added to the functionality
    '***

    cmd04.Text = mchrToken
    cmd04.Enabled = False
    CheckWinner()
  End Sub
  ' cmd04_Click(System.Object, System.EventArgs) Handles cmd04.Click

  Private Sub cmd05_Click(ByVal theSender As System.Object, ByVal theEventArguments As System.EventArgs) Handles cmd05.Click
    '***
    ' Action
    '   - Set the text of the button to a character (mchrToken)
    '   - Disable the button
    ' Called by
    '   - User action (Clicking a button)
    ' Calls
    '   - CheckWinner()
    ' Created
    '   - CopyPaste – 20070226 – VVDW
    ' Changed
    '   - Organisation – yyyymmdd – Initials of programmer – What changed
    ' Tested
    '   - CopyPaste – 20070226 – VVDW
    ' Keyboard key
    '   - 
    ' Proposal (To Do)
    '   - List of actions that can be added to the functionality
    '***

    cmd05.Text = mchrToken
    cmd05.Enabled = False
    CheckWinner()
  End Sub
  ' cmd05_Click(System.Object, System.EventArgs) Handles cmd05.Click

  Private Sub cmd06_Click(ByVal theSender As System.Object, ByVal theEventArguments As System.EventArgs) Handles cmd06.Click
    '***
    ' Action
    '   - Set the text of the button to a character (mchrToken)
    '   - Disable the button
    ' Called by
    '   - User action (Clicking a button)
    ' Calls
    '   - CheckWinner()
    ' Created
    '   - CopyPaste – 20070226 – VVDW
    ' Changed
    '   - Organisation – yyyymmdd – Initials of programmer – What changed
    ' Tested
    '   - CopyPaste – 20070226 – VVDW
    ' Keyboard key
    '   - 
    ' Proposal (To Do)
    '   - List of actions that can be added to the functionality
    '***

    cmd06.Text = mchrToken
    cmd06.Enabled = False
    CheckWinner()
  End Sub
  ' cmd06_Click(System.Object, System.EventArgs) Handles cmd06.Click

  Private Sub cmd07_Click(ByVal theSender As System.Object, ByVal theEventArguments As System.EventArgs) Handles cmd07.Click
    '***
    ' Action
    '   - Set the text of the button to a character (mchrToken)
    '   - Disable the button
    ' Called by
    '   - User action (Clicking a button)
    ' Calls
    '   - CheckWinner()
    ' Created
    '   - CopyPaste – 20070226 – VVDW
    ' Changed
    '   - Organisation – yyyymmdd – Initials of programmer – What changed
    ' Tested
    '   - CopyPaste – 20070226 – VVDW
    ' Keyboard key
    '   - 
    ' Proposal (To Do)
    '   - List of actions that can be added to the functionality
    '***

    cmd07.Text = mchrToken
    cmd07.Enabled = False
    CheckWinner()
  End Sub
  ' cmd07_Click(System.Object, System.EventArgs) Handles cmd07.Click

  Private Sub cmd08_Click(ByVal theSender As System.Object, ByVal theEventArguments As System.EventArgs) Handles cmd08.Click
    '***
    ' Action
    '   - Set the text of the button to a character (mchrToken)
    '   - Disable the button
    ' Called by
    '   - User action (Clicking a button)
    ' Calls
    '   - CheckWinner()
    ' Created
    '   - CopyPaste – 20070226 – VVDW
    ' Changed
    '   - Organisation – yyyymmdd – Initials of programmer – What changed
    ' Tested
    '   - CopyPaste – 20070226 – VVDW
    ' Keyboard key
    '   - 
    ' Proposal (To Do)
    '   - List of actions that can be added to the functionality
    '***

    cmd08.Text = mchrToken
    cmd08.Enabled = False
    CheckWinner()
  End Sub
  ' cmd08_Click(System.Object, System.EventArgs) Handles cmd08.Click

  Private Sub cmd09_Click(ByVal theSender As System.Object, ByVal theEventArguments As System.EventArgs) Handles cmd09.Click
    '***
    ' Action
    '   - Set the text of the button to a character (mchrToken)
    '   - Disable the button
    ' Called by
    '   - User action (Clicking a button)
    ' Calls
    '   - CheckWinner()
    ' Created
    '   - CopyPaste – 20070226 – VVDW
    ' Changed
    '   - Organisation – yyyymmdd – Initials of programmer – What changed
    ' Tested
    '   - CopyPaste – 20070226 – VVDW
    ' Keyboard key
    '   - 
    ' Proposal (To Do)
    '   - List of actions that can be added to the functionality
    '***

    cmd09.Text = mchrToken
    cmd09.Enabled = False
    CheckWinner()
  End Sub
  ' cmd09_Click(System.Object, System.EventArgs) Handles cmd09.Click

  Private Sub cmdQuit_Click(ByVal theSender As System.Object, ByVal theEventArguments As System.EventArgs) Handles cmdQuit.Click
    '***
    ' Action
    '   - End the program
    ' Called by
    '   - User action (Clicking a button)
    ' Calls
    '   - 
    ' Created
    '   - CopyPaste – 20070226 – VVDW
    ' Changed
    '   - Organisation – yyyymmdd – Initials of programmer – What changed
    ' Tested
    '   - CopyPaste – 20070226 – VVDW
    ' Keyboard key
    '   - 
    ' Proposal (To Do)
    '   - List of actions that can be added to the functionality
    '***

    End
  End Sub
  ' cmdQuit_Click(System.Object, System.EventArgs) Handles cmdQuit.Click

  Private Sub cmdRestart_Click(ByVal theSender As System.Object, ByVal theEventArguments As System.EventArgs) Handles cmdRestart.Click
    '***
    ' Action
    '   - Restart the game
    ' Called by
    '   - User action (Clicking a button)
    ' Calls
    '   - RestartGame()
    ' Created
    '   - CopyPaste – 20070226 – VVDW
    ' Changed
    '   - Organisation – yyyymmdd – Initials of programmer – What changed
    ' Tested
    '   - CopyPaste – 20070226 – VVDW
    ' Keyboard key
    '   - 
    ' Proposal (To Do)
    '   - List of actions that can be added to the functionality
    '***

    RestartGame()
  End Sub
  ' cmdRestart_Click(System.Object, System.EventArgs) Handles cmdRestart.Click

  Private Sub frmTicTacToe_Load(ByVal theSender As System.Object, ByVal theEventArguments As System.EventArgs) Handles MyBase.Load
    '***
    ' Action
    '   - Start the game
    ' Called by
    '   - User action (Loading the form)
    ' Calls
    '   - RestartGame()
    ' Created
    '   - CopyPaste – 20070226 – VVDW
    ' Changed
    '   - Organisation – yyyymmdd – Initials of programmer – What changed
    ' Tested
    '   - CopyPaste – 20070226 – VVDW
    ' Keyboard key
    '   - 
    ' Proposal (To Do)
    '   - List of actions that can be added to the functionality
    '***

    RestartGame()
  End Sub
  ' frmTicTacToe_Load(System.Object, System.EventArgs) Handles MyBase.Load

#End Region

#Region "Functionality"

  Private Sub CheckWinner()
    '***
    ' Action
    '   - Check each row, column and diagonal if the 3 texts are filled
    '   - Check each row, column and diagonal if the 3 texts are equal
    '   - If so
    '     - Put all with a yellow background
    '     - Display the info of the winner
    '   - If not
    '     - Game is not ended yet, and a next player is assigned
    ' Called by
    '   - cmd01_Click(System.Object, System.EventArgs) Handles cmd01.Click
    '   - cmd02_Click(System.Object, System.EventArgs) Handles cmd02.Click
    '   - cmd03_Click(System.Object, System.EventArgs) Handles cmd03.Click
    '   - cmd04_Click(System.Object, System.EventArgs) Handles cmd04.Click
    '   - cmd05_Click(System.Object, System.EventArgs) Handles cmd05.Click
    '   - cmd06_Click(System.Object, System.EventArgs) Handles cmd06.Click
    '   - cmd07_Click(System.Object, System.EventArgs) Handles cmd07.Click
    '   - cmd08_Click(System.Object, System.EventArgs) Handles cmd08.Click
    '   - cmd09_Click(System.Object, System.EventArgs) Handles cmd09.Click
    ' Calls
    '   - DisplayWinner()
    '   - NextPlayer()
    ' Created
    '   - CopyPaste – 20070226 – VVDW
    ' Changed
    '   - Organisation – yyyymmdd – Initials of programmer – What changed
    ' Tested
    '   - CopyPaste – 20070226 – VVDW
    ' Keyboard key
    '   - 
    ' Proposal (To Do)
    '   - List of actions that can be added to the functionality
    '***

    If (cmd01.Text & cmd02.Text & cmd03.Text).Length > 0 And _
        cmd01.Text = cmd02.Text And cmd02.Text = cmd03.Text Then
      cmd01.BackColor = Color.Yellow
      cmd02.BackColor = Color.Yellow
      cmd03.BackColor = Color.Yellow
      DisplayWinner()
    ElseIf (cmd04.Text & cmd05.Text & cmd06.Text).Length > 0 And _
        cmd04.Text = cmd05.Text And cmd05.Text = cmd06.Text Then
      cmd04.BackColor = Color.Yellow
      cmd05.BackColor = Color.Yellow
      cmd06.BackColor = Color.Yellow
      DisplayWinner()
    ElseIf (cmd07.Text & cmd08.Text & cmd09.Text).Length > 0 And _
        cmd07.Text = cmd08.Text And cmd08.Text = cmd09.Text Then
      cmd07.BackColor = Color.Yellow
      cmd08.BackColor = Color.Yellow
      cmd09.BackColor = Color.Yellow
      DisplayWinner()
    ElseIf (cmd01.Text & cmd04.Text & cmd07.Text).Length > 0 And _
        cmd01.Text = cmd04.Text And cmd04.Text = cmd07.Text Then
      cmd01.BackColor = Color.Yellow
      cmd04.BackColor = Color.Yellow
      cmd07.BackColor = Color.Yellow
      DisplayWinner()
    ElseIf (cmd02.Text & cmd05.Text & cmd08.Text).Length > 0 And _
        cmd02.Text = cmd05.Text And cmd05.Text = cmd08.Text Then
      cmd02.BackColor = Color.Yellow
      cmd05.BackColor = Color.Yellow
      cmd08.BackColor = Color.Yellow
      DisplayWinner()
    ElseIf (cmd03.Text & cmd06.Text & cmd09.Text).Length > 0 And _
        cmd03.Text = cmd06.Text And cmd06.Text = cmd09.Text Then
      cmd03.BackColor = Color.Yellow
      cmd06.BackColor = Color.Yellow
      cmd09.BackColor = Color.Yellow
      DisplayWinner()
    ElseIf (cmd01.Text & cmd05.Text & cmd09.Text).Length > 0 And _
        cmd01.Text = cmd05.Text And cmd05.Text = cmd09.Text Then
      cmd01.BackColor = Color.Yellow
      cmd05.BackColor = Color.Yellow
      cmd09.BackColor = Color.Yellow
      DisplayWinner()
    ElseIf (cmd03.Text & cmd05.Text & cmd07.Text).Length > 0 And _
        cmd03.Text = cmd05.Text And cmd05.Text = cmd07.Text Then
      cmd03.BackColor = Color.Yellow
      cmd05.BackColor = Color.Yellow
      cmd07.BackColor = Color.Yellow
      DisplayWinner()
    Else
      ' Each row, each column and diagonals haven't equal text
      NextPlayer()
    End If
    ' Check each row, each column and diagonals on equal text

  End Sub
  ' CheckWinner()

  Private Sub DisplayWinner()
    '***
    ' Action
    '   - When the first player has played
    '     - Second player gets an "X"
    '   - When the second player has played
    '     - First player gets an "O"
    '   - Set information on who plays on the screen
    ' Called by
    '   - CheckWinner()
    ' Calls
    '   - 
    ' Created
    '   - CopyPaste – 20070226 – VVDW
    ' Changed
    '   - Organisation – yyyymmdd – Initials of programmer – What changed
    ' Tested
    '   - CopyPaste – 20070226 – VVDW
    ' Keyboard key
    '   - 
    ' Proposal (To Do)
    '   - List of actions that can be added to the functionality
    '***

    lblPlayer.Text = "Speler '" & mchrToken & "' is de winnaar!"
    cmd01.Enabled = False
    cmd02.Enabled = False
    cmd03.Enabled = False
    cmd04.Enabled = False
    cmd05.Enabled = False
    cmd06.Enabled = False
    cmd07.Enabled = False
    cmd08.Enabled = False
    cmd09.Enabled = False
  End Sub
  ' DisplayWinner()

  Private Sub NextPlayer()
    '***
    ' Action
    '   - When the first player has played
    '     - Second player gets an "X"
    '   - When the second player has played
    '     - First player gets an "O"
    '   - Set information on who plays on the screen
    ' Called by
    '   - CheckWinner()
    '   - RestartGame()
    ' Calls
    '   - 
    ' Created
    '   - CopyPaste – 20070226 – VVDW
    ' Changed
    '   - Organisation – yyyymmdd – Initials of programmer – What changed
    ' Tested
    '   - CopyPaste – 20070226 – VVDW
    ' Keyboard key
    '   - 
    ' Proposal (To Do)
    '   - List of actions that can be added to the functionality
    '***

    If mlngPlayer = 1 Then
      mchrToken = "X"c
      mlngPlayer = 2
    Else
      ' mlngPlayer <> 1
      mchrToken = "O"c
      mlngPlayer = 1
    End If
    ' mlngPlayer = 1

    lblPlayer.Text = "Speler " & mlngPlayer & " : '" & mchrToken & "'"
  End Sub
  ' NextPlayer()

  Private Sub RestartGame()
    '***
    ' Action
    '   - Start or Restart the game by setting the defaults
    '   - All buttons enabled
    '   - All texts are cleared
    '   - All backcolors are set to default
    '   - Second player starts with an "O"
    '   - Run the next player routine
    ' Called by
    '   - cmdRestart_Click(System.Object, System.EventArgs) Handles cmdRestart.Click
    '   - frmTicTacToe_Load(System.Object, System.EventArgs) Handles MyBase.Load
    ' Calls
    '   - NextPlayer()
    ' Created
    '   - CopyPaste – 20070226 – VVDW
    ' Changed
    '   - Organisation – yyyymmdd – Initials of programmer – What changed
    ' Tested
    '   - CopyPaste – 20070226 – VVDW
    ' Keyboard key
    '   - 
    ' Proposal (To Do)
    '   - List of actions that can be added to the functionality
    '***

    cmd01.Enabled = True
    cmd02.Enabled = True
    cmd03.Enabled = True
    cmd04.Enabled = True
    cmd05.Enabled = True
    cmd06.Enabled = True
    cmd07.Enabled = True
    cmd08.Enabled = True
    cmd09.Enabled = True
    cmd01.Text = ""
    cmd02.Text = ""
    cmd03.Text = ""
    cmd04.Text = ""
    cmd05.Text = ""
    cmd06.Text = ""
    cmd07.Text = ""
    cmd08.Text = ""
    cmd09.Text = ""
    cmd01.BackColor = Color.LimeGreen
    cmd02.BackColor = Color.LimeGreen
    cmd03.BackColor = Color.LimeGreen
    cmd04.BackColor = Color.LimeGreen
    cmd05.BackColor = Color.LimeGreen
    cmd06.BackColor = Color.LimeGreen
    cmd07.BackColor = Color.LimeGreen
    cmd08.BackColor = Color.LimeGreen
    cmd09.BackColor = Color.LimeGreen

    mlngPlayer = 2
    mchrToken = "O"c
    NextPlayer()
  End Sub
  ' RestartGame()

  '#Region "Event"
  '#End Region

  '#Region "Sub / Function"
  '#End Region

#End Region

#End Region

End Class
' frmTicTacToe