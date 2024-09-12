<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form은 Dispose를 재정의하여 구성 요소 목록을 정리합니다.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows Form 디자이너에 필요합니다.
    Private components As System.ComponentModel.IContainer

    '참고: 다음 프로시저는 Windows Form 디자이너에 필요합니다.
    '수정하려면 Windows Form 디자이너를 사용하십시오.  
    '코드 편집기에서는 수정하지 마세요.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.button_mdb_read = New System.Windows.Forms.Button()
        Me.testText = New System.Windows.Forms.RichTextBox()
        Me.button_end_work = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.mdbsFilePath = New System.Windows.Forms.TextBox()
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.mdbPassword = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.odbsDsnText = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.ContextMenuStrip2 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.IntervalSelected = New System.Windows.Forms.ListBox()
        Me.DateTimePicker = New System.Windows.Forms.DateTimePicker()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Timer1
        '
        '
        'button_mdb_read
        '
        Me.button_mdb_read.Location = New System.Drawing.Point(73, 541)
        Me.button_mdb_read.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.button_mdb_read.Name = "button_mdb_read"
        Me.button_mdb_read.Size = New System.Drawing.Size(229, 40)
        Me.button_mdb_read.TabIndex = 0
        Me.button_mdb_read.Text = "작업 시작"
        Me.button_mdb_read.UseVisualStyleBackColor = True
        '
        'testText
        '
        Me.testText.Location = New System.Drawing.Point(597, 152)
        Me.testText.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.testText.Name = "testText"
        Me.testText.Size = New System.Drawing.Size(629, 428)
        Me.testText.TabIndex = 1
        Me.testText.Text = ""
        '
        'button_end_work
        '
        Me.button_end_work.Location = New System.Drawing.Point(337, 541)
        Me.button_end_work.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.button_end_work.Name = "button_end_work"
        Me.button_end_work.Size = New System.Drawing.Size(223, 40)
        Me.button_end_work.TabIndex = 2
        Me.button_end_work.Text = "작업 종료"
        Me.button_end_work.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(70, 155)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(109, 15)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "MDB 파일 경로"
        '
        'mdbsFilePath
        '
        Me.mdbsFilePath.Location = New System.Drawing.Point(258, 155)
        Me.mdbsFilePath.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.mdbsFilePath.Name = "mdbsFilePath"
        Me.mdbsFilePath.Size = New System.Drawing.Size(302, 25)
        Me.mdbsFilePath.TabIndex = 4
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(61, 4)
        '
        'mdbPassword
        '
        Me.mdbPassword.Location = New System.Drawing.Point(258, 230)
        Me.mdbPassword.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.mdbPassword.Name = "mdbPassword"
        Me.mdbPassword.Size = New System.Drawing.Size(302, 25)
        Me.mdbPassword.TabIndex = 6
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(70, 232)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(94, 15)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "MDB 암호키 "
        '
        'odbsDsnText
        '
        Me.odbsDsnText.Location = New System.Drawing.Point(258, 304)
        Me.odbsDsnText.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.odbsDsnText.Name = "odbsDsnText"
        Me.odbsDsnText.Size = New System.Drawing.Size(302, 25)
        Me.odbsDsnText.TabIndex = 8
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(70, 311)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(98, 15)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "ODBC DSN명"
        '
        'ContextMenuStrip2
        '
        Me.ContextMenuStrip2.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.ContextMenuStrip2.Name = "ContextMenuStrip2"
        Me.ContextMenuStrip2.Size = New System.Drawing.Size(61, 4)
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("굴림", 25.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label4.Location = New System.Drawing.Point(395, 35)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(690, 43)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "KT 출퇴근 연동 서비스　２ｖｅｒ"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(70, 389)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(99, 15)
        Me.Label5.TabIndex = 12
        Me.Label5.Text = "주기 시간(초)"
        '
        'IntervalSelected
        '
        Me.IntervalSelected.FormattingEnabled = True
        Me.IntervalSelected.ItemHeight = 15
        Me.IntervalSelected.Items.AddRange(New Object() {"1", "5", "10", "30", "60", "1800", "3600"})
        Me.IntervalSelected.Location = New System.Drawing.Point(258, 379)
        Me.IntervalSelected.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.IntervalSelected.Name = "IntervalSelected"
        Me.IntervalSelected.Size = New System.Drawing.Size(302, 94)
        Me.IntervalSelected.TabIndex = 14
        '
        'DateTimePicker
        '
        Me.DateTimePicker.Location = New System.Drawing.Point(258, 489)
        Me.DateTimePicker.Name = "DateTimePicker"
        Me.DateTimePicker.Size = New System.Drawing.Size(200, 25)
        Me.DateTimePicker.TabIndex = 15
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(70, 496)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(37, 15)
        Me.Label6.TabIndex = 16
        Me.Label6.Text = "날짜"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1271, 651)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.DateTimePicker)
        Me.Controls.Add(Me.IntervalSelected)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.odbsDsnText)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.mdbPassword)
        Me.Controls.Add(Me.mdbsFilePath)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.button_end_work)
        Me.Controls.Add(Me.testText)
        Me.Controls.Add(Me.button_mdb_read)
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Timer1 As Timer
    Friend WithEvents button_mdb_read As Button
    Friend WithEvents testText As RichTextBox
    Friend WithEvents button_end_work As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents mdbsFilePath As TextBox
    Friend WithEvents ContextMenuStrip1 As ContextMenuStrip
    Friend WithEvents mdbPassword As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents odbsDsnText As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents ContextMenuStrip2 As ContextMenuStrip
    Friend WithEvents Label4 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents IntervalSelected As ListBox
    Friend WithEvents DateTimePicker As DateTimePicker
    Friend WithEvents Label6 As Label
End Class
