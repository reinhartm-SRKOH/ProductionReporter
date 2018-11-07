Imports Oracle.ManagedDataAccess.Client
Imports Oracle.ManagedDataAccess.Types

Public Class MainForm



    Dim currentActivity As String = ""

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Call Reset_Form()

    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click

        Call Reset_Form()

    End Sub

    Private Sub Reset_Form()
        ButtonReport.Hide()

        LabelInput1.Hide()
        LabelInput2.Hide()
        LabelInput3.Hide()
        LabelInput4.Hide()
        LabelInput5.Hide()
        LabelInput6.Hide()
        LabelInput7.Hide()
        LabelInput8.Hide()
        LabelInput9.Hide()
        LabelInput10.Hide()
        LabelInput11.Hide()
        LabelInput12.Hide()

        TextBoxInput1.Hide()
        TextBoxInput2.Hide()
        TextBoxInput3.Hide()
        TextBoxInput4.Hide()
        TextBoxInput5.Hide()
        TextBoxInput6.Hide()
        TextBoxInput7.Hide()
        TextBoxInput8.Hide()
        TextBoxInput9.Hide()
        TextBoxInput10.Hide()
        TextBoxInput11.Hide()
        TextBoxInput12.Hide()

        TextBoxInput1.Enabled = True
        TextBoxInput2.Enabled = True
        TextBoxInput3.Enabled = True
        TextBoxInput4.Enabled = True
        TextBoxInput5.Enabled = True
        TextBoxInput6.Enabled = True
        TextBoxInput7.Enabled = True
        TextBoxInput8.Enabled = True
        TextBoxInput9.Enabled = True
        TextBoxInput10.Enabled = True
        TextBoxInput11.Enabled = True
        TextBoxInput12.Enabled = True

        RadioButtonShift1.Checked = False
        RadioButtonShift2.Checked = False
        RadioButtonShift3.Checked = False

        ' Change the header to the company name.
        ' The company name is determined by the value you enter for the Company setting in the app.config file.
        If My.Settings.Company = "SRKOH" Then
            LabelMenuDesc.Text = "SumiRiko Ohio, Inc."
            PictureBox1.Image = Image.FromFile("C:\Users\reinhartm\source\repos\ProductionReporting\ProductionReporting\images\SRK-OH.jpg")
        ElseIf My.Settings.Company = "SRKTN" Then
            LabelMenuDesc.Text = "SumiRiko Tennessee, Inc."
            PictureBox1.Image = Image.FromFile("C:\Users\reinhartm\source\repos\ProductionReporting\ProductionReporting\images\SRK-TN.jpg")
        End If

    End Sub

    Private Sub ButtonProduction_Click(sender As Object, e As EventArgs) Handles ButtonProduction.Click

        Call Reset_Form()

        currentActivity = "production"

        LabelInput1.Text = "Site:"
        LabelInput2.Text = "Associate:"
        LabelInput3.Text = "Part Number:"
        LabelInput4.Text = "Machine Number:"
        LabelInput5.Text = "Cavity Count:"
        LabelInput7.Text = "Shots:"
        LabelInput8.Text = "Free Shots:"
        LabelInput9.Text = "Total Parts Made:"
        LabelInput10.Text = "Scrapped Pieces:"
        LabelInput11.Text = "Total Good Production:"

        LabelMenuDesc.Text = "Production"

        ButtonReport.Show()

        LabelInput1.Show()
        LabelInput2.Show()
        LabelInput3.Show()
        LabelInput4.Show()
        LabelInput5.Show()
        LabelInput7.Show()
        LabelInput8.Show()
        LabelInput9.Show()
        LabelInput10.Show()
        LabelInput11.Show()

        TextBoxInput1.Show()
        TextBoxInput2.Show()
        TextBoxInput3.Show()
        TextBoxInput4.Show()
        TextBoxInput5.Show()
        TextBoxInput7.Show()
        TextBoxInput8.Show()
        TextBoxInput9.Show()
        TextBoxInput10.Show()
        TextBoxInput11.Show()

        TextBoxInput2.Select()

        Default_Shift()

        TextBoxInput1.Text = My.Settings.Site
        TextBoxInput1.Enabled = False
        TextBoxInput9.Enabled = False
        TextBoxInput11.Enabled = False

    End Sub

    Private Sub ButtonDowntime_Click(sender As Object, e As EventArgs) Handles ButtonDowntime.Click

        Call Reset_Form()

        currentActivity = "downtime"

        LabelInput1.Text = "Site:"
        LabelInput2.Text = "Associate:"
        LabelInput3.Text = "Part Number:"
        LabelInput4.Text = "Production Line:"
        LabelInput7.Text = "Downtime Code:"
        LabelInput8.Text = "Downtime Reason:"
        LabelInput9.Text = "Downtime:"

        LabelMenuDesc.Text = "Downtime"

        ButtonReport.Show()

        LabelInput1.Show()
        LabelInput2.Show()
        LabelInput3.Show()
        LabelInput4.Show()
        LabelInput7.Show()
        LabelInput8.Show()
        LabelInput9.Show()

        TextBoxInput1.Show()
        TextBoxInput2.Show()
        TextBoxInput3.Show()
        TextBoxInput4.Show()
        TextBoxInput7.Show()
        TextBoxInput8.Show()
        TextBoxInput9.Show()

        TextBoxInput2.Select()

        Default_Shift()

        TextBoxInput1.Text = My.Settings.Site
        TextBoxInput1.Enabled = False

    End Sub

    Private Sub ButtonScrap_Click(sender As Object, e As EventArgs) Handles ButtonScrap.Click

        Call Reset_Form()

        currentActivity = "scrap"

        LabelInput1.Text = "Site:"
        LabelInput2.Text = "Associate:"
        LabelInput3.Text = "Part Number:"
        LabelInput4.Text = "Production Line:"
        LabelInput7.Text = "Scrap Code:"
        LabelInput8.Text = "Scrap Reason:"
        LabelInput9.Text = "Scrap Quantity:"

        LabelMenuDesc.Text = "Scrap"

        ButtonReport.Show()

        LabelInput1.Show()
        LabelInput2.Show()
        LabelInput3.Show()
        LabelInput4.Show()
        LabelInput7.Show()
        LabelInput8.Show()
        LabelInput9.Show()

        TextBoxInput1.Show()
        TextBoxInput2.Show()
        TextBoxInput3.Show()
        TextBoxInput4.Show()
        TextBoxInput7.Show()
        TextBoxInput8.Show()
        TextBoxInput9.Show()

        TextBoxInput2.Select()

        Default_Shift()

        TextBoxInput1.Text = My.Settings.Site
        TextBoxInput1.Enabled = False

    End Sub

    Private Sub ButtonReport_Click(sender As Object, e As EventArgs) Handles ButtonReport.Click

        Select Case currentActivity
            Case "production"
                Call Report_Production()
            Case "downtime"
                Call Report_Downtime()
            Case "scrap"
                Call Report_Scrap()
        End Select

    End Sub

    ' Automatically select the correct radio button shift based on the current time.
    Private Sub Default_Shift()
        Dim currentTime As String = DateTime.Now.ToString("HH:mm:ss")

        Select Case currentTime
            Case "070000" To "150000"
                'First shift
                RadioButtonShift1.Checked = True
            Case "150000" To "230000"
                'Second shift
                RadioButtonShift2.Checked = True
            Case Else
                'Third shift
                RadioButtonShift3.Checked = True
        End Select
    End Sub

    Function Shift_Check()
        Dim shift As String

        If RadioButtonShift1.Checked Then
            shift = "1"
        ElseIf RadioButtonShift2.Checked Then
            shift = "2"
        ElseIf RadioButtonShift3.Checked Then
            shift = "3"
        Else
            shift = "0"
            MsgBox("ERROR: No Shift Was Selected!")
        End If

        Return shift
    End Function

    Private Sub Validate_Part_Num(part_num As String)
        ' Connect to oracle database and validate the user inputed part number and production line.



    End Sub

    Private Sub TextBoxInput5_TextChanged(sender As Object, e As EventArgs) Handles TextBoxInput5.Leave
        Select Case currentActivity
            Case "production"
                Update_Production_Fields()
            Case "downtime"
                '
            Case "scrap"
                '
        End Select
    End Sub

    Private Sub TextBoxInput7_TextChanged(sender As Object, e As EventArgs) Handles TextBoxInput7.Leave
        Select Case currentActivity
            Case "production"
                Update_Production_Fields()
            Case "downtime"
                '
            Case "scrap"
                '
        End Select
    End Sub

    Private Sub TextBoxInput8_TextChanged(sender As Object, e As EventArgs) Handles TextBoxInput8.Leave
        Select Case currentActivity
            Case "production"
                Update_Production_Fields()
            Case "downtime"
                '
            Case "scrap"
                '
        End Select
    End Sub

    Private Sub TextBoxInput10_TextChanged(sender As Object, e As EventArgs) Handles TextBoxInput10.Leave
        Select Case currentActivity
            Case "production"
                Update_Production_Fields()
            Case "downtime"
                '
            Case "scrap"
                '
        End Select
    End Sub

    Private Sub Update_Production_Fields()
        Dim totalPartsMade As Integer
        Dim totalGoodPartsMade As Integer
        Dim cavityCount As Integer
        Dim shots As Integer
        Dim freeShots As Integer
        Dim scrappedPieces As Integer

        ' Validate textbox data.  If none is entered, treat it as 0.
        If String.IsNullOrEmpty(TextBoxInput5.Text) Then
            cavityCount = 0
        Else
            cavityCount = Convert.ToInt32(TextBoxInput5.Text)
        End If

        If String.IsNullOrEmpty(TextBoxInput7.Text) Then
            shots = 0
        Else
            shots = Convert.ToInt32(TextBoxInput7.Text)
        End If

        If String.IsNullOrEmpty(TextBoxInput8.Text) Then
            freeShots = 0
        Else
            freeShots = Convert.ToInt32(TextBoxInput8.Text)
        End If

        If String.IsNullOrEmpty(TextBoxInput10.Text) Then
            scrappedPieces = 0
        Else
            scrappedPieces = Convert.ToInt32(TextBoxInput10.Text)
        End If


        totalPartsMade = (cavityCount * shots) - (cavityCount * freeShots)
        totalGoodPartsMade = totalPartsMade - scrappedPieces

        TextBoxInput9.Text = totalPartsMade
        TextBoxInput11.Text = totalGoodPartsMade

    End Sub

    Private Sub Report_Production()
        ' Implement reporting production code here

        Dim shift As String
        Dim reportDate As String
        Dim associate As String
        Dim prodLine As String
        Dim partNumber As String
        Dim quantity As String
        Dim reportTime As String
        Dim radleyString As String

        shift = Shift_Check()
        If shift = "0" Then
            MsgBox("ERROR: No Shift Was Selected")
            Exit Sub
        End If

        'If RadioButtonShift1.Checked Then
        '    shift = "1"
        'ElseIf RadioButtonShift2.Checked Then
        '    shift = "2"
        'ElseIf RadioButtonShift3.Checked Then
        '    shift = "3"
        'Else
        '    MsgBox("ERROR: A shift must be selected!")
        '    Exit Sub
        'End If

        reportDate = Date.Now.ToString("MMddyyyy")
        associate = TextBoxInput3.Text
        prodLine = TextBoxInput2.Text
        partNumber = TextBoxInput5.Text
        quantity = TextBoxInput12.Text
        reportTime = Date.Now.ToString("MMddyyyy HH:mm")

        ' Example Data: 1|11052018|2829|AS11|21671A1-AD|76|11052018 10:42|||||<CR>
        radleyString = shift + "|" + reportDate + "|" + associate + "|" + prodLine + "|" + partNumber + "|" + quantity + "|" + reportTime + "|" + "|" + "|" + "|" + "|" + "<CR>"

        ' create a text file containing the radleyString and place it into the appropriate Radley Production folder for watchdog to pick it up.
        Call WriteToFile(radleyString)

    End Sub

    Private Sub Report_Downtime()
        ' Implement reporting downtime code here
        MessageBox.Show("Feature Coming Soon!")

    End Sub

    Private Sub Report_Scrap()
        ' Implement reporting scrap code here
        MessageBox.Show("Feature Coming Soon!")

    End Sub

    Sub WriteToFile(ByVal radleyString As String)
        Dim filePath As String = My.Settings.DropFolder
        Dim oFile As System.IO.FileStream = Nothing
        Dim oWrite As System.IO.StreamWriter = Nothing
        Dim fileName As String
        Dim shift As String
        Dim reportDateTime As String

        reportDateTime = Date.Now.ToString("MM.dd.yyyy.HHmmss")

        shift = Shift_Check()
        If shift = "0" Then
            MsgBox("ERROR: No Shift Was Selected")
            Exit Sub
        End If

        ' Example File Name: 32001.AS11LiveProdReporter.1.11.05.2018.101242.txt
        fileName = My.Settings.Site + "." + My.Settings.ProductionLine + "LiveProdReporter." + shift + "." + reportDateTime + ".txt"

        filePath = filePath + "\" + fileName

        oFile = New System.IO.FileStream(filePath, IO.FileMode.Create, IO.FileAccess.Write)
        oWrite = New System.IO.StreamWriter(oFile)

        ' You can either use WRITE or WRITELINE method
        oWrite.WriteLine(radleyString)


        ' Close the filestream and streamwriter
        oWrite.Close()
        oFile.Close()
    End Sub

End Class
