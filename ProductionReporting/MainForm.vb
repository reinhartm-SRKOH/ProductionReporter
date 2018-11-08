Imports System.Data.OracleClient

Public Class MainForm

    Dim prodData As New FormData
    Dim downtimeData As New FormData
    Dim scrapData As New FormData

    Dim currentActivity As String
    Private Const production As String = "production"
    Private Const downtime As String = "downtime"
    Private Const scrap As String = "scrap"


    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Call Reset_Form()

    End Sub


    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click

        ' This will prevent the saved FormData from being cleared in the copy_form_data().  Without it, it's cleared
        '   because we end up copying the textbox data to the class, but all fields are empty on the home screen.
        currentActivity = "home"

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

        TextBoxInput1.Text = ""
        TextBoxInput2.Text = ""
        TextBoxInput3.Text = ""
        TextBoxInput4.Text = ""
        TextBoxInput5.Text = ""
        TextBoxInput6.Text = ""
        TextBoxInput7.Text = ""
        TextBoxInput8.Text = ""
        TextBoxInput9.Text = ""
        TextBoxInput10.Text = ""
        TextBoxInput11.Text = ""
        TextBoxInput12.Text = ""

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

        ' For maintaining current textbox data.
        '   Must be called before Reset_Form() and before currentActivity value modified.
        Call Copy_Form_Data()

        Call Reset_Form()

        currentActivity = production

        Call Update_Form_Fields()

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

        Call Default_Shift()

        TextBoxInput1.Text = My.Settings.Site
        TextBoxInput1.Enabled = False
        TextBoxInput9.Enabled = False
        TextBoxInput11.Enabled = False

    End Sub


    Private Sub ButtonDowntime_Click(sender As Object, e As EventArgs) Handles ButtonDowntime.Click

        ' For maintaining current textbox data.
        '   Must be called before Reset_Form() and before currentActivity value modified.
        Call Copy_Form_Data()

        Call Reset_Form()

        currentActivity = downtime

        Call Update_Form_Fields()

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

        Call Default_Shift()

        TextBoxInput1.Text = My.Settings.Site
        TextBoxInput1.Enabled = False

    End Sub


    Private Sub ButtonScrap_Click(sender As Object, e As EventArgs) Handles ButtonScrap.Click

        ' For maintaining current textbox data.
        '   Must be called before Reset_Form() and before currentActivity value modified.
        Call Copy_Form_Data()

        Call Reset_Form()

        currentActivity = scrap

        Call Update_Form_Fields()

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

        Call Default_Shift()

        TextBoxInput1.Text = My.Settings.Site
        TextBoxInput1.Enabled = False

    End Sub


    Private Sub ButtonReport_Click(sender As Object, e As EventArgs) Handles ButtonReport.Click
        Dim radleyString As String = ""

        If Shift_Check() = "0" Then
            MsgBox("ERROR: No Shift Was Selected")
            Exit Sub
        End If

        Call Copy_Form_Data()

        '    Below creates the string of data that watchdog processes to report production, downtime, and scrap.  The string is a little different
        ' depending on which one you are reporting.  Below explains each section of the string.  Production Line, Part Number, and Prodcution Quantity
        ' are required fields for watchdog.  Since downtime & scrap do not report production & it's a required field, we just put 0 in for that section.
        '    SHIFT|DATE|ASSOCIATE|PRODUCTION LINE|PART NUMBER|PRODUCTION QTY|RUNTIME|DOWNTIME|DOWNTIME REASON|SCRAP|SCRAP REASON|<CR>
        Select Case currentActivity
            Case production
                ' Production string example: 1|11052018|2829|AS11|21671A1-AD|76|11052018 10:42|||||<CR>
                radleyString = Shift_Check() + "|" + Date.Now.ToString("MMddyyyy") + "|" + prodData.Associate + "|" + prodData.ProdLine + "|" + prodData.PartNumber + "|" + prodData.ProductionQty + "|" + Date.Now.ToString("MMddyyyy HH:mm") + "|" + "|" + "|" + "|" + "|" + "<CR>"

            Case downtime
                ' Downtime string example: 1|11052018|2829|AS11|21671A1-AD|0|11052018 10:42|5|66|||<CR>
                radleyString = Shift_Check() + "|" + Date.Now.ToString("MMddyyyy") + "|" + downtimeData.Associate + "|" + downtimeData.ProdLine + "|" + downtimeData.PartNumber + "|" + "0" + "|" + Date.Now.ToString("MMddyyyy HH:mm") + "|" + downtimeData.DowntimeQty + "|" + downtimeData.DowntimeCode + "|" + "|" + "|" + "<CR>"

            Case scrap
                ' Scrap string example: 1|11052018|2829|AS11|21671A1-AD|0|11052018 10:42|||2|66|<CR>
                radleyString = Shift_Check() + "|" + Date.Now.ToString("MMddyyyy") + "|" + scrapData.Associate + "|" + scrapData.ProdLine + "|" + scrapData.PartNumber + "|" + "0" + "|" + Date.Now.ToString("MMddyyyy HH:mm") + "|" + "|" + "|" + scrapData.ScrapQty + "|" + scrapData.ScrapCode + "|" + "<CR>"
        End Select

        Call WriteToFile(radleyString)
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


    Private Sub Copy_Form_Data()
        If currentActivity = production Then
            prodData.Associate = TextBoxInput2.Text
            prodData.PartNumber = TextBoxInput3.Text
            prodData.ProdLine = TextBoxInput4.Text
            prodData.CavityCount = TextBoxInput5.Text
            prodData.Shots = TextBoxInput7.Text
            prodData.FreeShots = TextBoxInput8.Text
            prodData.ScrappedPieces = TextBoxInput10.Text
            prodData.ProductionQty = TextBoxInput11.Text

        ElseIf currentActivity = downtime Then
            downtimeData.Associate = TextBoxInput2.Text
            downtimeData.PartNumber = TextBoxInput3.Text
            downtimeData.ProdLine = TextBoxInput4.Text
            downtimeData.DowntimeCode = TextBoxInput7.Text
            downtimeData.DowntimeReason = TextBoxInput8.Text
            downtimeData.DowntimeQty = TextBoxInput9.Text

        ElseIf currentActivity = scrap Then
            scrapData.Associate = TextBoxInput2.Text
            scrapData.PartNumber = TextBoxInput3.Text
            scrapData.ProdLine = TextBoxInput4.Text
            scrapData.ScrapCode = TextBoxInput7.Text
            scrapData.ScrapReason = TextBoxInput8.Text
            scrapData.ScrapQty = TextBoxInput9.Text

        End If

    End Sub


    Private Sub TextBoxInput5_TextChanged(sender As Object, e As EventArgs) Handles TextBoxInput5.Leave
        Dim result As Integer

        Select Case currentActivity
            Case production
                If Not String.IsNullOrEmpty(TextBoxInput5.Text) Then
                    If Integer.TryParse(TextBoxInput5.Text, result) Then
                        Call Update_Quantity_Fields()
                    Else
                        MsgBox("ERROR: Entry is not a valid format.")
                        TextBoxInput5.Select()
                    End If
                End If

            Case downtime
                ' N/A
            Case scrap
                ' N/A
        End Select
    End Sub


    Private Sub TextBoxInput7_TextChanged(sender As Object, e As EventArgs) Handles TextBoxInput7.Leave
        Dim result As Integer

        Select Case currentActivity
            Case production
                If Not String.IsNullOrEmpty(TextBoxInput7.Text) Then
                    If Integer.TryParse(TextBoxInput7.Text, result) Then
                        Call Update_Quantity_Fields()
                    Else
                        MsgBox("ERROR: Entry is not a valid format.")
                        TextBoxInput7.Select()
                    End If
                End If

            Case downtime
                ' N/A
            Case scrap
                ' N/A
        End Select
    End Sub


    Private Sub TextBoxInput8_TextChanged(sender As Object, e As EventArgs) Handles TextBoxInput8.Leave
        Dim result As Integer

        Select Case currentActivity
            Case production
                If Not String.IsNullOrEmpty(TextBoxInput8.Text) Then
                    If Integer.TryParse(TextBoxInput8.Text, result) Then
                        Call Update_Quantity_Fields()
                    Else
                        MsgBox("ERROR: Entry is not a valid format.")
                        TextBoxInput8.Select()
                    End If
                End If

            Case downtime
                ' N/A
            Case scrap
                ' N/A
        End Select
    End Sub


    Private Sub TextBoxInput9_TextChanged(sender As Object, e As EventArgs) Handles TextBoxInput9.Leave
        Dim result As Integer

        Select Case currentActivity
            Case production
                ' N/A
            Case downtime
                If Not String.IsNullOrEmpty(TextBoxInput9.Text) Then
                    If Not Integer.TryParse(TextBoxInput9.Text, result) Then
                        MsgBox("ERROR: Entry is not a valid format.")
                        TextBoxInput9.Select()
                    End If
                End If

            Case scrap
                If Not String.IsNullOrEmpty(TextBoxInput9.Text) Then
                    If Not Integer.TryParse(TextBoxInput9.Text, result) Then
                        MsgBox("ERROR: Entry is not a valid format.")
                        TextBoxInput9.Select()
                    End If
                End If

        End Select
    End Sub


    Private Sub TextBoxInput10_TextChanged(sender As Object, e As EventArgs) Handles TextBoxInput10.Leave
        Dim result As Integer

        Select Case currentActivity
            Case production
                If Not String.IsNullOrEmpty(TextBoxInput10.Text) Then
                    If Integer.TryParse(TextBoxInput10.Text, result) Then
                        Call Update_Quantity_Fields()
                    Else
                        MsgBox("ERROR: Entry is not a valid format.")
                        TextBoxInput10.Select()
                    End If
                End If

            Case downtime
                ' N/A
            Case scrap
                ' N/A
        End Select
    End Sub


    Private Sub Update_Quantity_Fields()
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


    Private Sub Update_Form_Fields()
        If currentActivity = production Then
            TextBoxInput2.Text = prodData.Associate()
            TextBoxInput3.Text = prodData.PartNumber()
            TextBoxInput4.Text = prodData.ProdLine()
            TextBoxInput5.Text = prodData.CavityCount()
            TextBoxInput7.Text = prodData.Shots()
            TextBoxInput8.Text = prodData.FreeShots()
            TextBoxInput10.Text = prodData.ScrappedPieces()

            Call Update_Quantity_Fields()

        ElseIf currentActivity = downtime Then
            TextBoxInput2.Text = downtimeData.Associate()
            TextBoxInput3.Text = downtimeData.PartNumber()
            TextBoxInput4.Text = downtimeData.ProdLine()
            TextBoxInput7.Text = downtimeData.DowntimeCode()
            TextBoxInput8.Text = downtimeData.DowntimeReason()
            TextBoxInput9.Text = downtimeData.DowntimeQty()

        ElseIf currentActivity = scrap Then
            TextBoxInput2.Text = scrapData.Associate()
            TextBoxInput3.Text = scrapData.PartNumber()
            TextBoxInput4.Text = scrapData.ProdLine()
            TextBoxInput7.Text = scrapData.ScrapCode()
            TextBoxInput8.Text = scrapData.ScrapReason()
            TextBoxInput9.Text = scrapData.ScrapQty()

        End If
    End Sub


    ' Create a text file, name it, write the radleyString data to it, and place it into the Radley Production folder for watchdog to pick it up.
    Sub WriteToFile(ByVal radleyString As String)
        Dim oFile As System.IO.FileStream = Nothing
        Dim oWrite As System.IO.StreamWriter = Nothing
        Dim filePath As String = My.Settings.DropFolder
        Dim fileName As String

        ' Create the file name.
        '   Example: "32001.AS11LiveProdReporter.1.11.05.2018.101242.txt"
        If currentActivity = production Then
            fileName = My.Settings.Site + "." + prodData.ProdLine + "LiveProdReporter." + Shift_Check() + "." + Date.Now.ToString("MM.dd.yyyy.HHmmss") + ".txt"
        ElseIf currentActivity = downtime Then
            fileName = My.Settings.Site + "." + downtimeData.ProdLine + "LiveProdReporter." + Shift_Check() + "." + Date.Now.ToString("MM.dd.yyyy.HHmmss") + ".txt"
        ElseIf currentActivity = scrap Then
            fileName = My.Settings.Site + "." + scrapData.ProdLine + "LiveProdReporter." + Shift_Check() + "." + Date.Now.ToString("MM.dd.yyyy.HHmmss") + ".txt"
        End If

        filePath = filePath + "\" + fileName

        oFile = New System.IO.FileStream(filePath, IO.FileMode.Create, IO.FileAccess.Write)
        oWrite = New System.IO.StreamWriter(oFile)

        ' Write to the file.
        oWrite.WriteLine(radleyString)

        ' Close the filestream and streamwriter
        oWrite.Close()
        oFile.Close()
    End Sub

End Class
