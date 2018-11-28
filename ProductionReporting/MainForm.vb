Imports System.Data.OracleClient
Imports System.IO

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
        ComboBox1.Hide()

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
        ComboBox1.Enabled = True

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
        ComboBox1.Items.Clear()

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
        LabelInput4.Text = "Production Line:"
        If My.Settings.EnterCuringProductionManually Then
            LabelInput5.Text = "Cavity Count:"
            LabelInput7.Text = "Shots:"
            LabelInput8.Text = "Free Shots:"
            LabelInput9.Text = "Total Parts Made:"
            LabelInput10.Text = "Scrapped Pieces:"
            LabelInput11.Text = "Total Good Production:"
        Else
            LabelInput5.Text = "Quantity:"
        End If

        LabelMenuDesc.Text = "Production"

        ButtonReport.Show()


        LabelInput1.Show()
        LabelInput2.Show()
        LabelInput3.Show()
        LabelInput4.Show()
        LabelInput5.Show()
        If My.Settings.EnterCuringProductionManually Then
            LabelInput7.Show()
            LabelInput8.Show()
            LabelInput9.Show()
            LabelInput10.Show()
            LabelInput11.Show()
        End If

        TextBoxInput1.Show()
        TextBoxInput2.Show()
        TextBoxInput3.Show()
        TextBoxInput4.Show()
        TextBoxInput5.Show()
        If My.Settings.EnterCuringProductionManually Then
            TextBoxInput7.Show()
            TextBoxInput8.Show()
            TextBoxInput9.Show()
            TextBoxInput10.Show()
            TextBoxInput11.Show()
        End If

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
        LabelInput9.Text = "Downtime Quantity:"

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
        ComboBox1.Show()
        TextBoxInput9.Show()

        TextBoxInput2.Select()

        Call Default_Shift()

        TextBoxInput1.Text = My.Settings.Site
        TextBoxInput1.Enabled = False

        Call Populate_ComboBox()

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
        ComboBox1.Show()
        TextBoxInput9.Show()

        TextBoxInput2.Select()

        Call Default_Shift()

        TextBoxInput1.Text = My.Settings.Site
        TextBoxInput1.Enabled = False

        Call Populate_ComboBox()

    End Sub


    Private Sub ButtonReport_Click(sender As Object, e As EventArgs) Handles ButtonReport.Click
        Dim radleyString As String = ""
        Dim shift As String

        If Shift_Check() = "0" Then
            MsgBox("ERROR: No Shift Was Selected")
            Exit Sub
        Else
            shift = Shift_Check()
        End If

        Call Copy_Form_Data()

        '    Below creates the string of data that watchdog processes to report production, downtime, and scrap.  The string is a little different
        ' depending on which one you are reporting.  Below explains each section of the string.  Production Line, Part Number, and Prodcution Quantity
        ' are required fields for watchdog.  Since downtime & scrap do not report production & it's a required field, we just put 0 in for that section.
        '    SHIFT|DATE|ASSOCIATE|PRODUCTION LINE|PART NUMBER|PRODUCTION QTY|RUNTIME|DOWNTIME|DOWNTIME REASON|SCRAP|SCRAP REASON|<CR>
        Select Case currentActivity
            Case production
                ' Production string example: 1|11052018|2829|AS11|21671A1-AD|76|11052018 10:42|||||<CR>
                radleyString = shift + "|" + Date.Now.ToString("MMddyyyy") + "|" + prodData.Associate + "|" + prodData.ProdLine + "|" + prodData.PartNumber + "|" + prodData.ProductionQty + "|" + Date.Now.ToString("MMddyyyy HH:mm") + "|" + "|" + "|" + "|" + "|" + "<CR>"

            Case downtime
                ' Downtime string example: 1|11052018|2829|AS11|21671A1-AD|0|11052018 10:42|5|66|||<CR>
                radleyString = shift + "|" + Date.Now.ToString("MMddyyyy") + "|" + downtimeData.Associate + "|" + downtimeData.ProdLine + "|" + downtimeData.PartNumber + "|" + "0" + "|" + Date.Now.ToString("MMddyyyy HH:mm") + "|" + downtimeData.DowntimeQty + "|" + downtimeData.DowntimeCode + "|" + "|" + "|" + "<CR>"

            Case scrap
                ' Scrap string example: 1|11052018|2829|AS11|21671A1-AD|0|11052018 10:42|||2|66|<CR>
                radleyString = shift + "|" + Date.Now.ToString("MMddyyyy") + "|" + scrapData.Associate + "|" + scrapData.ProdLine + "|" + scrapData.PartNumber + "|" + "0" + "|" + Date.Now.ToString("MMddyyyy HH:mm") + "|" + "|" + "|" + scrapData.ScrapQty + "|" + scrapData.ScrapCode + "|" + "<CR>"
        End Select


        If My.Settings.Debugging Then
            MsgBox(radleyString)
        Else
            Call WriteToFile(radleyString)
        End If

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
            If My.Settings.EnterCuringProductionManually Then
                prodData.CavityCount = TextBoxInput5.Text
                prodData.Shots = TextBoxInput7.Text
                prodData.FreeShots = TextBoxInput8.Text
                prodData.ScrappedPieces = TextBoxInput10.Text
                prodData.ProductionQty = TextBoxInput11.Text
            Else
                prodData.ProductionQty = TextBoxInput5.Text
            End If

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


    Private Sub Populate_ComboBox()
        Dim fileName As String = ""
        Dim site As String = ""
        Select Case currentActivity
            Case downtime
                fileName = "InterruptionCauses"
                site = My.Settings.Site
            Case scrap
                fileName = "ScrappingCauses"
        End Select

        Dim reader = New StreamReader(My.Settings.PathToDataValidationFiles + fileName + site + ".txt")
        Dim lineData As String
        Dim index As Integer
        Dim downDesc As String

        ComboBox1.Items.Clear()

        While Not reader.EndOfStream
            lineData = reader.ReadLine()
            index = lineData.IndexOf("|")
            downDesc = lineData.Substring(index + 1)

            ComboBox1.Items.Add(downDesc)
        End While
    End Sub


    Private Sub Populate_Downtime_Code()
        Dim reader = New StreamReader(My.Settings.PathToDataValidationFiles + "InterruptionCauses" + My.Settings.Site + ".txt")
        Dim lineData As String
        Dim index As Integer
        Dim dtCode As String
        Dim dtDesc As String

        While Not reader.EndOfStream
            lineData = reader.ReadLine()
            index = lineData.IndexOf("|")
            dtCode = lineData.Substring(0, index)
            dtDesc = lineData.Substring(index + 1)

            If dtDesc = ComboBox1.Text Then
                TextBoxInput7.Text = dtCode
                Exit While
            End If

        End While

    End Sub


    Private Sub Populate_Scrap_Code()
        Dim reader = New StreamReader(My.Settings.PathToDataValidationFiles + "ScrapCauses.txt")
        Dim lineData As String
        Dim index As Integer
        Dim dtCode As String
        Dim dtDesc As String

        While Not reader.EndOfStream
            lineData = reader.ReadLine()
            index = lineData.IndexOf("|")
            dtCode = lineData.Substring(0, index)
            dtDesc = lineData.Substring(index + 1)

            If dtDesc = ComboBox1.Text Then
                TextBoxInput7.Text = dtCode
                Exit While
            End If

        End While

    End Sub


    ' Part Number Textbox.
    Private Sub TextBoxInput3_Leave(sender As Object, e As EventArgs) Handles TextBoxInput3.Leave
        If Not String.IsNullOrEmpty(TextBoxInput3.Text) Then
            TextBoxInput3.Text = TextBoxInput3.Text.ToUpper()

            If Validate_Part_Number(TextBoxInput3.Text) Then
                ' Part Number is valid.

                If Not String.IsNullOrEmpty(TextBoxInput4.Text) Then
                    ' Production line field has data, verify it with the Part Number.

                    If Not Validate_Prod_Line_Part() Then
                        ' Part number is NOT setup for the entered production line.

                        MsgBox("ERROR: This Part Number is NOT setup for the entered Production Line.")
                        TextBoxInput3.Select()
                    End If
                End If
            Else
                ' Part Number is NOT valid.
                MsgBox("ERROR: Part Number is NOT Valid.")
                TextBoxInput3.Select()
            End If

        End If
    End Sub


    ' Production Line Textbox.
    Private Sub TextBoxInput4_Leave(sender As Object, e As EventArgs) Handles TextBoxInput4.Leave
        If Not String.IsNullOrEmpty(TextBoxInput4.Text) Then
            TextBoxInput4.Text = TextBoxInput4.Text.ToUpper()

            If Validate_Production_Line(TextBoxInput4.Text) Then
                ' Production Line is valid.

                If Not String.IsNullOrEmpty(TextBoxInput3.Text) Then
                    ' Part Number field has data, verify it with the Production Line.

                    If Not Validate_Prod_Line_Part() Then
                        ' Production Line does not have the entered Part Number setup.

                        MsgBox("ERROR: This Production Line does not have the entered Part Number setup.")
                        TextBoxInput4.Select()
                    End If
                End If
            Else
                ' Production Line is NOT valid.
                MsgBox("ERROR: Production Line is NOT Valid.")
                TextBoxInput4.Select()
            End If

        End If
    End Sub


    Private Sub TextBoxInput5_Leave(sender As Object, e As EventArgs) Handles TextBoxInput5.Leave
        Dim result As Integer

        Select Case currentActivity
            Case production ' Cavity Count Textbox OR Quantity Textbox   (depending on the value of My.Settings.EnterCuringProductionManually)
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


    Private Sub TextBoxInput7_Leave(sender As Object, e As EventArgs) Handles TextBoxInput7.Leave
        Dim result As Integer

        If Not String.IsNullOrEmpty(TextBoxInput7.Text) Then
            If Integer.TryParse(TextBoxInput7.Text, result) Then

                Select Case currentActivity
                    Case production ' Field: Shots
                        Call Update_Quantity_Fields()

                    Case downtime ' Field: Downtime Code
                        ' Call procedure to check if the entered downtime code is in the downtime validation file.
                        ' Call Validate_Scrap_Code(TextBoxInput7.Text)

                    Case scrap ' Field: Scrap Code
                        ' Call procedure to check if the entered scrap code is in the scrap validation file.
                        ' Call Validate_Scrap_Code(TextBoxInput7.Text)

                End Select

            Else
                ' Entry is not a valid integer.
                MsgBox("ERROR: Entry is not a valid format.")
                TextBoxInput7.Select()
            End If
        End If

    End Sub


    Private Sub TextBoxInput8_Leave(sender As Object, e As EventArgs) Handles TextBoxInput8.Leave
        Dim result As Integer

        Select Case currentActivity
            Case production ' Field: Free Shots
                If Not String.IsNullOrEmpty(TextBoxInput8.Text) Then
                    If Integer.TryParse(TextBoxInput8.Text, result) Then
                        Call Update_Quantity_Fields()
                    Else
                        MsgBox("ERROR: Entry is not a valid format.")
                        TextBoxInput8.Select()
                    End If
                End If

            Case downtime
                ' N/A - Downtime utilizes a combobox in the place of this.
            Case scrap
                ' N/A - Scrap utilizes a combobox in the place of this.
        End Select
    End Sub


    Private Sub TextBoxInput9_Leave(sender As Object, e As EventArgs) Handles TextBoxInput9.Leave
        Dim result As Integer

        Select Case currentActivity
            Case production
                ' N/A
            Case downtime ' Field: Downtime Quantity
                If Not String.IsNullOrEmpty(TextBoxInput9.Text) Then
                    If Not Integer.TryParse(TextBoxInput9.Text, result) Then
                        MsgBox("ERROR: Entry is not a valid format.")
                        TextBoxInput9.Select()
                    End If
                End If

            Case scrap ' Field: Scrap Quantity
                If Not String.IsNullOrEmpty(TextBoxInput9.Text) Then
                    If Not Integer.TryParse(TextBoxInput9.Text, result) Then
                        MsgBox("ERROR: Entry is not a valid format.")
                        TextBoxInput9.Select()
                    End If
                End If

        End Select
    End Sub


    Private Sub TextBoxInput10_Leave(sender As Object, e As EventArgs) Handles TextBoxInput10.Leave
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

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Select Case currentActivity
            Case downtime
                Call Populate_Downtime_Code()
            Case scrap
                'Call Populate_Scrap_Code()
        End Select
    End Sub


    Private Sub Update_Quantity_Fields()
        If My.Settings.EnterCuringProductionManually Then

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

        End If
    End Sub


    Private Sub Update_Form_Fields()
        If currentActivity = production Then
            TextBoxInput2.Text = prodData.Associate()
            TextBoxInput3.Text = prodData.PartNumber()
            TextBoxInput4.Text = prodData.ProdLine()

            If My.Settings.EnterCuringProductionManually Then
                TextBoxInput5.Text = prodData.CavityCount()
                TextBoxInput7.Text = prodData.Shots()
                TextBoxInput8.Text = prodData.FreeShots()
                TextBoxInput10.Text = prodData.ScrappedPieces()

                Call Update_Quantity_Fields()
            Else
                TextBoxInput5.Text = prodData.ProductionQty()
            End If


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
    Private Sub WriteToFile(ByVal radleyString As String)
        Dim oFile As FileStream = Nothing
        Dim oWrite As StreamWriter = Nothing
        Dim filePath As String = My.Settings.DropFolder
        Dim fileName As String = ""

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

        oFile = New FileStream(filePath, FileMode.Create, FileAccess.Write)
        oWrite = New StreamWriter(oFile)

        ' Write to the file.
        oWrite.WriteLine(radleyString)

        ' Close the filestream and streamwriter
        oWrite.Close()
        oFile.Close()

        MsgBox("SUCCESS: " + currentActivity + " was reported!")

    End Sub


    Private Function Validate_Part_Number(ByVal partNum As String)
        ' Create schedule task in IFS to export the data from the Inventory Part window to a txt file.  Then use that data to validate
        '    the user entered data.  ONLY export the part number column.
        Dim reader = New StreamReader(My.Settings.PathToDataValidationFiles + "InventoryParts" + My.Settings.Site + ".txt")

        While Not reader.EndOfStream
            If reader.ReadLine() = partNum Then
                ' Valid part number
                Return True
            End If
        End While
        Return False

    End Function


    Private Function Validate_Production_Line(ByVal prodLine As String)
        ' Create schedule task in IFS to export the data from the Inventory Part window to a txt file.  Then use that data to validate
        '    the user entered data.  ONLY export the part number column.
        Dim reader = New StreamReader(My.Settings.PathToDataValidationFiles + "ProductionLines" + My.Settings.Site + ".txt")

        While Not reader.EndOfStream
            If reader.ReadLine() = prodLine Then
                Return True
            End If
        End While
        Return False

    End Function


    Private Function Validate_Prod_Line_Part()
        ' Create schedule task in IFS to export the data from the Inventory Part window to a txt file.  Then use that data to validate
        '    the user entered data.  ONLY export the part number column.
        Dim reader = New StreamReader(My.Settings.PathToDataValidationFiles + "ProductionLineByPart" + My.Settings.Site + ".txt")
        Dim lineData As String
        Dim index As Integer
        Dim partNum As String
        Dim prodLine As String


        While Not reader.EndOfStream
            lineData = reader.ReadLine()
            index = lineData.IndexOf("|")
            prodLine = lineData.Substring(0, index)
            partNum = lineData.Substring(index + 1)

            If partNum = TextBoxInput3.Text And prodLine = TextBoxInput4.Text Then
                Return True
            End If
        End While
        Return False

    End Function


    Private Function Validate_Downtime_Reason()

        Dim reader = New StreamReader(My.Settings.PathToDataValidationFiles + "InterruptionCauses" + My.Settings.Site + ".txt")
        Dim lineData As String
        Dim index As Integer
        Dim downCode As String
        Dim downDesc As String

        While Not reader.EndOfStream
            lineData = reader.ReadLine()
            index = lineData.IndexOf("|")
            downCode = lineData.Substring(0, index)
            downDesc = lineData.Substring(index + 1)

            If reader.ReadLine() = downDesc Then

                Return True
            End If
        End While
        Return False

    End Function


End Class
