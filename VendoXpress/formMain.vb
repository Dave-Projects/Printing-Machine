Imports System.IO
Imports System.Drawing.Printing
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.PowerPoint
Imports asposeWords = Aspose.Words

Imports System.IO.Ports
Imports Viscomsoft.PDFViewer
Imports Aspose.Pdf.Facades
Imports System.Printing

Imports QRCoder

Imports System.Threading
Imports System.Net
Imports System.Net.Sockets
Imports WIA
Imports asposePdf = Aspose.Pdf

Public Class formMain

    'For Arduino-Coin_Acceptor
    Dim comPORT As String
    Dim coinCount As Integer

    'Shared Variables
    Dim insertflashdrive_loadfrom As String
    Dim printPropertiesForm_loadFrom As String
    Dim printPropertiesForm_selectedFile As String


    'Variables FOR LANDING FORM
    Dim landingfocusedControl As Control
    Dim landingCodeDirectory As String
    Dim landingCode As String

    'Variables FOR LOADEDFILES FORM
    Dim loadcountDown As Integer
    Private WithEvents loadpanelItem As Panel
    Private WithEvents loadlblFileName As Label
    Private WithEvents loadpanelItemElipse As Bunifu.Framework.UI.BunifuElipse
    Private WithEvents loadpicItem As Bunifu.Framework.UI.BunifuImageButton

    Dim loadfocusedControl As Control

    Dim loadprintOption As String
    Dim loadpptFile As String
    Dim loadpptApp As Application
    Dim loadpresentation As Microsoft.Office.Interop.PowerPoint.Presentation
    Dim loadtotalSlides As Integer

    'Variables FOR PRINT FORM

    Dim print_pdfviewerCalc As Viscomsoft.PDFViewer.PDFView
    Dim print_pdfviewer As Viscomsoft.PDFViewer.PDFView
    Dim printtotalPages As Integer

    Dim printdoc As Viscomsoft.PDFViewer.PDFDocument
    Dim printpageOrientation As String
    Dim printcurrentPage As Integer
    Dim printfocusedControl As Control
    Dim printstartPage As Integer
    Dim printendPage As Integer
    Dim printnumberCopies As Integer
    Dim printoutputColor As String
    Dim printtotalPrice As Decimal
    Dim printtimeDelay As Integer

    Dim printcoloredPagePrice As Decimal
    Dim printbwPagePrice As Decimal
    Dim printblankPagePrice As Decimal
    Dim printdarkPercentage As Double

    Dim printcountThanksDown As Integer


    'Variables FOR WIFI FORM
    Dim wificountDown As Integer

    Private WithEvents wifipanelItem As Panel
    Private WithEvents wifilblFileName As Label
    Private WithEvents wifipanelItemElipse As Bunifu.Framework.UI.BunifuElipse
    Private WithEvents wifipicItem As Bunifu.Framework.UI.BunifuImageButton

    Dim wififilesReceived As Integer
    Dim wifiloadFormat As String

    Dim wififocusedControl As Control

    Dim wifiprintOption As String
    Dim wifipptFile As String
    Dim wifipptApp As Application
    Dim wifipresentation As Microsoft.Office.Interop.PowerPoint.Presentation
    Dim wifitotalSlides As Integer


    'Variables for SCANNER FORM
    Dim scancountDown As Integer
    Dim scancountThanksDown As Integer
    Dim scanflashdriveDirectory As String
    Dim scantotalPages As Integer
    Dim scanscanPagePrice As Decimal
    Dim scanscannedPages As Integer

    'Variables for ADMIN FORM
    Dim adminfocusedControl As Control
    Dim admincoinCount As Integer
    Dim admincomPORT As String
    Dim adminoutputCoin As String
    Dim adminwithdraw As Boolean


    Dim copyscannedPages As Integer

    Dim dispensing As Boolean
    Dim remaining5 As Integer
    Dim remaining1 As Integer
    Dim hopperElapsed As Integer
    Dim changeAvailable As Boolean
    Dim RefundPrintStatus As Boolean
    '=========================================================================
    Private Sub formMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataFetcher.FetchData()
        landingFormLoad()
        setupPort()
        loadingForm.Hide()
        'Me.TopMost = True

        printerList()
        If printforLong.Items.Count > 0 Then
            printforLong.SelectedIndex = 2
        End If

        If printforA4.Items.Count > 0 Then
            printforA4.SelectedIndex = 0
        End If


    End Sub

    Private Sub dateTime_Tick(sender As Object, e As EventArgs) Handles dateTimeTimer.Tick
        Dim currentDate As Date = Date.Now

        mainFormTImeL.Text = currentDate.ToString("hh:mm tt")
        mainFormTImeL.Text = currentDate.ToString("hh:mm tt")
        mainFormDate.Text = currentDate.ToString("dddd, dd MMMM")
        mainFormDateL.Text = currentDate.ToString("dddd, dd MMMM")
        mainFormDateC.Text = currentDate.ToString("dddd, dd MMMM")
        mainFormDateS.Text = currentDate.ToString("dddd, dd MMMM")
        mainFormDateW.Text = currentDate.ToString("dddd, dd MMMM")
        mainFormDateLF.Text = currentDate.ToString("dddd, dd MMMM")
        mainFormDateINS.Text = currentDate.ToString("dddd, dd MMMM")
        mainFormDateO.Text = currentDate.ToString("dddd, dd MMMM")
        changeTime.Text = currentDate.ToString("hh:mm tt")
        changeDate.Text = currentDate.ToString("dddd, dd MMMM")
    End Sub

    'CODES FOR LANDING FORM
    '======================================================================================
    Private Sub landingbtnPrint_Click(sender As Object, e As EventArgs) Handles landingbtnPrint.Click
        TabControl1.Visible = False
        TabControl1.SelectedTab.SuspendLayout()
        TabControl1.SelectedIndex = 1
        TabControl1.SelectedTab.ResumeLayout(True)
        TabControl1.Visible = True
        insertflashdrive_loadfrom = "Print"

        If SerialPort1.IsOpen Then
            SerialPort1.WriteLine("reset")
            coinCount = 0
            printlblCoins.Text = Format(coinCount, "0.00")
            scanlblCoins.Text = Format(coinCount, "0.00")
            printPanel14.BackColor = Color.Red
            scanPanel14.BackColor = Color.Red
        End If
    End Sub

    Private Sub landingbtnCopy_Click(sender As Object, e As EventArgs) Handles landingbtnCopy.Click
        TabControl1.Visible = False
        TabControl1.SelectedTab.SuspendLayout()
        TabControl1.SelectedIndex = 5
        TabControl1.SelectedTab.ResumeLayout(True)
        copyPicScanned.Image = Nothing
        TabControl1.Visible = True
        loadDeleteFilesInFolder(DataFetcher.ScannedImages)

        If SerialPort1.IsOpen Then
            SerialPort1.WriteLine("reset")
            coinCount = 0
            printlblCoins.Text = Format(coinCount, "0.00")
            scanlblCoins.Text = Format(coinCount, "0.00")
            printPanel14.BackColor = Color.Red
            scanPanel14.BackColor = Color.Red
        End If
    End Sub

    Private Sub landingbtnScan_Click(sender As Object, e As EventArgs) Handles landingbtnScan.Click
        scanscanPagePrice = DataFetcher.ScanPagePrice
        TabControl1.Visible = False
        TabControl1.SelectedTab.SuspendLayout()
        TabControl1.SelectedIndex = 2
        TabControl1.SelectedTab.ResumeLayout(True)
        scanPicScanned.Image = Nothing
        TabControl1.Visible = True
        insertflashdrive_loadfrom = "Scanner"
        flashdriveFormLoad()

        If SerialPort1.IsOpen Then
            SerialPort1.WriteLine("reset")
            coinCount = 0
            printlblCoins.Text = Format(coinCount, "0.00")
            scanlblCoins.Text = Format(coinCount, "0.00")
            printPanel14.BackColor = Color.Red
            scanPanel14.BackColor = Color.Red
        End If
    End Sub

    Private Sub ladingbtnQuit_Click(sender As Object, e As EventArgs) Handles ladingbtnQuit.Click
        landingCodeDirectory = "Exit"
        landingpanelCode.Visible = True
        For Each ctrl As Control In landingTab.Controls
            If ctrl IsNot landingpanelCode Then
                ctrl.Enabled = False
            End If
        Next
    End Sub

    Private Sub landingbtnSettings_Click(sender As Object, e As EventArgs) Handles landingbtnSettings.Click
        landingCodeDirectory = "Settings"
        landingpanelCode.Visible = True
        landingpanelCode.BringToFront()
        For Each ctrl As Control In landingTab.Controls
            If ctrl IsNot landingpanelCode Then
                ctrl.Enabled = False
            End If
        Next
    End Sub

    Private Sub landingNumericTextBox_KeyPress(sender As Object, e As KeyPressEventArgs) Handles landingtxtSystemPin.KeyPress
        If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub landingNumberButton_Click(sender As Object, e As EventArgs) Handles landingbtnZero.Click, landingbtnOne.Click, landingbtnTwo.Click, landingbtnThree.Click, landingbtnFour.Click, landingbtnFive.Click, landingbtnSix.Click, landingbtnSeven.Click, landingbtnEight.Click, landingbtnNine.Click
        Dim digit As String = DirectCast(sender, Guna.UI.WinForms.GunaCircleButton).Text
        If TypeOf landingfocusedControl Is Bunifu.UI.WinForms.BunifuTextbox.BunifuTextBox Then
            Dim focusedTextBox As Bunifu.UI.WinForms.BunifuTextbox.BunifuTextBox = DirectCast(landingfocusedControl, Bunifu.UI.WinForms.BunifuTextbox.BunifuTextBox)
            If focusedTextBox.Text.Length < 8 Then
                focusedTextBox.Text &= digit
            End If
        End If
    End Sub

    Private Sub landingbtnErase_Click(sender As Object, e As EventArgs) Handles landingbtnErase.Click
        If TypeOf landingfocusedControl Is Bunifu.UI.WinForms.BunifuTextbox.BunifuTextBox Then
            Dim focusedTextBox As Bunifu.UI.WinForms.BunifuTextbox.BunifuTextBox = CType(landingfocusedControl, Bunifu.UI.WinForms.BunifuTextbox.BunifuTextBox)

            If focusedTextBox.Text.Length > 0 Then
                focusedTextBox.Text = focusedTextBox.Text.Substring(0, focusedTextBox.Text.Length - 1)
            End If
        End If
    End Sub

    Private Sub txtSystemPin_Enter(sender As Object, e As EventArgs) Handles landingtxtSystemPin.Enter
        landingfocusedControl = landingtxtSystemPin
    End Sub

    Private Sub landingbtnConfirmNo_Click(sender As Object, e As EventArgs) Handles landingbtnConfirmNo.Click

        landingpanelCode.Visible = False
        For Each ctrl As Control In landingTab.Controls
            ctrl.Enabled = True
        Next
        landinglblEnterpin.Text = "Enter Pin"
        landinglblEnterpin.ForeColor = Color.Black
        landingtxtSystemPin.ForeColor = Color.Black
        landingtxtSystemPin.PlaceholderForeColor = Color.Black
        landingtxtSystemPin.BorderColorIdle = Color.Silver
        landingtxtSystemPin.BorderColorActive = Color.Black
        landingtxtSystemPin.BorderColorHover = Color.Silver
        landingtxtSystemPin.Text = ""
    End Sub

    Private Sub checkCode()
        landingCode = DataFetcher.SystemPin
        If landingtxtSystemPin.Text = landingCode Then
            If landingCodeDirectory = "Exit" Then
                System.Windows.Forms.Application.Exit()
            ElseIf landingCodeDirectory = "Settings" Then
                TabControl1.Visible = False
                TabControl1.SelectedTab.SuspendLayout()
                TabControl1.SelectedIndex = 7
                TabControl1.SelectedTab.ResumeLayout(True)
                TabControl1.Visible = True
                adminSettingsForm_Load()
            End If
            For Each ctrl As Control In landingTab.Controls
                If ctrl IsNot landingpanelCode Then
                    ctrl.Enabled = True
                End If
            Next
            landingpanelCode.Visible = False
            For Each ctrl As Control In landingTab.Controls
                If ctrl IsNot landingpanelCode Then
                    ctrl.Enabled = True
                End If
            Next
            landinglblEnterpin.Text = "Enter Pin"
            landinglblEnterpin.ForeColor = Color.Black
            landingtxtSystemPin.ForeColor = Color.Black
            landingtxtSystemPin.PlaceholderForeColor = Color.Black
            landingtxtSystemPin.BorderColorIdle = Color.Silver
            landingtxtSystemPin.BorderColorActive = Color.Black
            landingtxtSystemPin.BorderColorHover = Color.Silver
            landingtxtSystemPin.Text = ""
        Else
            landinglblEnterpin.Text = "Wrong Pin"
            landinglblEnterpin.ForeColor = Color.Red
            landingtxtSystemPin.ForeColor = Color.Red
            landingtxtSystemPin.PlaceholderForeColor = Color.Tomato
            landingtxtSystemPin.BorderColorIdle = Color.Tomato
            landingtxtSystemPin.BorderColorActive = Color.Red
            landingtxtSystemPin.BorderColorHover = Color.Tomato
            landingtxtSystemPin.Text = ""
        End If

    End Sub
    Private Sub landingbtnConfirmYes_Click(sender As Object, e As EventArgs) Handles landingbtnConfirmYes.Click, landingbtnOk.Click
        checkCode()
    End Sub

    Private Sub landingFormLoad()
        landingtxtSystemPin.MaxLength = 8
        landingCode = DataFetcher.SystemPin
    End Sub


    'CODES FOR PRINTING OPTION
    '======================================================================================

    Private Sub btnFlashdrive_Click(sender As Object, e As EventArgs) Handles optnbtnFlashdrive.Click
        insertflashdrive_loadfrom = "Print"
        TabControl1.Visible = False
        TabControl1.SelectedTab.SuspendLayout()
        TabControl1.SelectedIndex = 2
        TabControl1.SelectedTab.ResumeLayout(True)
        TabControl1.Visible = True
        flashdriveFormLoad()
        loadPanel1.BringToFront()
        loadbtnAllFiles.BringToFront()
        loadbringNotifConvertPaneltoFront()
    End Sub

    Private Sub btnBluetooth_Click(sender As Object, e As EventArgs) Handles optnbtnWifi.Click
        TabControl1.Visible = False
        TabControl1.SelectedTab.SuspendLayout()
        TabControl1.SelectedIndex = 3
        TabControl1.SelectedTab.ResumeLayout(True)
        TabControl1.Visible = True
        wifiReceiveForm_Load()
        wifiPanel1.BringToFront()
        wifibtnAllFiles.BringToFront()
        bringNotifConvertPaneltoFront()
    End Sub

    Private Sub btnBack_Click_1(sender As Object, e As EventArgs) Handles optnbtnBack.Click
        TabControl1.Visible = False
        TabControl1.SelectedTab.SuspendLayout()
        TabControl1.SelectedIndex = 0
        TabControl1.SelectedTab.ResumeLayout(True)
        TabControl1.Visible = True
        GC.Collect()
        EnableDisableButton()
    End Sub

    'INSERT FLAHDRIVE 
    '======================================================================================
    Private Sub flashdriveFormLoad()
        flashTimer1.Start()
    End Sub
    Private Sub flashCheckDriveAvailability()
        Dim dDriveInfo As DriveInfo = New DriveInfo(DataFetcher.FlashDrivePath)
        flashlblLoading.Visible = True
        If dDriveInfo.IsReady Then
            flashpanelFlahdrivegif.Visible = False
            flashlblLoading.Text = "Loading..."
            flashpreLoader.Visible = True
            flashlblflashNotif.Visible = True

            flashTimer2.Start()

        Else
            flashpanelFlahdrivegif.Visible = True
            flashlblLoading.Text = "Insert flash drive."
            flashpreLoader.Visible = False
            flashlblflashNotif.Visible = False
        End If
    End Sub

    Private Sub flashTimer1_Tick(sender As Object, e As EventArgs) Handles flashTimer1.Tick
        flashCheckDriveAvailability()
    End Sub

    Private Sub flashTimer2_Tick(sender As Object, e As EventArgs) Handles flashTimer2.Tick
        If insertflashdrive_loadfrom = "Scanner" Then
            TabControl1.Visible = False
            TabControl1.SelectedTab.SuspendLayout()
            TabControl1.SelectedIndex = 4
            TabControl1.SelectedTab.ResumeLayout(True)
            TabControl1.Visible = True
            flashTimer2.Stop()
            flashTimer1.Stop()
            ScannerForm_Load()

        ElseIf insertflashdrive_loadfrom = "Print" Then
            TabControl1.Visible = False
            TabControl1.SelectedTab.SuspendLayout()
            TabControl1.SelectedIndex = 8
            TabControl1.SelectedTab.ResumeLayout(True)
            TabControl1.Visible = True
            flashTimer2.Stop()
            flashTimer1.Stop()
            loadedfilesForm_Load()

        End If
    End Sub

    Private Sub flashbtnBack_Click(sender As Object, e As EventArgs) Handles flashbtnBack.Click
        If insertflashdrive_loadfrom = "Scanner" Then
            flashTimer1.Stop()
            flashTimer2.Stop()
            TabControl1.Visible = False
            TabControl1.SelectedTab.SuspendLayout()
            TabControl1.SelectedIndex = 0
            TabControl1.SelectedTab.ResumeLayout(True)
            TabControl1.Visible = True
            EnableDisableButton()
        ElseIf insertflashdrive_loadfrom = "Print" Then
            flashTimer1.Stop()
            flashTimer2.Stop()
            TabControl1.Visible = False
            TabControl1.SelectedTab.SuspendLayout()
            TabControl1.SelectedIndex = 1
            TabControl1.SelectedTab.ResumeLayout(True)
            TabControl1.Visible = True
        End If
    End Sub

    'INSERT LOADEDFILES FORM
    '======================================================================================
    Public Sub loadLoadFilesByType(files As FileInfo(), imageResource As Image)
        Array.Sort(files, Function(file1, file2) String.Compare(file1.Name, file2.Name))

        For Each file As FileInfo In files
            loadpanelItem = New Panel With {
                .BackColor = Color.White,
                .Size = New Size(185, 229),
                .Location = New Drawing.Point(4, 5),
                .Tag = file.FullName
            }

            loadpicItem = New Bunifu.Framework.UI.BunifuImageButton With {
                .Image = imageResource,
                .Location = New Drawing.Point(18, 5),
                .Name = "BunifuImageButton1",
                .Size = New Size(148, 171),
                .SizeMode = PictureBoxSizeMode.StretchImage,
                .TabStop = False,
                .Tag = file.FullName
            }

            loadlblFileName = New Label With {
                .BackColor = Color.Transparent,
                .Font = New Drawing.Font("Segoe UI Semibold", 12.0!),
                .ForeColor = Color.DimGray,
                .Location = New Drawing.Point(2, 185),
                .Margin = New Padding(2, 0, 2, 0),
                .Name = "lblFileName",
                .Size = New Size(181, 46),
                .Text = file.Name,
                .TextAlign = ContentAlignment.TopCenter,
                .Visible = True,
                .Tag = file.FullName
            }

            loadpanelItemElipse = New Bunifu.Framework.UI.BunifuElipse With {
               .ElipseRadius = 20,
               .TargetControl = Me.loadpanelItem
            }

            loadpanelItem.Controls.Add(loadpicItem)
            loadpanelItem.Controls.Add(loadlblFileName)
            loadFlowLayoutPanel1.Controls.Add(loadpanelItem)

            AddHandler loadpanelItem.Click, AddressOf loadFile_Click
            AddHandler loadpicItem.Click, AddressOf loadFile_Click
            AddHandler loadlblFileName.Click, AddressOf loadFile_Click
        Next
    End Sub

    Public Function loadGetFilesByExtension(directoryInfo As DirectoryInfo, extensions As String()) As FileInfo()
        Dim files As New List(Of FileInfo)()
        Dim dDriveInfo As DriveInfo = New DriveInfo(DataFetcher.FlashDrivePath)

        Try
            If dDriveInfo.IsReady Then
                For Each file As FileInfo In directoryInfo.GetFiles("*.*")
                    Dim extension As String = file.Extension.ToLower()
                    Dim fileName As String = file.Name

                    If Not fileName.StartsWith("~") AndAlso extensions.Contains(extension) Then
                        files.Add(file)
                    End If
                Next

                For Each subdirectory As DirectoryInfo In directoryInfo.GetDirectories()
                    If subdirectory.Name <> "System Volume Information" Then
                        files.AddRange(loadGetFilesByExtension(subdirectory, extensions))
                    End If
                Next
            End If
        Catch ex As UnauthorizedAccessException

        End Try

        Return files.ToArray()
    End Function

    Public Sub loadLoadAllFiles()
        loadFlowLayoutPanel1.Controls.Clear()
        Dim directoryInfo As New DirectoryInfo(DataFetcher.FlashDrivePath)

        Dim allFiles As FileInfo() = loadGetFilesByExtension(directoryInfo, {".pdf", ".doc", ".docx", ".ppt", ".pptx"})

        Dim pdfFiles = allFiles.Where(Function(f) f.Extension.ToLower() = ".pdf").ToArray()
        Dim wordFiles = allFiles.Where(Function(f) f.Extension.ToLower() = ".doc" Or f.Extension.ToLower() = ".docx").ToArray()
        Dim pptFiles = allFiles.Where(Function(f) f.Extension.ToLower() = ".ppt" Or f.Extension.ToLower() = ".pptx").ToArray()

        loadLoadFilesByType(pdfFiles, My.Resources.Resources.PDF)
        loadLoadFilesByType(wordFiles, My.Resources.Resources.DOC)
        loadLoadFilesByType(pptFiles, My.Resources.Resources.PPT)
    End Sub

    Public Async Sub loadFile_Click(sender As Object, e As EventArgs)
        Dim filePath As String = sender.tag.ToString
        Dim fileExtension As String = Path.GetExtension(filePath)
        If fileExtension = ".pdf" Then

            Dim newfile As String = Path.Combine(DataFetcher.CachePath, Path.GetFileName(filePath))

            Try
                loadPanel1.Visible = False
                loadPanel3.Visible = False
                loadbtnBack.Visible = False
                loadbtnAllFiles.Visible = False
                loadbtnPPT.Visible = False
                loadbtnPDF.Visible = False
                loadbtnDOC.Visible = False
                loadpanelLoading.Visible = True
                loadtimerLoading.Start()

                File.Copy(filePath, newfile)
                Await Task.Delay(500)

                loadWaitForFileCreation(newfile)
                printPropertiesForm_loadFrom = "loadFromFlashDrive"
                printPropertiesForm_selectedFile = newfile
                TabControl1.Visible = False
                TabControl1.SelectedTab.SuspendLayout()
                TabControl1.SelectedIndex = 6
                TabControl1.SelectedTab.ResumeLayout(True)
                TabControl1.Visible = True
                printPropertiesForm_Load()
            Catch ex As Exception

            End Try
        Else
            Dim pdfPath As String = Path.ChangeExtension(filePath, ".pdf")
            loadConvertToPDF(filePath, fileExtension)

        End If


    End Sub
    Private loadinputFilePathWord2Pdf As String
    Private loaddestinationDirectoryWord2Pdf As String
    Public Sub loadConvertToPDF(inputFilePath As String, fileExtension As String)
        If fileExtension = ".docx" OrElse fileExtension = ".doc" Then
            loadDisableComponents()
            loadpicConvert.Image = My.Resources.Doc_to_Pdf
            loadHideComponents()
            loadinputFilePathWord2Pdf = inputFilePath
            loaddestinationDirectoryWord2Pdf = DataFetcher.CachePath
            loadConvertWordToPDF.Start()
        ElseIf fileExtension = ".pptx" OrElse fileExtension = ".ppt" Then
            loadpptFile = inputFilePath
            'loadpanelPrintLayout.Visible = True
            Guna2Transition1.ShowSync(loadpanelPrintLayout)
            loadpanelPrintLayout.BringToFront()
            loadpptApp = New Microsoft.Office.Interop.PowerPoint.Application()
            loadpresentation = loadpptApp.Presentations.Open(inputFilePath, WithWindow:=MsoTriState.msoFalse)
            loadtotalSlides = loadpresentation.Slides.Count
            loadtxtTo.Text = loadtotalSlides.ToString
        Else
            MessageBox.Show("Unsupported file format.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If
    End Sub

    Private Sub loadConvertWordToPDF_Tick(sender As Object, e As EventArgs) Handles loadConvertWordToPDF.Tick
        Try
            Dim pdfFilePath As String = Path.Combine(loaddestinationDirectoryWord2Pdf, Path.GetFileNameWithoutExtension(loadinputFilePathWord2Pdf) & ".pdf")
            Dim doc As New asposeWords.Document(loadinputFilePathWord2Pdf)

            doc.Save(pdfFilePath, asposeWords.SaveFormat.Pdf)
            loadWaitForFileCreation(pdfFilePath)
            printPropertiesForm_selectedFile = pdfFilePath
            printPropertiesForm_loadFrom = "loadFromFlashDrive"

            TabControl1.Visible = False
            TabControl1.SelectedTab.SuspendLayout()
            TabControl1.SelectedIndex = 6
            TabControl1.SelectedTab.ResumeLayout(True)
            TabControl1.Visible = True
            printPropertiesForm_Load()
            loadConvertWordToPDF.Stop()
            loadShowinitialComponents()
        Catch ex As Exception
            loadShowComponents()
            loadConvertWordToPDF.Stop()
        End Try
    End Sub
    Public Sub loadConvertPPTToPDF(pptFilePath As String, destinationDirectory As String, FromPage As Integer, ToPage As Integer)
        loadDisableComponents()
        loadpicConvert.Image = My.Resources.Ppt_to_Pdf
        loadHideComponents()
        Dim OutputTypePage As PpPrintOutputType
        Try
            Select Case loadprintOption
                Case "Full Page Slides"
                    OutputTypePage = PpPrintOutputType.ppPrintOutputSlides
                Case "Notes Page"
                    OutputTypePage = PpPrintOutputType.ppPrintOutputNotesPages
                Case "Outline"
                    OutputTypePage = PpPrintOutputType.ppPrintOutputOutline
                Case "1 Slide Per Page"
                    OutputTypePage = PpPrintOutputType.ppPrintOutputOneSlideHandouts
                Case "2 Slides Per Page"
                    OutputTypePage = PpPrintOutputType.ppPrintOutputTwoSlideHandouts
                Case "3 Slides Per Page"
                    OutputTypePage = PpPrintOutputType.ppPrintOutputThreeSlideHandouts
                Case "4 Slides Per Page"
                    OutputTypePage = PpPrintOutputType.ppPrintOutputFourSlideHandouts
                Case "6 Slides Per Page"
                    OutputTypePage = PpPrintOutputType.ppPrintOutputSixSlideHandouts
                Case "9 Slides Per Page"
                    OutputTypePage = PpPrintOutputType.ppPrintOutputNineSlideHandouts
            End Select
            Dim pdfoutputPath As String = destinationDirectory & "\" & Path.GetFileNameWithoutExtension(pptFilePath) & ".pdf"
            Dim oRange As Microsoft.Office.Interop.PowerPoint.PrintRange = loadpresentation.PrintOptions.Ranges.Add(FromPage, ToPage)
            loadpresentation.ExportAsFixedFormat(Path:=pdfoutputPath, FixedFormatType:=PpFixedFormatType.ppFixedFormatTypePDF, FrameSlides:=MsoTriState.msoFalse,
                                 OutputType:=OutputTypePage, PrintHiddenSlides:=MsoTriState.msoFalse, PrintRange:=oRange,
                                 RangeType:=PpPrintRangeType.ppPrintSlideRange, HandoutOrder:=PpPrintHandoutOrder.ppPrintHandoutHorizontalFirst)

            loadWaitForFileCreation(pdfoutputPath)
            printPropertiesForm_selectedFile = pdfoutputPath
            printPropertiesForm_loadFrom = "loadFromFlashDrive"
            TabControl1.Visible = False
            TabControl1.SelectedTab.SuspendLayout()
            TabControl1.SelectedIndex = 6
            TabControl1.SelectedTab.ResumeLayout(True)
            TabControl1.Visible = True
            printPropertiesForm_Load()
        Catch ex As Exception
            'MessageBox.Show(ex.Message)
            loadShowComponents()
        Finally
            If loadpresentation IsNot Nothing Then
                loadpresentation.Close()
                Marshal.ReleaseComObject(loadpresentation)
                loadpresentation = Nothing
            End If
            loadpptApp.Quit()
            Marshal.ReleaseComObject(loadpptApp)

            loadShowinitialComponents()
        End Try
    End Sub
    Private Sub loadWaitForFileCreation(filePath As String)
        Do While Not (File.Exists(filePath) AndAlso New FileInfo(filePath).Length > 0)
            System.Threading.Thread.Sleep(500)
        Loop
    End Sub
    Public Sub loadDeleteFilesInFolder(folderPath As String)
        Try
            Dim files As String() = Directory.GetFiles(folderPath)

            For Each filePath As String In files
                File.Delete(filePath)
            Next

        Catch ex As Exception

        End Try
    End Sub

    Private Sub loadbtnAllFiles_Click(sender As Object, e As EventArgs) Handles loadbtnAllFiles.Click
        loadFlowLayoutPanel1.Controls.Clear()
        loadLoadAllFiles()
        loadPanel1.BringToFront()
        loadbtnAllFiles.BringToFront()
        loadbringNotifConvertPaneltoFront()

    End Sub

    Private Sub loadbtnPDF_Click(sender As Object, e As EventArgs) Handles loadbtnPDF.Click
        loadFlowLayoutPanel1.Controls.Clear()
        loadLoadFilesByType(loadGetFilesByExtension(New DirectoryInfo(DataFetcher.FlashDrivePath), {".pdf"}), My.Resources.Resources.PDF)
        loadPanel1.BringToFront()
        loadbtnPDF.BringToFront()
        loadbringNotifConvertPaneltoFront()
    End Sub

    Private Sub loadbtnDOC_Click(sender As Object, e As EventArgs) Handles loadbtnDOC.Click
        loadFlowLayoutPanel1.Controls.Clear()
        loadLoadFilesByType(loadGetFilesByExtension(New DirectoryInfo(DataFetcher.FlashDrivePath), {".doc", ".docx"}), My.Resources.Resources.DOC)
        loadPanel1.BringToFront()
        loadbtnDOC.BringToFront()
        loadbringNotifConvertPaneltoFront()
    End Sub

    Private Sub loadbtnPPT_Click(sender As Object, e As EventArgs) Handles loadbtnPPT.Click
        loadFlowLayoutPanel1.Controls.Clear()
        loadLoadFilesByType(loadGetFilesByExtension(New DirectoryInfo(DataFetcher.FlashDrivePath), {".ppt", ".pptx"}), My.Resources.Resources.PPT)
        loadPanel1.BringToFront()
        loadbtnPPT.BringToFront()
        loadbringNotifConvertPaneltoFront()
    End Sub

    Private Sub loadedfilesForm_Load()
        loadpresentation = Nothing
        loadTimer1.Start()
        Dim dDriveInfo As DriveInfo = New DriveInfo(DataFetcher.FlashDrivePath)

        If dDriveInfo.IsReady Then
            loadtimerLoadAll.Start()
        End If
        loadShowinitialComponents()
    End Sub

    Sub loadbringNotifConvertPaneltoFront()
        loadpanelConvert.BringToFront()
        loadpanelNotif.BringToFront()
    End Sub

    Private Sub loadTimer2_Tick(sender As Object, e As EventArgs) Handles loadTimer2.Tick
        loadcountDown = loadcountDown - 1
        If loadcountDown = 3 Then
            loadlbldown.Text = "Flash drive disconnected. Closing in... 3s."
        ElseIf loadcountDown = 2 Then
            loadlbldown.Text = "Flash drive disconnected. Closing in... 2s."
        ElseIf loadcountDown = 1 Then
            loadlbldown.Text = "Flash drive disconnected. Closing in... 1s."
        ElseIf loadcountDown = 0 Then
            loadTimer1.Stop()
            loadTimer2.Stop()

            loadDeleteFilesInFolder(DataFetcher.CachePath)
            TabControl1.Visible = False
            TabControl1.SelectedTab.SuspendLayout()
            TabControl1.SelectedIndex = 1
            TabControl1.SelectedTab.ResumeLayout(True)
            TabControl1.Visible = True
        End If
    End Sub
    Private Sub loadCheckDriveAvailability()
        Dim dDriveInfo As DriveInfo = New DriveInfo(DataFetcher.FlashDrivePath)

        If dDriveInfo.IsReady Then
            loadpanelNotif.Visible = False
            loadTimer2.Enabled = False
            loadFlowLayoutPanel1.Enabled = True
        Else
            loadpanelNotif.Visible = True
            loadTimer2.Enabled = True
            loadFlowLayoutPanel1.Enabled = False
        End If
    End Sub

    Private Sub loadTimer1_Tick(sender As Object, e As EventArgs) Handles loadTimer1.Tick
        loadCheckDriveAvailability()
    End Sub

    Private Sub loadpanelNotif1_VisibleChanged(sender As Object, e As EventArgs) Handles loadpanelNotif1.VisibleChanged
        loadlbldown.Text = "Flash drive disconnected. Closing in... 3s."
        loadcountDown = 3
    End Sub

    Private Sub loadtimerLoadAll_Tick(sender As Object, e As EventArgs) Handles loadtimerLoadAll.Tick
        loadLoadAllFiles()
        loadtimerLoadAll.Stop()
    End Sub

    Private Sub loadbtnBack_Click(sender As Object, e As EventArgs) Handles loadbtnBack.Click
        loadTimer1.Stop()
        loadTimer2.Stop()

        loadDeleteFilesInFolder(DataFetcher.CachePath)
        TabControl1.Visible = False
        TabControl1.SelectedTab.SuspendLayout()
        TabControl1.SelectedIndex = 1
        TabControl1.SelectedTab.ResumeLayout(True)
        TabControl1.Visible = True
    End Sub

    Private Sub loadtimerLoading_Tick(sender As Object, e As EventArgs) Handles loadtimerLoading.Tick
        loadpanelLoading.Visible = False
        loadtimerLoading.Stop()
        loadShowinitialComponents()
    End Sub
    Public Sub loadHideComponents()
        loadpanelConvert.Visible = True
        loadPanel1.Visible = False
        loadPanel3.Visible = False
        loadbtnBack.Visible = False
        loadbtnAllFiles.Visible = False
        loadbtnPPT.Visible = False
        loadbtnPDF.Visible = False
        loadbtnDOC.Visible = False
        loadpanelPrintLayout.Visible = False
    End Sub
    Public Sub loadShowinitialComponents()
        loadPanel1.Visible = True
        loadPanel3.Visible = True
        loadbtnBack.Visible = True
        loadbtnAllFiles.Visible = True
        loadbtnPPT.Visible = True
        loadbtnPDF.Visible = True
        loadbtnDOC.Visible = True

        loadpanelSlidestoPrint.Visible = False
        loadpanelPrintLayout.Visible = False
        loadpanelNotif.Visible = False
        loadpanelConvert.Visible = False

        For Each ctrl As Control In LoadedTab.Controls
            ctrl.Enabled = True
        Next
    End Sub
    Public Sub loadShowComponents()
        loadpanelConvert.Visible = False
        loadPanel1.Visible = True
        loadPanel3.Visible = True
        loadbtnBack.Visible = True
        loadbtnAllFiles.Visible = True
        loadbtnPPT.Visible = True
        loadbtnPDF.Visible = True
        loadbtnDOC.Visible = True
        loadpanelPrintLayout.Visible = True
    End Sub
    Public Sub loadEnableComponents()
        loadPanel1.Enabled = True
        loadPanel3.Enabled = True
        loadbtnBack.Enabled = True
        loadbtnAllFiles.Enabled = True
        loadbtnPPT.Enabled = True
        loadbtnPDF.Enabled = True
        loadbtnDOC.Enabled = True
        loadpanelPrintLayout.Enabled = True
    End Sub
    Public Sub loadDisableComponents()
        loadPanel1.Enabled = False
        loadPanel3.Enabled = False
        loadbtnBack.Enabled = False
        loadbtnAllFiles.Enabled = False
        loadbtnPPT.Enabled = False
        loadbtnPDF.Enabled = False
        loadbtnDOC.Enabled = False
        loadpanelPrintLayout.Enabled = False
    End Sub

    Private Sub loadbtnCancel_Click(sender As Object, e As EventArgs) Handles loadbtnCancel.Click
        loadEnableComponents()
        'loadpanelPrintLayout.Visible = False
        Guna2Transition1.HideSync(loadpanelPrintLayout)
        If loadpresentation IsNot Nothing Then
            loadpresentation.Close()
            Marshal.ReleaseComObject(loadpresentation)
            loadpresentation = Nothing
        End If
        loadpptApp.Quit()
        Marshal.ReleaseComObject(loadpptApp)
    End Sub
    Private Sub loadbtnFullPage_Click(sender As Object, e As EventArgs) Handles loadbtnFullPage.Click
        loadprintOption = "Full Page Slides"
        loadpanelSlidestoPrint.Visible = True
        loadpanelSlidestoPrint.BringToFront()
        loadpanelPrintLayout.Enabled = False
    End Sub

    Private Sub loadbtnNotesPage_Click(sender As Object, e As EventArgs) Handles loadbtnNotesPage.Click
        loadprintOption = "Notes Page"
        loadpanelSlidestoPrint.Visible = True
        loadpanelSlidestoPrint.BringToFront()
        loadpanelPrintLayout.Enabled = False
    End Sub

    Private Sub loadbtnOutline_Click(sender As Object, e As EventArgs) Handles loadbtnOutline.Click
        loadprintOption = "Outline"
        loadpanelSlidestoPrint.Visible = True
        loadpanelSlidestoPrint.BringToFront()
        loadpanelPrintLayout.Enabled = False
    End Sub

    Private Sub loadbtn1Slide_Click(sender As Object, e As EventArgs) Handles loadbtn1Slide.Click
        loadprintOption = "1 Slide Per Page"
        loadpanelSlidestoPrint.Visible = True
        loadpanelSlidestoPrint.BringToFront()
        loadpanelPrintLayout.Enabled = False
    End Sub

    Private Sub loadbtn2Slide_Click(sender As Object, e As EventArgs) Handles loadbtn2Slide.Click
        loadprintOption = "2 Slides Per Page"
        loadpanelSlidestoPrint.Visible = True
        loadpanelSlidestoPrint.BringToFront()
        loadpanelPrintLayout.Enabled = False
    End Sub

    Private Sub loadbtn3Slide_Click(sender As Object, e As EventArgs) Handles loadbtn3Slide.Click
        loadprintOption = "3 Slides Per Page"
        loadpanelSlidestoPrint.Visible = True
        loadpanelSlidestoPrint.BringToFront()
        loadpanelPrintLayout.Enabled = False
    End Sub

    Private Sub loadbtn4Slide_Click(sender As Object, e As EventArgs) Handles loadbtn4Slide.Click
        loadprintOption = "4 Slides Per Page"
        loadpanelSlidestoPrint.Visible = True
        loadpanelSlidestoPrint.BringToFront()
        loadpanelPrintLayout.Enabled = False
    End Sub

    Private Sub loadbtn6Slide_Click(sender As Object, e As EventArgs) Handles loadbtn6Slide.Click
        loadprintOption = "6 Slides Per Page"
        loadpanelSlidestoPrint.Visible = True
        loadpanelSlidestoPrint.BringToFront()
        loadpanelPrintLayout.Enabled = False
    End Sub

    Private Sub loadbtn9Slide_Click(sender As Object, e As EventArgs) Handles loadbtn9Slide.Click
        loadprintOption = "9 Slides Per Page"
        loadpanelSlidestoPrint.Visible = True
        loadpanelSlidestoPrint.BringToFront()
        loadpanelPrintLayout.Enabled = False
    End Sub

    Private Sub loadNumericTextBox_KeyPress(sender As Object, e As KeyPressEventArgs) Handles loadtxtFrom.KeyPress, loadtxtTo.KeyPress
        If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub loadNumberButton_Click(sender As Object, e As EventArgs) Handles loadbtnZero.Click, loadbtnOne.Click, loadbtnTwo.Click, loadbtnThree.Click, loadbtnFour.Click, loadbtnFive.Click, loadbtnSix.Click, loadbtnSeven.Click, loadbtnEight.Click, loadbtnNine.Click
        Dim digit As String = DirectCast(sender, Guna.UI.WinForms.GunaCircleButton).Text
        If TypeOf loadfocusedControl Is Bunifu.UI.WinForms.BunifuTextbox.BunifuTextBox Then
            Dim focusedTextBox As Bunifu.UI.WinForms.BunifuTextbox.BunifuTextBox = DirectCast(loadfocusedControl, Bunifu.UI.WinForms.BunifuTextbox.BunifuTextBox)
            If focusedTextBox.Text.Length < 3 Then
                focusedTextBox.Text &= digit
            End If
        End If
    End Sub

    Private Sub loadbtnErase_Click(sender As Object, e As EventArgs) Handles loadbtnErase.Click
        If TypeOf loadfocusedControl Is Bunifu.UI.WinForms.BunifuTextbox.BunifuTextBox Then
            Dim focusedTextBox As Bunifu.UI.WinForms.BunifuTextbox.BunifuTextBox = CType(loadfocusedControl, Bunifu.UI.WinForms.BunifuTextbox.BunifuTextBox)

            If focusedTextBox.Text.Length > 0 Then
                focusedTextBox.Text = focusedTextBox.Text.Substring(0, focusedTextBox.Text.Length - 1)
            End If
        End If
    End Sub

    Private Sub loadrdAll_Click(sender As Object, e As EventArgs) Handles loadrdAll.Click
        loadpanelKeypad.Enabled = False
        loadpanelSpecific.Enabled = False
        loadtxtFrom.Text = "1"
        loadtxtTo.Text = loadtotalSlides.ToString
    End Sub

    Private Sub loadrdSpecific_Click(sender As Object, e As EventArgs) Handles loadrdSpecific.Click
        loadpanelKeypad.Enabled = True
        loadpanelSpecific.Enabled = True
    End Sub

    Private Sub loadlblAllpages_Click(sender As Object, e As EventArgs) Handles loadlblAllpages.Click
        loadrdAll.Checked = True
        loadrdSpecific.Checked = False
        loadpanelKeypad.Enabled = False
        loadpanelSpecific.Enabled = False
        loadtxtFrom.Text = "1"
        loadtxtTo.Text = loadtotalSlides.ToString
    End Sub

    Private Sub loadlblSpecific_Click(sender As Object, e As EventArgs) Handles loadlblSpecific.Click
        loadrdSpecific.Checked = True
        loadrdAll.Checked = False
        loadpanelKeypad.Enabled = True
        loadpanelSpecific.Enabled = True
    End Sub

    Private Sub loadFromToChecker()
        If (loadtxtFrom.Text = "" Or Val(loadtxtFrom.Text) = 0) And (loadtxtTo.Text = "" Or Val(loadtxtTo.Text) = 0) Then
            loadlblError.Visible = True
            loadlblError.Text = "'From' and 'To' value is required."
            loadbtnOk.Enabled = False
            loadbtnOkay.Enabled = False
        ElseIf loadtxtFrom.Text = "" Or Val(loadtxtFrom.Text) = 0 Then
            loadlblError.Visible = True
            loadlblError.Text = "'From' value is required."
            loadbtnOk.Enabled = False
            loadbtnOkay.Enabled = False
        ElseIf loadtxtTo.Text = "" Or Val(loadtxtTo.Text) = 0 Then
            loadlblError.Visible = True
            loadlblError.Text = "'To' value is required."
            loadbtnOk.Enabled = False
            loadbtnOkay.Enabled = False
        Else
            If Val(loadtxtFrom.Text) > Val(loadtxtTo.Text) Then
                loadlblError.Visible = True
                loadlblError.Text = "'From' value must be greater than the 'To' value."
                loadbtnOk.Enabled = False
                loadbtnOkay.Enabled = False
            ElseIf Val(loadtxtFrom.Text) > loadtotalSlides Then
                loadlblError.Visible = True
                loadlblError.Text = "'From' value has exceeded the total number of pages"
                loadbtnOk.Enabled = False
                loadbtnOkay.Enabled = False
            ElseIf Val(loadtxtTo.Text) > loadtotalSlides Then
                loadlblError.Visible = True
                loadlblError.Text = "'To' value has exceeded the total number of pages"
                loadbtnOk.Enabled = False
                loadbtnOkay.Enabled = False
            Else
                loadlblError.Visible = False
                loadbtnOk.Enabled = True
                loadbtnOkay.Enabled = True
            End If
        End If

    End Sub

    Private Sub loadbtnCancelSpecific_Click(sender As Object, e As EventArgs) Handles loadbtnCancelSpecific.Click
        loadpanelSlidestoPrint.Visible = False
        loadpanelPrintLayout.Enabled = True
        loadpanelKeypad.Enabled = False
        loadpanelSpecific.Enabled = False
        loadrdAll.Checked = True
        loadrdSpecific.Checked = False
        loadtxtFrom.Text = "1"
        loadtxtTo.Text = loadtotalSlides.ToString
    End Sub

    Private Sub loadtxtFrom_Enter(sender As Object, e As EventArgs) Handles loadtxtFrom.Enter
        loadpanelKeypad.Visible = True
        loadfocusedControl = loadtxtFrom
        loadFromToChecker()
    End Sub

    Private Sub loadtxtTo_Enter(sender As Object, e As EventArgs) Handles loadtxtTo.Enter
        loadpanelKeypad.Visible = True
        loadfocusedControl = loadtxtTo
        loadFromToChecker()
    End Sub

    Private Sub loadtxtFrom_TextChanged(sender As Object, e As EventArgs) Handles loadtxtFrom.TextChanged
        loadFromToChecker()
    End Sub

    Private Sub loadtxtTo_TextChanged(sender As Object, e As EventArgs) Handles loadtxtTo.TextChanged
        loadFromToChecker()
    End Sub

    Private Sub loadbtnOkay_Click(sender As Object, e As EventArgs) Handles loadbtnOkay.Click, loadbtnOk.Click
        loadpanelSlidestoPrint.Visible = False
        loadConvertPPTToPDF(loadpptFile, DataFetcher.CachePath, Val(loadtxtFrom.Text), Val(loadtxtTo.Text))
    End Sub

    Private Sub loadbtnFromInc_Click(sender As Object, e As EventArgs) Handles loadbtnFromInc.Click
        If loadtxtFrom.Text.Length > 0 Then
            Dim currentValue As Integer = Integer.Parse(loadtxtFrom.Text)
            Dim newValue As Integer = (currentValue Mod loadtotalSlides) + 1

            ' Ensure newValue is less than or equal to txtTo value
            If newValue <= Integer.Parse(loadtxtTo.Text) Then
                loadtxtFrom.Text = newValue.ToString()
            End If
        Else
            loadtxtFrom.Text = "1"
        End If
    End Sub

    Private Sub loadbtnFromDec_Click(sender As Object, e As EventArgs) Handles loadbtnFromDec.Click
        If loadtxtFrom.Text.Length > 0 Then
            Dim currentValue As Integer = Integer.Parse(loadtxtFrom.Text)
            Dim newValue As Integer = If(currentValue > 1, currentValue - 1, loadtotalSlides)

            ' Ensure newValue is less than or equal to txtTo value
            If newValue <= Integer.Parse(loadtxtTo.Text) Then
                loadtxtFrom.Text = newValue.ToString()
            End If
        Else
            loadtxtFrom.Text = loadtotalSlides
        End If
    End Sub

    Private Sub loadbtnToInc_Click(sender As Object, e As EventArgs) Handles loadbtnToInc.Click
        If loadtxtTo.Text.Length > 0 Then
            Dim currentValue As Integer = Integer.Parse(loadtxtTo.Text)
            Dim newValue As Integer = (currentValue Mod loadtotalSlides) + 1

            ' Ensure newValue is greater than or equal to txtFrom value
            If newValue >= Integer.Parse(loadtxtFrom.Text) Then
                loadtxtTo.Text = newValue.ToString()
            End If
        Else
            loadtxtTo.Text = "1"
        End If
    End Sub

    Private Sub loadbtnToDec_Click(sender As Object, e As EventArgs) Handles loadbtnToDec.Click
        If loadtxtTo.Text.Length > 0 Then
            Dim currentValue As Integer = Integer.Parse(loadtxtTo.Text)
            Dim newValue As Integer = If(currentValue > 1, currentValue - 1, loadtotalSlides)

            ' Ensure newValue is greater than or equal to txtFrom value
            If newValue >= Integer.Parse(loadtxtFrom.Text) Then
                loadtxtTo.Text = newValue.ToString()
            End If
        Else
            loadtxtTo.Text = loadtotalSlides
        End If
    End Sub
    'INSERT PrintProperties FORM
    '======================================================================================

    Public Sub printDeleteFilesInFolder(folderPath As String)
        Try
            Dim files As String() = Directory.GetFiles(folderPath)

            For Each filePath As String In files
                File.Delete(filePath)
            Next

        Catch ex As Exception

        End Try
    End Sub

    Public Sub printloadFile()
        print_pdfviewer = New Viscomsoft.PDFViewer.PDFView
        printdoc = New Viscomsoft.PDFViewer.PDFDocument

        If printdoc.open(printPropertiesForm_selectedFile) Then
            print_pdfviewer.Canvas.Parent = Me.printpanelViewer
            print_pdfviewer.Canvas.BackColor = Color.WhiteSmoke

            ' Load the page and check orientation
            Dim page As New PDFPage(printdoc)
            page.load(1)
            Dim isLandscape As Boolean = page.Width > page.Height
            If isLandscape Then
                Landscape()
            Else
                Portrait()
            End If
            page.unload(printdoc, True)

            print_pdfviewer.Canvas.Size = New Size(Me.printpanelViewer.ClientSize.Width, Me.printpanelViewer.ClientSize.Height)
            print_pdfviewer.Document = printdoc
            print_pdfviewer.Zoom = Zoom.FitPage

            printtotalPages = printdoc.PageCount
            print_pdfviewer.gotoPage(1)
            printcurrentPage = 1

            ' Handle trackbars for multiple pages
            If printtotalPages > 1 Then
                printportraitTrackBar.Maximum = printtotalPages
                printlandscapeTrackBar.Maximum = printtotalPages
                printlandscapeTrackBar.Minimum = 1
                printportraitTrackBar.Minimum = 1
            End If
            printportraitTrackBar.Value = printcurrentPage
            printlandscapeTrackBar.Value = printcurrentPage

            ' Show file name and page details
            Dim fileName As String = printPropertiesForm_selectedFile.Substring(printPropertiesForm_selectedFile.LastIndexOf("\") + 1)
            printlblFileName.Text = fileName
            printlblPages.Text = printtotalPages
            printtxtFrom.Text = "1"
            printtxtTo.Text = printtotalPages

            ' Extract paper size using Aspose.Pdf
            Try
                Dim pdfDocument As New asposePdf.Document(printPropertiesForm_selectedFile)
                Dim firstPage As asposePdf.Page = pdfDocument.Pages(1)
                Dim pageWidth As Integer = firstPage.MediaBox.Width
                Dim pageHeight As Integer = firstPage.MediaBox.Height
                Dim pageWidthInch As Double = firstPage.MediaBox.Width / 72
                Dim pageHeightInch As Double = firstPage.MediaBox.Height / 72
                ' PARA READABLE 
                Dim paperSize As String
                If pageWidth = 612 And pageHeight = 1008 Then
                    paperSize = "Legal"
                ElseIf pageWidth = 612 And pageHeight = 792 Then
                    paperSize = "Letter"
                ElseIf pageWidth = 595 And pageHeight = 842 Then
                    paperSize = "A4"
                Else
                    paperSize = "Custom Size: " & pageWidth & "x" & pageHeight & "pt"
                End If

                ' Display the paper size in label98
                Label98.Text = paperSize
            Catch ex As Exception
                ' Handle any errors in reading the PDF or displaying the paper size
                Label98.Text = "Unable to determine paper size"
            End Try

            If Label98.Text = "Letter" Then
                MessageBox.Show("This Printer only Allowed Long & A4 Size", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)

            End If
        Else
            printdoc.close()
        End If

    End Sub
    Private Sub Portrait()
        printpageOrientation = "Portrait"
        printpanelHolder.Size = New Size(851, 1117)
        printpanelHolder.Location = New System.Drawing.Point(930, 75)
        printbtnNext.Location = New System.Drawing.Point(1800, 686)
        printbtnPrev.Location = New System.Drawing.Point(830, 686)
        printportraitTrackBar.Visible = True
        printlandscapeTrackBar.Visible = False
        printpanelPrice.Location = New System.Drawing.Point(60, 200)
        printpanelProps.Location = New System.Drawing.Point(60, 200)
        printpanelKeypad.Location = New System.Drawing.Point(195, 680)
        mainFormDate.Visible = False
        mainFormTImeL.Visible = False


        printtimeDelay = 750
    End Sub
    Private Sub Landscape()
        printpageOrientation = "Landscape"
        printpanelHolder.Size = New Size(1117, 851)
        printpanelHolder.Location = New System.Drawing.Point(800, 200)
        printbtnNext.Location = New System.Drawing.Point(1790, 1100)
        printbtnPrev.Location = New System.Drawing.Point(840, 1100)
        printlandscapeTrackBar.Visible = True
        printportraitTrackBar.Visible = False

        printpanelPrice.Location = New System.Drawing.Point(40, 200)
        printpanelProps.Location = New System.Drawing.Point(40, 200)
        printpanelKeypad.Location = New System.Drawing.Point(185, 680)

        printtimeDelay = 750

        mainFormDateLF.Visible = False
        mainFormTImeL.Visible = False
    End Sub
    Private Sub printPropertiesForm_Load()

        setupPort()
        printloadFile()
        printtxtCopies.Text = 1
        printcurrentPage = 1

        printcoloredPagePrice = DataFetcher.ColoredPagePrice
        printbwPagePrice = DataFetcher.BWPagePrice
        printblankPagePrice = DataFetcher.BlankPagePrice
        coinCount = 0

    End Sub
    Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles printbtnNext.Click
        printcurrentPage += 1
        If printcurrentPage > printtotalPages Then
            printcurrentPage = 1
        End If

        printportraitTrackBar.Value = printcurrentPage
        printlandscapeTrackBar.Value = printcurrentPage
        print_pdfviewer.gotoPage(printcurrentPage)
        printlblpage.Text = "Page: " + printcurrentPage.ToString
    End Sub

    Private Sub btnPrev_Click(sender As Object, e As EventArgs) Handles printbtnPrev.Click
        printcurrentPage -= 1
        If printcurrentPage < 1 Then
            printcurrentPage = printtotalPages
        End If

        printportraitTrackBar.Value = printcurrentPage
        printlandscapeTrackBar.Value = printcurrentPage
        print_pdfviewer.gotoPage(printcurrentPage)
        printlblpage.Text = "Page: " + printcurrentPage.ToString
    End Sub

    Private Sub portraitTrackBar_Scroll_1(sender As Object, e As ScrollEventArgs) Handles printportraitTrackBar.Scroll
        Dim newPage As Integer = printportraitTrackBar.Value

        If newPage <> printcurrentPage Then
            printcurrentPage = newPage
            printlblpage.Text = "Page: " + printcurrentPage.ToString
            print_pdfviewer.gotoPage(printcurrentPage)
        End If
    End Sub

    Private Sub landscapeTrackBar_Scroll_1(sender As Object, e As ScrollEventArgs) Handles printlandscapeTrackBar.Scroll
        Dim newPage As Integer = printlandscapeTrackBar.Value

        If newPage <> printcurrentPage Then
            printcurrentPage = newPage
            printlblpage.Text = "Page: " + printcurrentPage.ToString
            print_pdfviewer.gotoPage(printcurrentPage)
        End If
    End Sub

    Private Sub printlblAllpages_Click(sender As Object, e As EventArgs) Handles printlblAllpages.Click
        printrdAll.Checked = True
        printrdSpecific.Checked = False
        printpanelFromTo.Enabled = False
    End Sub

    Private Sub printlblSpecific_Click(sender As Object, e As EventArgs) Handles printlblSpecific.Click
        printrdSpecific.Checked = True
        printrdAll.Checked = False
        printpanelFromTo.Enabled = True
    End Sub

    Private Sub printbtnFromInc_Click(sender As Object, e As EventArgs) Handles printbtnFromInc.Click
        If printtxtFrom.Text.Length > 0 Then
            Dim currentValue As Integer = Integer.Parse(printtxtFrom.Text)
            Dim newValue As Integer = (currentValue Mod printtotalPages) + 1

            ' Ensure newValue is less than or equal to txtTo value
            If newValue <= Integer.Parse(printtxtTo.Text) Then
                printtxtFrom.Text = newValue.ToString()
            End If
        Else
            printtxtFrom.Text = "1"
        End If
    End Sub

    Private Sub printbtnFromDec_Click(sender As Object, e As EventArgs) Handles printbtnFromDec.Click
        If printtxtFrom.Text.Length > 0 Then
            Dim currentValue As Integer = Integer.Parse(printtxtFrom.Text)
            Dim newValue As Integer = If(currentValue > 1, currentValue - 1, printtotalPages)

            ' Ensure newValue is less than or equal to txtTo value
            If newValue <= Integer.Parse(printtxtTo.Text) Then
                printtxtFrom.Text = newValue.ToString()
            End If
        Else
            printtxtFrom.Text = printtotalPages
        End If

    End Sub

    Private Sub printbtnToInc_Click(sender As Object, e As EventArgs) Handles printbtnToInc.Click
        If printtxtTo.Text.Length > 0 Then
            Dim currentValue As Integer = Integer.Parse(printtxtTo.Text)
            Dim newValue As Integer = (currentValue Mod printtotalPages) + 1

            ' Ensure newValue is greater than or equal to txtFrom value
            If newValue >= Integer.Parse(printtxtFrom.Text) Then
                printtxtTo.Text = newValue.ToString()
            End If
        Else
            printtxtTo.Text = "1"
        End If

    End Sub

    Private Sub printbtnToDec_Click(sender As Object, e As EventArgs) Handles printbtnToDec.Click
        If printtxtTo.Text.Length > 0 Then
            Dim currentValue As Integer = Integer.Parse(printtxtTo.Text)
            Dim newValue As Integer = If(currentValue > 1, currentValue - 1, printtotalPages)

            ' Ensure newValue is greater than or equal to txtFrom value
            If newValue >= Integer.Parse(printtxtFrom.Text) Then
                printtxtTo.Text = newValue.ToString()
            End If
        Else
            printtxtTo.Text = printtotalPages
        End If
    End Sub

    Private Sub printrdAll_Click(sender As Object, e As EventArgs) Handles printrdAll.Click
        printpanelFromTo.Enabled = False
    End Sub

    Private Sub printrdSpecific_Click(sender As Object, e As EventArgs) Handles printrdSpecific.Click
        printpanelFromTo.Enabled = True
    End Sub

    Private Sub btnCopiesInc_Click(sender As Object, e As EventArgs) Handles printbtnCopiesInc.Click
        If printtxtCopies.Text.Length > 0 Then
            Dim currentValue As Integer = Integer.Parse(printtxtCopies.Text)
            Dim newValue As Integer = (currentValue Mod 99) + 1
            printtxtCopies.Text = newValue.ToString()
        Else
            printtxtCopies.Text = "1"
        End If
    End Sub

    Private Sub btnCopiesDec_Click(sender As Object, e As EventArgs) Handles printbtnCopiesDec.Click
        If printtxtCopies.Text.Length > 0 Then
            Dim currentValue As Integer = Integer.Parse(printtxtCopies.Text)
            Dim newValue As Integer = If(currentValue > 1, currentValue - 1, 1)
            printtxtCopies.Text = newValue.ToString()
        Else
            printtxtCopies.Text = "1"
        End If
    End Sub



    Private Sub printNumericTextBox_KeyPress(sender As Object, e As KeyPressEventArgs) Handles printtxtFrom.KeyPress, printtxtTo.KeyPress, printtxtCopies.KeyPress
        If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub printNumberButton_Click(sender As Object, e As EventArgs) Handles printbtnZero.Click, printbtnOne.Click, printbtnTwo.Click, printbtnThree.Click, printbtnFour.Click, printbtnFive.Click, printbtnSix.Click, printbtnSeven.Click, printbtnEight.Click, printbtnNine.Click
        Dim digit As String = DirectCast(sender, Guna.UI.WinForms.GunaCircleButton).Text
        If TypeOf printfocusedControl Is Bunifu.UI.WinForms.BunifuTextbox.BunifuTextBox Then
            Dim focusedTextBox As Bunifu.UI.WinForms.BunifuTextbox.BunifuTextBox = DirectCast(printfocusedControl, Bunifu.UI.WinForms.BunifuTextbox.BunifuTextBox)
            If focusedTextBox.Text.Length < 3 Then
                focusedTextBox.Text &= digit
            End If
        End If
    End Sub

    Private Sub printbtnErase_Click(sender As Object, e As EventArgs) Handles printbtnErase.Click
        If TypeOf printfocusedControl Is Bunifu.UI.WinForms.BunifuTextbox.BunifuTextBox Then
            Dim focusedTextBox As Bunifu.UI.WinForms.BunifuTextbox.BunifuTextBox = CType(printfocusedControl, Bunifu.UI.WinForms.BunifuTextbox.BunifuTextBox)

            If focusedTextBox.Text.Length > 0 Then
                focusedTextBox.Text = focusedTextBox.Text.Substring(0, focusedTextBox.Text.Length - 1)
            End If
        End If
    End Sub

    Private Sub btnOk_Click(sender As Object, e As EventArgs) Handles printbtnOk.Click
        printpanelKeypad.Visible = False
    End Sub

    Private Sub printFromToChecker()
        If (printtxtFrom.Text = "" Or Val(printtxtFrom.Text) = 0) And (printtxtTo.Text = "" Or Val(printtxtTo.Text) = 0) Then
            printlblError.Visible = True
            printlblError.Text = "'From' and 'To' value is required."
            disableButton()
        ElseIf printtxtFrom.Text = "" Or Val(printtxtFrom.Text) = 0 Then
            printlblError.Visible = True
            printlblError.Text = "'From' value is required."
            disableButton()
        ElseIf printtxtTo.Text = "" Or Val(printtxtTo.Text) = 0 Then
            printlblError.Visible = True
            printlblError.Text = "'To' value is required."
            disableButton()
        Else
            If Val(printtxtFrom.Text) > Val(printtxtTo.Text) Then
                printlblError.Visible = True
                printlblError.Text = "'From' value must be greater than the 'To' value."
                disableButton()
            ElseIf Val(printtxtFrom.Text) > printtotalPages Then
                printlblError.Visible = True
                printlblError.Text = "'From' value has exceeded the total number of pages"
                disableButton()
            ElseIf Val(printtxtTo.Text) > printtotalPages Then
                printlblError.Visible = True
                printlblError.Text = "'To' value has exceeded the total number of pages"
                disableButton()
            Else
                printlblError.Visible = False
                printbtnGreyScale.Enabled = True
                printbtnSmartPrice.Enabled = True
                printbtnColored.Enabled = True
            End If
        End If

    End Sub
    Private Sub disableButton()
        printbtnGreyScale.Enabled = False
        printbtnColored.Enabled = False
        printbtnSmartPrice.Enabled = False
    End Sub

    Private Sub printtxtFrom_TextChanged(sender As Object, e As EventArgs) Handles printtxtFrom.TextChanged
        printFromToChecker()
    End Sub

    Private Sub printtxtTo_TextChanged(sender As Object, e As EventArgs) Handles printtxtTo.TextChanged
        printFromToChecker()
    End Sub

    Private Sub txtCopies_TextChanged(sender As Object, e As EventArgs) Handles printtxtCopies.TextChanged
        If printtxtCopies.Text = "" Or Val(printtxtCopies.Text) = 0 Then
            printlblCopiesError.Visible = True
            printlblCopiesError.Text = "'Copies' value is required."
            disableButton()
        Else
            printlblCopiesError.Visible = False
            printbtnGreyScale.Enabled = True
            printbtnSmartPrice.Enabled = True
            printbtnColored.Enabled = True
        End If
    End Sub

    Private Sub printtxtFrom_Enter(sender As Object, e As EventArgs) Handles printtxtFrom.Enter
        printpanelKeypad.Visible = True
        printfocusedControl = printtxtFrom
        printFromToChecker()
    End Sub

    Private Sub printtxtTo_Enter(sender As Object, e As EventArgs) Handles printtxtTo.Enter
        printpanelKeypad.Visible = True
        printfocusedControl = printtxtTo
        printFromToChecker()
    End Sub

    Private Sub txtCopies_Enter(sender As Object, e As EventArgs) Handles printtxtCopies.Enter
        printpanelKeypad.Visible = True
        printfocusedControl = printtxtCopies
        If printtxtCopies.Text = "" Or Val(printtxtCopies.Text) = 0 Then
            printlblCopiesError.Visible = True
            printlblCopiesError.Text = "'Copies' value is required."
            disableButton()
        Else
            printlblCopiesError.Visible = False
            printbtnGreyScale.Enabled = True
            printbtnSmartPrice.Enabled = True
        End If
    End Sub

    Private Async Function ProcessPagesAsync() As Task
        print_pdfviewerCalc.gotoPage(printstartPage)
        Dim bwPages As New List(Of Integer)
        Dim coloredPages As New List(Of Integer)
        printtotalPrice = 0
        Dim coloredPercentage As Double
        Dim message As String = "Pricing summary"
        printListBox1.Items.Clear()
        printListBox1.Items.Add(message)
        Await Task.Delay(2500)

        picSmartPrice.Image = Nothing
        Dim messages As New List(Of String)
        For i As Integer = printstartPage - 1 To printendPage - 1
            Await Task.Delay(printtimeDelay)
            Dim bmp As New Bitmap(printpanelViewCalc.Width, printpanelViewCalc.Height)

            ' Draw the panel onto the bitmap
            printpanelViewCalc.DrawToBitmap(bmp, New Rectangle(0, 0, printpanelViewCalc.Width, printpanelViewCalc.Height))
            picSmartPrice.Image = bmp
            While picSmartPrice.Image Is Nothing
                Await Task.Delay(100)
            End While

            coloredPercentage = CalculateColoredPercentage(bmp)
            Dim isCompletelyWhite As Boolean = CheckIfPageIsCompletelyWhite(bmp)
            Dim pagePrice As Decimal = If(isCompletelyWhite, printblankPagePrice, CalculatePriceForColoredPage(printbwPagePrice, printcoloredPagePrice, coloredPercentage, isCompletelyWhite))
            message = "Page " & i + 1 & " price: ₱" & pagePrice.ToString("0") & vbCrLf
            messages.Add(message)
            printtotalPrice += Math.Round(pagePrice)
            print_pdfviewerCalc.nextPage()
        Next
        Await Task.Delay(printtimeDelay)
        printListBox1.Items.AddRange(messages.ToArray())
        print_pdfviewerCalc.Document.close()
        ShowDetails()
        picSmartPrice.Image = Nothing
    End Function
    ' Function to calculate the percentage of colored pixels in the image
    Private Function CalculateColoredPercentage(image As Bitmap) As Double
        Dim totalPixels As Integer = image.Width * image.Height
        Dim coloredPixels As Integer = 0
        Dim BlackValue As Integer = 0
        Dim Diff As Integer = 1
        For x As Integer = 0 To image.Width - 1
            For y As Integer = 0 To image.Height - 1
                Dim pixel As Color = image.GetPixel(x, y)

                ' Calculate the absolute differences between R, G, and B values
                Dim RVal As Integer = CInt(pixel.R)
                Dim GVal As Integer = CInt(pixel.G)
                Dim BVal As Integer = CInt(pixel.B)
                Dim rDiff As Integer = RVal - GVal
                Dim gDiff As Integer = GVal - BVal
                Dim bDiff As Integer = BVal - RVal

                If rDiff > Diff OrElse gDiff > Diff OrElse bDiff > Diff Then
                    coloredPixels += 1
                ElseIf pixel.R = 0 OrElse pixel.G = 0 OrElse pixel.B = 0 Then
                    BlackValue += 1
                End If
            Next
        Next

        printdarkPercentage = (BlackValue / totalPixels) * 100
        Dim coloredPercentage As Double = (coloredPixels / totalPixels) * 100
        'MessageBox.Show(coloredPercentage)
        Return coloredPercentage
    End Function

    Private Function CalculatePriceForColoredPage(bwPagePrice As Decimal, coloredPagePrice As Decimal, coloredPercentage As Double, isCompletelyWhite As Boolean) As Decimal
        If isCompletelyWhite Then
            Return 1
        ElseIf coloredPercentage < 0.5 Then
            Return bwPagePrice
        ElseIf coloredPercentage >= 0.5 AndAlso coloredPercentage < 25 Then
            Dim priceDifference As Double = coloredPagePrice - bwPagePrice
            Dim priceIncrease As Decimal = priceDifference / 14
            Dim calculatedPrice As Decimal = bwPagePrice + (priceIncrease * coloredPercentage)
            Return Math.Min(calculatedPrice, coloredPagePrice)
        Else
            Return coloredPagePrice
        End If
    End Function

    Private Function CheckIfPageIsCompletelyWhite(image As Bitmap) As Boolean
        For x As Integer = 0 To image.Width - 1
            For y As Integer = 0 To image.Height - 1
                Dim pixel As Color = image.GetPixel(x, y)
                If pixel.R <> 255 OrElse pixel.G <> 255 OrElse pixel.B <> 255 Then
                    Return False ' If any non-white pixel is found, the page is not completely white
                End If
            Next
        Next
        Return True ' If no non-white pixel is found, the page is completely white
    End Function
    Private Sub ShowDetails()

        Dim fileName As String = printPropertiesForm_selectedFile.Substring(printPropertiesForm_selectedFile.LastIndexOf("\") + 1)
        If fileName.Length > 15 Then
            fileName = fileName.Substring(0, 13) & "..."
        End If
        printlblSFileName.Text = fileName
        printlblPagestoPrint.Text = printstartPage.ToString + "-" + printendPage.ToString
        printlblCopies.Text = printnumberCopies.ToString
        printlblOutput.Text = printoutputColor

        If printoutputColor = "Smart Pricing" Then
            printlblTotalPrice.Text = Format(printtotalPrice * printnumberCopies, "0.00")
            printbtnShowPrices.Visible = True
        ElseIf printoutputColor = "Colored" Then
            Dim price As Integer = (((printendPage - printstartPage) + 1) * printcoloredPagePrice)
            printlblTotalPrice.Text = Format(price * printnumberCopies, "0.00")
            printbtnShowPrices.Visible = False
        ElseIf printoutputColor = "Greyscale" Then
            Dim price As Integer = (((printendPage - printstartPage) + 1) * printbwPagePrice)
            printlblTotalPrice.Text = Format(price * printnumberCopies, "0.00")
            printbtnShowPrices.Visible = False
        End If

        printpanelPrice.Visible = True
        printpanelHolder.Visible = True
        printbtnBack.Visible = True
        printbtnPrev.Visible = True
        printbtnNext.Visible = True
        printpanelCalculating.Visible = False
        printpanelViewCalc.Visible = False
    End Sub

    Private Sub btnSmartPrice_Click(sender As Object, e As EventArgs) Handles printbtnSmartPrice.Click
        Panel14.Visible = True
        greyscalePanelConfirm.Visible = False
        coloredPanelConfirm.Visible = False

        printpanelConfirmation.Visible = True
        printpanelConfirmation.BringToFront()
        For Each ctrl As Control In printTab.Controls
            If ctrl IsNot printpanelConfirmation Then
                ctrl.Enabled = False
            End If
        Next
        printpanelConfirmation.Enabled = True
    End Sub


    Private Sub btnGreyScale_Click(sender As Object, e As EventArgs) Handles printbtnGreyScale.Click

        Panel14.Visible = False
        greyscalePanelConfirm.Visible = True
        coloredPanelConfirm.Visible = False

        printpanelConfirmation.Visible = True
        printpanelConfirmation.BringToFront()
        For Each ctrl As Control In printTab.Controls
            If ctrl IsNot printpanelConfirmation Then
                ctrl.Enabled = False
            End If
        Next
        printpanelConfirmation.Enabled = True
    End Sub

    Private Sub btnColored_Click(sender As Object, e As EventArgs) Handles printbtnColored.Click

        Panel14.Visible = False
        greyscalePanelConfirm.Visible = False
        coloredPanelConfirm.Visible = True

        printpanelConfirmation.Visible = True
        printpanelConfirmation.BringToFront()
        For Each ctrl As Control In printTab.Controls
            If ctrl IsNot printpanelConfirmation Then
                ctrl.Enabled = False
            End If
        Next
        printpanelConfirmation.Enabled = True
    End Sub

    Private Sub timerCoinReceiver_Tick(sender As Object, e As EventArgs) Handles printtimerCoinReceiver.Tick
        Dim receivedData As String = ReceiveSerialData()
        If receivedData.StartsWith("Coin: ") Then
            Dim coinCountText As String = receivedData.Substring(6).Trim()
            If Integer.TryParse(coinCountText, coinCount) Then
                printlblCoins.Text = Format(coinCount, "0.00")
                scanlblCoins.Text = Format(coinCount, "0.00")
            End If
        End If
        'If receivedData.StartsWith("dispensing") Then
        '    dispensing = True
        'Else

        If receivedData.StartsWith("idle") Then
            dispensing = False
        End If
        If receivedData.StartsWith("D5: ") Then
            dispensing = True
            Dim remaining5pesos As String = receivedData.Substring(4).Trim()
            Dim remainingInt As Integer
            If Integer.TryParse(remaining5pesos, remainingInt) Then
                remaining5 = remainingInt * 5
            End If
        End If
        If receivedData.StartsWith("D1: ") Then
            dispensing = True
            Dim remaining1pesos As String = receivedData.Substring(4).Trim()
            Dim remainingInt As Integer
            If Integer.TryParse(remaining1pesos, remainingInt) Then
                remaining1 = remainingInt
            End If
        End If

    End Sub
    Function ReceiveSerialData() As String
        Try
            Dim Incoming As String = SerialPort1.ReadExisting()
            Return Incoming
        Catch ex As TimeoutException
            Return "Error: Serial Port read timed out."
        End Try
    End Function

    Private Sub setupPort()
        Try
            SerialPort1.Close()
            comPORT = DataFetcher.CoinSlotPort
            SerialPort1.PortName = comPORT
            SerialPort1.BaudRate = 9600
            SerialPort1.DataBits = 8
            SerialPort1.Parity = Parity.None
            SerialPort1.StopBits = StopBits.One
            SerialPort1.Handshake = Handshake.None
            SerialPort1.Encoding = System.Text.Encoding.Default
            SerialPort1.ReadTimeout = 10000
            If Not SerialPort1.IsOpen Then
                SerialPort1.Open()
                printtimerCoinReceiver.Start()
            End If
        Catch ex As Exception
            MsgBox("Coin Slot is not Connected, Set up the port first!", MsgBoxStyle.Exclamation, "Port Error")
        End Try

    End Sub

    Private Sub lblCoins_TextChanged(sender As Object, e As EventArgs) Handles printlblCoins.TextChanged
        If Val(printlblCoins.Text) >= Val(printlblTotalPrice.Text) Then
            printbtnPrint.Visible = True
        Else
            printbtnPrint.Visible = False
        End If
    End Sub

    Private Sub printbtnBack_Click(sender As Object, e As EventArgs) Handles printbtnBack.Click
        mainFormDateLF.Visible = True
        mainFormTImeL.Visible = True
        printdoc.close()
        printpanelViewer.Controls.Clear()
        printpanelPrice.Visible = False
        printlblpage.Text = "Page: 0"
        Dim coins As Decimal = Decimal.Parse(printlblCoins.Text)
        Dim withdrawCoins As Integer = CInt(coins)
        changefunction(withdrawCoins)

        printDeleteFilesInFolder(DataFetcher.CachePath) 'Clear Folder
        If printPropertiesForm_loadFrom = "loadFromFlashDrive" Then
            TabControl1.Visible = False
            TabControl1.SelectedTab.SuspendLayout()
            TabControl1.SelectedIndex = 8
            TabControl1.SelectedTab.ResumeLayout(True)
            TabControl1.Visible = True
        ElseIf printPropertiesForm_loadFrom = "loadFromWifi" Then
            TabControl1.Visible = False
            TabControl1.SelectedTab.SuspendLayout()
            TabControl1.SelectedIndex = 3
            TabControl1.SelectedTab.ResumeLayout(True)
            TabControl1.Visible = True
        ElseIf printPropertiesForm_loadFrom = "loadFromCopy" Then
            TabControl1.Visible = False
            TabControl1.SelectedTab.SuspendLayout()
            TabControl1.SelectedIndex = 0
            TabControl1.SelectedTab.ResumeLayout(True)
            TabControl1.Visible = True
            EnableDisableButton()
        End If
        printpanelKeypad.Visible = False
        printrdAll.Checked = True
        printrdSpecific.Checked = False
        printpanelFromTo.Enabled = False
    End Sub

    Private Sub printbtnConfirmYes_Click(sender As Object, e As EventArgs) Handles printbtnConfirmYes.Click
        Try
            DataFetcher.FetchData()
            printpanelPrintStatus.Visible = True
            printpanelPrintStatus.BringToFront()
            printlblPrintStatus.Text = "Loading..."
            printlblPrintStatus.ForeColor = Color.Black
            printpreloaderStatus.Visible = True
            printlblContactAdmin.Visible = False
            Dim change As Integer
            printlblPrintStatus.ForeColor = Color.Green

            Dim coins As Decimal = Decimal.Parse(printlblCoins.Text)
            Dim totalPrice As Decimal = Decimal.Parse(printlblTotalPrice.Text)
            change = CInt(coins) - CInt(totalPrice)
            If change > 0 Then
                changefunction(change)
            End If
            TimeDispense.Start()

        Catch ex As Exception
            'MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub changefunction(change As Integer)
        DataFetcher.FetchData()
        Dim change5pesos As Integer = 0
        Dim change1pesos As Integer = 0
        ' Calculate change using a greedy algorithm
        If change > 0 Then
            While change > 0
                If change >= 5 Then
                    change -= 5
                    change5pesos += 1
                ElseIf change >= 1 Then
                    change -= 1
                    change1pesos += 1
                Else
                    ' If there are not enough coins of any denomination to make change, break the loop
                    Exit While
                End If
            End While
            Try
                If change5pesos > 0 Then
                    SerialPort1.WriteLine("Change5Peso: " & change5pesos.ToString())
                End If
                If change1pesos > 0 Then
                    SerialPort1.WriteLine("Change1Peso: " & change1pesos.ToString())
                End If
                If SerialPort1.IsOpen Then
                    SerialPort1.WriteLine("reset")
                    coinCount = 0
                    printlblCoins.Text = Format(coinCount, "0.00")
                    scanlblCoins.Text = Format(coinCount, "0.00")
                    printPanel14.BackColor = Color.Red
                    scanPanel14.BackColor = Color.Red
                End If
            Catch ex As Exception
                'MessageBox.Show(ex.Message)
            End Try
        End If
    End Sub
    Private Async Function PrintDocs() As Task
        Try
            Await Task.Delay(500)
            Dim fromPage As Integer = Integer.Parse(printtxtFrom.Text)
            Dim toPage As Integer = Integer.Parse(printtxtTo.Text)
            Dim copies As Integer = Integer.Parse(printtxtCopies.Text)

            Dim viewer As New PdfViewer()
            viewer.BindPdf(printPropertiesForm_selectedFile)

            Dim printerSettings As New PrinterSettings()
            printerSettings.Copies = copies


            '2nd CHANGES FOR PRINTER'


            ' Check if the printer combo box has a valid selected item
            If BunifuRadioButton1.Checked Then
                If printforLong.SelectedItem IsNot Nothing Then
                    printerSettings.PrinterName = printforLong.SelectedItem.ToString()
                Else
                    MessageBox.Show("Please select a printer for Long paper before proceeding. Contact Administrator first", "Printer Selection Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit Try
                End If
            ElseIf BunifuRadioButton2.Checked Then
                If printforA4.SelectedItem IsNot Nothing Then
                    printerSettings.PrinterName = printforA4.SelectedItem.ToString()
                Else
                    MessageBox.Show("Please select a printer for A4 paper before proceeding. Contact Administrator first", "Printer Selection Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit Try
                End If
            End If

            printerSettings.FromPage = fromPage
            printerSettings.ToPage = toPage
            printerSettings.PrintRange = System.Drawing.Printing.PrintRange.SomePages

            Dim PageSettings As New PageSettings()
            PageSettings.Margins = New Margins(0, 0, 0, 0)
            PageSettings.Color = printoutputColor <> "Greyscale"
            PageSettings.Landscape = printpageOrientation = "Landscape"

            viewer.AutoResize = True
            viewer.VerticalAlignment = Aspose.Pdf.VerticalAlignment.Center
            viewer.PrintDocumentWithSettings(PageSettings, printerSettings)
            viewer.Close()
        Catch ex As Exception
            ' Handle the exception as necessary
            MessageBox.Show($"An error occurred while printing: {ex.Message}", "Print Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function
    Private Sub printTimer3_Tick(sender As Object, e As EventArgs) Handles printTimer3.Tick
        printcountThanksDown = printcountThanksDown - 1
        If printcountThanksDown = 3 Then
            printlblThanksCount.ForeColor = Color.Red
            printlblThanksCount.Text = "Closing in... 3s."
        ElseIf printcountThanksDown = 2 Then
            printlblThanksCount.Text = "Closing in... 2s."
        ElseIf printcountThanksDown = 1 Then
            printlblThanksCount.Text = "Closing in... 1s."
        ElseIf printcountThanksDown = 0 Then
            TabControl1.Visible = False
            TabControl1.SelectedTab.SuspendLayout()
            TabControl1.SelectedIndex = 0
            TabControl1.SelectedTab.ResumeLayout(True)
            TabControl1.Visible = True
            copyPicScanned.Image = Nothing
            mainFormDateLF.Visible = True
            mainFormTImeL.Visible = True
            printdoc.close()
            loadDeleteFilesInFolder(DataFetcher.ScannedImages)
            printDeleteFilesInFolder(DataFetcher.CachePath) 'Clear Folder
            loadDeleteFilesInFolder(DataFetcher.WifiStoragePath)
            wifitimerChecker.Stop()
            printTimer3.Stop()
            printpanelViewer.Controls.Clear()
            printpanelPrice.Visible = False
            printpanelThanks.Visible = False
            printpanelPrintStatus.Visible = False
            For Each ctrl As Control In printTab.Controls
                ctrl.Enabled = True
            Next
            EnableDisableButton()
        End If
    End Sub
    Private Sub printpanelThanks_VisibleChanged(sender As Object, e As EventArgs) Handles printpanelThanks.VisibleChanged
        printlblThanksCount.Text = "Closing in... 3s."
        printcountThanksDown = 3
    End Sub
    Private Sub printbtnConfirmNo_Click(sender As Object, e As EventArgs) Handles printbtnConfirmNo.Click
        printpanelConfirmation.Visible = False
        For Each ctrl As Control In printTab.Controls
            ctrl.Enabled = True
        Next
    End Sub

    Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles printbtnPrint.Click
        Panel14.Visible = False
        coloredPanelConfirm.Visible = False
        greyscalePanelConfirm.Visible = False
        printpanelConfirmation.Visible = True
        printpanelConfirmation.BringToFront()
        For Each ctrl As Control In printTab.Controls
            If ctrl IsNot printpanelConfirmation Then
                ctrl.Enabled = False
            End If
        Next
        printpanelConfirmation.Enabled = True
    End Sub
    Private Sub TriggerRefund()
        Try
            ' Get the current coin balance
            Dim refundAmount As Integer = CInt(coinCount)

            If refundAmount > 0 Then
                ' Send refund command to Arduino
                If SerialPort1.IsOpen Then
                    SerialPort1.WriteLine("Refund: " & refundAmount.ToString())
                End If

                ' Reset the coin count and update UI
                coinCount = 0
                printlblCoins.Text = Format(coinCount, "0.00")
                scanlblCoins.Text = Format(coinCount, "0.00")
                printPanel14.BackColor = Color.Red
                scanPanel14.BackColor = Color.Red
            End If
        Catch ex As Exception
            MessageBox.Show($"Error during refund: {ex.Message}", "Refund Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub printtimerPrinterStatus_Tick(sender As Object, e As EventArgs) Handles printtimerPrinterStatus.Tick
        Dim printServer As New LocalPrintServer()
        Dim printQueue As PrintQueue = printServer.GetPrintQueue(DataFetcher.PrinterName)
        printQueue.Refresh()
        If printQueue.IsOutOfPaper OrElse printQueue.IsDoorOpened OrElse printQueue.HasPaperProblem OrElse printQueue.IsPaperJammed Then
            printlblPrintStatus.Text = "Printer Error (Refunding Coins)"
            printlblPrintStatus.ForeColor = Color.Red
            printpreloaderStatus.Visible = False
            printlblContactAdmin.Visible = True

            ' Trigger the refund process
            ' Await'
            TriggerRefund()

        ElseIf printQueue.IsTonerLow Then
            printlblPrintStatus.Text = "Low Ink"
            printlblPrintStatus.ForeColor = Color.Red
            printpreloaderStatus.Visible = False
            printlblContactAdmin.Visible = True
        ElseIf printQueue.IsOffline Then
            printlblPrintStatus.Text = "Printer is Offline"
            printlblPrintStatus.ForeColor = Color.Red
            printpreloaderStatus.Visible = False
            printlblContactAdmin.Visible = True
        ElseIf printQueue.IsPrinting Then
            printlblPrintStatus.Text = "Printing..."
            printlblPrintStatus.ForeColor = Color.Green
            printpreloaderStatus.Visible = True
            printlblContactAdmin.Visible = False
        ElseIf printQueue.GetPrintJobInfoCollection().Count = 0 Then
            printpanelConfirmation.Visible = False
            For Each ctrl As Control In printTab.Controls
                If ctrl IsNot printpanelThanks Then
                    ctrl.Enabled = False
                End If
            Next
            printpanelThanks.Enabled = True
            printpanelThanks.Visible = True
            printpanelThanks.BringToFront()
            Me.TopMost = False
            printTimer3.Start()
            printtimerPrinterStatus.Stop()
        End If
    End Sub


    'CODES FOR WIFI
    '==============================================================================
    Public Sub LoadFilesByType(files As FileInfo(), imageResource As Image)
        Array.Sort(files, Function(file1, file2) String.Compare(file1.Name, file2.Name))

        For Each file As FileInfo In files
            wifipanelItem = New Panel With {
                .BackColor = Color.White,
                .Size = New Size(185, 229),
                .Location = New Drawing.Point(4, 5),
                .Tag = file.FullName
            }

            wifipicItem = New Bunifu.Framework.UI.BunifuImageButton With {
                .Image = imageResource,
                .Location = New Drawing.Point(18, 5),
                .Name = "BunifuImageButton1",
                .Size = New Size(148, 171),
                .SizeMode = PictureBoxSizeMode.StretchImage,
                .TabStop = False,
                .Tag = file.FullName
            }

            wifilblFileName = New Label With {
                .BackColor = Color.Transparent,
                .Font = New Drawing.Font("Segoe UI Semibold", 12.0!),
                .ForeColor = Color.DimGray,
                .Location = New Drawing.Point(2, 185),
                .Margin = New Padding(2, 0, 2, 0),
                .Name = "lblFileName",
                .Size = New Size(181, 46),
                .Text = file.Name,
                .TextAlign = ContentAlignment.TopCenter,
                .Visible = True,
                .Tag = file.FullName
            }

            wifipanelItemElipse = New Bunifu.Framework.UI.BunifuElipse With {
               .ElipseRadius = 20,
               .TargetControl = Me.wifipanelItem
            }

            wifipanelItem.Controls.Add(wifipicItem)
            wifipanelItem.Controls.Add(wifilblFileName)
            wifiFlowLayoutPanel1.Controls.Add(wifipanelItem)

            AddHandler wifipanelItem.Click, AddressOf File_Click
            AddHandler wifipicItem.Click, AddressOf File_Click
            AddHandler wifilblFileName.Click, AddressOf File_Click
        Next
    End Sub

    Public Function GetFilesByExtension(directoryInfo As DirectoryInfo, extensions As String()) As FileInfo()
        Dim files As New List(Of FileInfo)()
        Dim dDriveInfo As DriveInfo = New DriveInfo(DataFetcher.WifiStoragePath)

        Try
            If dDriveInfo.IsReady Then
                For Each file As FileInfo In directoryInfo.GetFiles("*.*")
                    Dim extension As String = file.Extension.ToLower()
                    Dim fileName As String = file.Name

                    If Not fileName.StartsWith("~") AndAlso extensions.Contains(extension) Then
                        files.Add(file)
                    End If
                Next

                For Each subdirectory As DirectoryInfo In directoryInfo.GetDirectories()
                    If subdirectory.Name <> "System Volume Information" Then
                        files.AddRange(GetFilesByExtension(subdirectory, extensions))
                    End If
                Next
            End If
        Catch ex As UnauthorizedAccessException

        End Try

        Return files.ToArray()
    End Function

    Public Sub LoadAllFiles()
        Dim directoryInfo As New DirectoryInfo(DataFetcher.WifiStoragePath)

        Dim allFiles As FileInfo() = GetFilesByExtension(directoryInfo, {".pdf", ".doc", ".docx", ".ppt", ".pptx"})

        Dim pdfFiles = allFiles.Where(Function(f) f.Extension.ToLower() = ".pdf").ToArray()
        Dim wordFiles = allFiles.Where(Function(f) f.Extension.ToLower() = ".doc" Or f.Extension.ToLower() = ".docx").ToArray()
        Dim pptFiles = allFiles.Where(Function(f) f.Extension.ToLower() = ".ppt" Or f.Extension.ToLower() = ".pptx").ToArray()

        LoadFilesByType(pdfFiles, My.Resources.Resources.PDF)
        LoadFilesByType(wordFiles, My.Resources.Resources.DOC)
        LoadFilesByType(pptFiles, My.Resources.Resources.PPT)
    End Sub

    Public Async Sub File_Click(sender As Object, e As EventArgs)
        Dim filePath As String = sender.tag.ToString
        Dim fileExtension As String = Path.GetExtension(filePath)
        If fileExtension = ".pdf" Then

            Dim newfile As String = Path.Combine(DataFetcher.CachePath, Path.GetFileName(filePath))

            Try
                For Each ctrl As Control In wifiTab.Controls
                    ctrl.Visible = False
                Next
                wifipanelLoading.Visible = True
                wifitimerLoading.Start()

                File.Copy(filePath, newfile)
                Await Task.Delay(500)
                wifiWaitForFileCreation(newfile)
                printPropertiesForm_loadFrom = "loadFromWifi"
                printPropertiesForm_selectedFile = newfile
                TabControl1.Visible = False
                TabControl1.SelectedTab.SuspendLayout()
                TabControl1.SelectedIndex = 6
                TabControl1.SelectedTab.ResumeLayout(True)
                TabControl1.Visible = True
                printPropertiesForm_Load()
            Catch ex As Exception

            End Try
        Else
            Dim pdfPath As String = Path.ChangeExtension(filePath, ".pdf")
            ConvertToPDF(filePath, fileExtension)

        End If


    End Sub
    Private wifiinputFilePathWord2Pdf As String
    Private wifidestinationDirectoryWord2Pdf As String
    Public Sub ConvertToPDF(inputFilePath As String, fileExtension As String)
        If fileExtension = ".docx" OrElse fileExtension = ".doc" Then
            For Each ctrl As Control In wifiTab.Controls
                If ctrl IsNot wifipanelConvert Then
                    ctrl.Enabled = True
                End If
            Next
            wifipicConvert.Image = My.Resources.Doc_to_Pdf
            For Each ctrl As Control In wifiTab.Controls
                If ctrl IsNot wifipanelConvert Then
                    ctrl.Visible = False
                End If
            Next
            wifipanelConvert.Visible = False
            wifiinputFilePathWord2Pdf = inputFilePath
            wifidestinationDirectoryWord2Pdf = DataFetcher.CachePath
            wifiConvertWordToPDF.Start()
        ElseIf fileExtension = ".pptx" OrElse fileExtension = ".ppt" Then
            wifipptFile = inputFilePath
            'wifipanelPrintLayout.Visible = True
            Guna2Transition1.ShowSync(wifipanelPrintLayout)
            wifipanelPrintLayout.BringToFront()
            wifipptApp = New Microsoft.Office.Interop.PowerPoint.Application()
            wifipresentation = wifipptApp.Presentations.Open(inputFilePath, WithWindow:=MsoTriState.msoFalse)
            wifitotalSlides = wifipresentation.Slides.Count
            wifitxtTo.Text = wifitotalSlides.ToString
        Else
            MessageBox.Show("Unsupported file format.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If
    End Sub

    Public Sub ConvertPPTToPDF(pptFilePath As String, destinationDirectory As String, FromPage As Integer, ToPage As Integer)
        For Each ctrl As Control In wifiTab.Controls
            If ctrl IsNot wifipanelConvert Then
                ctrl.Enabled = True
            End If
        Next
        wifipicConvert.Image = My.Resources.Ppt_to_Pdf
        For Each ctrl As Control In wifiTab.Controls
            If ctrl IsNot wifipanelConvert Then
                ctrl.Visible = False
            End If
        Next
        wifipanelConvert.Visible = True
        Dim OutputTypePage As PpPrintOutputType
        Try
            Select Case wifiprintOption
                Case "Full Page Slides"
                    OutputTypePage = PpPrintOutputType.ppPrintOutputSlides
                Case "Notes Page"
                    OutputTypePage = PpPrintOutputType.ppPrintOutputNotesPages
                Case "Outline"
                    OutputTypePage = PpPrintOutputType.ppPrintOutputOutline
                Case "1 Slide Per Page"
                    OutputTypePage = PpPrintOutputType.ppPrintOutputOneSlideHandouts
                Case "2 Slides Per Page"
                    OutputTypePage = PpPrintOutputType.ppPrintOutputTwoSlideHandouts
                Case "3 Slides Per Page"
                    OutputTypePage = PpPrintOutputType.ppPrintOutputThreeSlideHandouts
                Case "4 Slides Per Page"
                    OutputTypePage = PpPrintOutputType.ppPrintOutputFourSlideHandouts
                Case "6 Slides Per Page"
                    OutputTypePage = PpPrintOutputType.ppPrintOutputSixSlideHandouts
                Case "9 Slides Per Page"
                    OutputTypePage = PpPrintOutputType.ppPrintOutputNineSlideHandouts
            End Select
            Dim pdfoutputPath As String = destinationDirectory & "\" & Path.GetFileNameWithoutExtension(pptFilePath) & ".pdf"
            Dim oRange As Microsoft.Office.Interop.PowerPoint.PrintRange = wifipresentation.PrintOptions.Ranges.Add(FromPage, ToPage)
            wifipresentation.ExportAsFixedFormat(Path:=pdfoutputPath, FixedFormatType:=PpFixedFormatType.ppFixedFormatTypePDF, FrameSlides:=MsoTriState.msoFalse,
                                 OutputType:=OutputTypePage, PrintHiddenSlides:=MsoTriState.msoFalse, PrintRange:=oRange,
                                 RangeType:=PpPrintRangeType.ppPrintSlideRange, HandoutOrder:=PpPrintHandoutOrder.ppPrintHandoutHorizontalFirst)

            wifiWaitForFileCreation(pdfoutputPath)
            printPropertiesForm_selectedFile = pdfoutputPath
            printPropertiesForm_loadFrom = "loadFromWifi"
            TabControl1.Visible = False
            TabControl1.SelectedTab.SuspendLayout()
            TabControl1.SelectedIndex = 6
            TabControl1.SelectedTab.ResumeLayout(True)
            TabControl1.Visible = True
            printPropertiesForm_Load()
        Catch ex As Exception
            'MessageBox.Show(ex.Message)
            wifiinitialView()

        Finally
            If wifipresentation IsNot Nothing Then
                wifipresentation.Close()
                Marshal.ReleaseComObject(wifipresentation)
                wifipresentation = Nothing
            End If
            wifipptApp.Quit()
            Marshal.ReleaseComObject(wifipptApp)
            wifiinitialView()
            wifipanelConvert.Visible = False
        End Try
    End Sub
    Private Sub wifiWaitForFileCreation(filePath As String)
        Do While Not (File.Exists(filePath) AndAlso New FileInfo(filePath).Length > 0)
            System.Threading.Thread.Sleep(100)
        Loop
    End Sub
    Public Sub wifiDeleteFilesInFolder(folderPath As String)
        Try
            Dim files As String() = Directory.GetFiles(folderPath)

            For Each filePath As String In files
                File.Delete(filePath)
            Next

        Catch ex As Exception

        End Try
    End Sub

    Private Sub btnAllFiles_Click(sender As Object, e As EventArgs) Handles wifibtnAllFiles.Click
        wifiFlowLayoutPanel1.Controls.Clear()
        LoadAllFiles()
        wifiPanel1.BringToFront()
        wifibtnAllFiles.BringToFront()
        bringNotifConvertPaneltoFront()

    End Sub

    Private Sub btnPDF_Click(sender As Object, e As EventArgs) Handles wifibtnPDF.Click
        wifiFlowLayoutPanel1.Controls.Clear()
        LoadFilesByType(GetFilesByExtension(New DirectoryInfo(DataFetcher.WifiStoragePath), {".pdf"}), My.Resources.Resources.PDF)
        wifiPanel1.BringToFront()
        wifibtnPDF.BringToFront()
        bringNotifConvertPaneltoFront()
    End Sub

    Private Sub btnDOC_Click(sender As Object, e As EventArgs) Handles wifibtnDOC.Click
        wifiFlowLayoutPanel1.Controls.Clear()
        LoadFilesByType(GetFilesByExtension(New DirectoryInfo(DataFetcher.WifiStoragePath), {".doc", ".docx"}), My.Resources.Resources.DOC)
        wifiPanel1.BringToFront()
        wifibtnDOC.BringToFront()
        bringNotifConvertPaneltoFront()
    End Sub

    Private Sub btnPPT_Click(sender As Object, e As EventArgs) Handles wifibtnPPT.Click
        wifiFlowLayoutPanel1.Controls.Clear()
        LoadFilesByType(GetFilesByExtension(New DirectoryInfo(DataFetcher.WifiStoragePath), {".ppt", ".pptx"}), My.Resources.Resources.PPT)
        wifiPanel1.BringToFront()
        wifibtnPPT.BringToFront()
        bringNotifConvertPaneltoFront()
    End Sub

    Private Sub loadedfilesForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        wifipresentation = Nothing
        Dim dDriveInfo As DriveInfo = New DriveInfo(DataFetcher.WifiStoragePath)

        If dDriveInfo.IsReady Then
            wifitimerLoadAll.Start()
        End If
    End Sub

    Sub bringNotifConvertPaneltoFront()
        wifipanelConvert.BringToFront()
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles wifibtnCancel.Click
        For Each ctrl As Control In wifiTab.Controls
            ctrl.Enabled = True
        Next
        ''wifipanelPrintLayout.Visible = False
        Guna2Transition1.HideSync(wifipanelPrintLayout)
        If wifipresentation IsNot Nothing Then
            wifipresentation.Close()
            Marshal.ReleaseComObject(wifipresentation)
            wifipresentation = Nothing
        End If
        wifipptApp.Quit()
        Marshal.ReleaseComObject(wifipptApp)
    End Sub
    Private Sub btnFullPage_Click(sender As Object, e As EventArgs) Handles wifibtnFullPage.Click
        wifiprintOption = "Full Page Slides"
        wifipanelSlidestoPrint.Visible = True
        wifipanelSlidestoPrint.BringToFront()
        wifipanelPrintLayout.Enabled = False

    End Sub

    Private Sub btnNotesPage_Click(sender As Object, e As EventArgs) Handles wifibtnNotesPage.Click
        wifiprintOption = "Notes Page"
        wifipanelSlidestoPrint.Visible = True
        wifipanelSlidestoPrint.BringToFront()
        wifipanelPrintLayout.Enabled = False
    End Sub

    Private Sub btnOutline_Click(sender As Object, e As EventArgs) Handles wifibtnOutline.Click
        wifiprintOption = "Outline"
        wifipanelSlidestoPrint.Visible = True
        wifipanelSlidestoPrint.BringToFront()
        wifipanelPrintLayout.Enabled = False
    End Sub

    Private Sub btn1Slide_Click(sender As Object, e As EventArgs) Handles wifibtn1Slide.Click
        wifiprintOption = "1 Slide Per Page"
        wifipanelSlidestoPrint.Visible = True
        wifipanelSlidestoPrint.BringToFront()
        wifipanelPrintLayout.Enabled = False
    End Sub

    Private Sub btn2Slide_Click(sender As Object, e As EventArgs) Handles wifibtn2Slide.Click
        wifiprintOption = "2 Slides Per Page"
        wifipanelSlidestoPrint.Visible = True
        wifipanelSlidestoPrint.BringToFront()
        wifipanelPrintLayout.Enabled = False
    End Sub

    Private Sub btn3Slide_Click(sender As Object, e As EventArgs) Handles wifibtn3Slide.Click
        wifiprintOption = "3 Slides Per Page"
        wifipanelSlidestoPrint.Visible = True
        wifipanelSlidestoPrint.BringToFront()
        wifipanelPrintLayout.Enabled = False
    End Sub

    Private Sub btn4Slide_Click(sender As Object, e As EventArgs) Handles wifibtn4Slide.Click
        wifiprintOption = "4 Slides Per Page"
        wifipanelSlidestoPrint.Visible = True
        wifipanelSlidestoPrint.BringToFront()
        wifipanelPrintLayout.Enabled = False
    End Sub

    Private Sub btn6Slide_Click(sender As Object, e As EventArgs) Handles wifibtn6Slide.Click, loadbtn6Slide.Click
        wifiprintOption = "6 Slides Per Page"
        wifipanelSlidestoPrint.Visible = True
        wifipanelSlidestoPrint.BringToFront()
        wifipanelPrintLayout.Enabled = False
    End Sub

    Private Sub btn9Slide_Click(sender As Object, e As EventArgs) Handles wifibtn9Slide.Click
        wifiprintOption = "9 Slides Per Page"
        wifipanelSlidestoPrint.Visible = True
        wifipanelSlidestoPrint.BringToFront()
        wifipanelPrintLayout.Enabled = False
    End Sub

    Private Sub wifiNumericTextBox_KeyPress(sender As Object, e As KeyPressEventArgs) Handles wifitxtTo.KeyPress, wifitxtFrom.KeyPress
        If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub wifiNumberButton_Click(sender As Object, e As EventArgs) Handles wifibtnZero.Click, wifibtnTwo.Click, wifibtnThree.Click, wifibtnSix.Click, wifibtnSeven.Click, wifibtnOne.Click, wifibtnNine.Click, wifibtnFour.Click, wifibtnFive.Click, wifibtnEight.Click, wifibtnZero.Click, wifibtnTwo.Click, wifibtnThree.Click, wifibtnSix.Click, wifibtnSeven.Click, wifibtnOne.Click, wifibtnNine.Click, wifibtnFour.Click, wifibtnFive.Click, wifibtnEight.Click
        Dim digit As String = DirectCast(sender, Guna.UI.WinForms.GunaCircleButton).Text
        If TypeOf wififocusedControl Is Bunifu.UI.WinForms.BunifuTextbox.BunifuTextBox Then
            Dim focusedTextBox As Bunifu.UI.WinForms.BunifuTextbox.BunifuTextBox = DirectCast(wififocusedControl, Bunifu.UI.WinForms.BunifuTextbox.BunifuTextBox)
            If focusedTextBox.Text.Length < 3 Then
                focusedTextBox.Text &= digit
            End If
        End If
    End Sub

    Private Sub wifibtnErase_Click(sender As Object, e As EventArgs) Handles wifibtnErase.Click
        If TypeOf wififocusedControl Is Bunifu.UI.WinForms.BunifuTextbox.BunifuTextBox Then
            Dim focusedTextBox As Bunifu.UI.WinForms.BunifuTextbox.BunifuTextBox = CType(wififocusedControl, Bunifu.UI.WinForms.BunifuTextbox.BunifuTextBox)

            If focusedTextBox.Text.Length > 0 Then
                focusedTextBox.Text = focusedTextBox.Text.Substring(0, focusedTextBox.Text.Length - 1)
            End If
        End If
    End Sub

    Private Sub rdAll_Click(sender As Object, e As EventArgs) Handles wifirdAll.Click, printrdAll.Click
        wifipanelKeypad.Enabled = False
        wifipanelSpecific.Enabled = False
        wifitxtFrom.Text = "1"
        wifitxtTo.Text = wifitotalSlides.ToString
    End Sub

    Private Sub rdSpecific_Click(sender As Object, e As EventArgs) Handles wifirdSpecific.Click, printrdSpecific.Click
        wifipanelKeypad.Enabled = True
        wifipanelSpecific.Enabled = True
    End Sub

    Private Sub lblAllpages_Click(sender As Object, e As EventArgs) Handles wifilblAllpages.Click, printlblAllpages.Click
        wifirdAll.Checked = True
        wifirdSpecific.Checked = False
        wifipanelKeypad.Enabled = False
        wifipanelSpecific.Enabled = False
        wifitxtFrom.Text = "1"
        wifitxtTo.Text = wifitotalSlides.ToString
    End Sub

    Private Sub lblSpecific_Click(sender As Object, e As EventArgs) Handles wifilblSpecific.Click, printlblSpecific.Click
        wifirdSpecific.Checked = True
        wifirdAll.Checked = False
        wifipanelKeypad.Enabled = True
        wifipanelSpecific.Enabled = True
    End Sub

    Private Sub FromToChecker()
        If (wifitxtFrom.Text = "" Or Val(wifitxtFrom.Text) = 0) And (wifitxtTo.Text = "" Or Val(wifitxtTo.Text) = 0) Then
            wifilblError.Visible = True
            wifilblError.Text = "'From' and 'To' value is required."
            wifibtnOk.Enabled = False
            wifibtnOkay.Enabled = False
        ElseIf wifitxtFrom.Text = "" Or Val(wifitxtFrom.Text) = 0 Then
            wifilblError.Visible = True
            wifilblError.Text = "'From' value is required."
            wifibtnOk.Enabled = False
            wifibtnOkay.Enabled = False
        ElseIf wifitxtTo.Text = "" Or Val(wifitxtTo.Text) = 0 Then
            wifilblError.Visible = True
            wifilblError.Text = "'To' value is required."
            wifibtnOk.Enabled = False
            wifibtnOkay.Enabled = False
        Else
            If Val(wifitxtFrom.Text) > Val(wifitxtTo.Text) Then
                wifilblError.Visible = True
                wifilblError.Text = "'From' value must be greater than the 'To' value."
                wifibtnOk.Enabled = False
                wifibtnOkay.Enabled = False
            ElseIf Val(wifitxtFrom.Text) > wifitotalSlides Then
                wifilblError.Visible = True
                wifilblError.Text = "'From' value has exceeded the total number of pages"
                wifibtnOk.Enabled = False
                wifibtnOkay.Enabled = False
            ElseIf Val(wifitxtTo.Text) > wifitotalSlides Then
                wifilblError.Visible = True
                wifilblError.Text = "'To' value has exceeded the total number of pages"
                wifibtnOk.Enabled = False
                wifibtnOkay.Enabled = False
            Else
                wifilblError.Visible = False
                wifibtnOk.Enabled = True
                wifibtnOkay.Enabled = True
            End If
        End If

    End Sub

    Private Sub btnCancelSpecific_Click(sender As Object, e As EventArgs) Handles wifibtnCancelSpecific.Click
        wifipanelSlidestoPrint.Visible = False
        wifipanelPrintLayout.Enabled = True
        wifipanelKeypad.Enabled = False
        wifipanelSpecific.Enabled = False
        wifirdAll.Checked = True
        wifirdSpecific.Checked = False
        wifitxtFrom.Text = "1"
        wifitxtTo.Text = wifitotalSlides.ToString
    End Sub

    Private Sub txtFrom_Enter(sender As Object, e As EventArgs) Handles wifitxtFrom.Enter, printtxtFrom.Enter
        wifipanelKeypad.Visible = True
        wififocusedControl = wifitxtFrom
        FromToChecker()
    End Sub

    Private Sub txtTo_Enter(sender As Object, e As EventArgs) Handles wifitxtTo.Enter, printtxtTo.Enter
        wifipanelKeypad.Visible = True
        wififocusedControl = wifitxtTo
        FromToChecker()
    End Sub

    Private Sub txtFrom_TextChanged(sender As Object, e As EventArgs) Handles wifitxtFrom.TextChanged, printtxtFrom.TextChanged
        FromToChecker()
    End Sub

    Private Sub txtTo_TextChanged(sender As Object, e As EventArgs) Handles wifitxtTo.TextChanged, printtxtTo.TextChanged
        FromToChecker()
    End Sub

    Private Sub btnOkay_Click(sender As Object, e As EventArgs) Handles wifibtnOkay.Click, wifibtnOk.Click
        wifipanelSlidestoPrint.Visible = False
        ConvertPPTToPDF(wifipptFile, DataFetcher.CachePath, Val(wifitxtFrom.Text), Val(wifitxtTo.Text))
    End Sub

    Private Sub btnFromInc_Click(sender As Object, e As EventArgs) Handles wifibtnFromInc.Click
        If wifitxtFrom.Text.Length > 0 Then
            Dim currentValue As Integer = Integer.Parse(wifitxtFrom.Text)
            Dim newValue As Integer = (currentValue Mod wifitotalSlides) + 1

            ' Ensure newValue is less than or equal to txtTo value
            If newValue <= Integer.Parse(wifitxtTo.Text) Then
                wifitxtFrom.Text = newValue.ToString()
            End If
        Else
            wifitxtFrom.Text = "1"
        End If
    End Sub

    Private Sub btnFromDec_Click(sender As Object, e As EventArgs) Handles wifibtnFromDec.Click
        If wifitxtFrom.Text.Length > 0 Then
            Dim currentValue As Integer = Integer.Parse(wifitxtFrom.Text)
            Dim newValue As Integer = If(currentValue > 1, currentValue - 1, wifitotalSlides)

            ' Ensure newValue is less than or equal to txtTo value
            If newValue <= Integer.Parse(wifitxtTo.Text) Then
                wifitxtFrom.Text = newValue.ToString()
            End If
        Else
            wifitxtFrom.Text = wifitotalSlides
        End If
    End Sub

    Private Sub btnToInc_Click(sender As Object, e As EventArgs) Handles wifibtnToInc.Click
        If wifitxtTo.Text.Length > 0 Then
            Dim currentValue As Integer = Integer.Parse(wifitxtTo.Text)
            Dim newValue As Integer = (currentValue Mod wifitotalSlides) + 1

            ' Ensure newValue is greater than or equal to txtFrom value
            If newValue >= Integer.Parse(wifitxtFrom.Text) Then
                wifitxtTo.Text = newValue.ToString()
            End If
        Else
            wifitxtTo.Text = "1"
        End If
    End Sub

    Private Sub btnToDec_Click(sender As Object, e As EventArgs) Handles wifibtnToDec.Click
        If wifitxtTo.Text.Length > 0 Then
            Dim currentValue As Integer = Integer.Parse(wifitxtTo.Text)
            Dim newValue As Integer = If(currentValue > 1, currentValue - 1, wifitotalSlides)

            ' Ensure newValue is greater than or equal to txtFrom value
            If newValue >= Integer.Parse(wifitxtFrom.Text) Then
                wifitxtTo.Text = newValue.ToString()
            End If
        Else
            wifitxtTo.Text = wifitotalSlides
        End If
    End Sub

    Private Sub timerLoadAll_Tick(sender As Object, e As EventArgs) Handles wifitimerLoadAll.Tick
        LoadAllFiles()
        wifitimerLoadAll.Stop()
    End Sub
    Private Sub timerLoading_Tick(sender As Object, e As EventArgs) Handles wifitimerLoading.Tick
        wifiinitialView()
        wifitimerLoading.Stop()
    End Sub


    Public Sub QrGenerator()
        Dim qrText As String = DataFetcher.ServerLink
        Dim gen As New QRCodeGenerator
        Dim data = gen.CreateQrCode(qrText, QRCodeGenerator.ECCLevel.Q)
        Dim code As New QRCode(data)
        wifipicQr.Image = code.GetGraphic(6)

        wifilblWebsite.Text = "Or visit: " + qrText
    End Sub

    Private Sub wifiReceiveForm_Load()
        wifiloadFormat = "ALL"
        QrGenerator()
        wifitimerChecker.Start()
    End Sub

    Private Sub timerChecker_Tick(sender As Object, e As EventArgs) Handles wifitimerChecker.Tick
        Dim filesInFolder As String() = Directory.GetFiles(DataFetcher.WifiStoragePath)
        If filesInFolder.Length > 0 Then
            wififilesReceived = filesInFolder.Length
            wifiFlowLayoutPanel1.Visible = True
            wifilblNoFiles.Visible = False
            wifiFlowLayoutPanel1.Enabled = True
            wifitimerChecker.Stop()

            wifiFlowLayoutPanel1.Controls.Clear()
            LoadAllFiles()
            wifiPanel1.BringToFront()
            wifibtnAllFiles.BringToFront()
            bringNotifConvertPaneltoFront()
        Else
            wifiFlowLayoutPanel1.Visible = False
            wifilblNoFiles.Visible = True
        End If
    End Sub

    Private Sub btnShowPrices_Click(sender As Object, e As EventArgs) Handles wifibtnYes.Click
        wifiDeleteFilesInFolder(DataFetcher.WifiStoragePath)
        TabControl1.Visible = False
        TabControl1.SelectedTab.SuspendLayout()
        TabControl1.SelectedIndex = 1
        TabControl1.SelectedTab.ResumeLayout(True)
        TabControl1.Visible = True
        wifipanelConfirm.Visible = False
    End Sub

    Private Sub BunifuButton1_Click(sender As Object, e As EventArgs) Handles wifiBunifuButton1.Click
        wifipanelConfirm.Visible = False
        wifiFlowLayoutPanel1.Enabled = True
        wifipanelConfirm.SendToBack()
    End Sub

    Private Sub timerReceive_Tick(sender As Object, e As EventArgs) Handles wifitimerReceive.Tick
        Dim filesInFolder As String() = Directory.GetFiles(DataFetcher.WifiStoragePath)
        If filesInFolder.Length > wififilesReceived Then
            wififilesReceived = filesInFolder.Length
            wifiFlowLayoutPanel1.Controls.Clear()
            If wifiloadFormat = "ALL" Then
                LoadAllFiles()
            ElseIf wifiloadFormat = "PDF" Then
                LoadFilesByType(GetFilesByExtension(New DirectoryInfo(DataFetcher.WifiStoragePath), {".pdf"}), My.Resources.Resources.PDF)
            ElseIf wifiloadFormat = "DOC" Then
                LoadFilesByType(GetFilesByExtension(New DirectoryInfo(DataFetcher.WifiStoragePath), {".doc", ".docx"}), My.Resources.Resources.DOC)
            ElseIf wifiloadFormat = "PPT" Then
                LoadFilesByType(GetFilesByExtension(New DirectoryInfo(DataFetcher.WifiStoragePath), {".ppt", ".pptx"}), My.Resources.Resources.PPT)
            End If
        End If
    End Sub
    Private Sub wifibtnBack_Click(sender As Object, e As EventArgs) Handles wifibtnBack.Click
        wifipanelConfirm.Visible = True
        wifiFlowLayoutPanel1.Enabled = False
        wifipanelConfirm.BringToFront()
    End Sub

    Private Sub wifiConvertWordToPDF_Tick(sender As Object, e As EventArgs) Handles wifiConvertWordToPDF.Tick
        Try
            Dim pdfFilePath As String = Path.Combine(wifidestinationDirectoryWord2Pdf, Path.GetFileNameWithoutExtension(wifiinputFilePathWord2Pdf) & ".pdf")
            Dim doc As New asposeWords.Document(wifiinputFilePathWord2Pdf)

            doc.Save(pdfFilePath, asposeWords.SaveFormat.Pdf)
            wifiWaitForFileCreation(pdfFilePath)
            printPropertiesForm_selectedFile = pdfFilePath
            printPropertiesForm_loadFrom = "loadFromWifi"

            TabControl1.Visible = False
            TabControl1.SelectedTab.SuspendLayout()
            TabControl1.SelectedIndex = 6
            TabControl1.SelectedTab.ResumeLayout(True)
            TabControl1.Visible = True
            printPropertiesForm_Load()
            wifiConvertWordToPDF.Stop()
        Catch ex As Exception
            wifiinitialView()
            wifiConvertWordToPDF.Stop()
        End Try
        wifiinitialView()

    End Sub

    Private Sub wifiinitialView()
        For Each ctrl As Control In wifiTab.Controls
            If ctrl IsNot wifipanelPrintLayout AndAlso ctrl IsNot wifipanelSlidestoPrint AndAlso ctrl IsNot wifipanelLoading AndAlso ctrl IsNot wifipanelConfirm AndAlso ctrl IsNot wifipanelConvert Then
                ctrl.Visible = True
            End If
        Next
        wifipanelConvert.Visible = False
        wifipanelConfirm.Visible = False
        wifipanelConfirm.SendToBack()
    End Sub

    'CODES FOR SCANNER
    '===================================================

    Private Sub ScannerForm_Load()
        setupPort()
        scanTimer1.Start()
        scanflashdriveDirectory = DataFetcher.FlashDrivePath
        loadDeleteFilesInFolder(DataFetcher.ScannedImages)
    End Sub
    Private Sub scanSaveImagesToPDF()
        scanlblstatus.Text = "Saving..."
        ' Get all image paths from the folder
        Dim imagePaths As String() = Directory.GetFiles(DataFetcher.ScannedImages)

        ' Check if the folder contains any images
        If imagePaths.Length = 0 Then
            MessageBox.Show("The folder does not contain any images.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
        Try
            ' Create a new PDF document
            Dim pdfDocument As New asposePdf.Document()

            ' Loop through all images
            For Each imagePath As String In imagePaths
                ' Create a new page and add it to the PDF document
                Dim page As asposePdf.Page = pdfDocument.Pages.Add()

                ' Add an image to the page
                Dim image As New Aspose.Pdf.Image()
                page.Paragraphs.Add(image)

                ' Set the image file stream
                image.File = imagePath
            Next
            For Each page As asposePdf.Page In pdfDocument.Pages
                ' Remove margins
                page.PageInfo.Margin.Bottom = 0
                page.PageInfo.Margin.Top = 0
                page.PageInfo.Margin.Left = 0
                page.PageInfo.Margin.Right = 0
            Next

            Dim timestamp As String = DateTime.Now.ToString("mmddyyhhmmss")
            Dim pdfFilePath As String = DataFetcher.FlashDrivePath & "PrintVendoScan_" & timestamp & ".pdf"
            ' Save the PDF document
            pdfDocument.Save(pdfFilePath)

            pdfDocument.Dispose()
            ' Release resources associated with the image files
            For Each imagePath As String In imagePaths
                File.Delete(imagePath)
            Next
            loadDeleteFilesInFolder(DataFetcher.ScannedImages)
            scanPicScanned.Image = Nothing
            scanlbltotalpage.Text = "0"
            scanscannedPages = 0
            scanpanelThanks.Visible = True
            scanpanelThanks.Enabled = True
            scanpanelThanks.BringToFront()
            scanbtnSave.Visible = False
            scanbtnScan.Enabled = False
            scanPanel14.BackColor = Color.Red
            scanTimer3.Start()
            scanpanelConfirmation.Visible = False
            For Each ctrl As Control In scannerTab.Controls
                ctrl.Enabled = True
            Next

            scanTimer1.Stop()
        Catch ex As Exception
            MessageBox.Show($"Error saving file: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TabControl1.Visible = False
            TabControl1.SelectedTab.SuspendLayout()
            TabControl1.SelectedIndex = 0
            TabControl1.SelectedTab.ResumeLayout(True)
            TabControl1.Visible = True
            EnableDisableButton()
        End Try
        scanPanelstat.Visible = False
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles scanTimer2.Tick
        scancountDown = scancountDown - 1
        If scancountDown = 3 Then
            scanlbldown.Text = "Flash drive disconnected. Closing in... 3s."
        ElseIf scancountDown = 2 Then
            scanlbldown.Text = "Flash drive disconnected. Closing in... 2s."
        ElseIf scancountDown = 1 Then
            scanlbldown.Text = "Flash drive disconnected. Closing in... 1s."
        ElseIf scancountDown = 0 Then
            TabControl1.Visible = False
            TabControl1.SelectedTab.SuspendLayout()
            TabControl1.SelectedIndex = 0
            TabControl1.SelectedTab.ResumeLayout(True)
            TabControl1.Visible = True
            scanTimer1.Stop()
            scanTimer2.Stop()
            EnableDisableButton()
        End If
    End Sub
    Private Sub Timer3_Tick(sender As Object, e As EventArgs) Handles scanTimer3.Tick
        scancountThanksDown = scancountThanksDown - 1
        If scancountThanksDown = 3 Then
            scanlblThanksCount.Text = "Closing in... 3s."
        ElseIf scancountThanksDown = 2 Then
            scanlblThanksCount.Text = "Closing in... 2s."
        ElseIf scancountThanksDown = 1 Then
            scanlblThanksCount.Text = "Closing in... 1s."
        ElseIf scancountThanksDown = 0 Then
            TabControl1.Visible = False
            TabControl1.SelectedTab.SuspendLayout()
            TabControl1.SelectedIndex = 0
            TabControl1.SelectedTab.ResumeLayout(True)
            TabControl1.Visible = True
            scanTimer3.Stop()
            scanpanelThanks.Visible = False
            scanPicScanned.Image = Nothing
            EnableDisableButton()
        End If
    End Sub
    Private Sub scanCheckDriveAvailability()
        Dim dDriveInfo As DriveInfo = New DriveInfo(scanflashdriveDirectory)

        If dDriveInfo.IsReady Then
            scanpanelNotif.Visible = False
            scanTimer2.Enabled = False
            If (Val(scanlbltotalpage.Text) > 0) And (Val(scanlblCoins.Text) >= Val(scanlblTotalPrice.Text)) Then
                scanbtnSave.Visible = True
            Else
                scanbtnSave.Visible = False
            End If

            If scanpanelConfirmation.Visible = True Then
                scanbtnScan.Enabled = False
            Else
                scanbtnScan.Enabled = True
            End If
        Else
            scanpanelNotif.Visible = True
            scanpanelNotif.Enabled = True
            scanpanelNotif.BringToFront()
            scanTimer2.Enabled = True
            scanbtnSave.Visible = False
            scanbtnScan.Enabled = False
            scanPanel14.BackColor = Color.Red
        End If
    End Sub

    Private Sub scanTimer1_Tick(sender As Object, e As EventArgs) Handles scanTimer1.Tick
        scanCheckDriveAvailability()
    End Sub

    Private Sub panelNotif1_VisibleChanged(sender As Object, e As EventArgs) Handles scanpanelNotif1.VisibleChanged
        scanlbldown.Text = "Flash drive disconnected. Closing in... 3s."
        scancountDown = 3
    End Sub
    Private Sub panelThanks_VisibleChanged(sender As Object, e As EventArgs) Handles scanpanelThanks.VisibleChanged
        scanlbldown.Text = "Closing in... 3s."
        scancountThanksDown = 3
    End Sub


    Private Sub scanbtnBack_Click(sender As Object, e As EventArgs) Handles scanbtnBack.Click
        loadDeleteFilesInFolder(DataFetcher.ScannedImages)
        scanPicScanned.Image = Nothing
        TabControl1.Visible = False
        scanlbltotalpage.Text = "0"
        scanscannedPages = 0
        TabControl1.SelectedTab.SuspendLayout()
        TabControl1.SelectedIndex = 0
        TabControl1.SelectedTab.ResumeLayout(True)
        TabControl1.Visible = True
        scanTimer1.Stop()
        Dim coins As Decimal = Decimal.Parse(scanlblCoins.Text)
        Dim withdrawCoins As Integer = CInt(coins)
        changefunction(withdrawCoins)
        EnableDisableButton()
    End Sub

    Private Sub scanbtnConfirmYes_Click(sender As Object, e As EventArgs) Handles scanbtnConfirmYes.Click
        Try

            'scanlblstatus.Text = "Saving..."
            Dim change As Integer
            Dim coins As Decimal = Decimal.Parse(scanlblCoins.Text)
            Dim totalPrice As Decimal = Decimal.Parse(scanlblTotalPrice.Text)
            change = CInt(coins) - CInt(totalPrice)
            If change > 0 Then
                changefunction(change)
            End If
            timeDispenseScan.Start()
            scanPanelstat.Visible = True

        Catch ex As Exception
            MessageBox.Show("Flashdrive is out of space or corrupted")
        End Try

    End Sub

    Private Sub scanbtnConfirmNo_Click(sender As Object, e As EventArgs) Handles scanbtnConfirmNo.Click
        scanpanelConfirmation.Visible = False
        For Each ctrl As Control In scannerTab.Controls
            ctrl.Enabled = True
        Next
    End Sub
    Private Sub scanlbltotalpage_TextChanged(sender As Object, e As EventArgs) Handles scanlbltotalpage.TextChanged
        Dim totalPrice As Decimal = scanscanPagePrice * Val(scanlbltotalpage.Text)
        scanlblTotalPrice.Text = Format(totalPrice, "0.00")
    End Sub

    Private Sub scanbtnScan_Click(sender As Object, e As EventArgs) Handles scanbtnScan.Click
        Try
            scanPanelScanning.Visible = True
            scanPanelScanning.BringToFront()
            scanPicScanned.Image = Nothing
            Dim deviceManager As New WIA.DeviceManager()

            Dim deviceInfo As DeviceInfo = Nothing
            For Each info As DeviceInfo In deviceManager.DeviceInfos
                If info.Type = WiaDeviceType.ScannerDeviceType Then
                    deviceInfo = info
                    Exit For
                End If
            Next

            ' Check if a scanner is found
            If deviceInfo IsNot Nothing Then
                Dim device As Device = deviceInfo.Connect()

                Dim item As Item = device.Items(1)

                Dim propertyIdHorizontal As Object = 6147
                Dim propertyIdVertical As Object = 6148

                For Each prop As WIA.Property In item.Properties
                    If prop.PropertyID = propertyIdVertical Then
                        prop.Value = 200
                    End If
                    If prop.PropertyID = propertyIdHorizontal Then
                        prop.Value = 200
                    End If
                Next

                Dim imageFile As ImageFile = DirectCast(item.Transfer(WIA.FormatID.wiaFormatJPEG), ImageFile)

                Dim imageBytes() As Byte = DirectCast(imageFile.FileData.BinaryData, Byte())

                Using memoryStream As New IO.MemoryStream(imageBytes)
                    Dim bitmap As New System.Drawing.Bitmap(memoryStream)
                    scanscannedPages += 1
                    scanPicScanned.Image = bitmap
                    Dim savePath As String = DataFetcher.ScannedImages
                    bitmap.Save(savePath & "\scannedImage_" & scanscannedPages & ".jpeg")
                End Using

                scanlbltotalpage.Text = scanscannedPages
            Else
                MessageBox.Show("No scanner found", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
            scanPanelScanning.Visible = False
        Catch ex As Exception
            MessageBox.Show($"Error scanning: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            scanPanelScanning.Visible = False
        End Try
    End Sub

    Private Sub scanbtnSave_Click(sender As Object, e As EventArgs) Handles scanbtnSave.Click
        scanpanelConfirmation.BringToFront()
        scanpanelConfirmation.Visible = True
        For Each ctrl As Control In scannerTab.Controls
            If ctrl IsNot scanpanelConfirmation Then
                ctrl.Enabled = False
            End If
        Next
    End Sub

    'CODES FOR COPY FORM

    Private Sub btnScan_Click(sender As Object, e As EventArgs) Handles copybtnScan.Click
        Try
            copyPanelScanning.Visible = True
            copyPanelScanning.BringToFront()
            copyPicScanned.Image = Nothing
            Dim deviceManager As New WIA.DeviceManager()

            Dim deviceInfo As DeviceInfo = Nothing
            For Each info As DeviceInfo In deviceManager.DeviceInfos
                If info.Type = WiaDeviceType.ScannerDeviceType Then
                    deviceInfo = info
                    Exit For
                End If
            Next

            ' Check if a scanner is found
            If deviceInfo IsNot Nothing Then
                Dim device As Device = deviceInfo.Connect()

                Dim item As Item = device.Items(1)

                Dim propertyIdHorizontal As Object = 6147
                Dim propertyIdVertical As Object = 6148

                For Each prop As WIA.Property In item.Properties
                    If prop.PropertyID = propertyIdVertical Then
                        prop.Value = 200
                    End If
                    If prop.PropertyID = propertyIdHorizontal Then
                        prop.Value = 200
                    End If
                Next

                Dim imageFile As ImageFile = DirectCast(item.Transfer(WIA.FormatID.wiaFormatJPEG), ImageFile)

                Dim imageBytes() As Byte = DirectCast(imageFile.FileData.BinaryData, Byte())

                Using memoryStream As New IO.MemoryStream(imageBytes)
                    Dim bitmap As New System.Drawing.Bitmap(memoryStream)
                    copyscannedPages += 1
                    copyPicScanned.Image = bitmap
                    Dim savePath As String = DataFetcher.ScannedImages
                    bitmap.Save(savePath & "\scannedImage_" & copylbltotalpage.Text & ".jpeg")
                End Using

                copylbltotalpage.Text = copyscannedPages
            Else
                MessageBox.Show("No scanner found", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
            copyPanelScanning.Visible = False
        Catch ex As Exception
            MessageBox.Show($"Error scanning: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            copyPanelScanning.Visible = False
        End Try
    End Sub
    Private Sub btnBack_Click(sender As Object, e As EventArgs) Handles copybtnBack.Click
        loadDeleteFilesInFolder(DataFetcher.ScannedImages)
        copyPicScanned.Image = Nothing
        TabControl1.Visible = False
        copylbltotalpage.Text = "0"
        copyscannedPages = 0
        TabControl1.SelectedTab.SuspendLayout()
        TabControl1.SelectedIndex = 0
        TabControl1.SelectedTab.ResumeLayout(True)
        TabControl1.Visible = True
        EnableDisableButton()
    End Sub

    Private Sub copybtnSave_Click(sender As Object, e As EventArgs) Handles copybtnSave.Click
        copypanelConfirmation.BringToFront()
        copypanelConfirmation.Visible = True
        For Each ctrl As Control In copyTab.Controls
            If ctrl IsNot copypanelConfirmation Then
                ctrl.Enabled = False
            End If
        Next
    End Sub

    Private Sub copySaveImagesToPDF()

        ' Get all image paths from the folder
        Dim imagePaths As String() = Directory.GetFiles(DataFetcher.ScannedImages)

        ' Check if the folder contains any images
        If imagePaths.Length = 0 Then
            MessageBox.Show("The folder does not contain any images.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
        Try
            ' Create a new PDF document
            Dim pdfDocument As New asposePdf.Document()

            ' Loop through all images
            For Each imagePath As String In imagePaths
                ' Create a new page and add it to the PDF document
                Dim page As asposePdf.Page = pdfDocument.Pages.Add()

                ' Add an image to the page
                Dim image As New Aspose.Pdf.Image()
                page.Paragraphs.Add(image)

                ' Set the image file stream
                image.File = imagePath
            Next
            For Each page As asposePdf.Page In pdfDocument.Pages
                ' Remove margins
                page.PageInfo.Margin.Bottom = 0
                page.PageInfo.Margin.Top = 0
                page.PageInfo.Margin.Left = 0
                page.PageInfo.Margin.Right = 0
            Next

            Dim timestamp As String = DateTime.Now.ToString("mmddyyhhmmss")
            Dim pdfFilePath As String = DataFetcher.CachePath & "\PrintVendoScan_" & timestamp & ".pdf"
            ' Save the PDF document
            pdfDocument.Save(pdfFilePath)
            pdfDocument.Dispose()

            ' Release resources associated with the image files
            For Each imagePath As String In imagePaths
                File.Delete(imagePath)
            Next
            loadDeleteFilesInFolder(DataFetcher.ScannedImages)
            copyPicScanned.Image = Nothing
            copylbltotalpage.Text = "0"
            copyscannedPages = 0
            printPropertiesForm_selectedFile = pdfFilePath
            printPropertiesForm_loadFrom = "loadFromCopy"
            TabControl1.Visible = False
            TabControl1.SelectedTab.SuspendLayout()
            TabControl1.SelectedIndex = 6
            TabControl1.SelectedTab.ResumeLayout(True)
            TabControl1.Visible = True
            printPropertiesForm_Load()
        Catch ex As Exception
            MessageBox.Show($"Error saving file: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TabControl1.Visible = False
            TabControl1.SelectedTab.SuspendLayout()
            TabControl1.SelectedIndex = 0
            TabControl1.SelectedTab.ResumeLayout(True)
            TabControl1.Visible = True
            EnableDisableButton()
        End Try
        copypanelConfirmation.Visible = False
        For Each ctrl As Control In copyTab.Controls
            ctrl.Enabled = True
        Next
    End Sub

    Private Sub copybtnConfirmYes_Click(sender As Object, e As EventArgs) Handles copybtnConfirmYes.Click
        Try
            copySaveImagesToPDF()
        Catch ex As Exception
            'MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub copybtnConfirmNo_Click(sender As Object, e As EventArgs) Handles copybtnConfirmNo.Click
        copypanelConfirmation.Visible = False
        For Each ctrl As Control In copyTab.Controls
            ctrl.Enabled = True
        Next
    End Sub

    Private Sub lbltotalpage_TextChanged(sender As Object, e As EventArgs) Handles copylbltotalpage.TextChanged
        If Val(copylbltotalpage.Text) = 0 Then
            copybtnSave.Visible = False
        Else
            copybtnSave.Visible = True
        End If
    End Sub

    'Code FOR ADMIN
    Private Sub adminbtnBack_Click(sender As Object, e As EventArgs) Handles adminbtnBack.Click
        TabControl1.Visible = False
        TabControl1.SelectedTab.SuspendLayout()
        TabControl1.SelectedIndex = 0
        TabControl1.SelectedTab.ResumeLayout(True)
        TabControl1.Visible = True
        mainFormDateL.Visible = True
        mainFormTImeL.Visible = True
        EnableDisableButton()
    End Sub

    Private Sub adminbtnBrowseFlashC_Click(sender As Object, e As EventArgs) Handles adminbtnBrowseCache.Click
        Dim folderBrowserDialog As New FolderBrowserDialog()
        If folderBrowserDialog.ShowDialog() = DialogResult.OK Then
            Dim selectedFolderPath As String = folderBrowserDialog.SelectedPath
            admintxtCache.Text = selectedFolderPath
        End If
    End Sub

    Private Sub adminbtnBrowseWifi_Click(sender As Object, e As EventArgs) Handles adminbtnBrowseWifi.Click
        Dim folderBrowserDialog As New FolderBrowserDialog()
        If folderBrowserDialog.ShowDialog() = DialogResult.OK Then
            Dim selectedFolderPath As String = folderBrowserDialog.SelectedPath
            admintxtWifi.Text = selectedFolderPath
        End If
    End Sub

    Private Sub adminSettingsForm_Load()
        printerList()
        portList()
        textboxSize()
        FetchData()
    End Sub
    Private Function CheckPrinter(ByVal printerName As String) As Boolean
        Try
            Dim printDocument As PrintDocument = New PrintDocument
            printDocument.PrinterSettings.PrinterName = printerName
            Return printDocument.PrinterSettings.IsValid
        Catch ex As System.Exception
            Return False
        End Try
    End Function
    Sub printerList()
        printforLong.Items.Clear()
        printforA4.Items.Clear()
        Dim InstalledPrinters As String
        For Each InstalledPrinters In System.Drawing.Printing.PrinterSettings.InstalledPrinters
            printforLong.Items.Add(InstalledPrinters)
            printforA4.Items.Add(InstalledPrinters)
        Next InstalledPrinters
    End Sub
    Sub portList()
        admincomboPorts.Items.Clear()
        For Each sp As String In My.Computer.Ports.SerialPortNames
            admincomboPorts.Items.Add(sp)
        Next
    End Sub
    Sub textboxSize()
        admintxtCache.Size = New Size(393, 46)
        admintxtWifi.Size = New Size(393, 46)
        admintxtServer.Size = New Size(308, 46)
        admintxtBlank.Size = New Size(160, 46)
        admintxtBW.Size = New Size(160, 46)
        admintxtColored.Size = New Size(160, 46)
        admintxtBlank.MaxLength = 3
        admintxtBW.MaxLength = 3
        admintxtColored.MaxLength = 3
        admintxtSystemPin.MaxLength = 8
    End Sub

    Private Sub btnBlankDec_Click(sender As Object, e As EventArgs) Handles adminbtnBlankInc.Click
        If admintxtBlank.Text.Length > 0 Then
            Dim currentValue As Integer = Integer.Parse(admintxtBlank.Text)
            Dim newValue As Integer = currentValue + 1
            admintxtBlank.Text = newValue.ToString()
        Else
            admintxtBlank.Text = "1"
        End If
    End Sub

    Private Sub btnBlankInc_Click(sender As Object, e As EventArgs) Handles adminbtnBlankDec.Click
        If admintxtBlank.Text.Length > 0 Then
            Dim currentValue As Integer = Integer.Parse(admintxtBlank.Text)
            Dim newValue As Integer = If(currentValue > 1, currentValue - 1, 1)
            admintxtBlank.Text = newValue.ToString()
        Else
            admintxtBlank.Text = "1"
        End If
    End Sub

    Private Sub btnBWDec_Click(sender As Object, e As EventArgs) Handles adminbtnBWInc.Click
        If admintxtBW.Text.Length > 0 Then
            Dim currentValue As Integer = Integer.Parse(admintxtBW.Text)
            Dim newValue As Integer = currentValue + 1
            admintxtBW.Text = newValue.ToString()
        Else
            admintxtBW.Text = "1"
        End If
    End Sub

    Private Sub btnBWInc_Click(sender As Object, e As EventArgs) Handles adminbtnBWDec.Click
        If admintxtBW.Text.Length > 0 Then
            Dim currentValue As Integer = Integer.Parse(admintxtBW.Text)
            Dim newValue As Integer = If(currentValue > 1, currentValue - 1, 1)
            admintxtBW.Text = newValue.ToString()
        Else
            admintxtBW.Text = "1"
        End If
    End Sub

    Private Sub btnColoredDec_Click(sender As Object, e As EventArgs) Handles adminbtnColoredInc.Click
        If admintxtColored.Text.Length > 0 Then
            Dim currentValue As Integer = Integer.Parse(admintxtColored.Text)
            Dim newValue As Integer = currentValue + 1
            admintxtColored.Text = newValue.ToString()
        Else
            admintxtColored.Text = "1"
        End If
    End Sub

    Private Sub btnColoredInc_Click(sender As Object, e As EventArgs) Handles adminbtnColoredDec.Click
        If admintxtColored.Text.Length > 0 Then
            Dim currentValue As Integer = Integer.Parse(admintxtColored.Text)
            Dim newValue As Integer = If(currentValue > 1, currentValue - 1, 1)
            admintxtColored.Text = newValue.ToString()
        Else
            admintxtColored.Text = "1"
        End If
    End Sub
    Private Sub btnScanDec_Click(sender As Object, e As EventArgs) Handles adminbtnScanInc.Click
        If admintxtScan.Text.Length > 0 Then
            Dim currentValue As Integer = Integer.Parse(admintxtScan.Text)
            Dim newValue As Integer = currentValue + 1
            admintxtScan.Text = newValue.ToString()
        Else
            admintxtScan.Text = "1"
        End If
    End Sub

    Private Sub btnScanInc_Click(sender As Object, e As EventArgs) Handles adminbtnScanDec.Click
        If admintxtScan.Text.Length > 0 Then
            Dim currentValue As Integer = Integer.Parse(admintxtScan.Text)
            Dim newValue As Integer = If(currentValue > 1, currentValue - 1, 1)
            admintxtScan.Text = newValue.ToString()
        Else
            admintxtScan.Text = "1"
        End If
    End Sub
    Private Sub FetchData()
        Try
            admincombodrive.Text = DataFetcher.FlashDrivePath
            admintxtCache.Text = DataFetcher.CachePath
            admintxtWifi.Text = DataFetcher.WifiStoragePath
            admintxtBlank.Text = DataFetcher.BlankPagePrice
            admintxtBW.Text = DataFetcher.BWPagePrice
            admintxtColored.Text = DataFetcher.ColoredPagePrice
            admintxtScan.Text = DataFetcher.ScanPagePrice
            admincomboPorts.Text = DataFetcher.CoinSlotPort
            printforLong.Text = My.Settings.LongPrinterName
            printforA4.Text = My.Settings.A4PrinterName
            admintxtSystemPin.Text = DataFetcher.SystemPin
            admintxtServer.Text = DataFetcher.ServerLink

            settingsChanged = False

            If admincomboPorts.Text = "" Then
                adminbtn1.Enabled = False
                adminbtn5.Enabled = False
            Else
                adminbtn1.Enabled = True
                adminbtn5.Enabled = True
            End If
        Catch ex As Exception
            'MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub Setting_Changed(sender As Object, e As EventArgs) Handles admintxtCache.TextChanged, admintxtWifi.TextChanged, admintxtBlank.TextChanged, admintxtBW.TextChanged, admintxtColored.TextChanged, admintxtScan.TextChanged, admincomboPorts.SelectedIndexChanged, printforLong.SelectedIndexChanged, printforA4.SelectedIndexChanged, admintxtSystemPin.TextChanged, admintxtServer.TextChanged
        settingsChanged = True
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles adminbtnSave.Click
        Dim cachePath As String = admintxtCache.Text
        Dim wifiPath As String = admintxtWifi.Text
        If admincombodrive.Text = "" Or admintxtCache.Text = "" Or admintxtWifi.Text = "" Or admintxtBlank.Text = "" Or printforLong.Text = "" Or admintxtSystemPin.Text = "" Or admintxtServer.Text = "" Or admintxtBW.Text = "" Or admintxtColored.Text = "" Or admintxtScan.Text = "" Then
            MsgBox("Dont leave any of the fields blank!", MsgBoxStyle.Exclamation, "Save Error")
        ElseIf Val(admintxtBW.Text) > (admintxtColored.Text) Then
            MsgBox("Black and White price must be lesst than Colored price", MsgBoxStyle.Exclamation, "Save Error")
        ElseIf Not Directory.Exists(cachePath) Then
            MsgBox("Entered cache file path is not a valid path", MsgBoxStyle.Exclamation, "Save Error")
        ElseIf Not Directory.Exists(wifiPath) Then
            MsgBox("Entered wifi file path is not a valid path", MsgBoxStyle.Exclamation, "Save Error")
        Else
            adminpanelConfirmation.Visible = True

            For Each ctrl As Control In adminTab.Controls
                If ctrl IsNot adminpanelConfirmation Then
                    ctrl.Enabled = False
                End If
            Next
        End If
    End Sub

    Private Sub btnShowHide_Click(sender As Object, e As EventArgs) Handles adminbtnShowHide.Click
        If admintxtSystemPin.PasswordChar = "*" Then
            admintxtSystemPin.PasswordChar = ""
            adminbtnShowHide.Text = "Hide"
        Else
            admintxtSystemPin.PasswordChar = "*"
            adminbtnShowHide.Text = "Show"
        End If
    End Sub

    Private Sub btnConfirmNo_Click(sender As Object, e As EventArgs) Handles adminbtnConfirmNo.Click
        adminpanelConfirmation.Visible = False
        For Each ctrl As Control In adminTab.Controls
            If ctrl IsNot adminpanelConfirmation Then
                ctrl.Enabled = True
            End If
        Next
    End Sub

    Private Sub btnConfirmYes_Click(sender As Object, e As EventArgs) Handles adminbtnConfirmYes.Click
        If settingsChanged Then
            adminpanelSaving.Visible = True
            adminTimer1.Start()

            ' Save settings
            My.Settings.flash_Drive = admincombodrive.SelectedItem.ToString()
            My.Settings.cache_Path = admintxtCache.Text.ToString()
            My.Settings.wifi_Storage_Path = admintxtWifi.Text.ToString()
            My.Settings.server_Link = admintxtServer.Text.ToString()
            My.Settings.blank_Page_Price = Val(admintxtBlank.Text)
            My.Settings.bw_Page_Price = Val(admintxtBW.Text)
            My.Settings.colored_Page_Price = Val(admintxtColored.Text)
            My.Settings.scan_Page_Price = Val(admintxtScan.Text)

            If admincomboPorts.SelectedItem IsNot Nothing Then
                My.Settings.coin_Slot_Port = admincomboPorts.SelectedItem.ToString()
            Else
                My.Settings.coin_Slot_Port = ""
            End If

            If printforA4.SelectedItem IsNot Nothing Then
                My.Settings.A4PrinterName = printforA4.SelectedItem.ToString()
            Else
                My.Settings.A4PrinterName = ""
            End If

            If printforLong.SelectedItem IsNot Nothing Then
                My.Settings.LongPrinterName = printforLong.SelectedItem.ToString()
            Else
                My.Settings.LongPrinterName = ""
            End If

            My.Settings.system_Pin = admintxtSystemPin.Text.ToString()
            My.Settings.Save()
            DataFetcher.FetchData()
            FetchData()


            settingsChanged = False
        Else
            MsgBox("No changes detected, settings are already up to date.", MsgBoxStyle.Information, "No Changes")
        End If
    End Sub

    Dim time As Integer
    Private settingsChanged As Boolean

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles adminTimer1.Tick
        time += 1
        If time = 3 Then
            adminlblSaving.Text = "Saved!"
        ElseIf time = 4 Then
            adminpanelSaving.Visible = False
            adminTimer1.Stop()
            adminpanelConfirmation.Visible = False
            time = 0
            For Each ctrl As Control In adminTab.Controls
                ctrl.Enabled = True
            Next
            adminlblSaving.Text = "Saving..."
        End If
    End Sub

    Private Sub NumericTextBox_KeyPress(sender As Object, e As KeyPressEventArgs) Handles admintxtSystemPin.KeyPress, admintxtColored.KeyPress, admintxtBW.KeyPress, admintxtBlank.KeyPress
        If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub NumberButton_Click(sender As Object, e As EventArgs)
        Dim digit As String = DirectCast(sender, Guna.UI.WinForms.GunaCircleButton).Text
        If TypeOf adminfocusedControl Is Bunifu.UI.WinForms.BunifuTextbox.BunifuTextBox Then
            Dim focusedTextBox As Bunifu.UI.WinForms.BunifuTextbox.BunifuTextBox = DirectCast(adminfocusedControl, Bunifu.UI.WinForms.BunifuTextbox.BunifuTextBox)
            If focusedTextBox.Text.Length < 8 Then
                focusedTextBox.Text &= digit
            End If
        End If
    End Sub

    Private Sub btnErase_Click(sender As Object, e As EventArgs)
        If TypeOf adminfocusedControl Is Bunifu.UI.WinForms.BunifuTextbox.BunifuTextBox Then
            Dim focusedTextBox As Bunifu.UI.WinForms.BunifuTextbox.BunifuTextBox = CType(adminfocusedControl, Bunifu.UI.WinForms.BunifuTextbox.BunifuTextBox)

            If focusedTextBox.Text.Length > 0 Then
                focusedTextBox.Text = focusedTextBox.Text.Substring(0, focusedTextBox.Text.Length - 1)
            End If
        End If
    End Sub

    Private Sub txtCoinCustom_Enter(sender As Object, e As EventArgs)
        adminfocusedControl = admintxtSystemPin
    End Sub

    Private Sub btn20_Click(sender As Object, e As EventArgs)
        adminoutputCoin = "20"
        adminlabelWithdrawText.Text = "Withdraw 20 peso coins"
        adminpanelWithdrawing.Visible = True
        For Each ctrl As Control In adminTab.Controls
            If ctrl IsNot adminpanelWithdrawing Then
                ctrl.Enabled = False
            End If
        Next
    End Sub

    Private Sub btn10_Click(sender As Object, e As EventArgs)
        adminoutputCoin = "10"
        adminlabelWithdrawText.Text = "Withdraw 10 peso coins"
        adminpanelWithdrawing.Visible = True
        For Each ctrl As Control In adminTab.Controls
            If ctrl IsNot adminpanelWithdrawing Then
                ctrl.Enabled = False
            End If
        Next
    End Sub

    Private Sub btn5_Click(sender As Object, e As EventArgs) Handles adminbtn5.Click
        adminoutputCoin = "5"
        adminlabelWithdrawText.Text = "Withdraw 5 peso coins"
        adminpanelWithdrawing.Visible = True
        For Each ctrl As Control In adminTab.Controls
            If ctrl IsNot adminpanelWithdrawing Then
                ctrl.Enabled = False
            End If
        Next
    End Sub

    Private Sub btn1_Click(sender As Object, e As EventArgs) Handles adminbtn1.Click
        adminoutputCoin = "1"
        adminlabelWithdrawText.Text = "Withdraw 1 peso coins"
        adminpanelWithdrawing.Visible = True
        For Each ctrl As Control In adminTab.Controls
            If ctrl IsNot adminpanelWithdrawing Then
                ctrl.Enabled = False
            End If
        Next
    End Sub

    Private Sub btnWithdrawYes_Click(sender As Object, e As EventArgs)
        WithdrawCoin()
    End Sub
    Private Sub WithdrawCoin()
        adminpanelWithdrawing.Visible = True
        adminpanelWithdrawing.Enabled = True
        adminpanelWithdrawing.BringToFront()

        Try
            If adminoutputCoin = "1" Then
                SerialPort1.WriteLine("Withdraw1")
            ElseIf adminoutputCoin = "5" Then
                SerialPort1.WriteLine("Withdraw5")
            End If
        Catch ex As Exception
            MsgBox("Coin Slot is not Connected, Set up the port first!", MsgBoxStyle.Exclamation, "Port Error")
            For Each ctrl As Control In adminTab.Controls
                ctrl.Enabled = True
            Next
        End Try
    End Sub

    Private Sub Guna2CircleButton1_Click(sender As Object, e As EventArgs) Handles adminbtnrefetch.Click
        printerList()
        portList()
        textboxSize()
        FetchData()
        setupPort()

        admintimerPrinterStatus.Start()
    End Sub


    Public Function GetLocalIPAddress() As String
        Dim hostName As String = Dns.GetHostName()
        Dim ipEntry As IPHostEntry = Dns.GetHostEntry(hostName)

        For Each ipAddress As IPAddress In ipEntry.AddressList
            If ipAddress.AddressFamily = AddressFamily.InterNetwork Then
                Return ipAddress.ToString()
            End If
        Next
        Return "No IPv4 address found"
    End Function

    Private Sub btngetServer_Click(sender As Object, e As EventArgs) Handles adminbtngetServer.Click
        Dim localIPAddress As String = GetLocalIPAddress()
        admintxtServer.Text = localIPAddress
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles adminbtnClose.Click
        adminwithdraw = False
        Try
            SerialPort1.WriteLine("WithdrawStop")
        Catch ex As Exception
            MsgBox("Coin Slot is not Connected, Set up the port first!", MsgBoxStyle.Exclamation, "Port Error")
            For Each ctrl As Control In adminTab.Controls
                ctrl.Enabled = True
            Next
        End Try

        adminpanelWithdrawing.Visible = False
        adminpanelWithdrawing.Enabled = False
        adminpanelOngoingWithdraw.Visible = False
        For Each ctrl As Control In adminTab.Controls
            ctrl.Enabled = True
        Next
    End Sub

    Private Sub btnWithdrawYes_Click_2(sender As Object, e As EventArgs) Handles adminbtnWithdrawYes.Click
        adminpanelOngoingWithdraw.Visible = True
        WithdrawCoin()
    End Sub

    Private Sub btnWithdrawNo_Click(sender As Object, e As EventArgs) Handles adminbtnWithdrawNo.Click
        adminpanelWithdrawing.Visible = False
        adminpanelWithdrawing.Enabled = False
        For Each ctrl As Control In adminTab.Controls
            ctrl.Enabled = True
        Next
    End Sub

    Private Sub timerPrinterStatus_Tick(sender As Object, e As EventArgs) Handles admintimerPrinterStatus.Tick
        Try
            Dim printServer As New LocalPrintServer()
            Dim printQueue As PrintQueue = printServer.GetPrintQueue(DataFetcher.PrinterName)
            printQueue.Refresh()
            If printQueue.IsOutOfPaper Then
                adminlblPrinterStatus.Text = "Out of Paper"
                adminlblPrinterStatus.ForeColor = Color.Red
            ElseIf printQueue.IsPaperJammed Then
                adminlblPrinterStatus.Text = "Paper Jammed"
                adminlblPrinterStatus.ForeColor = Color.Red
            ElseIf printQueue.HasPaperProblem Then
                adminlblPrinterStatus.Text = "Paper Jammed"
                adminlblPrinterStatus.ForeColor = Color.Red
            ElseIf printQueue.IsTonerLow Then
                adminlblPrinterStatus.Text = "Low Ink"
                adminlblPrinterStatus.ForeColor = Color.Red
            ElseIf printQueue.IsDoorOpened Then
                adminlblPrinterStatus.Text = "Door Open"
                adminlblPrinterStatus.ForeColor = Color.Red
            ElseIf printQueue.IsOutOfMemory Then
                adminlblPrinterStatus.Text = "Out of Memory"
                adminlblPrinterStatus.ForeColor = Color.Red
            ElseIf printQueue.IsInError Then
                adminlblPrinterStatus.Text = "Error Occured"
                adminlblPrinterStatus.ForeColor = Color.Red
            ElseIf printQueue.IsOffline Then
                adminlblPrinterStatus.Text = "Printer is offline"
                adminlblPrinterStatus.ForeColor = Color.Red
            ElseIf printQueue.IsPrinting Then
                adminlblPrinterStatus.Text = "Printing"
                adminlblPrinterStatus.ForeColor = Color.Green
            Else
                adminlblPrinterStatus.Text = "Idle"
                adminlblPrinterStatus.ForeColor = Color.Green
            End If
        Catch ex As Exception
            'MessageBox.Show(ex.Message)
            admintimerPrinterStatus.Stop()
        End Try

    End Sub

    Private Sub adminPrintSwitch_CheckedChanged(sender As Object, e As EventArgs) Handles adminPrintSwitch.CheckedChanged
        EnableDisableButton()
    End Sub

    Private Sub adminCopySwitch_CheckedChanged(sender As Object, e As EventArgs) Handles adminCopySwitch.CheckedChanged
        EnableDisableButton()
    End Sub

    Private Sub adminScanSwitch_CheckedChanged(sender As Object, e As EventArgs) Handles adminScanSwitch.CheckedChanged
        EnableDisableButton()
    End Sub
    Public Sub EnableDisableButton()
        If adminPrintSwitch.Checked = True Then
            landingbtnPrint.Enabled = True
        Else
            landingbtnPrint.Enabled = False
        End If
        If adminCopySwitch.Checked = True Then
            landingbtnCopy.Enabled = True
        Else
            landingbtnCopy.Enabled = False
        End If
        If adminScanSwitch.Checked = True Then
            landingbtnScan.Enabled = True
        Else
            landingbtnScan.Enabled = False
        End If
    End Sub

    Private Sub timerColor_Tick(sender As Object, e As EventArgs) Handles timerColor.Tick
        If Val(printlblCoins.Text) >= Val(printlblTotalPrice.Text) Then
            printPanel14.BackColor = Color.Green
            printlblCoins.ForeColor = Color.Green
            Label42.ForeColor = Color.Green
            Label41.ForeColor = Color.Green
        Else
            printPanel14.BackColor = Color.Red
            printlblCoins.ForeColor = Color.Red
            Label42.ForeColor = Color.Red
            Label41.ForeColor = Color.Red
        End If


        If (Val(scanlbltotalpage.Text) > 0) And (Val(scanlblCoins.Text) >= Val(scanlblTotalPrice.Text)) Then
            scanPanel14.BackColor = Color.Green
            scanlblCoins.ForeColor = Color.Green
            Label24.ForeColor = Color.Green
            Label23.ForeColor = Color.Green
        Else
            scanPanel14.BackColor = Color.Red
            scanlblCoins.ForeColor = Color.Red
            Label24.ForeColor = Color.Red
            Label23.ForeColor = Color.Red
        End If
    End Sub

    Private Async Sub TimeDispense_Tick(sender As Object, e As EventArgs) Handles TimeDispense.Tick
        If dispensing = True Then
            printlblPrintStatus.Text = "Dispensing..."
        Else
            TimeDispense.Stop()
            Me.TopMost = True
            printlblPrintStatus.Text = "Printing"
            Await PrintDocs()
            changebtnClose.Visible = True

            printpreloaderStatus.Visible = True
            'Await Task.Delay(2000)
            printtimerPrinterStatus.Start()
            copylbltotalpage.Text = 0
            loadDeleteFilesInFolder(DataFetcher.ScannedImages)
        End If
    End Sub

    Private Sub timerCheckingDispense_Tick(sender As Object, e As EventArgs) Handles timerCheckingDispense.Tick
        If dispensing = True Then
            hopperElapsed += 1
        Else
            hopperElapsed = 0
        End If

        If hopperElapsed = 10 Then
            dispensing = False
            changeAvailable = True

            panelNoChange.Visible = True
            adminChangeSwitch.Checked = False
            Dim totalBal As Integer
            totalBal = remaining1 + remaining5
            changePanelBal.Text = "Remaining balance: " & totalBal.ToString() & "PHP"
            If SerialPort1.IsOpen Then
                SerialPort1.WriteLine("SA")
            End If
        End If

        If changeAvailable = True Then
            lblnoChange.Visible = False
        ElseIf changeAvailable = False Then
            lblnoChange.Visible = True
        End If
    End Sub

    Private Sub timeDispenseScan_Tick(sender As Object, e As EventArgs) Handles timeDispenseScan.Tick
        If dispensing = True Then
            scanlblstatus.Text = "Dispensing..."
        Else
            timeDispenseScan.Stop()
            scanSaveImagesToPDF()
            changebtnClose.Visible = True
        End If
    End Sub

    Private Sub changebtnClose_Click(sender As Object, e As EventArgs) Handles changebtnClose.Click
        panelNoChange.Visible = False
        changebtnClose.Visible = False
        remaining1 = 0
        remaining5 = 0
        hopperElapsed = 0
    End Sub

    Private Sub adminChangeSwitch_CheckedChanged(sender As Object, e As EventArgs) Handles adminChangeSwitch.CheckedChanged
        If adminChangeSwitch.Checked = True Then
            changeAvailable = True
        Else
            changeAvailable = False
        End If
    End Sub

    Private Sub printbtnShowPrices_Click(sender As Object, e As EventArgs) Handles printbtnShowPrices.Click
        If printpanelPriceDetails.Visible = False Then
            printpanelPriceDetails.Size = New Size(302, 589)
            printpanelPriceDetails.Visible = True
            printpanelPriceDetails.Location = New System.Drawing.Point(376, 166)
            printbtnShowPrices.Text = "Hide price details"
        Else
            printpanelPriceDetails.Size = New Size(10, 589)
            printpanelPriceDetails.Visible = False
            printpanelPriceDetails.Location = New System.Drawing.Point(668, 166)
            printbtnShowPrices.Text = "Show price details"
        End If
    End Sub

    Private Sub BunifuButton2_Click(sender As Object, e As EventArgs) Handles greyscaleBtnYes.Click
        ' Add the condition to check if the label's text is "Letter"
        If Label98.Text = "Letter" Then
            MessageBox.Show("This Printer only Allows Long & A4 Size", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        ' The rest of the code runs only if the above condition is not met
        printoutputColor = "Greyscale"
        printpanelHolder.Visible = False
        printpanelKeypad.Visible = False

        If printrdAll.Checked = True And BunifuRadioButton1.Checked Then
            printstartPage = 1
            printendPage = printtotalPages
            Label101.Text = "Long"
        Else
            printstartPage = Val(printtxtFrom.Text)
            printendPage = Val(printtxtTo.Text)
            Label101.Text = "A4"
        End If

        printnumberCopies = Val(printtxtCopies.Text)

        printpanelConfirmation.Visible = False
        For Each ctrl As Control In printTab.Controls
            ctrl.Enabled = True
        Next


        ShowDetails()
    End Sub

    Private Sub BunifuButton1_Click_1(sender As Object, e As EventArgs) Handles greyscaleBtnNo.Click

        printpanelConfirmation.Visible = False
        For Each ctrl As Control In printTab.Controls
            ctrl.Enabled = True
        Next
    End Sub

    Private Sub BunifuButton4_Click(sender As Object, e As EventArgs) Handles coloredBtnYes.Click
        If Label98.Text = "Letter" Then
            MessageBox.Show("This Printer only Allows Long & A4 Size", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If


        printoutputColor = "Colored"
        printpanelHolder.Visible = False
        printpanelKeypad.Visible = False

        If printrdAll.Checked = True And BunifuRadioButton1.Checked Then
            printstartPage = 1
            printendPage = printtotalPages
            Label101.Text = "Long"
        Else
            printstartPage = Val(printtxtFrom.Text)
            printendPage = Val(printtxtTo.Text)
            Label101.Text = "A4"
        End If
        printnumberCopies = Val(printtxtCopies.Text)

        printpanelConfirmation.Visible = False
        For Each ctrl As Control In printTab.Controls
            ctrl.Enabled = True
        Next


        ShowDetails()
    End Sub

    Private Sub BunifuButton3_Click(sender As Object, e As EventArgs) Handles coloredBtnNo.Click
        printpanelConfirmation.Visible = False
        For Each ctrl As Control In printTab.Controls
            ctrl.Enabled = True
        Next
    End Sub

    Private Async Sub BunifuButton2_Click_1(sender As Object, e As EventArgs) Handles BunifuButton2.Click
        printoutputColor = "Smart Pricing"
        printpanelConfirmation.Visible = False
        printpanelProps.Visible = False
        printpanelHolder.Visible = False
        printpanelKeypad.Visible = False
        printbtnBack.Visible = False
        printbtnPrev.Visible = False
        printbtnNext.Visible = False
        printpanelCalculating.Visible = True
        printpanelViewCalc.Visible = True

        If printrdAll.Checked = True And BunifuRadioButton1.Checked Then
            printstartPage = 1
            printendPage = printtotalPages
            Label101.Text = "Long"
        Else
            printstartPage = Val(printtxtFrom.Text)
            printendPage = Val(printtxtTo.Text)
            Label101.Text = "A4"
        End If
        printnumberCopies = Val(printtxtCopies.Text)


        If printpageOrientation = "Portrait" Then
            printpanelCalculating.Size = New Size(367, 552)
            printpanelCalculating.Location = New System.Drawing.Point(777, 389)
            printpanelViewCalc.Size = New Size(367, 475)
            printpanelViewCalc.Location = New System.Drawing.Point(777, 389)
        ElseIf printpageOrientation = "Landscape" Then
            printpanelCalculating.Size = New Size(461, 423)
            printpanelCalculating.Location = New System.Drawing.Point(730, 448)
            printpanelViewCalc.Size = New Size(461, 346)
            printpanelViewCalc.Location = New System.Drawing.Point(730, 448)
        End If

        Dim doc As New Viscomsoft.PDFViewer.PDFDocument
        If doc.open(printPropertiesForm_selectedFile) Then
            Await Task.Delay(100)
            print_pdfviewerCalc = New Viscomsoft.PDFViewer.PDFView
            print_pdfviewerCalc.Document = doc
            print_pdfviewerCalc.Canvas.Parent = Me.printpanelViewCalc
            print_pdfviewerCalc.Canvas.BackColor = Color.White
            print_pdfviewerCalc.Canvas.Size = New Size(Me.printpanelViewCalc.ClientSize.Width, Me.printpanelViewCalc.ClientSize.Height)
            print_pdfviewerCalc.Zoom = Zoom.FitPage
            Await ProcessPagesAsync()
        Else
            doc.close()
        End If
        printpanelProps.Visible = True
        printpanelHolder.Visible = True
        printbtnBack.Visible = True
        printbtnPrev.Visible = True
        printbtnNext.Visible = True

        printpanelConfirmation.Visible = False
        For Each ctrl As Control In printTab.Controls
            ctrl.Enabled = True
        Next
    End Sub

    Private Sub BunifuButton1_Click_2(sender As Object, e As EventArgs) Handles BunifuButton1.Click
        printpanelConfirmation.Visible = False
        For Each ctrl As Control In printTab.Controls
            ctrl.Enabled = True
        Next
    End Sub
End Class