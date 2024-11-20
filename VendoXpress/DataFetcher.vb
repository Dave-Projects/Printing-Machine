Imports asposePDF = Aspose.Pdf
Imports asposeWords = Aspose.Words
Public Class DataFetcher
    Public Shared Property FlashDrivePath As String
    Public Shared Property CachePath As String
    Public Shared Property ScannedImages As String
    Public Shared Property WifiStoragePath As String
    Public Shared Property BlankPagePrice As Decimal
    Public Shared Property BWPagePrice As Decimal
    Public Shared Property ColoredPagePrice As Decimal
    Public Shared Property ScanPagePrice As Decimal
    Public Shared Property CoinSlotPort As String
    Public Shared Property PrinterName As String
    Public Shared Property NumberOfPapers As Integer
    Public Shared Property SystemPin As String
    Public Shared Property ServerLink As String

    Public Shared Sub FetchData()
        Try
            FlashDrivePath = My.Settings.flash_Drive
            CachePath = My.Settings.cache_Path
            ScannedImages = My.Settings.scanned_Images
            WifiStoragePath = My.Settings.wifi_Storage_Path
            ServerLink = My.Settings.server_Link
            BlankPagePrice = My.Settings.blank_Page_Price
            BWPagePrice = My.Settings.bw_Page_Price
            ColoredPagePrice = My.Settings.colored_Page_Price
            CoinSlotPort = My.Settings.coin_Slot_Port
            PrinterName = My.Settings.printer_Name
            SystemPin = My.Settings.system_Pin
            ScanPagePrice = My.Settings.scan_Page_Price
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Public Shared Sub ApplyAsposeLicense()
        Dim licenseWords As New asposeWords.License()
        licenseWords.SetLicense("C:\License.lic")
        Dim licensePDF As New asposePDF.License()
        licensePDF.SetLicense("C:\License.lic")
    End Sub
End Class
