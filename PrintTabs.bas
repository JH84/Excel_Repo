Attribute VB_Name = "PrintTabs"
Sub PrintPDF()

' PrintExReport PDF Macro for Netcare Sub billing

Application.ScreenUpdating = False

Dim pdfNameSite As String
Dim EndDate As String
Dim Page As String


    EndDate = Sheets("Montana_Ampath").Range("d1").Value
    pdfNameSite = ["NCR"]
    
    SheetNames = Array("Montana_Ampath", "Montana_Coffee", "Montana_Renal_Care_Normal", "Montana_Renal_Care_Emergency", "Montana_Rad_MRI_AC", "Montana_Rad_Emergency")
    
    For Each tennant In SheetNames
    
    Page = tennant
        
    Sheets(tennant).Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
            "N:\NETCARE (NCR)\AEM\Billing & Tariff\Sub-Billing\PDF\" + pdfNameSite & " " & Page & "_" & EndDate & ".pdf" _
            , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
            :=False, OpenAfterPublish:=True
            
    Next tennant
    
            
End Sub

