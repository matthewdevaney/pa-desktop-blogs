# Import drivers
Add-Type -Path "C:\rpa\fillpdf\itextsharp.dll"
Add-Type -Path "C:\rpa\fillpdf\BouncyCastle.Crypto.dll"

# Variables for PDF locations
$PdfFileInput = "C:\rpa\fillpdf\CertificateOfOrigin.pdf"
$PdfFileOutput = "C:\rpa\fillpdf\CertificateOfOrigin_1HGCG1659WA030328.pdf"


# Create PDF Reader & Stamper Objects
$PdfReader = New-Object iTextSharp.text.pdf.PdfReader($PdfFileInput)
$PdfStamper = New-Object iTextSharp.text.pdf.PdfStamper($PdfReader, [System.IO.File]::Create($PdfFileOutput))

# Fill PDF Fields With These Values
$PdfFields = @{
 IssueDate = "5/1/2022"
 InvoiceNumber = "103743"
 VehicleIDNumber = "1HGCG1659WA030328"
 ModelYear = "2019"
 Model = "D4500"
 Manufacturer = "Company Name"
}

# Fill Each PDF Field And Set To Read-Only
ForEach ($PdfField in $PdfFields.GetEnumerator()) {
    $PdfStamper.AcroFields.SetField($PdfField.Key, $PdfField.Value)
    $PdfStamper.AcroFields.SetFieldProperty($PdfField.Key, "setfflags", [iTextSharp.text.pdf.PdfFormField]::FF_READ_ONLY, 0)
}

# Close PDF
$PdfStamper.Close()
$PdfReader.Close()



































<#
# Display PDF Fields
 $PdfFields = $PdfReader.AcroFields.Fields
 echo $PdfFields

# Write data to PDF fields
$PdfStamper.AcroFields.SetField("IssueDate", "5/1/2022")
$PdfStamper.AcroFields.SetField("InvoiceNumber", "103743")
$PdfStamper.AcroFields.SetField("VehicleIDNumber", "1HGCG1659WA030328")
$PdfStamper.AcroFields.SetField("ModelYear", "2019")
$PdfStamper.AcroFields.SetField("Model", "D4500")
$PdfStamper.AcroFields.SetField("Manufacturer", "Company Name")

# Set fields to read-only
$PdfStamper.AcroFields.SetFieldProperty("IssueDate", "setfflags", [iTextSharp.text.pdf.PdfFormField]::FF_READ_ONLY, 0)
$PdfStamper.AcroFields.SetFieldProperty("InvoiceNumber", "setfflags", [iTextSharp.text.pdf.PdfFormField]::FF_READ_ONLY, 0)
$PdfStamper.AcroFields.SetFieldProperty("VehicleIDNumber", "setfflags", [iTextSharp.text.pdf.PdfFormField]::FF_READ_ONLY, 0)
$PdfStamper.AcroFields.SetFieldProperty("ModelYear", "setfflags", [iTextSharp.text.pdf.PdfFormField]::FF_READ_ONLY, 0)
$PdfStamper.AcroFields.SetFieldProperty("Model", "setfflags", [iTextSharp.text.pdf.PdfFormField]::FF_READ_ONLY, 0)
$PdfStamper.AcroFields.SetFieldProperty("Manufacturer", "setfflags", [iTextSharp.text.pdf.PdfFormField]::FF_READ_ONLY, 0)
#>