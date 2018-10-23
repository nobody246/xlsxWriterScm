(use xlsxwriterscm)

(create-workbook "doc-custom-properties.xlsx")
(add-worksheet "")
(workbook-set-custom-property-string "Checked By" "Eve")
(workbook-set-custom-property-date-time "Date Completed" 2016 12 12 0 0 0)
(workbook-set-custom-property-number "Document Number" 12345)
(workbook-set-custom-property-number "Reference Number" 1.2345)
(workbook-set-custom-property-number "Has Review" 1)
(workbook-set-custom-property-number "Signed Off" 0)

(worksheet-set-column 50 0 0)
(worksheet-write-string "Select 'Workbook Properties' to see properties.")

(close-workbook)
(exit)
