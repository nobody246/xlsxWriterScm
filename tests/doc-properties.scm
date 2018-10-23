(use xlsxwriterscm)

(create-workbook "doc-properties.xlsx")
(add-worksheet "")
(set-doc-properties "This is an example spreadsheet"
                    "With document properties"
                    "Alex S"
                    "Dr. Heinz Doofenshmirtz"
                    "of Wolves"
                    "Example spreadsheets"
                    "Sample, Example, Properties"
                    "Created with libxlsxwriter & xlsxwriterscm"
                    "Quo")
(worksheet-set-column 50 0 0)
(worksheet-write-string "Select 'Workbook Properties' to see properties.")
(close-workbook)
(exit)
