(use xlsxwriterscm)

(create-workbook "worksheet-protection.xlsx")
(add-worksheet "")

(define-formats '((hidden ((format-set-hidden)))
                  (unlocked ((format-set-unlocked)))
                  (none ())))

(worksheet-set-column 40 0 0)
(worksheet-protect-settings password: "")

(worksheet-write-string "B1 is locked, and cannot be edited")
(set-row 1)
(worksheet-write-string "B2 is unlocked, and can be edited")
(set-row 2)
(worksheet-write-string "B3 formula is hidden")

(set-pos 0 1)
(worksheet-write-formula "=1+2")
(set-row 1)
(set-format 'unlocked)
(worksheet-write-formula "=1+2")
(set-row 2)
(set-format 'hidden)
(worksheet-write-formula "=1+2")
(set-row 2)
(close-workbook)
(exit)



                  
