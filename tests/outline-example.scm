(use xlsxwriterscm)
(create-workbook "outline-example.xlsx")
(define-formats '((bold ((set-bold)))
                  (italic ((set-italic)))
                  (center ((set-align ($align 'center))))
                  (superscript ((set-superscript)))
                  (none ())))

(add-worksheet "Outlined rows")
(create-row-col-opt 1 0 0)
(create-row-col-opt 2 0 0)

(set-col-option-index 0)
(worksheet-set-column 20 0 0)

(set-col-option-index 0)
(worksheet-set-row-opt $lxw-def-row-height 0)
(worksheet-set-row-opt $lxw-def-row-height 1)
(worksheet-set-row-opt $lxw-def-row-height 2)
(worksheet-set-row-opt $lxw-def-row-height 3)
(set-col-option-index 1)
(worksheet-set-row-opt $lxw-def-row-height 4)

(set-col-option-index 0)
(worksheet-set-row-opt $lxw-def-row-height 6)
(worksheet-set-row-opt $lxw-def-row-height 7)
(worksheet-set-row-opt $lxw-def-row-height 8)
(worksheet-set-row-opt $lxw-def-row-height 9)
(set-col-option-index 1)
(worksheet-set-row-opt $lxw-def-row-height 10)



(set-format 'bold)
(worksheet-write-string "Region")
(set-format 'none)
(set-row 1)
(worksheet-write-string "North")
(set-row 2)
(worksheet-write-string "North")
(set-row 3)
(worksheet-write-string "North")
(set-row 4)
(set-format 'bold)
(worksheet-write-string "North Total")

(set-pos 0 1)
(worksheet-write-string "Sales")
(set-format 'none)
(set-row 1)
(worksheet-write-number 1000)
(set-row 2)
(worksheet-write-number 1200)
(set-row 3)
(worksheet-write-number 1200)
(set-row 4)
(set-format 'bold)
(worksheet-write-formula "=SUBTOTAL(9,B2:B5)")

(set-pos 6 0)
(worksheet-write-string "Region")
(set-format 'none)
(set-row 7)
(worksheet-write-string "South")
(set-row 8)
(worksheet-write-string "South")
(set-row 9)
(worksheet-write-string "South")
(set-row 10)
(set-format 'bold)
(worksheet-write-string "South Total")

(set-format 'none)
(set-pos 6 1)
(worksheet-write-number 400)
(set-row 7)
(worksheet-write-number 600)
(set-row 8)
(worksheet-write-number 500)
(set-row 9)
(worksheet-write-number 600)
(set-row 10)
(set-format 'bold)
(worksheet-write-formula "=SUBTOTAL(9,B7:B10)")

(add-worksheet "Collapsed rows")
(create-row-col-opt 2 1 0)
(create-row-col-opt 1 1 0)
(create-row-col-opt 0 0 1)

(set-col-option-index 0)
(worksheet-set-column 20 0 0)

(set-col-option-index 2)
(worksheet-set-row-opt $lxw-def-row-height 0)
(worksheet-set-row-opt $lxw-def-row-height 1)
(worksheet-set-row-opt $lxw-def-row-height 2)
(worksheet-set-row-opt $lxw-def-row-height 3)
(set-col-option-index 3)
(worksheet-set-row-opt $lxw-def-row-height 4)

(set-col-option-index 2)
(worksheet-set-row-opt $lxw-def-row-height 6)
(worksheet-set-row-opt $lxw-def-row-height 7)
(worksheet-set-row-opt $lxw-def-row-height 8)
(worksheet-set-row-opt $lxw-def-row-height 9)
(set-col-option-index 3)
(worksheet-set-row-opt $lxw-def-row-height 10)
(set-col-option-index 4)
(worksheet-set-row-opt $lxw-def-row-height 11)



(set-pos 0 0)
(set-format 'bold)
(worksheet-write-string "Region")
(set-format 'none)
(set-row 1)
(worksheet-write-string "North")
(set-row 2)
(worksheet-write-string "North")
(set-row 3)
(worksheet-write-string "North")
(set-row 4)
(set-format 'bold)
(worksheet-write-string "North Total")

(set-pos 0 1)
(worksheet-write-string "Sales")
(set-format 'none)
(set-row 1)
(worksheet-write-number 1000)
(set-row 2)
(worksheet-write-number 1200)
(set-row 3)
(worksheet-write-number 1200)
(set-row 4)
(set-format 'bold)
(worksheet-write-formula "=SUBTOTAL(9,B2:B5)")

(set-pos 6 0)
(worksheet-write-string "Region")
(set-format 'none)
(set-row 7)
(worksheet-write-string "South")
(set-row 8)
(worksheet-write-string "South")
(set-row 9)
(worksheet-write-string "South")
(set-row 10)
(set-format 'bold)
(worksheet-write-string "South Total")

(set-format 'none)
(set-pos 6 1)
(worksheet-write-number 400)
(set-row 7)
(worksheet-write-number 600)
(set-row 8)
(worksheet-write-number 500)
(set-row 9)
(worksheet-write-number 600)
(set-row 10)
(set-format 'bold)
(worksheet-write-formula "=SUBTOTAL(9,B7:B10)")



(close-workbook)
(exit)

