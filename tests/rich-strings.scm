(use xlsxwriterscm)

(create-workbook "rich-strings.xlsx")
(define-formats '((bold ((set-bold)))
                  (italic ((set-italic)))
                  (red ((set-font-color #xFF0000)))
                  (blue ((set-font-color #x0000FF)
                         (set-bold)
                         (set-italic)))
                  (center ((set-align ($align 'center))))
                  (superscript ((set-superscript)))
                  (none ())))
                  
;make first column wider for clarity
(worksheet-set-column 30 0 0)

(add-worksheet "")
(set-pos 0 0)
(worksheet-write-rich-string `((,($formats 'none) "This is ")
                               (,($formats 'bold) "bold ")
                               (,($formats 'none) "and this is ")
                               (,($formats 'italic)  "italic.")))
(set-pos 2 0)
(worksheet-write-rich-string `((,($formats 'none) "This is ")
                               (,($formats 'red) "red ")
                               (,($formats 'none) "and this is ")
                               (,($formats 'blue) "blue.")))

(set-pos 4 0)
(set-format 'center)
(worksheet-write-rich-string `((,($formats 'none) "Some ")
                               (,($formats 'bold) "bold text")
                               (,($formats 'none) " centered")))


(set-pos 6 0)
(set-format 'center)
(worksheet-write-rich-string `((,($formats 'italic) "j =k")
                               (,($formats 'superscript) "(n-1)")))

(close-workbook)


(exit)
