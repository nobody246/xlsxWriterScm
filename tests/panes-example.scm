(use xlsxwriterscm)


(create-workbook "panes-example.xlsx")
(add-worksheet "")

(define-formats '((header ((set-align ($align 'center))
                           (set-align ($align 'vertical-center))
                           (set-fg-color #xD7E4BC)
                           (set-bold)
                           (set-border ($border 'thin))))
                  (center ((set-align ($align 'center))))
                  (none ())))

;freeze top row
(worksheet-freeze-panes 1 0)
(worksheet-set-column 16 0 8)
(worksheet-set-row 20 0)
(worksheet-set-selection 4 3 4 3)

(set-format 'header)
(define (x y #!optional #!key (maximum 9))
  (if (< y maximum)
      (begin
        (set-col y)
        (worksheet-write-string "Scroll Down")
        (x (add1 y) maximum: maximum))
      (set-format 'none)))
(x 0)

(define (o x y maximum)
  (when (< x maximum)
    (begin
      (set-col x)
      (worksheet-write-number y)
      (o (add1 x) y maximum))))

(set-format 'center)
(define (z y)
  (if (< y 100)
      (begin
        (set-row y)
        (o 0 y 9)
        (z (add1 y)))
      (set-format 'none)))
(z 1)


;freeze left column
(add-worksheet "")
(worksheet-freeze-panes 0 1)
(worksheet-set-column 16 0 0)

(define (q x)
  (if (< x 50)
      (begin
        (set-pos x 0)
        (set-format 'header)
        (worksheet-write-string "scroll right")
        (set-format 'center)
        (o 1 x 26)
        (q (add1 x)))
      (set-format 'none)))
(q 0)


;freeze left column and top row
(add-worksheet "")
(worksheet-freeze-panes 1 1)
(worksheet-set-column 16 0 0)
(q 0)
(set-row 0)
(set-format 'header)
(x 0 maximum: 26)


;split pane on left column and top row
(add-worksheet "")
(worksheet-split-panes 15 8.43)
(q 0)
(set-row 0)
(set-format 'header)
(x 0 maximum: 26)
          
  
(close-workbook)
(exit)
