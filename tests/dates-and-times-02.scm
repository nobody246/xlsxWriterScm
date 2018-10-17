(use xlsxwriterscm)

(define number '(2013 1 23 12 30 5.123))
(define formats '("dd/mm/yy"
                  "mm/dd/yy"
                  "dd m yy"
                  "d mm yy"
                  "d mmm yy"
                  "d mmmm yy"
                  "d mmmm yyy"
                  "d mmmm yyyy"
                  "dd/mm/yy hh:mm"
                  "dd/mm/yy hh:mm:ss"
                  "dd/mm/yy hh:mm:ss.000"
                  "hh:mm"
                  "hh:mm:ss"
                  "hh:mm:ss.000"))
(define row 0)
(define col 0)
(init-formats (add1 (length formats)))
(create-workbook "dates-and-times-02.xlsx")
(add-worksheet "")
(create-format)
(set-bold)
(worksheet-write-string "Formatted Date")
(set-col (add1 col))
(worksheet-write-string "Format")
(set! row (add1 row))
(worksheet-set-column 20 0 1)
(define (x y)
  (when (not (null? y))
    (set-pos row col)
    (create-format)
    (set-cell-num-format (car y))
    (worksheet-write-datetime (car number)
                              (cadr number)
                              (caddr number)
                              (cadddr number)
                              (car (cddddr number))
                              (car (cdr (cddddr number))))
    (set-col (add1 col))
    (worksheet-write-string (car y))
    (set! row (add1 row))
    (x (cdr y))))
(x formats)  
(close-workbook)
(exit)


