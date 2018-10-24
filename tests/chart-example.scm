(use xlsxwriterscm)

(create-workbook "chart-example.xlsx")
(add-worksheet "")
(define (write-worksheet-data)
  (let ((d `((1 2 3)
             (2 4 6)
             (3 6 9)
             (4 8 12)
             (5 10 15)))
        (row 0))
    (for-each
     (lambda(x)
       (set-pos row 0)
       (worksheet-write-number (car x))
       (set-col 1)
       (worksheet-write-number (cadr x))
       (set-col 2)
       (worksheet-write-number (caddr x))
       (set! row (add1 row)))
     d)))
(write-worksheet-data)
(create-chart ($chart-type 'column))
(create-chart-series "" "Sheet1!$A$1:$A$5")
(create-chart-series "" "Sheet1!$B$1:$B$5")
(create-chart-series "" "Sheet1!$C$1:$C$5")
(set-pos 6 1)
(worksheet-insert-chart)
(close-workbook)
(exit)
       

    
              
