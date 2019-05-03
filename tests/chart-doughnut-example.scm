;not working 100%..
(use xlsxwriterscm)

(define (write-worksheet-data)
  (add-worksheet "")
  (set-pos 0 0)
  (worksheet-write-string "Category")
  (set-row 1)
  (worksheet-write-string "Glazed")
  (set-row 2)
  (worksheet-write-string "Chocolate")
  (set-row 3)
  (worksheet-write-string "Cream")
  (set-pos 0 1)
  (worksheet-write-string "Values")
  (set-row 1)
  (worksheet-write-number 50)
  (set-row 2)
  (worksheet-write-number 35)
  (set-row 3)
  (worksheet-write-number 15))

(create-workbook "chart-doughnut-example.xlsx")
(format-set-bold)
(write-worksheet-data)

(create-doughnut-chart chart-title: "Popular Doughnut Types"
                       chart-series-title: "Doughnut sales data"
                       chart-series-name-range:  "=Sheet1!$A$2:$A$4"
                       chart-series-value-range: "=Sheet1!$B$2:$B$4")
(set-pos 1 4)
(worksheet-insert-chart)

(create-doughnut-chart chart-title: "Doughnut Chart with user defined colors"
                       chart-series-title: "Doughnut sales data"
                       chart-series-name-range:  "=Sheet1!$A$2:$A$4"
                       chart-series-value-range: "=Sheet1!$B$2:$B$4"
                       chart-point-definitions:
                       '(((fill-color #xFA58D0))
                         ((fill-color #x61210B))
                         ((fill-color #xF5F6CE))))
(set-pos 17 4)
(worksheet-insert-chart)

(create-doughnut-chart chart-title: "Doughnut Chart with segment rotation"
                       chart-series-title: "Doughnut sales data"
                       chart-series-name-range:  "=Sheet1!$A$2:$A$4"
                       chart-series-value-range: "=Sheet1!$B$2:$B$4"
                       rotation: 90)
(set-pos 33 4)
(worksheet-insert-chart)

(create-doughnut-chart chart-title: "Doughnut Chart with options applied."
                       chart-series-title: "Doughnut sales data"
                       chart-series-name-range:  "=Sheet1!$A$2:$A$4"
                       chart-series-value-range: "=Sheet1!$B$2:$B$4")
(chart-set-hole-size 33)
(set-pos 49 4)
(worksheet-insert-chart)

(close-workbook)
(exit)
