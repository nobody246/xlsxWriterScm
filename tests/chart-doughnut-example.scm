(use xlsxwriterscm)

(create-workbook "chart-doughnut-example.xlsx")
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
(worksheet-write-number 15)


(create-doughnut-chart chart-title: "Pie Chart with user define colors"
                       chart-series-title: "Pie sales data"
                       chart-series-name-range:  "=Sheet1!$A$2:$A$4"
                       chart-series-value-range: "=Sheet1!$B$2:$B$4")
(set-pos 7 0)
(worksheet-insert-chart)
(set-row 0)
(add-worksheet "")
(create-doughnut-chart chart-title: "Foo"
                       chart-series-title: "Bar"
                       chart-series-name-range:  "=Sheet1!$A$2:$A$4"
                       chart-series-value-range: "=Sheet1!$B$2:$B$4"
                       chart-point-definitions:
                       '(((fill-color #xFF0000)
                          (line-color #xFF00FF))
                         ((fill-color #x0000FF))))
(worksheet-insert-chart)
(add-worksheet "")
(create-doughnut-chart chart-title: "Bang"
                       chart-series-title: "Bat"
                       chart-series-name-range:  "=Sheet1!$A$2:$A$4"
                       chart-series-value-range: "=Sheet1!$B$2:$B$4"
                       rotation: 70
                       style: 26
                       chart-point-definitions:
                       `(((fill-color #x00FF00))
                         ((fill-color #xFF00FF))
                         ((fill-color #xFFABAB))))
(worksheet-insert-chart)

(add-worksheet "")
(create-doughnut-chart chart-title: "Bang"
                       chart-series-title: "Bat"
                       chart-series-name-range:  "=Sheet1!$A$2:$A$4"
                       chart-series-value-range: "=Sheet1!$B$2:$B$4"
                       style: 26)
(worksheet-insert-chart)

(close-workbook)
(exit)
