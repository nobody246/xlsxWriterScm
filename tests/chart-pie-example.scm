(use xlsxwriterscm)

(create-workbook "chart-pie-example.xlsx")
(add-worksheet "")
(set-pos 0 0)
(worksheet-write-string "Pass")
(set-col 1)
(worksheet-write-string "Fail")
(set-pos 1 0)
(worksheet-write-number 56)
(set-col 1)
(worksheet-write-number 10)
(set-pos 2 0)
(worksheet-write-number 11)
(set-col 1)
(worksheet-write-number 11)
(set-pos 3 0)
(worksheet-write-number 6)
(set-col 1)
(worksheet-write-number 6)
(create-pie-chart chart-title: "Pie Chart with user define colors"
                  chart-series-title: "Pie sales data"
                  chart-series-name-range:  "=Sheet1!$A$2:$A$4"
                  chart-series-value-range: "=Sheet1!$B$2:$B$4"
                  chart-point-definitions:
                  '(((fill-color #x5ABA10))
                    ((fill-color #xFF0000))))
(set-pos 7 0)
(worksheet-insert-chart)
(set-row 0)
(add-worksheet "")
(create-pie-chart chart-title: "Foo"
                  chart-series-title: "Bar"
                  chart-series-name-range:  "=Sheet1!$A$2:$A$4"
                  chart-series-value-range: "=Sheet1!$B$2:$B$4"
                  chart-point-definitions:
                  '(((fill-color #xFF0000)
                     (line-color #xFF00FF))
                    ((fill-color #x0000FF))))
(worksheet-insert-chart)
(add-worksheet "")
(create-pie-chart chart-title: "Bang"
                  chart-series-title: "Bat"
                  chart-series-name-range:  "=Sheet1!$A$2:$A$4"
                  chart-series-value-range: "=Sheet1!$B$2:$B$4"
                  rotation: 70
                  chart-point-definitions:
                  `(((fill-color #x00FF00)
                     (line-none ,($lxw-bool 'false))
                     (line-color #x000002)
                     (line-width 2))
                    ((fill-color #xFF00FF))
                    ((fill-color #xFFABAB))))
(worksheet-insert-chart)
(close-workbook)
(exit)
