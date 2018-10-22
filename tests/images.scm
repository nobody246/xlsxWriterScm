(use xlsxwriterscm)
(create-workbook "images.xlsx")
(add-worksheet "")
(set-pos 0 0)
(worksheet-write-string "Insert an image in a cell:")
(set-col 1) 
(worksheet-insert-image "i.jpg")

(set-pos 11 0)
(worksheet-write-string "Insert an offset image:")
(set-col 1)
(worksheet-insert-image-opt 15 10 1 1 "i.jpg")

(set-pos 20 0)
(worksheet-write-string "Insert a scaled image:")
(set-col 1)
(worksheet-insert-image-opt 0 0 .5 .5 "i.jpg")

(close-workbook)
(exit)



