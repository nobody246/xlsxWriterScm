;xlsxWriterScm (c) 2019-2025 Alexander Semotan
; Released under the BSD 2 Software License 
(define user-formats '())
(define data-validations '())
(define format-c 0)
(define data-validation-c 0)

(define ($formats format-name)
  (car (alist-ref format-name user-formats)))

(define ($data-validations data-validation-name)
  (car (alist-ref data-validation-name data-validations)))

(define (set-format format-name)
  (set-format-index ($formats format-name)))

(define (set-data-validation data-validation-name)
  (set-data-validation-index ($data-validations
                              data-validation-name)))

(define (define-formats format-definitions)
  (let ((allowed-format-definition-functions
         '(set-bold
           set-italic
           set-superscript
           set-cell-num-format
           set-cell-num-format-index
           format-hide-formula
           format-set-unlocked
           format-set-hidden
           format-set-bold
           format-set-italic
           set-underline-single
           set-underline-double
           set-underline-single-acct
           set-underline-double-acct
           set-strikeout
           set-superscript
           set-subscript
           set-font-color
           set-font-name
           set-rotation
           set-indent
           set-bold
           set-italic
           set-shrink
           set-pattern
           set-align
           set-border
           set-border-bottom
           set-border-rop
           set-border-reft
           set-border-right
           set-bg-color
           set-fg-color
           set-border-color
           set-border-bottom-color
           set-border-top-color
           set-border-left-color
           set-border-right-color
           turn-on-text-wrap)))
  (when (not (null? format-definitions))
    (init-formats (add1 (length format-definitions)))
    (define (d x)
      (when (not (null? x))
        (let ((y (cdr x)))
          (set! user-formats
            (append user-formats
                    `((,(car x) ,format-c))))
          (set! format-c (add1 format-c))
          (when (not (null? y))
            (for-each
             (lambda(a)
               (if (memq
                   (car a)
                   allowed-format-definition-functions)
                   (eval a)
                   (print (car a)
                          " is not a valid format function.")))
             (car y)))
          (d (cddr x)))))
    (define (e x)
      (when (not (null? x))
        (create-format)
        (d (car x))
        (e (cdr x))))
    (e format-definitions))))

(define (worksheet-write-rich-string rs-segments)
  (when (not (null? rs-segments))
    (init-rich-string (add1 (length rs-segments)))
    (define (w x)
        (when (not (null? x))
          (let*
              ((x (car x))
               (z (car x))
               (y (cadr x)))
            (when (not (null? z))
              (set-format-index z))
            (create-rich-string-fragment y))
          (w (cdr x)))))
    (w rs-segments)
    (worksheet-write-rich-string-fragments))

(define (worksheet-protect-settings #!optional
                                    #!key
                                    (password "")
                                    (no-select-locked-cells 0)
                                    (no-select-unlocked-cells 0)
                                    (format-cells 0)
                                    (format-columns 0)
                                    (format-rows 0)
                                    (insert-columns 0)
                                    (insert-rows 0)
                                    (insert-hyperlinks 0)
                                    (delete-columns 0)
                                    (delete-rows 0)
                                    (sort 0)
                                    (autofilter 0)
                                    (pivot-tables 0)
                                    (scenarios 0)
                                    (objects 0)
                                    (no-content 0)
                                    (no-objects 0))
  (worksheet-protect password
                     no-select-locked-cells
                     no-select-unlocked-cells
                     format-cells
                     format-columns
                     format-rows
                     insert-columns
                     insert-rows
                     insert-hyperlinks
                     delete-columns
                     delete-rows
                     sort
                     autofilter
                     pivot-tables
                     scenarios
                     objects
                     no-content
                     no-objects))

(define (alist-ref-or-zero val list)
  (let ((x (alist-ref val list)))
    (if x (car x) 0)))

   

(define (create-pie-chart #!optional
                          #!key
                          (chart-title #f)
                          (chart-series-title #f)
                          (chart-series-name-range "")
                          (chart-series-value-range "")
                          (chart-point-definitions '())
                          (rotation 0)
                          (chart-type 'pie))
  (create-chart ($chart-type chart-type))
  (chart-set-rotation rotation)
  (create-chart-series chart-series-name-range chart-series-value-range)
  (when chart-title
    (chart-title-set-name chart-title))
  (when chart-series-title
    (chart-series-set-name chart-series-title))
  (let ((c (length chart-point-definitions)))
    (when (> c 0)
      (init-chart-points (length chart-point-definitions))
      (init-chart-fills (length chart-point-definitions))
      (init-chart-lines (length chart-point-definitions))
      (init-chart-patterns (length chart-point-definitions))
      (for-each
       (lambda(x)
         (let* ((fill-color         (alist-ref-or-zero 'fill-color       x))
                (fill-none          (alist-ref-or-zero 'fill-none        x))
                (fill-transparency  (alist-ref-or-zero 'fill-type        x))
                (line-color         (alist-ref-or-zero 'line-color       x))
                (line-none          (alist-ref-or-zero 'line-none        x))
                (line-width         (alist-ref-or-zero 'line-width       x))
                (line-dash-type     (alist-ref-or-zero 'line-dash-type   x))
                (pattern-fg-color   (alist-ref-or-zero 'pattern-fg-color x))
                (pattern-bg-color   (alist-ref-or-zero 'pattern-bg-color x))
                (pattern-type       (alist-ref-or-zero 'pattern-type     x)))
           (create-chart-fill fill-color fill-none fill-transparency)
           (create-chart-line line-color line-none line-width line-dash-type)
           (when (> pattern-type 0)
             (create-chart-pattern
              pattern-fg-color
              pattern-bg-color
              pattern-type))
           (create-chart-point)))
       chart-point-definitions)
      (chart-series-set-points))))
      
(define (create-doughnut-chart #!optional
                               #!key
                               (chart-title #f)
                               (chart-series-title #f)
                               (chart-series-name-range "")
                               (chart-series-value-range "")
                               (chart-point-definitions '())
                               (rotation 0))
  (create-pie-chart chart-title: chart-title
                    chart-series-title: chart-series-title
                    chart-series-name-range: chart-series-name-range
                    chart-series-value-range: chart-series-value-range
                    chart-point-definitions: chart-point-definitions
                    rotation: rotation
                    chart-type: 'doughnut))


        
        
