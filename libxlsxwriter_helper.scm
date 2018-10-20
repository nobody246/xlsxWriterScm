(define user-formats '())
(define format-c 0)

(define ($formats format-name)
  (car (alist-ref format-name user-formats)))

(define (set-format format-name)
  (set-format-index ($formats format-name)))

(define (define-formats format-definitions)
  (let ((allowed-format-definition-functions
         '(set-bold
           set-italic
           set-superscript
           set-cell-num-format
           set-cell-num-format-index
           format-hide-formula
           format-set-unlocked
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
           set-indentation
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
               (if(memq
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


  
      
