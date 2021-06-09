(use posix s)

(define project-header #<#EOF
(use data-structures easyffi)
[module 
    xlsxwriterscm
    *
  (import scheme chicken foreign data-structures)
  (foreign-declare "##include             \"xlsxwriter.h\"
                    ##include             \"libxlsxwriter_layer.h\"
                    ##include             \"nonwrapped.c\"
                    ##include             \"libxlsxwriter_layer.c\"")
  
  (define-syntax $n
    (syntax-rules ()
      (($n body)
       (foreign-value body unsigned-integer))))
  
  (define pattern
    (list (cons 'solid                    ($n LXW_PATTERN_SOLID))
          (cons 'medium-gray              ($n LXW_PATTERN_MEDIUM_GRAY))
          (cons 'light-vertical           ($n LXW_PATTERN_LIGHT_VERTICAL))
          (cons 'light-up                 ($n LXW_PATTERN_LIGHT_UP))
          (cons 'light-trellis            ($n LXW_PATTERN_LIGHT_TRELLIS))
          (cons 'light-horizontal         ($n LXW_PATTERN_LIGHT_HORIZONTAL))
          (cons 'light-grid               ($n LXW_PATTERN_LIGHT_GRID))
          (cons 'light-gray               ($n LXW_PATTERN_LIGHT_GRAY))
          (cons 'light-down               ($n LXW_PATTERN_LIGHT_DOWN))
          (cons 'gray-125                 ($n LXW_PATTERN_GRAY_125))
          (cons 'gray-0625                ($n LXW_PATTERN_GRAY_0625))
          (cons 'dark-vertical            ($n LXW_PATTERN_DARK_VERTICAL))
          (cons 'dark-up                  ($n LXW_PATTERN_DARK_UP))
          (cons 'dark-trellis             ($n LXW_PATTERN_DARK_TRELLIS))
          (cons 'dark-horizontal          ($n LXW_PATTERN_DARK_HORIZONTAL))
          (cons 'dark-grid                ($n LXW_PATTERN_DARK_GRID))
          (cons 'dark-gray                ($n LXW_PATTERN_DARK_GRAY))
          (cons 'dark-down                ($n LXW_PATTERN_DARK_DOWN))))
  
  
  
  (define align
    (list (cons 'left                     ($n LXW_ALIGN_LEFT))
          (cons 'right                    ($n LXW_ALIGN_RIGHT))
          (cons 'center                   ($n LXW_ALIGN_CENTER))
          (cons 'fill                     ($n LXW_ALIGN_FILL))
          (cons 'justify                  ($n LXW_ALIGN_JUSTIFY))
          (cons 'center-across            ($n LXW_ALIGN_CENTER_ACROSS))
          (cons 'distributed              ($n LXW_ALIGN_DISTRIBUTED))
          (cons 'vertical-top             ($n LXW_ALIGN_VERTICAL_TOP))
          (cons 'vertical-bottom          ($n LXW_ALIGN_VERTICAL_BOTTOM))
          (cons 'vertical-center          ($n LXW_ALIGN_VERTICAL_CENTER))
          (cons 'vertical-justify         ($n LXW_ALIGN_VERTICAL_JUSTIFY))
          (cons 'vertical-distributed     ($n LXW_ALIGN_VERTICAL_DISTRIBUTED))))
  
  (define border
    (list (cons 'thin                     ($n LXW_BORDER_THIN))
          (cons 'medium                   ($n LXW_BORDER_MEDIUM))
          (cons 'dashed                   ($n LXW_BORDER_DASHED))
          (cons 'dotted                   ($n LXW_BORDER_DOTTED))
          (cons 'thick                    ($n LXW_BORDER_THICK))
          (cons 'double                   ($n LXW_BORDER_DOUBLE))
          (cons 'hair                     ($n LXW_BORDER_HAIR))
          (cons 'medium-dashed            ($n LXW_BORDER_MEDIUM_DASHED))
          (cons 'dash-dot                 ($n LXW_BORDER_DASH_DOT))
          (cons 'medium-dash-dot          ($n LXW_BORDER_MEDIUM_DASH_DOT))
          (cons 'slant-dash-dot           ($n LXW_BORDER_SLANT_DASH_DOT))))
  
  (define chart-type
    (list (cons 'none                     ($n LXW_CHART_NONE))
          (cons 'area                     ($n LXW_CHART_AREA))
          (cons 'area-stacked             ($n LXW_CHART_AREA_STACKED))
          (cons 'percent-stacked          ($n LXW_CHART_AREA_STACKED_PERCENT))
          (cons 'bar                      ($n LXW_CHART_BAR))
          (cons 'bar-stacked              ($n LXW_CHART_BAR_STACKED))
          (cons 'bar-stacked-percent      ($n LXW_CHART_BAR_STACKED_PERCENT))
          (cons 'column                   ($n LXW_CHART_COLUMN))
          (cons 'column-stacked           ($n LXW_CHART_COLUMN_STACKED))
          (cons 'column-stacked-percent
		($n LXW_CHART_COLUMN_STACKED_PERCENT))
          (cons 'doughnut                 ($n LXW_CHART_DOUGHNUT))
          (cons 'line                     ($n LXW_CHART_LINE))
          (cons 'pie                      ($n LXW_CHART_PIE))
          (cons 'scatter                  ($n LXW_CHART_SCATTER))
          (cons 'scatter-straight         ($n LXW_CHART_SCATTER_STRAIGHT))
          (cons 'scatter-straight-with-markers
		($n LXW_CHART_SCATTER_STRAIGHT_WITH_MARKERS))
          (cons 'scatter-smooth           ($n LXW_CHART_SCATTER_SMOOTH))
          (cons 'scatter-smooth-with-markers
		($n LXW_CHART_SCATTER_SMOOTH_WITH_MARKERS))
          (cons 'radar                    ($n LXW_CHART_RADAR))
          (cons 'radar-with-markers       ($n LXW_CHART_RADAR_WITH_MARKERS))
          (cons 'radar-filled             ($n LXW_CHART_RADAR_FILLED))))
  
  (define legend-position
    (list (cons 'none                     ($n LXW_CHART_LEGEND_NONE))
          (cons 'right                    ($n LXW_CHART_LEGEND_RIGHT))
          (cons 'left                     ($n LXW_CHART_LEGEND_LEFT))
          (cons 'top                      ($n LXW_CHART_LEGEND_TOP))
          (cons 'bottom                   ($n LXW_CHART_LEGEND_BOTTOM))
          (cons 'overlay-right            ($n LXW_CHART_LEGEND_OVERLAY_RIGHT))
          (cons 'overlay-left             ($n LXW_CHART_LEGEND_OVERLAY_LEFT))))
  
  (define chart-line-dash-type
    (list (cons 'dash-solid               ($n LXW_CHART_LINE_DASH_SOLID))
          (cons 'dash-round-dot           ($n LXW_CHART_LINE_DASH_ROUND_DOT))
          (cons 'dash-square-dot          ($n LXW_CHART_LINE_DASH_SQUARE_DOT))
          (cons 'dash-dash                ($n LXW_CHART_LINE_DASH_DASH))
          (cons 'dash-dash-dot            ($n LXW_CHART_LINE_DASH_DASH_DOT))
          (cons 'dash-long-dash           ($n LXW_CHART_LINE_DASH_LONG_DASH))
          (cons 'dash-long-dash-dot       ($n LXW_CHART_LINE_DASH_LONG_DASH_DOT))
          (cons 'dash-long-dash-dot-dot 
		($n LXW_CHART_LINE_DASH_LONG_DASH_DOT_DOT))))
  
  (define chart-marker-type
    (list (cons 'automatic                ($n LXW_CHART_MARKER_AUTOMATIC))
          (cons 'none                     ($n LXW_CHART_MARKER_NONE))
          (cons 'square                   ($n LXW_CHART_MARKER_SQUARE))
          (cons 'diamond                  ($n LXW_CHART_MARKER_DIAMOND))
          (cons 'triangle                 ($n LXW_CHART_MARKER_TRIANGLE))
          (cons 'x                        ($n LXW_CHART_MARKER_X))
          (cons 'star                     ($n LXW_CHART_MARKER_STAR))
          (cons 'short-dash               ($n LXW_CHART_MARKER_SHORT_DASH))
          (cons 'long-dash                ($n LXW_CHART_MARKER_LONG_DASH))
          (cons 'circle                   ($n LXW_CHART_MARKER_CIRCLE))
          (cons 'plus                     ($n LXW_CHART_MARKER_PLUS))))
  
  (define chart-pattern-type
    (list (cons 'none                     ($n LXW_CHART_PATTERN_NONE ))
          (cons 'percent-5                ($n LXW_CHART_PATTERN_PERCENT_5 ))
          (cons 'percent-10               ($n LXW_CHART_PATTERN_PERCENT_10 ))
          (cons 'percent-20               ($n LXW_CHART_PATTERN_PERCENT_20 ))
          (cons 'percent-25               ($n LXW_CHART_PATTERN_PERCENT_25 ))
          (cons 'percent-30               ($n LXW_CHART_PATTERN_PERCENT_30 ))
          (cons 'percent-40               ($n LXW_CHART_PATTERN_PERCENT_40 ))
          (cons 'percent-50               ($n LXW_CHART_PATTERN_PERCENT_50 ))
          (cons 'percent-60               ($n LXW_CHART_PATTERN_PERCENT_60 ))
          (cons 'percent-70               ($n LXW_CHART_PATTERN_PERCENT_70 ))
          (cons 'percent-75               ($n LXW_CHART_PATTERN_PERCENT_75 ))
          (cons 'percent-80               ($n LXW_CHART_PATTERN_PERCENT_80 ))
          (cons 'percent-90               ($n LXW_CHART_PATTERN_PERCENT_90 ))
          (cons 'light-downward-diagonal  ($n LXW_CHART_PATTERN_LIGHT_DOWNWARD_DIAGONAL))
          (cons 'light-upward-diagonal    ($n LXW_CHART_PATTERN_LIGHT_UPWARD_DIAGONAL))
          (cons 'dark-downward-diagonal   ($n LXW_CHART_PATTERN_DARK_DOWNWARD_DIAGONAL))
          (cons 'dark-upward-diagonal     ($n LXW_CHART_PATTERN_DARK_UPWARD_DIAGONAL))
          (cons 'wide-downward-diagonal   ($n LXW_CHART_PATTERN_WIDE_DOWNWARD_DIAGONAL))
          (cons 'wide-upward-diagonal     ($n LXW_CHART_PATTERN_WIDE_UPWARD_DIAGONAL))
          (cons 'light-vertical           ($n LXW_CHART_PATTERN_LIGHT_VERTICAL))
          (cons 'light-horizontal         ($n LXW_CHART_PATTERN_LIGHT_HORIZONTAL))
          (cons 'narrow-vertical          ($n LXW_CHART_PATTERN_NARROW_VERTICAL))
          (cons 'narrow-horizontal        ($n LXW_CHART_PATTERN_NARROW_HORIZONTAL))
          (cons 'dark-vertical            ($n LXW_CHART_PATTERN_DARK_VERTICAL))
          (cons 'dark-horizontal          ($n LXW_CHART_PATTERN_DARK_HORIZONTAL))
          (cons 'dashed-downward-diagonal ($n LXW_CHART_PATTERN_DASHED_DOWNWARD_DIAGONAL))
          (cons 'dashed-upward-diagonal   ($n LXW_CHART_PATTERN_DASHED_UPWARD_DIAGONAL))
          (cons 'dashed-vertical          ($n LXW_CHART_PATTERN_DASHED_VERTICAL))
          (cons 'dashed-horizontal        ($n LXW_CHART_PATTERN_DASHED_HORIZONTAL))
          (cons 'small-confetti           ($n LXW_CHART_PATTERN_SMALL_CONFETTI))
          (cons 'large-confetti           ($n LXW_CHART_PATTERN_LARGE_CONFETTI))
          (cons 'zigzag                   ($n LXW_CHART_PATTERN_ZIGZAG))
          (cons 'wave                     ($n LXW_CHART_PATTERN_WAVE))
          (cons 'diagonal-brick           ($n LXW_CHART_PATTERN_DIAGONAL_BRICK))
          (cons 'horizontal-brick         ($n LXW_CHART_PATTERN_HORIZONTAL_BRICK))
          (cons 'weave                    ($n LXW_CHART_PATTERN_WEAVE))
          (cons 'plaid                    ($n LXW_CHART_PATTERN_PLAID))
          (cons 'divot                    ($n LXW_CHART_PATTERN_DIVOT))
          (cons 'dotted-grid              ($n LXW_CHART_PATTERN_DOTTED_GRID))
          (cons 'dotted-diamond           ($n LXW_CHART_PATTERN_DOTTED_DIAMOND))
          (cons 'shingle                  ($n LXW_CHART_PATTERN_SHINGLE))
          (cons 'trellis                  ($n LXW_CHART_PATTERN_TRELLIS))
          (cons 'sphere                   ($n LXW_CHART_PATTERN_SPHERE))
          (cons 'small-grid               ($n LXW_CHART_PATTERN_SMALL_GRID))
          (cons 'large-grid               ($n LXW_CHART_PATTERN_LARGE_GRID))
          (cons 'small-check              ($n LXW_CHART_PATTERN_SMALL_CHECK))
          (cons 'large-check              ($n LXW_CHART_PATTERN_LARGE_CHECK))
          (cons 'outlined-diamond         ($n LXW_CHART_PATTERN_OUTLINED_DIAMOND))
          (cons 'solid-diamond            ($n LXW_CHART_PATTERN_SOLID_DIAMOND))))
  
  (define chart-label-position
    (list (cons 'default                  ($n LXW_CHART_LABEL_POSITION_DEFAULT))
          (cons 'center                   ($n LXW_CHART_LABEL_POSITION_CENTER))
          (cons 'right                    ($n LXW_CHART_LABEL_POSITION_RIGHT))
          (cons 'left                     ($n LXW_CHART_LABEL_POSITION_LEFT))
          (cons 'above                    ($n LXW_CHART_LABEL_POSITION_ABOVE))
          (cons 'below                    ($n LXW_CHART_LABEL_POSITION_BELOW))
          (cons 'inside-base              ($n LXW_CHART_LABEL_POSITION_INSIDE_BASE))
          (cons 'inside-end               ($n LXW_CHART_LABEL_POSITION_INSIDE_END))
          (cons 'outside-end              ($n LXW_CHART_LABEL_POSITION_OUTSIDE_END))
          (cons 'best-fit                 ($n LXW_CHART_LABEL_POSITION_BEST_FIT))))
  
  (define chart-label-separator
    (list (cons 'comma                    ($n LXW_CHART_LABEL_SEPARATOR_COMMA))
          (cons 'semicolon                ($n LXW_CHART_LABEL_SEPARATOR_SEMICOLON))
          (cons 'period                   ($n LXW_CHART_LABEL_SEPARATOR_PERIOD))
          (cons 'newline                  ($n LXW_CHART_LABEL_SEPARATOR_NEWLINE))
          (cons 'space                    ($n LXW_CHART_LABEL_SEPARATOR_SPACE))))
  
  (define chart-axis-tick-position
    (list (cons 'on-tick                  ($n LXW_CHART_AXIS_POSITION_ON_TICK))
          (cons 'between                  ($n LXW_CHART_AXIS_POSITION_BETWEEN))))
  
  (define chart-axis-label-position
    (list (cons 'next-to                  ($n LXW_CHART_AXIS_LABEL_POSITION_NEXT_TO))
          (cons 'high                     ($n LXW_CHART_AXIS_LABEL_POSITION_HIGH))
          (cons 'low                      ($n LXW_CHART_AXIS_LABEL_POSITION_LOW))
          (cons 'none                     ($n LXW_CHART_AXIS_LABEL_POSITION_NONE))))
  
  (define chart-axis-display-unit
    (list (cons 'none                     ($n LXW_CHART_AXIS_UNITS_NONE))
          (cons 'hundreds                 ($n LXW_CHART_AXIS_UNITS_HUNDREDS))
          (cons 'thousands                ($n LXW_CHART_AXIS_UNITS_THOUSANDS))
          (cons 'ten-thousands            ($n LXW_CHART_AXIS_UNITS_TEN_THOUSANDS))
          (cons 'hundred-thousands        ($n LXW_CHART_AXIS_UNITS_HUNDRED_THOUSANDS))
          (cons 'millions                 ($n LXW_CHART_AXIS_UNITS_MILLIONS))
          (cons 'ten-millions             ($n LXW_CHART_AXIS_UNITS_TEN_MILLIONS))
          (cons 'hundred-millions         ($n LXW_CHART_AXIS_UNITS_HUNDRED_MILLIONS))
          (cons 'billions                 ($n LXW_CHART_AXIS_UNITS_BILLIONS))
          (cons 'trillions                ($n LXW_CHART_AXIS_UNITS_TRILLIONS))))
  
  (define chart-axis-tick-mark
    (list (cons 'default                  ($n LXW_CHART_AXIS_TICK_MARK_DEFAULT))
          (cons 'none                     ($n LXW_CHART_AXIS_TICK_MARK_NONE))
          (cons 'inside                   ($n LXW_CHART_AXIS_TICK_MARK_INSIDE))
          (cons 'outside                  ($n LXW_CHART_AXIS_TICK_MARK_OUTSIDE))
          (cons 'crossing                 ($n LXW_CHART_AXIS_TICK_MARK_CROSSING))))
  
  (define chart-blank
    (list (cons 'gap                      ($n LXW_CHART_BLANKS_AS_GAP))
          (cons 'zero                     ($n LXW_CHART_BLANKS_AS_ZERO))
          (cons 'connected                ($n LXW_CHART_BLANKS_AS_CONNECTED))))
  
  (define chart-error-bar-type
    (list (cons 'std-error                ($n LXW_CHART_ERROR_BAR_TYPE_STD_ERROR))
          (cons 'fixed                    ($n LXW_CHART_ERROR_BAR_TYPE_FIXED))
          (cons 'percentage               ($n LXW_CHART_ERROR_BAR_TYPE_PERCENTAGE))
          (cons 'std-dev                  ($n LXW_CHART_ERROR_BAR_TYPE_STD_DEV))))
  
  (define chart-error-bar-direction
    (list (cons 'dir-both                 ($n LXW_CHART_ERROR_BAR_DIR_BOTH))
          (cons 'dir-plus                 ($n LXW_CHART_ERROR_BAR_DIR_PLUS))
          (cons 'dir-minus                ($n LXW_CHART_ERROR_BAR_DIR_MINUS))))
  
  (define chart-error-bar-cap
    (list (cons 'end-cap                  ($n LXW_CHART_ERROR_BAR_END_CAP))
          (cons 'no-cap                   ($n LXW_CHART_ERROR_BAR_NO_CAP))))
  
  (define chart-trendline-type
    (list (cons 'linear                   ($n LXW_CHART_TRENDLINE_TYPE_LINEAR))
          (cons 'log                      ($n LXW_CHART_TRENDLINE_TYPE_LOG))
          (cons 'poly                     ($n LXW_CHART_TRENDLINE_TYPE_POLY))
          (cons 'power                    ($n LXW_CHART_TRENDLINE_TYPE_POWER))
          (cons 'exp                      ($n LXW_CHART_TRENDLINE_TYPE_EXP))
          (cons 'average                  ($n LXW_CHART_TRENDLINE_TYPE_AVERAGE))))

  (define validation-criteria
    (list (cons 'between                  ($n LXW_VALIDATION_CRITERIA_BETWEEN ))
          (cons 'not-between              ($n LXW_VALIDATION_CRITERIA_NOT_BETWEEN ))
          (cons 'equal                    ($n LXW_VALIDATION_CRITERIA_EQUAL_TO ))
          (cons 'not-equal                ($n LXW_VALIDATION_CRITERIA_NOT_EQUAL_TO ))
          (cons 'greater-than             ($n LXW_VALIDATION_CRITERIA_GREATER_THAN))
          (cons 'less-than                ($n LXW_VALIDATION_CRITERIA_LESS_THAN))
          (cons 'greater-than-or-equal    ($n LXW_VALIDATION_CRITERIA_GREATER_THAN_OR_EQUAL_TO))
          (cons 'less-than-or-equal       ($n LXW_VALIDATION_CRITERIA_LESS_THAN_OR_EQUAL_TO ))))

  (define validation-error-type
    (list (cons 'stop                     ($n LXW_VALIDATION_ERROR_TYPE_STOP ))
          (cons 'warning                  ($n LXW_VALIDATION_ERROR_TYPE_WARNING ))
          (cons 'information              ($n LXW_VALIDATION_ERROR_TYPE_INFORMATION ))))

  (define validation-type
    (list (cons 'integer                  ($n LXW_VALIDATION_TYPE_INTEGER  ))
          (cons 'integer-formula          ($n LXW_VALIDATION_TYPE_INTEGER_FORMULA ))
          (cons 'decimal                  ($n LXW_VALIDATION_TYPE_DECIMAL ))
          (cons 'decimal-formula          ($n LXW_VALIDATION_TYPE_DECIMAL_FORMULA ))
          (cons 'list                     ($n LXW_VALIDATION_TYPE_LIST))
          (cons 'list-formula             ($n LXW_VALIDATION_TYPE_LIST_FORMULA))
          (cons 'date                     ($n LXW_VALIDATION_TYPE_DATE ))
          (cons 'date-formula             ($n LXW_VALIDATION_TYPE_DATE_FORMULA  ))
          (cons 'time                     ($n LXW_VALIDATION_TYPE_TIME  ))
          (cons 'time-formula             ($n LXW_VALIDATION_TYPE_TIME_FORMULA ))
          (cons 'length                   ($n LXW_VALIDATION_TYPE_LENGTH ))
          (cons 'length-formula           ($n LXW_VALIDATION_TYPE_LENGTH_FORMULA))
          (cons 'custom-formula           ($n LXW_VALIDATION_TYPE_CUSTOM_FORMULA))
          (cons 'any                      ($n LXW_VALIDATION_TYPE_ANY))))  
  
  (define lxw-boolean
    (list (cons 'true    ($n LXW_TRUE))
          (cons 'false   ($n LXW_FALSE))))

  (define axis-type
    (list (cons 'x   ($n LXW_CHART_ERROR_BAR_AXIS_X))
          (cons 'y   ($n LXW_CHART_ERROR_BAR_AXIS_Y))))
  
  (define-syntax $pattern
    (syntax-rules ()
      (($pattern body)
       (alist-ref body pattern))))
  
  (define-syntax $align
    (syntax-rules ()
      (($align body)
       (alist-ref body align))))
  
  (define-syntax $border
    (syntax-rules ()
      (($border body)
       (alist-ref body border))))
  
  (define-syntax $chart-type
    (syntax-rules ()
      (($chart-type body)
       (alist-ref body chart-type))))
  
  (define-syntax $legend-position
    (syntax-rules ()
      (($legend-position body)
       (alist-ref body legend-position))))
  
  (define-syntax $chart-line-dash-type
    (syntax-rules ()
      (($chart-line-dash-type body)
       (alist-ref body chart-line-dash-type))))
  
  (define-syntax $chart-marker-type
    (syntax-rules ()
      (($chart-marker-type body)
       (alist-ref body chart-marker-type))))
  
  (define-syntax $chart-pattern-type
    (syntax-rules ()
      (($chart-pattern-type body)
       (alist-ref body chart-pattern-type))))
  
  (define-syntax $chart-label-position
    (syntax-rules ()
      (($chart-label-position body)
       (alist-ref body chart-label-position))))
  
  (define-syntax $chart-label-separator
    (syntax-rules ()
      (($chart-label-separator body)
       (alist-ref body chart-label-separator))))
  
  (define-syntax $chart-axis-tick-position
    (syntax-rules ()
      (($chart-axis-tick-position body)
       (alist-ref body chart-axis-tick-position))))
  
  (define-syntax $chart-axis-display-unit
    (syntax-rules ()
      (($chart-axis-display-unit body)
       (alist-ref body chart-axis-display-unit))))
  
  (define-syntax $chart-axis-tick-mark
    (syntax-rules ()
      (($chart-axis-tick-mark)
       (alist-ref body chart-axis-tick-mark))))
  
  (define-syntax $chart-blank
    (syntax-rules ()
      (($chart-blank body)
       (alist-ref body chart-blank))))
  
  (define-syntax $chart-error-bar-type
    (syntax-rules ()
      (($chart-error-bar-type body)
       (alist-ref body chart-error-bar-type))))
  
  (define-syntax $chart-error-bar-direction
    (syntax-rules ()
      (($chart-error-bar-direction body)
       (alist-ref body chart-error-bar-direction))))
  
  (define-syntax $chart-error-bar-cap
    (syntax-rules ()
      (($chart-error-bar-cap body)
       (alist-ref body chart-error-bar-cap))))
  
  (define-syntax $chart-trendline-type
    (syntax-rules ()
      (($chart-trendline-type body)
       (alist-ref body chart-trendline-type))))

  (define-syntax $lxw-bool
    (syntax-rules ()
      (($lxw-bool body)
       (alist-ref body lxw-boolean))))

  (define-syntax $axis
    (syntax-rules ()
      (($axis-type body)
       (alist-ref body axis-type))))

  (define-syntax $validation-criteria
    (syntax-rules ()
      (($validation-criteria body)
       (alist-ref body validation-criteria))))

  (define-syntax $validation-error-type
    (syntax-rules ()
      (($validation-error-type body)
       (alist-ref body validation-error-type))))

  (define-syntax $validation-type
    (syntax-rules ()
      (($validation-type body)
       (alist-ref body validation-type))))

  (define $lxw-def-row-height ($n LXW_DEF_ROW_HEIGHT))
  
  ~A

  ~A
  
  ]
EOF
)

(define wrap-process 
  (process "chicken-wrap libxlsxwriter_layer.c -to-stdout"))

;replace c notation with scheme notation
(define (read-lines-and-format y)
  (let ((l (read-lines y))
        (in-scheme-code #t))
    (s-join
     "\n"
     (map
      (lambda(x)
        (let ((z (string->list x))
              (new-list '()))
          (for-each
           (lambda(d)
             (when (eq? d #\")
               (set! in-scheme-code (not in-scheme-code)))
             (if (and
                  in-scheme-code
                  (char-upper-case? d))
                 (set! new-list
                   (append new-list
                           '(#\-)
                           `(,(char-downcase d))))
                 (set! new-list
                   (append new-list
                           `(,d)))))
           z)
          (list->string new-list)))
      l))))

(with-output-to-file "xlsxwriterscm.scm"
  (lambda()
    (let* ((g (read-lines-and-format wrap-process))
           (g (s-replace "\"uint8_t\""  "unsigned-byte" g))
           (g (s-replace "\"int8_t\""   "byte" g))
           (g (s-replace "\"uint\""     "unsigned-integer" g))
           (g (s-replace "\"uint16_t\"" "unsigned-short" g))
           (g (s-replace "\"uint32_t\"" "unsigned-integer32" g))
           (g (s-replace "\"int32_t\""  "integer32" g))
           (g (s-replace "\"ptrdiff_t\""  "integer" g))
           (a (s-join "\n" (read-lines "libxlsxwriter_helper.scm"))))
      (printf project-header
              g
              a))))
(exit)

