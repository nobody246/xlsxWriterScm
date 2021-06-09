(use data-structures easyffi)
[module 
    xlsxwriterscm
    *
  (import scheme chicken foreign data-structures)
  (foreign-declare "#include             \"xlsxwriter.h\"
                    #include             \"libxlsxwriter_layer.h\"
                    #include             \"nonwrapped.c\"
                    #include             \"libxlsxwriter_layer.c\"")
  
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
  
  
;;; generated by chicken-wrap from libxlsxwriter_layer.c

(begin
  (begin
    (define create-workbook (foreign-lambda void "createWorkbook" c-string)))
  (begin
    (define create-workbook-constant-memory
      (foreign-lambda void "createWorkbookConstantMemory" c-string)))
  (begin
    (define resize-data-validations
      (foreign-lambda void "resizeDataValidations")))
  (begin
    (define data-validations-bounds-check
      (foreign-lambda void "dataValidationsBoundsCheck")))
  (begin (define formats-cleanup (foreign-lambda void "formatsCleanup")))
  (begin (define row-number-cleanup (foreign-lambda void "rowNumberCleanup")))
  (begin (define col-number-cleanup (foreign-lambda void "colNumberCleanup")))
  (begin
    (define data-validations-cleanup
      (foreign-lambda void "dataValidationsCleanup")))
  (begin
    (define chart-points-cleanup (foreign-lambda void "chartPointsCleanup")))
  (begin (define chart-fills-cleanup (foreign-lambda void "chartFillsCleanup")))
  (begin
    (define chart-patterns-cleanup (foreign-lambda void "chartPatternsCleanup")))
  (begin (define chart-lines-cleanup (foreign-lambda void "chartLinesCleanup")))
  (begin (define series-cleanup (foreign-lambda void "seriesCleanup")))
  (begin (define charts-cleanup (foreign-lambda void "chartsCleanup")))
  (begin
    (define series-error-bars-cleanup
      (foreign-lambda void "seriesErrorBarsCleanup")))
  (begin (define chart-fonts-cleanup (foreign-lambda void "chartFontsCleanup")))
  (begin (define init-formats (foreign-lambda void "initFormats" integer)))
  (begin
    (define init-row-numbers (foreign-lambda void "initRowNumbers" integer)))
  (begin
    (define init-col-numbers (foreign-lambda void "initColNumbers" integer)))
  (begin
    (define init-data-validations
      (foreign-lambda void "initDataValidations" integer)))
  (begin
    (define init-data-validation-lists
      (foreign-lambda void "initDataValidationLists")))
  (begin
    (define init-chart-fills (foreign-lambda void "initChartFills" integer)))
  (begin
    (define init-chart-points
      (foreign-lambda void "initChartPoints" integer)))
  (begin
    (define init-chart-lines (foreign-lambda void "initChartLines" integer)))
  (begin
    (define init-chart-patterns
      (foreign-lambda void "initChartPatterns" integer)))
  (begin (define init-series (foreign-lambda void "initSeries" integer)))
  (begin (define init-charts (foreign-lambda void "initCharts" integer)))
  (begin
    (define init-chart-fonts (foreign-lambda void "initChartFonts" integer)))
  (begin (define col-option-cleanup (foreign-lambda void "colOptionCleanup")))
  (begin
    (define init-col-options (foreign-lambda void "initColOptions" integer)))
  (begin
    (define create-data-validation-list-entry
      (foreign-lambda void "createDataValidationListEntry" c-string)))
  (begin
    (define set-validation-validate
      (foreign-lambda void "setValidationValidate" unsigned-byte)))
  (begin
    (define set-validation-criteria
      (foreign-lambda void "setValidationCriteria" unsigned-byte)))
  (begin
    (define set-validation-ignore-blank
      (foreign-lambda void "setValidationIgnoreBlank" unsigned-byte)))
  (begin
    (define set-validation-show-input
      (foreign-lambda void "setValidationShowInput" unsigned-byte)))
  (begin
    (define set-validation-show-error
      (foreign-lambda void "setValidationShowError" unsigned-byte)))
  (begin
    (define set-validation-error-type
      (foreign-lambda void "setValidationErrorType" unsigned-byte)))
  (begin
    (define set-validation-dropdown
      (foreign-lambda void "setValidationDropdown" unsigned-byte)))
  (begin
    (define set-validation-value-number
      (foreign-lambda void "setValidationValueNumber" double)))
  (begin
    (define set-validation-value-formula
      (foreign-lambda void "setValidationValueFormula" c-string)))
  (begin
    (define set-validation-value-date-time
      (foreign-lambda
        void
        "setValidationValueDateTime"
        integer
        integer
        integer
        integer
        integer
        double)))
  (begin
    (define set-validation-min-number
      (foreign-lambda void "setValidationMinNumber" double)))
  (begin
    (define set-validation-min-formula
      (foreign-lambda void "setValidationMinFormula" c-string)))
  (begin
    (define set-validation-min-date-time
      (foreign-lambda
        void
        "setValidationMinDateTime"
        integer
        integer
        integer
        integer
        integer
        double)))
  (begin
    (define set-validation-max-number
      (foreign-lambda void "setValidationMaxNumber" double)))
  (begin
    (define set-validation-max-formula
      (foreign-lambda void "setValidationMaxFormula" c-string)))
  (begin
    (define set-validation-max-date-time
      (foreign-lambda
        void
        "setValidationMaxDateTime"
        integer
        integer
        integer
        integer
        integer
        double)))
  (begin
    (define set-validation-input-title
      (foreign-lambda void "setValidationInputTitle" c-string)))
  (begin
    (define set-validation-input-message
      (foreign-lambda void "setValidationInputMessage" c-string)))
  (begin
    (define set-validation-error-title
      (foreign-lambda void "setValidationErrorTitle" c-string)))
  (begin
    (define set-validation-error-message
      (foreign-lambda void "setValidationErrorMessage" c-string)))
  (begin
    (define set-validation-value-list
      (foreign-lambda void "setValidationValueList")))
  (begin
    (define create-data-validation (foreign-lambda void "createDataValidation")))
  (begin (define create-chart-point (foreign-lambda void "createChartPoint")))
  (begin
    (define create-chart-fill
      (foreign-lambda void "createChartFill" integer32 unsigned-byte unsigned-byte)))
  (begin
    (define create-chart-line
      (foreign-lambda
        void
        "createChartLine"
        integer32
        unsigned-byte
        float
        unsigned-byte)))
  (begin
    (define create-chart-pattern
      (foreign-lambda
        void
        "createChartPattern"
        integer32
        integer32
        unsigned-byte)))
  (begin
    (define create-chart-font
      (foreign-lambda
        void
        "createChartFont"
        c-string
        double
        unsigned-byte
        unsigned-byte
        unsigned-byte
        unsigned-byte
        unsigned-byte
        unsigned-byte
        unsigned-byte
        byte)))
  (begin
    (define add-to-row-number-list
      (foreign-lambda void "addToRowNumberList" unsigned-integer32)))
  (begin
    (define add-to-col-number-list
      (foreign-lambda void "addToColNumberList" unsigned-short)))
  (begin
    (define set-data-validation-index
      (foreign-lambda void "setDataValidationIndex" integer)))
  (begin
    (define set-chart-point-index
      (foreign-lambda void "setChartPointIndex" integer)))
  (begin
    (define set-chart-fill-index
      (foreign-lambda void "setChartFillIndex" integer)))
  (begin
    (define set-chart-line-index
      (foreign-lambda void "setChartLineIndex" integer)))
  (begin
    (define set-chart-pattern-index
      (foreign-lambda void "setChartPatternIndex" integer)))
  (begin
    (define set-format-index (foreign-lambda void "setFormatIndex" integer)))
  (begin
    (define set-series-index (foreign-lambda void "setSeriesIndex" integer)))
  (begin
    (define set-chart-index (foreign-lambda void "setChartIndex" integer)))
  (begin
    (define set-series-error-bars-index
      (foreign-lambda void "setSeriesErrorBarsIndex" integer)))
  (begin
    (define set-chart-font-index
      (foreign-lambda void "setChartFontIndex" integer)))
  (begin
    (define set-col-option-index
      (foreign-lambda void "setColOptionIndex" integer)))
  (begin (define add-worksheet (foreign-lambda void "addWorksheet" c-string)))
  (begin
    (define worksheet-write-string
      (foreign-lambda void "worksheetWriteString" c-string)))
  (begin
    (define worksheet-write-number
      (foreign-lambda void "worksheetWriteNumber" float)))
  (begin
    (define worksheet-write-formula
      (foreign-lambda void "worksheetWriteFormula" c-string)))
  (begin
    (define worksheet-write-array-formula
      (foreign-lambda
        void
        "worksheetWriteArrayFormula"
        c-string
        unsigned-integer32
        unsigned-short
        unsigned-integer32
        unsigned-short)))
  (begin
    (define worksheet-write-datetime
      (foreign-lambda
        void
        "worksheetWriteDatetime"
        integer
        integer
        integer
        integer
        integer
        double)))
  (begin
    (define worksheet-write-url
      (foreign-lambda void "worksheetWriteUrl" c-string)))
  (begin
    (define worksheet-write-boolean
      (foreign-lambda void "worksheetWriteBoolean" short)))
  (begin
    (define worksheet-write-blank (foreign-lambda void "worksheetWriteBlank")))
  (begin
    (define worksheet-set-row
      (foreign-lambda void "worksheetSetRow" double unsigned-integer32)))
  (begin
    (define worksheet-set-row-opt
      (foreign-lambda void "worksheetSetRowOpt" double unsigned-integer32)))
  (begin
    (define worksheet-set-column
      (foreign-lambda void "worksheetSetColumn" double unsigned-short unsigned-short)))
  (begin
    (define worksheet-set-column-opt
      (foreign-lambda
        void
        "worksheetSetColumnOpt"
        double
        unsigned-short
        unsigned-short)))
  (begin
    (define worksheet-insert-image
      (foreign-lambda void "worksheetInsertImage" c-string)))
  (begin
    (define worksheet-insert-image-opt
      (foreign-lambda
        void
        "worksheetInsertImageOpt"
        integer
        integer
        double
        double
        c-string)))
  (begin
    (define worksheet-insert-chart (foreign-lambda void "worksheetInsertChart")))
  (begin
    (define worksheet-insert-chart-opt
      (foreign-lambda
        void
        "worksheetInsertChartOpt"
        integer
        integer
        integer
        integer)))
  (begin
    (define worksheet-merge-range
      (foreign-lambda
        void
        "worksheetMergeRange"
        c-string
        unsigned-integer32
        unsigned-short
        unsigned-integer32
        unsigned-short)))
  (begin
    (define worksheet-auto-filter
      (foreign-lambda
        void
        "worksheetAutoFilter"
        unsigned-integer32
        unsigned-short
        unsigned-integer32
        unsigned-short)))
  (begin
    (define worksheet-data-validation-cell
      (foreign-lambda void "worksheetDataValidationCell")))
  (begin
    (define worksheet-data-validation-range
      (foreign-lambda
        void
        "worksheetDataValidationRange"
        unsigned-integer32
        unsigned-short
        unsigned-integer32
        unsigned-short)))
  (begin (define worksheet-activate (foreign-lambda void "worksheetActivate")))
  (begin (define worksheet-select (foreign-lambda void "worksheetSelect")))
  (begin (define worksheet-hide (foreign-lambda void "worksheetHide")))
  (begin
    (define worksheet-set-first-sheet
      (foreign-lambda void "worksheetSetFirstSheet")))
  (begin
    (define worksheet-freeze-panes
      (foreign-lambda void "worksheetFreezePanes" unsigned-integer32 unsigned-short)))
  (begin
    (define worksheet-split-panes
      (foreign-lambda void "worksheetSplitPanes" double double)))
  (begin
    (define worksheet-set-selection
      (foreign-lambda
        void
        "worksheetSetSelection"
        unsigned-integer32
        unsigned-short
        unsigned-integer32
        unsigned-short)))
  (begin
    (define worksheet-set-landscape
      (foreign-lambda void "worksheetSetLandscape")))
  (begin
    (define worksheet-set-portrait (foreign-lambda void "worksheetSetPortrait")))
  (begin
    (define worksheet-set-page-view (foreign-lambda void "worksheetSetPageView")))
  (begin
    (define worksheet-set-paper
      (foreign-lambda void "worksheetSetPaper" unsigned-char)))
  (begin
    (define worksheet-set-margins
      (foreign-lambda void "worksheetSetMargins" double double double double)))
  (begin
    (define worksheet-set-header
      (foreign-lambda void "worksheetSetHeader" c-string)))
  (begin
    (define worksheet-set-footer
      (foreign-lambda void "worksheetSetFooter" c-string)))
  (begin
    (define worksheet-set-header-opt
      (foreign-lambda void "worksheetSetHeaderOpt" c-string double)))
  (begin
    (define worksheet-set-footer-opt
      (foreign-lambda void "worksheetSetFooterOpt" c-string double)))
  (begin
    (define worksheet-set-h-page-breaks
      (foreign-lambda void "worksheetSetHPageBreaks")))
  (begin
    (define worksheet-set-v-page-breaks
      (foreign-lambda void "worksheetSetVPageBreaks")))
  (begin (define create-format (foreign-lambda void "createFormat")))
  (begin
    (define set-cell-num-format
      (foreign-lambda void "setCellNumFormat" c-string)))
  (begin
    (define set-cell-num-format-index
      (foreign-lambda void "setCellNumFormatIndex" unsigned-byte)))
  (begin (define format-set-hidden (foreign-lambda void "formatSetHidden")))
  (begin (define format-set-unlocked (foreign-lambda void "formatSetUnlocked")))
  (begin (define format-set-bold (foreign-lambda void "formatSetBold")))
  (begin (define format-set-italic (foreign-lambda void "formatSetItalic")))
  (begin
    (define set-underline-single (foreign-lambda void "setUnderlineSingle")))
  (begin
    (define set-underline-double (foreign-lambda void "setUnderlineDouble")))
  (begin
    (define set-underline-single-acct
      (foreign-lambda void "setUnderlineSingleAcct")))
  (begin
    (define set-underline-double-acct
      (foreign-lambda void "setUnderlineDoubleAcct")))
  (begin (define set-strikeout (foreign-lambda void "setStrikeout")))
  (begin (define set-superscript (foreign-lambda void "setSuperscript")))
  (begin (define set-subscript (foreign-lambda void "setSubscript")))
  (begin (define set-font-color (foreign-lambda void "setFontColor" integer32)))
  (begin (define set-font-name (foreign-lambda void "setFontName" c-string)))
  (begin (define set-rotation (foreign-lambda void "setRotation" integer)))
  (begin (define set-indent (foreign-lambda void "setIndent" integer)))
  (begin (define set-bold (foreign-lambda void "setBold")))
  (begin (define set-italic (foreign-lambda void "setItalic")))
  (begin (define set-shrink (foreign-lambda void "setShrink")))
  (begin (define set-pattern (foreign-lambda void "setPattern" unsigned-byte)))
  (begin (define set-align (foreign-lambda void "setAlign" unsigned-byte)))
  (begin (define set-border (foreign-lambda void "setBorder" unsigned-byte)))
  (begin
    (define set-border-bottom (foreign-lambda void "setBorderBottom" unsigned-byte)))
  (begin (define set-border-top (foreign-lambda void "setBorderTop" unsigned-byte)))
  (begin
    (define set-border-left (foreign-lambda void "setBorderLeft" unsigned-byte)))
  (begin
    (define set-border-right (foreign-lambda void "setBorderRight" unsigned-byte)))
  (begin (define set-bg-color (foreign-lambda void "setBgColor" integer32)))
  (begin (define set-fg-color (foreign-lambda void "setFgColor" integer32)))
  (begin
    (define set-border-color (foreign-lambda void "setBorderColor" integer32)))
  (begin
    (define set-border-bottom-color
      (foreign-lambda void "setBorderBottomColor" integer32)))
  (begin
    (define set-border-top-color
      (foreign-lambda void "setBorderTopColor" integer32)))
  (begin
    (define set-border-left-color
      (foreign-lambda void "setBorderLeftColor" integer32)))
  (begin
    (define set-border-right-color
      (foreign-lambda void "setBorderRightColor" integer32)))
  (begin (define turn-on-text-wrap (foreign-lambda void "turnOnTextWrap")))
  (begin
    (define set-doc-properties
      (foreign-lambda
        void
        "setDocProperties"
        c-string
        c-string
        c-string
        c-string
        c-string
        c-string
        c-string
        c-string
        c-string)))
  (begin
    (define workbook-set-custom-property-string
      (foreign-lambda
        void
        "workbookSetCustomPropertyString"
        c-string
        c-string)))
  (begin
    (define workbook-set-custom-property-number
      (foreign-lambda void "workbookSetCustomPropertyNumber" c-string double)))
  (begin
    (define workbook-set-custom-property-boolean
      (foreign-lambda
        void
        "workbookSetCustomPropertyBoolean"
        c-string
        unsigned-char)))
  (begin
    (define workbook-set-custom-property-date-time
      (foreign-lambda
        void
        "workbookSetCustomPropertyDateTime"
        c-string
        integer
        integer
        integer
        integer
        integer
        double)))
  (begin
    (define workbook-define-name
      (foreign-lambda void "workbookDefineName" c-string c-string)))
  (begin
    (define workbook-get-worksheet-by-name
      (foreign-lambda void "workbookGetWorksheetByName" c-string)))
  (begin
    (define workbook-validate-worksheet-name
      (foreign-lambda void "workbookValidateWorksheetName" c-string)))
  (begin (define set-col (foreign-lambda void "setCol" unsigned-short)))
  (begin (define set-row (foreign-lambda void "setRow" unsigned-integer32)))
  (begin
    (define set-pos (foreign-lambda void "setPos" unsigned-integer32 unsigned-short)))
  (begin
    (define create-chart (foreign-lambda void "createChart" unsigned-short)))
  (begin
    (define create-chart-series
      (foreign-lambda void "createChartSeries" c-string c-string)))
  (begin
    (define create-row-col-opt
      (foreign-lambda
        void
        "createRowColOpt"
        unsigned-short
        unsigned-short
        unsigned-short)))
  (begin
    (define chart-series-set-name
      (foreign-lambda void "chartSeriesSetName" c-string)))
  (begin
    (define chart-series-set-line (foreign-lambda void "chartSeriesSetLine")))
  (begin
    (define chart-series-set-fill
      (foreign-lambda void "chartSeriesSetFill" integer)))
  (begin
    (define chart-series-set-pattern
      (foreign-lambda
        void
        "chartSeriesSetPattern"
        unsigned-byte
        integer32
        unsigned-integer32)))
  (begin
    (define chart-series-set-marker-type
      (foreign-lambda void "chartSeriesSetMarkerType" unsigned-char)))
  (begin
    (define chart-series-set-marker-size
      (foreign-lambda void "chartSeriesSetMarkerSize" unsigned-char)))
  (begin
    (define chart-series-set-marker-line
      (foreign-lambda void "chartSeriesSetMarkerLine")))
  (begin
    (define chart-series-set-marker-fill
      (foreign-lambda void "chartSeriesSetMarkerFill" integer)))
  (begin
    (define chart-series-set-marker-pattern
      (foreign-lambda
        void
        "chartSeriesSetMarkerPattern"
        unsigned-byte
        integer32
        integer32)))
  (begin
    (define chart-series-invert-if-negative
      (foreign-lambda void "chartSeriesInvertIfNegative")))
  (begin
    (define chart-series-set-points (foreign-lambda void "chartSeriesSetPoints")))
  (begin
    (define chart-series-set-smooth
      (foreign-lambda void "chartSeriesSetSmooth" unsigned-byte)))
  (begin
    (define chart-series-set-labels (foreign-lambda void "chartSeriesSetLabels")))
  (begin
    (define chart-series-set-labels-opt
      (foreign-lambda
        void
        "chartSeriesSetLabelsOpt"
        unsigned-byte
        unsigned-byte
        unsigned-byte)))
  (begin
    (define chart-series-set-labels-separator
      (foreign-lambda void "chartSeriesSetLabelsSeparator" unsigned-byte)))
  (begin
    (define chart-series-set-labels-position
      (foreign-lambda void "chartSeriesSetLabelsPosition" unsigned-byte)))
  (begin
    (define chart-series-set-labels-leader-line
      (foreign-lambda void "chartSeriesSetLabelsLeaderLine")))
  (begin
    (define chart-series-set-labels-legend
      (foreign-lambda void "chartSeriesSetLabelsLegend")))
  (begin
    (define chart-series-set-labels-percentage
      (foreign-lambda void "chartSeriesSetLabelsPercentage")))
  (begin
    (define chart-series-set-labels-num-format
      (foreign-lambda void "chartSeriesSetLabelsNumFormat" c-string)))
  (begin
    (define chart-series-set-labels-font
      (foreign-lambda void "chartSeriesSetLabelsFont" c-string)))
  (begin
    (define chart-series-set-trendline
      (foreign-lambda void "chartSeriesSetTrendline" unsigned-byte unsigned-byte)))
  (begin
    (define chart-series-set-trendline-forecast
      (foreign-lambda void "chartSeriesSetTrendlineForecast" double double)))
  (begin
    (define chart-series-set-trendline-equation
      (foreign-lambda void "chartSeriesSetTrendlineEquation")))
  (begin
    (define chart-series-set-trendline-r-squared
      (foreign-lambda void "chartSeriesSetTrendlineRSquared")))
  (begin
    (define chart-series-set-trendline-intercept
      (foreign-lambda void "chartSeriesSetTrendlineIntercept" double)))
  (begin
    (define chart-series-set-trendline-name
      (foreign-lambda void "chartSeriesSetTrendlineName" c-string)))
  (begin
    (define chart-series-set-trendline-line
      (foreign-lambda void "chartSeriesSetTrendlineLine")))
  (begin
    (define chart-series-get-error-bars
      (foreign-lambda void "chartSeriesGetErrorBars" unsigned-byte)))
  (begin
    (define chart-series-set-error-bars
      (foreign-lambda
        void
        "chartSeriesSetErrorBars"
        unsigned-byte
        double
        unsigned-short)))
  (begin
    (define chart-series-set-error-bars-direction
      (foreign-lambda void "chartSeriesSetErrorBarsDirection" unsigned-byte)))
  (begin
    (define chart-series-set-error-bars-end-cap
      (foreign-lambda void "chartSeriesSetErrorBarsEndCap" unsigned-byte)))
  (begin
    (define chart-series-set-error-bars-line
      (foreign-lambda void "chartSeriesSetErrorBarsLine")))
  (begin
    (define chart-toggle-major-grid-lines-x-axis
      (foreign-lambda void "chartToggleMajorGridLinesXAxis" unsigned-short)))
  (begin
    (define chart-toggle-major-grid-lines-y-axis
      (foreign-lambda void "chartToggleMajorGridLinesYAxis" unsigned-short)))
  (begin
    (define chart-toggle-minor-grid-lines-x-axis
      (foreign-lambda void "chartToggleMinorGridLinesXAxis" unsigned-short)))
  (begin
    (define chart-toggle-minor-grid-lines-y-axis
      (foreign-lambda void "chartToggleMinorGridLinesYAxis" unsigned-short)))
  (begin
    (define chart-x-axis-set-name
      (foreign-lambda void "chartXAxisSetName" c-string)))
  (begin
    (define chart-y-axis-set-name
      (foreign-lambda void "chartYAxisSetName" c-string)))
  (begin
    (define chart-x-axis-set-name-range
      (foreign-lambda void "chartXAxisSetNameRange" c-string)))
  (begin
    (define chart-y-axis-set-name-range
      (foreign-lambda void "chartYAxisSetNameRange" c-string)))
  (begin
    (define chart-x-axis-set-name-font
      (foreign-lambda void "chartXAxisSetNameFont")))
  (begin
    (define chart-y-axis-set-name-font
      (foreign-lambda void "chartYAxisSetNameFont")))
  (begin
    (define chart-x-axis-set-num-font (foreign-lambda void "chartXAxisSetNumFont")))
  (begin
    (define chart-y-axis-set-num-font (foreign-lambda void "chartYAxisSetNumFont")))
  (begin
    (define chart-x-axis-set-num-format
      (foreign-lambda void "chartXAxisSetNumFormat" c-string)))
  (begin
    (define chart-y-axis-set-num-format
      (foreign-lambda void "chartYAxisSetNumFormat" c-string)))
  (begin (define chart-x-axis-set-line (foreign-lambda void "chartXAxisSetLine")))
  (begin (define chart-y-axis-set-line (foreign-lambda void "chartYAxisSetLine")))
  (begin (define chart-x-axis-set-fill (foreign-lambda void "chartXAxisSetFill")))
  (begin (define chart-y-axis-set-fill (foreign-lambda void "chartYAxisSetFill")))
  (begin
    (define chart-x-axis-set-pattern (foreign-lambda void "chartXAxisSetPattern")))
  (begin
    (define chart-y-axis-set-pattern (foreign-lambda void "chartYAxisSetPattern")))
  (begin (define chart-reverse-x-axis (foreign-lambda void "chartReverseXAxis")))
  (begin (define chart-reverse-y-axis (foreign-lambda void "chartReverseYAxis")))
  (begin
    (define chart-x-axis-set-crossing
      (foreign-lambda void "chartXAxisSetCrossing" double)))
  (begin
    (define chart-y-axis-set-crossing
      (foreign-lambda void "chartYAxisSetCrossing" double)))
  (begin
    (define chart-x-axis-set-crossing-max
      (foreign-lambda void "chartXAxisSetCrossingMax")))
  (begin
    (define chart-y-axis-set-crossing-max
      (foreign-lambda void "chartYAxisSetCrossingMax")))
  (begin (define chart-x-axis-off (foreign-lambda void "chartXAxisOff")))
  (begin (define chart-y-axis-off (foreign-lambda void "chartYAxisOff")))
  (begin
    (define chart-x-axis-set-position
      (foreign-lambda void "chartXAxisSetPosition" unsigned-byte)))
  (begin
    (define chart-y-axis-set-position
      (foreign-lambda void "chartYAxisSetPosition" unsigned-byte)))
  (begin
    (define chart-x-axis-set-label-position
      (foreign-lambda void "chartXAxisSetLabelPosition" unsigned-byte)))
  (begin
    (define chart-y-axis-set-label-position
      (foreign-lambda void "chartYAxisSetLabelPosition" unsigned-byte)))
  (begin
    (define chart-x-axis-set-label-align
      (foreign-lambda void "chartXAxisSetLabelAlign" unsigned-byte)))
  (begin
    (define chart-y-axis-set-label-align
      (foreign-lambda void "chartYAxisSetLabelAlign" unsigned-byte)))
  (begin
    (define chart-x-axis-set-min (foreign-lambda void "chartXAxisSetMin" double)))
  (begin
    (define chart-y-axis-set-min (foreign-lambda void "chartYAxisSetMin" double)))
  (begin
    (define chart-x-axis-set-max (foreign-lambda void "chartXAxisSetMax" double)))
  (begin
    (define chart-y-axis-set-max (foreign-lambda void "chartYAxisSetMax" double)))
  (begin
    (define chart-x-axis-set-log-base
      (foreign-lambda void "chartXAxisSetLogBase" unsigned-short)))
  (begin
    (define chart-y-axis-set-log-base
      (foreign-lambda void "chartYAxisSetLogBase" unsigned-short)))
  (begin
    (define chart-x-axis-set-major-tick-mark
      (foreign-lambda void "chartXAxisSetMajorTickMark" unsigned-byte)))
  (begin
    (define chart-y-axis-set-major-tick-mark
      (foreign-lambda void "chartYAxisSetMajorTickMark" unsigned-byte)))
  (begin
    (define chart-x-axis-set-minor-tick-mark
      (foreign-lambda void "chartXAxisSetMinorTickMark" unsigned-byte)))
  (begin
    (define chart-y-axis-set-minor-tick-mark
      (foreign-lambda void "chartYAxisSetMinorTickMark" unsigned-byte)))
  (begin
    (define chart-x-axis-set-interval-unit
      (foreign-lambda void "chartXAxisSetIntervalUnit" unsigned-short)))
  (begin
    (define chart-y-axis-set-interval-unit
      (foreign-lambda void "chartYAxisSetIntervalUnit" unsigned-short)))
  (begin
    (define chart-x-axis-set-interval-tick
      (foreign-lambda void "chartXAxisSetIntervalTick" unsigned-short)))
  (begin
    (define chart-y-axis-set-interval-tick
      (foreign-lambda void "chartYAxisSetIntervalTick" unsigned-short)))
  (begin
    (define chart-x-axis-set-major-unit
      (foreign-lambda void "chartXAxisSetMajorUnit" double)))
  (begin
    (define chart-y-axis-set-major-unit
      (foreign-lambda void "chartYAxisSetMajorUnit" double)))
  (begin
    (define chart-x-axis-set-minor-unit
      (foreign-lambda void "chartXAxisSetMinorUnit" double)))
  (begin
    (define chart-y-axis-set-minor-unit
      (foreign-lambda void "chartYAxisSetMinorUnit" double)))
  (begin
    (define chart-x-axis-set-display-units
      (foreign-lambda void "chartXAxisSetDisplayUnits" unsigned-byte)))
  (begin
    (define chart-y-axis-set-display-units
      (foreign-lambda void "chartYAxisSetDisplayUnits" unsigned-byte)))
  (begin
    (define chart-x-axis-set-display-units-visible
      (foreign-lambda void "chartXAxisSetDisplayUnitsVisible" unsigned-byte)))
  (begin
    (define chart-y-axis-set-display-units-visible
      (foreign-lambda void "chartYAxisSetDisplayUnitsVisible" unsigned-byte)))
  (begin
    (define chart-x-axis-set-major-gridlines-visible
      (foreign-lambda void "chartXAxisSetMajorGridlinesVisible" unsigned-byte)))
  (begin
    (define chart-y-axis-set-major-gridlines-visible
      (foreign-lambda void "chartYAxisSetMajorGridlinesVisible" unsigned-byte)))
  (begin
    (define chart-x-axis-set-minor-gridlines-visible
      (foreign-lambda void "chartXAxisSetMinorGridlinesVisible" unsigned-byte)))
  (begin
    (define chart-y-axis-set-minor-gridlines-visible
      (foreign-lambda void "chartYAxisSetMinorGridlinesVisible" unsigned-byte)))
  (begin
    (define chart-x-axis-set-major-gridlines-set-line
      (foreign-lambda void "chartXAxisSetMajorGridlinesSetLine")))
  (begin
    (define chart-y-axis-set-major-gridlines-set-line
      (foreign-lambda void "chartYAxisSetMajorGridlinesSetLine")))
  (begin
    (define chart-x-axis-set-minor-gridlines-set-line
      (foreign-lambda void "chartXAxisSetMinorGridlinesSetLine")))
  (begin
    (define chart-y-axis-set-minor-gridlines-set-line
      (foreign-lambda void "chartYAxisSetMinorGridlinesSetLine")))
  (begin
    (define chart-title-set-name
      (foreign-lambda void "chartTitleSetName" c-string)))
  (begin
    (define chart-title-set-name-range
      (foreign-lambda void "chartTitleSetNameRange" c-string)))
  (begin
    (define chart-title-set-name-font
      (foreign-lambda void "chartTitleSetNameFont")))
  (begin (define chart-title-off (foreign-lambda void "chartTitleOff")))
  (begin
    (define chart-legend-set-position
      (foreign-lambda void "chartLegendSetPosition" unsigned-byte)))
  (begin
    (define chart-legend-set-font (foreign-lambda void "chartLegendSetFont")))
  (begin
    (define add-chart-legend-series-to-delete
      (foreign-lambda void "addChartLegendSeriesToDelete" unsigned-short)))
  (begin
    (define chart-legend-delete-series
      (foreign-lambda void "chartLegendDeleteSeries")))
  (begin (define chart-area-set-line (foreign-lambda void "chartAreaSetLine")))
  (begin (define chart-area-set-fill (foreign-lambda void "chartAreaSetFill")))
  (begin
    (define chart-area-set-pattern (foreign-lambda void "chartAreaSetPattern")))
  (begin
    (define chart-plot-area-set-line (foreign-lambda void "chartPlotAreaSetLine")))
  (begin
    (define chart-plot-area-set-fill (foreign-lambda void "chartPlotAreaSetFill")))
  (begin
    (define chart-plot-area-set-pattern
      (foreign-lambda void "chartPlotAreaSetPattern")))
  (begin
    (define chart-set-style (foreign-lambda void "chartSetStyle" unsigned-byte)))
  (begin (define chart-set-table (foreign-lambda void "chartSetTable")))
  (begin
    (define chart-set-table-grid
      (foreign-lambda
        void
        "chartSetTableGrid"
        unsigned-byte
        unsigned-byte
        unsigned-byte
        unsigned-byte)))
  (begin
    (define chart-set-up-down-bars (foreign-lambda void "chartSetUpDownBars")))
  (begin
    (define chart-set-up-bars-line-and-fill
      (foreign-lambda void "chartSetUpBarsLineAndFill")))
  (begin
    (define chart-set-down-bars-line-and-fill
      (foreign-lambda void "chartSetDownBarsLineAndFill")))
  (begin (define chart-set-drop-lines (foreign-lambda void "chartSetDropLines")))
  (begin
    (define chart-set-high-low-lines (foreign-lambda void "chartSetHighLowLines")))
  (begin
    (define chart-set-series-overlap
      (foreign-lambda void "chartSetSeriesOverlap" unsigned-byte)))
  (begin
    (define chart-set-series-gap
      (foreign-lambda void "chartSetSeriesGap" unsigned-short)))
  (begin
    (define chart-show-blanks-as
      (foreign-lambda void "chartShowBlanksAs" unsigned-byte)))
  (begin
    (define chart-show-hidden-data (foreign-lambda void "chartShowHiddenData")))
  (begin
    (define chart-set-rotation
      (foreign-lambda void "chartSetRotation" unsigned-short)))
  (begin
    (define chart-set-hole-size
      (foreign-lambda void "chartSetHoleSize" unsigned-byte)))
  (begin
    (define chart-series-set-name-range
      (foreign-lambda void "chartSeriesSetNameRange" c-string)))
  (begin
    (define chart-x-axis-set-reverse (foreign-lambda void "chartXAxisSetReverse")))
  (begin
    (define chart-y-axis-set-reverse (foreign-lambda void "chartYAxisSetReverse")))
  (begin
    (define create-chartsheet
      (foreign-lambda void "createChartsheet" c-string)))
  (begin
    (define chartsheet-set-chart (foreign-lambda void "chartsheetSetChart")))
  (begin
    (define chartsheet-activate (foreign-lambda void "chartsheetActivate")))
  (begin (define chartsheet-select (foreign-lambda void "chartsheetSelect")))
  (begin (define chartsheet-hide (foreign-lambda void "chartsheetHide")))
  (begin
    (define chartsheet-set-first-sheet
      (foreign-lambda void "chartsheetSetFirstSheet")))
  (begin
    (define chartsheet-set-tab-color
      (foreign-lambda void "chartsheetSetTabColor" integer32)))
  (begin
    (define worksheet-set-default-row
      (foreign-lambda void "worksheetSetDefaultRow" double unsigned-byte)))
  (begin
    (define worksheet-protect
      (foreign-lambda
        void
        "worksheetProtect"
        c-string
        unsigned-byte
        unsigned-byte
        unsigned-byte
        unsigned-byte
        unsigned-byte
        unsigned-byte
        unsigned-byte
        unsigned-byte
        unsigned-byte
        unsigned-byte
        unsigned-byte
        unsigned-byte
        unsigned-byte
        unsigned-byte
        unsigned-byte
        unsigned-byte
        unsigned-byte)))
  (begin
    (define chartsheet-protect
      (foreign-lambda
        void
        "chartsheetProtect"
        c-string
        unsigned-byte
        unsigned-byte
        unsigned-byte
        unsigned-byte
        unsigned-byte
        unsigned-byte
        unsigned-byte
        unsigned-byte
        unsigned-byte
        unsigned-byte
        unsigned-byte
        unsigned-byte
        unsigned-byte
        unsigned-byte
        unsigned-byte
        unsigned-byte
        unsigned-byte)))
  (begin
    (define chartsheet-set-zoom
      (foreign-lambda void "chartsheetSetZoom" unsigned-short)))
  (begin
    (define chartsheet-set-landscape
      (foreign-lambda void "chartsheetSetLandscape")))
  (begin
    (define chartsheet-set-portrait
      (foreign-lambda void "chartsheetSetPortrait")))
  (begin
    (define chartsheet-set-paper
      (foreign-lambda void "chartsheetSetPaper" unsigned-byte)))
  (begin
    (define chartsheet-set-margins
      (foreign-lambda
        void
        "chartsheetSetMargins"
        double
        double
        double
        double)))
  (begin
    (define chartsheet-set-header
      (foreign-lambda void "chartsheetSetHeader" c-string)))
  (begin
    (define chartsheet-set-footer
      (foreign-lambda void "chartsheetSetFooter" c-string)))
  (begin
    (define chartsheet-set-header-opt
      (foreign-lambda void "chartsheetSetHeaderOpt" c-string double)))
  (begin
    (define chartsheet-set-footer-opt
      (foreign-lambda void "chartsheetSetFooterOpt" c-string double)))
  (begin
    (define init-rich-string-list
      (foreign-lambda void "initRichStringList" integer)))
  (begin
    (define init-rich-string (foreign-lambda void "initRichString" integer)))
  (begin
    (define create-rich-string-fragment
      (foreign-lambda void "createRichStringFragment" c-string)))
  (begin (define rich-string-cleanup (foreign-lambda void "richStringCleanup")))
  (begin
    (define rich-string-list-cleanup
      (foreign-lambda void "richStringListCleanup")))
  (begin
    (define worksheet-set-tab-color
      (foreign-lambda void "worksheetSetTabColor" integer32)))
  (begin
    (define worksheet-write-rich-string-fragments
      (foreign-lambda void "worksheetWriteRichStringFragments")))
  (begin
    (define all-worksheets-set-column
      (foreign-lambda
        void
        "allWorksheetsSetColumn"
        double
        unsigned-short
        unsigned-short)))
  (begin
    (define all-worksheets-set-row
      (foreign-lambda void "allWorksheetsSetRow" double unsigned-integer32)))
  (begin
    (define all-worksheets-write-string
      (foreign-lambda void "allWorksheetsWriteString" c-string)))
  (begin
    (define all-worksheets-write-formula
      (foreign-lambda void "allWorksheetsWriteFormula" c-string)))
  (begin (define close-workbook (foreign-lambda void "closeWorkbook"))))

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


        
        
  
  ]