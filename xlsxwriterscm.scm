(use srfi-69 easyffi)
[module 
    xlsxwriterscm
    *
  (import scheme chicken foreign srfi-69)
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
  
  
;;; generated by chicken-wrap from libxlsxwriter_layer.c

(begin
  (begin
    (define createWorkbook (foreign-lambda void "createWorkbook" c-string)))
  (begin (define formatsCleanup (foreign-lambda void "formatsCleanup")))
  (begin (define rowNumberCleanup (foreign-lambda void "rowNumberCleanup")))
  (begin (define colNumberCleanup (foreign-lambda void "colNumberCleanup")))
  (begin
    (define dataValidationListsCleanup
      (foreign-lambda void "dataValidationListsCleanup")))
  (begin
    (define dataValidationsCleanup
      (foreign-lambda void "dataValidationsCleanup")))
  (begin
    (define chartPointsCleanup (foreign-lambda void "chartPointsCleanup")))
  (begin (define chartFillsCleanup (foreign-lambda void "chartFillsCleanup")))
  (begin
    (define chartPatternsCleanup (foreign-lambda void "chartPatternsCleanup")))
  (begin (define chartLinesCleanup (foreign-lambda void "chartLinesCleanup")))
  (begin (define seriesCleanup (foreign-lambda void "seriesCleanup")))
  (begin (define chartsCleanup (foreign-lambda void "chartsCleanup")))
  (begin
    (define seriesErrorBarsCleanup
      (foreign-lambda void "seriesErrorBarsCleanup")))
  (begin (define chartFontsCleanup (foreign-lambda void "chartFontsCleanup")))
  (begin
    (define debugFormat
      (foreign-lambda void "debugFormat" (c-pointer "lxw_format"))))
  (begin (define debugFormats (foreign-lambda void "debugFormats")))
  (begin
    (define debugChartPattern
      (foreign-lambda
        void
        "debugChartPattern"
        (c-pointer "lxw_chart_pattern"))))
  (begin
    (define debugChartLine
      (foreign-lambda void "debugChartLine" (c-pointer "lxw_chart_line"))))
  (begin (define initFormats (foreign-lambda void "initFormats" integer)))
  (begin
    (define initRowNumbers (foreign-lambda void "initRowNumbers" integer)))
  (begin
    (define initColNumbers (foreign-lambda void "initColNumbers" integer)))
  (begin
    (define initDataValidations
      (foreign-lambda void "initDataValidations" integer)))
  (begin
    (define initChartFills (foreign-lambda void "initChartFills" integer)))
  (begin
    (define initChartPoints
      (foreign-lambda void "initChartPoints" integer)))
  (begin
    (define initChartLines (foreign-lambda void "initChartLines" integer)))
  (begin
    (define initChartPatterns
      (foreign-lambda void "initChartPatterns" integer)))
  (begin (define initSeries (foreign-lambda void "initSeries" integer)))
  (begin (define initCharts (foreign-lambda void "initCharts" integer)))
  (begin
    (define initChartFonts (foreign-lambda void "initChartFonts" integer)))
  (begin
    (define createDataValidationList
      (foreign-lambda void "createDataValidationList" c-string)))
  (begin
    (define createChartPoint
      (foreign-lambda void "createChartPoint" unsigned-integer double)))
  (begin
    (define createChartFill
      (foreign-lambda void "createChartFill" integer32 unsigned-byte unsigned-byte)))
  (begin
    (define createChartLine
      (foreign-lambda
        void
        "createChartLine"
        integer32
        unsigned-byte
        float
        unsigned-byte)))
  (begin
    (define createChartPattern
      (foreign-lambda
        void
        "createChartPattern"
        integer32
        integer32
        unsigned-byte)))
  (begin
    (define createChartFont
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
    (define addToRowNumberList
      (foreign-lambda void "addToRowNumberList" unsigned-integer32)))
  (begin
    (define addToColNumberList
      (foreign-lambda void "addToColNumberList" unsigned-short)))
  (begin
    (define debugDataValidationList
      (foreign-lambda void "debugDataValidationList")))
  (begin
    (define debugDataValidation
      (foreign-lambda
        void
        "debugDataValidation"
        (c-pointer "lxw_data_validation"))))
  (begin
    (define debugDataValidations (foreign-lambda void "debugDataValidations")))
  (begin (define debugChartPoints (foreign-lambda void "debugChartPoints")))
  (begin
    (define debugChartFill
      (foreign-lambda void "debugChartFill" (c-pointer "lxw_chart_fill"))))
  (begin (define debugChartFills (foreign-lambda void "debugChartFills")))
  (begin (define debugChartLines (foreign-lambda void "debugChartLines")))
  (begin
    (define debugChartPatterns (foreign-lambda void "debugChartPatterns")))
  (begin
    (define setDataValidationIndex
      (foreign-lambda void "setDataValidationIndex" integer)))
  (begin
    (define setChartPointIndex
      (foreign-lambda void "setChartPointIndex" integer)))
  (begin
    (define setChartFillIndex
      (foreign-lambda void "setChartFillIndex" integer)))
  (begin
    (define setChartLineIndex
      (foreign-lambda void "setChartLineIndex" integer)))
  (begin
    (define setChartPatternIndex
      (foreign-lambda void "setChartPatternIndex" integer)))
  (begin
    (define setFormatIndex (foreign-lambda void "setFormatIndex" integer)))
  (begin
    (define setSeriesIndex (foreign-lambda void "setSeriesIndex" integer)))
  (begin
    (define setChartIndex (foreign-lambda void "setChartIndex" integer)))
  (begin
    (define setSeriesErrorBarsIndex
      (foreign-lambda void "setSeriesErrorBarsIndex" integer)))
  (begin
    (define setChartFontIndex
      (foreign-lambda void "setChartFontIndex" integer)))
  (begin (define addWorksheet (foreign-lambda void "addWorksheet" c-string)))
  (begin
    (define worksheetWriteString
      (foreign-lambda void "worksheetWriteString" c-string)))
  (begin
    (define worksheetWriteNumber
      (foreign-lambda void "worksheetWriteNumber" float)))
  (begin
    (define worksheetWriteFormula
      (foreign-lambda
        void
        "worksheetWriteFormula"
        c-string
        unsigned-integer32
        unsigned-short)))
  (begin
    (define worksheetWriteArrayFormula
      (foreign-lambda
        void
        "worksheetWriteArrayFormula"
        c-string
        unsigned-integer32
        unsigned-short
        unsigned-integer32
        unsigned-short)))
  (begin
    (define worksheetWriteDatetime
      (foreign-lambda
        void
        "worksheetWriteDatetime"
        integer
        integer
        integer
        integer
        integer
        double
        unsigned-integer32
        unsigned-short)))
  (begin
    (define worksheetWriteURL
      (foreign-lambda
        void
        "worksheetWriteURL"
        c-string
        unsigned-integer32
        unsigned-short)))
  (begin
    (define worksheetWriteBoolean
      (foreign-lambda
        void
        "worksheetWriteBoolean"
        short
        unsigned-integer32
        unsigned-short)))
  (begin
    (define worksheetWriteBlank
      (foreign-lambda void "worksheetWriteBlank" unsigned-integer32 unsigned-short)))
  (begin
    (define worksheetSetRow
      (foreign-lambda void "worksheetSetRow" double unsigned-integer32)))
  (begin
    (define worksheetSetRowOpt
      (foreign-lambda
        void
        "worksheetSetRowOpt"
        unsigned-char
        unsigned-char
        unsigned-char
        double
        unsigned-integer32)))
  (begin
    (define worksheetSetColumn
      (foreign-lambda
        void
        "worksheetSetColumn"
        double
        unsigned-short
        unsigned-short)))
  (begin
    (define worksheetSetColumnOpt
      (foreign-lambda
        void
        "worksheetSetColumnOpt"
        unsigned-char
        unsigned-char
        unsigned-char
        double
        unsigned-short
        unsigned-short)))
  (begin
    (define worksheetInsertImage
      (foreign-lambda
        void
        "worksheetInsertImage"
        c-string
        unsigned-integer32
        unsigned-short)))
  (begin
    (define worksheetInsertImageOpt
      (foreign-lambda
        void
        "worksheetInsertImageOpt"
        integer
        integer
        double
        double
        c-string
        unsigned-integer32
        unsigned-short)))
  (begin
    (define worksheetInsertChart
      (foreign-lambda void "worksheetInsertChart" unsigned-integer32 unsigned-short)))
  (begin
    (define worksheetInsertChartOpt
      (foreign-lambda
        void
        "worksheetInsertChartOpt"
        integer
        integer
        integer
        integer
        unsigned-integer32
        unsigned-short)))
  (begin
    (define worksheetMergeRange
      (foreign-lambda
        void
        "worksheetMergeRange"
        c-string
        unsigned-integer32
        unsigned-short
        unsigned-integer32
        unsigned-short)))
  (begin
    (define worksheetAutoFilter
      (foreign-lambda
        void
        "worksheetAutoFilter"
        unsigned-integer32
        unsigned-short
        unsigned-integer32
        unsigned-short)))
  (begin
    (define worksheetDataValidationCell
      (foreign-lambda
        void
        "worksheetDataValidationCell"
        unsigned-integer32
        unsigned-short)))
  (begin
    (define worksheetDataValidationRange
      (foreign-lambda
        void
        "worksheetDataValidationRange"
        unsigned-integer32
        unsigned-short
        unsigned-integer32
        unsigned-short)))
  (begin (define worksheetActivate (foreign-lambda void "worksheetActivate")))
  (begin (define worksheetSelect (foreign-lambda void "worksheetSelect")))
  (begin (define worksheetHide (foreign-lambda void "worksheetHide")))
  (begin
    (define worksheetSetFirstSheet
      (foreign-lambda void "worksheetSetFirstSheet")))
  (begin
    (define worksheetSplitPanes
      (foreign-lambda void "worksheetSplitPanes" double double)))
  (begin
    (define worksheetSetSelection
      (foreign-lambda
        void
        "worksheetSetSelection"
        unsigned-integer32
        unsigned-short
        unsigned-integer32
        unsigned-short)))
  (begin
    (define worksheetSetLandscape
      (foreign-lambda void "worksheetSetLandscape")))
  (begin
    (define worksheetSetPortrait (foreign-lambda void "worksheetSetPortrait")))
  (begin
    (define worksheetSetPageView (foreign-lambda void "worksheetSetPageView")))
  (begin
    (define worksheetSetPaper
      (foreign-lambda void "worksheetSetPaper" unsigned-char)))
  (begin
    (define worksheetSetMargins
      (foreign-lambda void "worksheetSetMargins" double double double double)))
  (begin
    (define worksheetSetHeader
      (foreign-lambda void "worksheetSetHeader" c-string)))
  (begin
    (define worksheetSetFooter
      (foreign-lambda void "worksheetSetFooter" c-string)))
  (begin
    (define worksheetSetHeaderOpt
      (foreign-lambda void "worksheetSetHeaderOpt" c-string double)))
  (begin
    (define worksheetSetFooterOpt
      (foreign-lambda void "worksheetSetFooterOpt" c-string double)))
  (begin
    (define worksheetSetHPageBreaks
      (foreign-lambda void "worksheetSetHPageBreaks")))
  (begin
    (define worksheetSetVPageBreaks
      (foreign-lambda void "worksheetSetVPageBreaks")))
  (begin (define createFormat (foreign-lambda void "createFormat")))
  (begin
    (define setCellNumFormat
      (foreign-lambda void "setCellNumFormat" c-string)))
  (begin
    (define setCellNumFormatIndex
      (foreign-lambda void "setCellNumFormatIndex" unsigned-byte)))
  (begin (define formatHideFormula (foreign-lambda void "formatHideFormula")))
  (begin (define formatSetUnlocked (foreign-lambda void "formatSetUnlocked")))
  (begin (define formatSetBold (foreign-lambda void "formatSetBold")))
  (begin (define formatSetItalic (foreign-lambda void "formatSetItalic")))
  (begin
    (define setUnderlineSingle (foreign-lambda void "setUnderlineSingle")))
  (begin
    (define setUnderlineDouble (foreign-lambda void "setUnderlineDouble")))
  (begin
    (define setUnderlineSingleAcct
      (foreign-lambda void "setUnderlineSingleAcct")))
  (begin
    (define setUnderlineDoubleAcct
      (foreign-lambda void "setUnderlineDoubleAcct")))
  (begin (define setStrikeout (foreign-lambda void "setStrikeout")))
  (begin (define setSuperscript (foreign-lambda void "setSuperscript")))
  (begin (define setSubscript (foreign-lambda void "setSubscript")))
  (begin (define setFontColor (foreign-lambda void "setFontColor" unsigned-integer32)))
  (begin (define setFontName (foreign-lambda void "setFontName" c-string)))
  (begin (define setRotation (foreign-lambda void "setRotation" integer)))
  (begin
    (define setIndentation (foreign-lambda void "setIndentation" integer)))
  (begin (define setBold (foreign-lambda void "setBold")))
  (begin (define setItalic (foreign-lambda void "setItalic")))
  (begin (define setShrink (foreign-lambda void "setShrink")))
  (begin (define setPattern (foreign-lambda void "setPattern" unsigned-byte)))
  (begin (define setAlign (foreign-lambda void "setAlign" unsigned-byte)))
  (begin (define setBorder (foreign-lambda void "setBorder" unsigned-byte)))
  (begin
    (define setBorderBottom (foreign-lambda void "setBorderBottom" unsigned-byte)))
  (begin (define setBorderTop (foreign-lambda void "setBorderTop" unsigned-byte)))
  (begin
    (define setBorderLeft (foreign-lambda void "setBorderLeft" unsigned-byte)))
  (begin
    (define setBorderRight (foreign-lambda void "setBorderRight" unsigned-byte)))
  (begin (define setBGColor (foreign-lambda void "setBGColor" unsigned-integer32)))
  (begin (define setFGColor (foreign-lambda void "setFGColor" unsigned-integer32)))
  (begin
    (define setBorderColor (foreign-lambda void "setBorderColor" unsigned-integer32)))
  (begin
    (define setBorderBottomColor
      (foreign-lambda void "setBorderBottomColor" unsigned-integer32)))
  (begin
    (define setBorderTopColor
      (foreign-lambda void "setBorderTopColor" unsigned-integer32)))
  (begin
    (define setBorderLeftColor
      (foreign-lambda void "setBorderLeftColor" unsigned-integer32)))
  (begin
    (define setBorderRightColor
      (foreign-lambda void "setBorderRightColor" unsigned-integer32)))
  (begin (define turnOnTextWrap (foreign-lambda void "turnOnTextWrap")))
  (begin
    (define setDocProperties
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
    (define workbookSetCustomPropertyString
      (foreign-lambda
        void
        "workbookSetCustomPropertyString"
        c-string
        c-string)))
  (begin
    (define workbookSetCustomPropertyNumber
      (foreign-lambda void "workbookSetCustomPropertyNumber" c-string double)))
  (begin
    (define workbookSetCustomPropertyBoolean
      (foreign-lambda
        void
        "workbookSetCustomPropertyBoolean"
        c-string
        unsigned-char)))
  (begin
    (define workbookSetCustomPropertyDatetime
      (foreign-lambda
        void
        "workbookSetCustomPropertyDatetime"
        c-string
        integer
        integer
        integer
        integer
        integer
        double)))
  (begin
    (define workbookDefineName
      (foreign-lambda void "workbookDefineName" c-string c-string)))
  (begin
    (define workbookGetWorksheetByName
      (foreign-lambda void "workbookGetWorksheetByName" c-string)))
  (begin
    (define workbookValidateWorksheetName
      (foreign-lambda void "workbookValidateWorksheetName" c-string)))
  (begin (define setCol (foreign-lambda void "setCol" unsigned-short)))
  (begin (define setRow (foreign-lambda void "setRow" unsigned-integer32)))
  (begin
    (define setPos (foreign-lambda void "setPos" unsigned-integer32 unsigned-short)))
  (begin
    (define createChart (foreign-lambda void "createChart" unsigned-short)))
  (begin
    (define createChartSeries
      (foreign-lambda void "createChartSeries" c-string c-string)))
  (begin
    (define chartSeriesSetName
      (foreign-lambda void "chartSeriesSetName" c-string)))
  (begin
    (define chartSeriesSetLine (foreign-lambda void "chartSeriesSetLine")))
  (begin
    (define chartSeriesSetFill
      (foreign-lambda void "chartSeriesSetFill" integer)))
  (begin
    (define chartSeriesSetPattern
      (foreign-lambda
        void
        "chartSeriesSetPattern"
        unsigned-byte
        unsigned-integer32
        unsigned-integer32)))
  (begin
    (define chartSeriesSetMarkerType
      (foreign-lambda void "chartSeriesSetMarkerType" unsigned-char)))
  (begin
    (define chartSeriesSetMarkerSize
      (foreign-lambda void "chartSeriesSetMarkerSize" unsigned-char)))
  (begin
    (define chartSeriesSetMarkerLine
      (foreign-lambda void "chartSeriesSetMarkerLine")))
  (begin
    (define chartSeriesSetMarkerFill
      (foreign-lambda void "chartSeriesSetMarkerFill" integer)))
  (begin
    (define chartSeriesSetMarkerPattern
      (foreign-lambda
        void
        "chartSeriesSetMarkerPattern"
        unsigned-byte
        unsigned-integer32
        unsigned-integer32)))
  (begin
    (define chartSeriesInvertIfNegative
      (foreign-lambda void "chartSeriesInvertIfNegative")))
  (begin
    (define chartSeriesSetPoints (foreign-lambda void "chartSeriesSetPoints")))
  (begin
    (define chartSeriesSetSmooth
      (foreign-lambda void "chartSeriesSetSmooth" unsigned-byte)))
  (begin
    (define chartSeriesSetLabels (foreign-lambda void "chartSeriesSetLabels")))
  (begin
    (define chartSeriesSetLabelsOpt
      (foreign-lambda
        void
        "chartSeriesSetLabelsOpt"
        unsigned-byte
        unsigned-byte
        unsigned-byte)))
  (begin
    (define chartSeriesSetLabelsSeparator
      (foreign-lambda void "chartSeriesSetLabelsSeparator" unsigned-byte)))
  (begin
    (define chartSeriesSetLabelsPosition
      (foreign-lambda void "chartSeriesSetLabelsPosition" unsigned-byte)))
  (begin
    (define chartSeriesSetLabelsLeaderLine
      (foreign-lambda void "chartSeriesSetLabelsLeaderLine")))
  (begin
    (define chartSeriesSetLabelsLegend
      (foreign-lambda void "chartSeriesSetLabelsLegend")))
  (begin
    (define chartSeriesSetLabelsPercentage
      (foreign-lambda void "chartSeriesSetLabelsPercentage")))
  (begin
    (define chartSeriesSetLabelsNumFormat
      (foreign-lambda void "chartSeriesSetLabelsNumFormat" c-string)))
  (begin
    (define chartSeriesSetLabelsFont
      (foreign-lambda void "chartSeriesSetLabelsFont" c-string)))
  (begin
    (define chartSeriesSetTrendline
      (foreign-lambda void "chartSeriesSetTrendline" unsigned-byte unsigned-byte)))
  (begin
    (define chartSeriesSetTrendlineForecast
      (foreign-lambda void "chartSeriesSetTrendlineForecast" double double)))
  (begin
    (define chartSeriesSetTrendlineEquation
      (foreign-lambda void "chartSeriesSetTrendlineEquation")))
  (begin
    (define chartSeriesSetTrendlineRSquared
      (foreign-lambda void "chartSeriesSetTrendlineRSquared")))
  (begin
    (define chartSeriesSetTrendlineIntercept
      (foreign-lambda void "chartSeriesSetTrendlineIntercept" double)))
  (begin
    (define chartSeriesSetTrendlineName
      (foreign-lambda void "chartSeriesSetTrendlineName" c-string)))
  (begin
    (define chartSeriesSetTrendlineLine
      (foreign-lambda void "chartSeriesSetTrendlineLine")))
  (begin
    (define chartSeriesGetErrorBars
      (foreign-lambda void "chartSeriesGetErrorBars" unsigned-byte)))
  (begin
    (define chartSeriesSetErrorBars
      (foreign-lambda
        void
        "chartSeriesSetErrorBars"
        unsigned-byte
        double
        unsigned-short)))
  (begin
    (define chartSeriesSetErrorBarsDirection
      (foreign-lambda void "chartSeriesSetErrorBarsDirection" unsigned-byte)))
  (begin
    (define chartSeriesSetErrorBarsEndCap
      (foreign-lambda void "chartSeriesSetErrorBarsEndCap" unsigned-byte)))
  (begin
    (define chartSeriesSetErrorBarsLine
      (foreign-lambda void "chartSeriesSetErrorBarsLine")))
  (begin
    (define chartToggleMajorGridLinesXAxis
      (foreign-lambda void "chartToggleMajorGridLinesXAxis" unsigned-short)))
  (begin
    (define chartToggleMajorGridLinesYAxis
      (foreign-lambda void "chartToggleMajorGridLinesYAxis" unsigned-short)))
  (begin
    (define chartToggleMinorGridLinesXAxis
      (foreign-lambda void "chartToggleMinorGridLinesXAxis" unsigned-short)))
  (begin
    (define chartToggleMinorGridLinesYAxis
      (foreign-lambda void "chartToggleMinorGridLinesYAxis" unsigned-short)))
  (begin
    (define chartXAxisSetName
      (foreign-lambda void "chartXAxisSetName" c-string)))
  (begin
    (define chartYAxisSetName
      (foreign-lambda void "chartYAxisSetName" c-string)))
  (begin
    (define chartXAxisSetNameRange
      (foreign-lambda
        void
        "chartXAxisSetNameRange"
        c-string
        unsigned-integer32
        unsigned-short)))
  (begin
    (define chartYAxisSetNameRange
      (foreign-lambda
        void
        "chartYAxisSetNameRange"
        c-string
        unsigned-integer32
        unsigned-short)))
  (begin
    (define chartXAxisSetNameFont
      (foreign-lambda void "chartXAxisSetNameFont")))
  (begin
    (define chartYAxisSetNameFont
      (foreign-lambda void "chartYAxisSetNameFont")))
  (begin
    (define chartXAxisSetNumFont (foreign-lambda void "chartXAxisSetNumFont")))
  (begin
    (define chartYAxisSetNumFont (foreign-lambda void "chartYAxisSetNumFont")))
  (begin
    (define chartXAxisSetNumFormat
      (foreign-lambda void "chartXAxisSetNumFormat" c-string)))
  (begin
    (define chartYAxisSetNumFormat
      (foreign-lambda void "chartYAxisSetNumFormat" c-string)))
  (begin (define chartXAxisSetLine (foreign-lambda void "chartXAxisSetLine")))
  (begin (define chartYAxisSetLine (foreign-lambda void "chartYAxisSetLine")))
  (begin (define chartXAxisSetFill (foreign-lambda void "chartXAxisSetFill")))
  (begin (define chartYAxisSetFill (foreign-lambda void "chartYAxisSetFill")))
  (begin
    (define chartXAxisSetPattern (foreign-lambda void "chartXAxisSetPattern")))
  (begin
    (define chartYAxisSetPattern (foreign-lambda void "chartYAxisSetPattern")))
  (begin (define chartReverseXAxis (foreign-lambda void "chartReverseXAxis")))
  (begin (define chartReverseYAxis (foreign-lambda void "chartReverseYAxis")))
  (begin
    (define chartXAxisSetCrossing
      (foreign-lambda void "chartXAxisSetCrossing" double)))
  (begin
    (define chartYAxisSetCrossing
      (foreign-lambda void "chartYAxisSetCrossing" double)))
  (begin
    (define chartXAxisSetCrossingMax
      (foreign-lambda void "chartXAxisSetCrossingMax")))
  (begin
    (define chartYAxisSetCrossingMax
      (foreign-lambda void "chartYAxisSetCrossingMax")))
  (begin (define chartXAxisOff (foreign-lambda void "chartXAxisOff")))
  (begin (define chartYAxisOff (foreign-lambda void "chartYAxisOff")))
  (begin
    (define chartXAxisSetPosition
      (foreign-lambda void "chartXAxisSetPosition" unsigned-byte)))
  (begin
    (define chartYAxisSetPosition
      (foreign-lambda void "chartYAxisSetPosition" unsigned-byte)))
  (begin
    (define chartXAxisSetLabelPosition
      (foreign-lambda void "chartXAxisSetLabelPosition" unsigned-byte)))
  (begin
    (define chartYAxisSetLabelPosition
      (foreign-lambda void "chartYAxisSetLabelPosition" unsigned-byte)))
  (begin
    (define chartXAxisSetLabelAlign
      (foreign-lambda void "chartXAxisSetLabelAlign" unsigned-byte)))
  (begin
    (define chartYAxisSetLabelAlign
      (foreign-lambda void "chartYAxisSetLabelAlign" unsigned-byte)))
  (begin
    (define chartXAxisSetMin (foreign-lambda void "chartXAxisSetMin" double)))
  (begin
    (define chartXAxisSetMax (foreign-lambda void "chartXAxisSetMax" double)))
  (begin
    (define chartYAxisSetMin (foreign-lambda void "chartYAxisSetMin" double)))
  (begin
    (define chartYAxisSetMax (foreign-lambda void "chartYAxisSetMax" double)))
  (begin
    (define chartXAxisSetLogBase
      (foreign-lambda void "chartXAxisSetLogBase" unsigned-short)))
  (begin
    (define chartYAxisSetLogBase
      (foreign-lambda void "chartYAxisSetLogBase" unsigned-short)))
  (begin
    (define chartXAxisSetMajorTickPoint
      (foreign-lambda void "chartXAxisSetMajorTickPoint" unsigned-byte)))
  (begin
    (define chartYAxisSetMajorTickPoint
      (foreign-lambda void "chartYAxisSetMajorTickPoint" unsigned-byte)))
  (begin
    (define chartXAxisSetMinorTickPoint
      (foreign-lambda void "chartXAxisSetMinorTickPoint" unsigned-byte)))
  (begin
    (define chartYAxisSetMinorTickPoint
      (foreign-lambda void "chartYAxisSetMinorTickPoint" unsigned-byte)))
  (begin
    (define chartXAxisSetIntervalUnit
      (foreign-lambda void "chartXAxisSetIntervalUnit" unsigned-short)))
  (begin
    (define chartYAxisSetIntervalUnit
      (foreign-lambda void "chartYAxisSetIntervalUnit" unsigned-short)))
  (begin
    (define chartXAxisSetIntervalUnitTick
      (foreign-lambda void "chartXAxisSetIntervalUnitTick" unsigned-short)))
  (begin
    (define chartYAxisSetIntervalUnitTick
      (foreign-lambda void "chartYAxisSetIntervalUnitTick" unsigned-short)))
  (begin
    (define chartXAxisSetMajorUnit
      (foreign-lambda void "chartXAxisSetMajorUnit" double)))
  (begin
    (define chartYAxisSetMajorUnit
      (foreign-lambda void "chartYAxisSetMajorUnit" double)))
  (begin
    (define chartXAxisSetMinorUnit
      (foreign-lambda void "chartXAxisSetMinorUnit" double)))
  (begin
    (define chartYAxisSetMinorUnit
      (foreign-lambda void "chartYAxisSetMinorUnit" double)))
  (begin
    (define chartXAxisSetDisplayUnits
      (foreign-lambda void "chartXAxisSetDisplayUnits" unsigned-byte)))
  (begin
    (define chartYAxisSetDisplayUnits
      (foreign-lambda void "chartYAxisSetDisplayUnits" unsigned-byte)))
  (begin
    (define chartXAxisSetDisplayUnitsVisible
      (foreign-lambda void "chartXAxisSetDisplayUnitsVisible" unsigned-byte)))
  (begin
    (define chartYAxisSetDisplayUnitsVisible
      (foreign-lambda void "chartYAxisSetDisplayUnitsVisible" unsigned-byte)))
  (begin
    (define chartXAxisSetMajorGridlinesVisible
      (foreign-lambda void "chartXAxisSetMajorGridlinesVisible" unsigned-byte)))
  (begin
    (define chartYAxisSetMajorGridlinesVisible
      (foreign-lambda void "chartYAxisSetMajorGridlinesVisible" unsigned-byte)))
  (begin
    (define chartXAxisSetMinorGridlinesVisible
      (foreign-lambda void "chartXAxisSetMinorGridlinesVisible" unsigned-byte)))
  (begin
    (define chartYAxisSetMinorGridlinesVisible
      (foreign-lambda void "chartYAxisSetMinorGridlinesVisible" unsigned-byte)))
  (begin
    (define chartXAxisSetMajorGridlinesSetLine
      (foreign-lambda void "chartXAxisSetMajorGridlinesSetLine")))
  (begin
    (define chartYAxisSetMajorGridlinesSetLine
      (foreign-lambda void "chartYAxisSetMajorGridlinesSetLine")))
  (begin
    (define chartXAxisSetMinorGridlinesSetLine
      (foreign-lambda void "chartXAxisSetMinorGridlinesSetLine")))
  (begin
    (define chartYAxisSetMinorGridlinesSetLine
      (foreign-lambda void "chartYAxisSetMinorGridlinesSetLine")))
  (begin
    (define chartSeriesSetNameRange
      (foreign-lambda
        void
        "chartSeriesSetNameRange"
        c-string
        unsigned-integer32
        unsigned-short)))
  (begin
    (define chartXAxisSetReverse (foreign-lambda void "chartXAxisSetReverse")))
  (begin
    (define chartYAxisSetReverse (foreign-lambda void "chartYAxisSetReverse")))
  (begin
    (define chartXAxisSetCrossingOff
      (foreign-lambda void "chartXAxisSetCrossingOff")))
  (begin
    (define chartYAxisSetCrossingOff
      (foreign-lambda void "chartYAxisSetCrossingOff")))
  (begin (define closeWorkbook (foreign-lambda void "closeWorkbook"))))]