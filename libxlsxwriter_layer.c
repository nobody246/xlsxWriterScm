
void createWorkbook(char* workbookName)
{
  workbook = workbook_new(workbookName);
}

void formatsCleanup()
{
  if (formats)
    {free(formats);}
  maxAllowedFormats = 0;
  formatIndex = 0;
  formatCount = 0;
}


void rowNumberCleanup()
{
  if (rowNumbers)
    {free(rowNumbers);}
  maxAllowedRowNumbers = 0;
}

void colNumberCleanup()
{
  if (colNumbers)
    {free(colNumbers);}
  maxAllowedColNumbers = 0;
}

void dataValidationListsCleanup()
{
  if (dataValidationList)
    {free(dataValidationList);}
}

void dataValidationsCleanup()
{
  if (dataValidations)
    {free(dataValidations);}
  maxAllowedDataValidations = 0;
  dataValidationCount = 0;
  dataValidationIndex = 0;
}

void chartPointsCleanup()
{
  if (chartPoints)
    {free(chartPoints);}
  maxAllowedChartPoints = 0;
  chartPointCount = 0;
  chartPointIndex = 0;
}

void chartFillsCleanup()
{
  if (chartFillCount)
    {
      int i, x;
      for (i = 0; i < chartFillCount; i++)
	{
	  lxw_chart_fill* f = &chartFills[i];
	  if (!f) continue;              
	  for (x = 0; x < chartPointCount; x++)
	    {
	      lxw_chart_point* p = &chartPoints[x];
	      if (p->fill == f)
                {
                  printf("Chart fill # %d is referenced in chart point objects, to free chart fills, "
                         "free the chart point objects (chart-points-cleanup) first, and call "
                         "(chart-fills-cleanup) again.\n",
                         i);
                  return;
                }
	    }
	}
      free(chartFills);
    }
  maxAllowedChartFills = 0;
  chartFillCount = 0;
  chartFillIndex = 0;
}
     
void chartPatternsCleanup()
{
  int i, x;
  for (i = 0; i < chartPatternCount; i++)
    {
      lxw_chart_pattern* pt = &chartPatterns[i];
      if (!pt) continue;
      for (x = 0; x < chartPointCount; x++)
	{
	  lxw_chart_point* p = &chartPoints[x];
	  if (p->pattern == pt)
	    {
	      printf("Chart pattern # %d is referenced in chart point objects, to free chart patterns, "
                     "free the chart point objects (chart-points-cleanup) first, and call "
                     "(chart-patterns-cleanup) again.\n",
		     i);
	      return;
	    }
	}
    }
  if (chartPatternCount)
    {free(chartPatterns);}
  maxAllowedChartPatterns = 0;
  chartPatternCount = 0;
  chartPatternIndex = 0;
}
     
void chartLinesCleanup()
{
  int i, x;
  for (i = 0; i < chartLineCount; i++)
    {
      lxw_chart_line* l = &chartLines[i];
      if (!l) continue;
      for (x = 0; x < chartPointCount; x++)
	{
	  lxw_chart_point* p = &chartPoints[x];
	  if (p->line == l)
	    {
	      printf("Chart line # %d is referenced in chart point objects, to free chart lines, "
                     "free the chart point objects (chart-points-cleanup) first, and call "
                     "(chart-lines-cleanup) again.\n",
		     i);
	      return;
	    }
	}
    }
  if (chartLineCount)
    {free(chartLines);}
  maxAllowedChartLines = 0;
  chartLineCount = 0;
  chartLineIndex = 0;
}     

void seriesCleanup()
{
  if (series)
    {free(series);}
  maxAllowedSeries = 0;
  seriesCount = 0;
  seriesIndex = 0;
}

void chartsCleanup()
{
  if (charts)
    {free(charts);}
  maxAllowedCharts = 0;
  chartCount = 0;
  chartIndex = 0;
}

void seriesErrorBarsCleanup()
{
  if (seriesErrorBars)
    {free(seriesErrorBars);}
  maxAllowedSeriesErrorBars = 0;
  seriesErrorBarsCount = 0;
  seriesErrorBarsIndex = 0;
}

void chartFontsCleanup()
{
  if (chartFonts)
    {free(chartFonts);}
  maxAllowedChartFonts = 0;
  chartFontCount = 0;
  chartFontIndex = 0;
}

void debugFormat(lxw_format* f)
{
}

void debugFormats()
{
  if (!formatCount)
    {return;}
  int i;
  for (i =0; i<formatCount; i++)
    {
      debugFormat(&formats[i]);
    }
}

void debugChartPattern (lxw_chart_pattern* p)
{
  if (p->fg_color)
    {
      printf("Pattern FG Color Set To: %lx, (uint32_t) %lu\n", 
	     (uint32_t) p->fg_color,
	     (uint32_t) p->fg_color);
    }
  if (p->bg_color)
    {
      printf("Pattern BG Color Set To: %lx, (int32) %lu\n",
	     (uint32_t) p->bg_color,
	     (uint32_t) p->bg_color);
    }
  printf("Pattern Type Set To: %u\n", p->type);
  printf("===========\n");
 
}
     
void debugChartLine (lxw_chart_line* l)
{
  if (l->color)
    {
      printf("Color Set To:  %lx, (int32) %lu\n", 
	     (uint32_t) l->color, 
	     (uint32_t) l->color);
    }
  if (l->none)
    {printf("`none` Set To: %u\n", l->none);}
  if (l->width)
    {printf("Width Set To: %f\n", l->width);}
  if (l->dash_type)
    {printf("Dash type Set To: %u\n", l->dash_type);}
  printf("===========\n");
}

void initFormats (ptrdiff_t allocateN)
{
  formatsCleanup();
  formats = calloc(allocateN, sizeof(lxw_format));
  maxAllowedFormats = allocateN;
}

void initRowNumbers (ptrdiff_t allocateN)
{
  rowNumberCleanup();
  rowNumbers = calloc(allocateN, sizeof(lxw_row_t));
  maxAllowedRowNumbers = allocateN;
}

void initColNumbers (ptrdiff_t allocateN)
{
  colNumberCleanup();
  colNumbers = calloc(allocateN, sizeof(lxw_col_t));
  maxAllowedColNumbers = allocateN;
}

void initDataValidations (ptrdiff_t allocateN)
{
  dataValidationsCleanup();
  dataValidations = calloc(allocateN, sizeof(lxw_data_validation));
  maxAllowedDataValidations = allocateN;
}

void initChartFills (ptrdiff_t allocateN)
{
  chartFillsCleanup();
  chartFills = calloc(allocateN, sizeof(lxw_chart_fill));
  maxAllowedChartFills = allocateN;
}

void initChartPoints (ptrdiff_t allocateN)
{
  chartPointsCleanup();
  chartPoints = calloc(allocateN, sizeof(lxw_chart_point));
  maxAllowedChartPoints = allocateN;
}

void initChartLines (ptrdiff_t allocateN)
{
  chartLinesCleanup();
  chartLines = calloc(allocateN, sizeof(lxw_chart_line));
  maxAllowedChartLines = allocateN;
}

void initChartPatterns (ptrdiff_t allocateN)
{
  chartPatternsCleanup();
  chartPatterns = calloc(allocateN, sizeof(lxw_chart_pattern));
  maxAllowedChartPatterns = allocateN;
}

void initSeries(ptrdiff_t allocateN)
{
  seriesCleanup();
  series = calloc(allocateN, sizeof(lxw_chart_series));
  maxAllowedSeries = allocateN;
}

void initCharts(ptrdiff_t allocateN)
{
  chartsCleanup();
  charts = calloc(allocateN, sizeof(lxw_chart));
  maxAllowedCharts = allocateN;
}

void initChartFonts(ptrdiff_t allocateN)
{
  chartFontsCleanup();
  chartFonts = callocl(allocateN, sizeof(lxw_chart_font));
  maxAllowedChartFonts = allocateN;
}

void createDataValidationList (char* dvl)
{
  char* r;
  dataValidationList = calloc(255, sizeof(char));
  char d[1]; d[0] = ',';
  *(dataValidationList) = strtok_r(dvl, d, &r);
  int chCnt = 0,
    x = 0,
    i = 0;
  //data validation list is restricted by Excel to 255 characters, including comma separators.
  char* currStr;
  while (currStr = strtok_r(NULL, d, &r))
    {
      if (chCnt >= 254)
	{break;}
      int c = 0;
      while(currStr[c])
	{c++;chCnt++;}
      if ((254 - chCnt) >= 0)
	{*(dataValidationList + i) = currStr;}
      else
	{
	  x = chCnt - 254;
	  c -= x;
	  chCnt -= x;
	  char *v = calloc(c, sizeof(char));
	  strncpy(v, currStr, c);
	  *(dataValidationList + i) = v;
	  *(dataValidationList + (++i)) = NULL;
	  break;
	}
      chCnt++;//inc for comma separator.
      i++;
    }
  if (chCnt < 255)
    {*(dataValidationList + (++i)) = NULL;}
}

void createChartPoint(uint none,
		      double width)
{
  if (chartPointCount < maxAllowedChartPoints
      && (chartFillIndex < maxAllowedChartFills && chartFillIndex < chartFillCount)
      && (chartLineIndex < maxAllowedChartLines && chartLineIndex < chartLineCount)
      && (chartPatternIndex < maxAllowedChartPatterns && chartPatternIndex < chartPatternCount))
    {
      *(chartPoints + chartPointCount) = (lxw_chart_point){.fill =    &chartFills[chartFillIndex],
							   .line =    &chartLines[chartLineIndex],
							   .pattern = &chartPatterns[chartPatternIndex]};
      chartPointCount++;
    }
  else
    {printf("Referencing invalid chart element indices (create-chart-point)\n");}
}

void createChartFill(int32_t color,
		     uint8_t none,
		     uint8_t transparency)
{
  if (chartFillCount < maxAllowedChartFills)
    {
      *(chartFills + chartFillCount) = (lxw_chart_fill) {.color = color, 
							 .none = none, 
							 .transparency = transparency};
      chartFillCount++;
    }
  else
    {printf("You are trying to allocate more chart fills than alotted.\n");}
}

void createChartLine(int32_t color,
		     uint8_t none,
		     float width,
		     uint8_t dashType)
{
  if (chartLineCount < maxAllowedChartLines)
    {
      *(chartLines + chartLineCount) = (lxw_chart_line){.color     = color,
							.none      = none,
							.width     = width,
							.dash_type = dashType};
      chartLineCount++;
    }
  else
    {
      printf("Trying to allocate more chart lines than alotted, or referencing invalid "
	     "chartFillIndex (create-chart-line)\n");
    }
}

void createChartPattern(int32_t fgColor,
			int32_t bgColor,
			uint8_t ty)
{
  if (chartPatternCount < maxAllowedChartPatterns)
    {
      *(chartPatterns + chartPatternCount) = (lxw_chart_pattern){.fg_color = fgColor,
								 .bg_color = bgColor,
								 .type     = ty};
      chartPatternCount++;
    }
  else
    {
      printf("Trying to allocate more chart patterns than alotted, or referencing invalid "
	     "chartPatternIndex (create-chart-pattern)\n");
    }
}

void createChartFont(char* name,
		     double size,
		     uint8_t bold,
		     uint8_t italic,
		     uint8_t underline,
		     uint8_t rotation,
		     uint8_t color,
		     uint8_t pitchFamily,
		     uint8_t charset,
		     int8_t baseline)
{
  if (chartFontCount < maxAllowedChartFonts)
    {
      *(chartFonts + chartFontCount) =
	(lxw_chart_font)
	{.name = name,
	 .size = size,
	 .bold = bold,
	 .italic = italic,
	 .underline = underline,
	 .rotation = rotation,
	 .color = color,
	 .pitch_family = pitchFamily,
	 .charset = charset,
	 .baseline = baseline};
      chartFontIndex = chartFontCount;
      chartFontCount++;
    }
}
		     

void addToRowNumberList(uint32_t val)
{
  if (rowNumberCount < maxAllowedRowNumbers)
    {*(rowNumbers + rowNumberCount) = (lxw_row_t) val;}
}

void addToColNumberList(unsigned short val)
{
  if (colNumberCount < maxAllowedColNumbers)
    {*(colNumbers + colNumberCount) = (lxw_col_t) val;}
}

void debugDataValidationList()
{
  int i = 0;
  while(i < 255 && dataValidationList[i])
    {
      printf("%s %d", dataValidationList[i], i);
      i++;
    }
}

void debugDataValidation(lxw_data_validation* dv)
{
  printf("validate: %u \n"
         "criteria: %u \n"
         "ignore_blank: %u \n"
         "show_input: %u \n"
         "show_error: %u \n"
         "error_type: %u \n"
         "dropdown: %u \n"
         "value_number: %f \n"
         "value_formula: %s \n"
         "value_datetime_year: %d \n"
         "value_datetime_month: %d \n"
         "value_datetime_day: %d \n"
         "value_datetime_min: %d \n"
         "value_datetime_sec: %f \n"
         "minimum_number: %f \n"
         "minimum_formula: %s \n"
         "minimum_datetime_year: %d \n"
         "minimum_datetime_month: %d \n"
         "minimum_datetime_day: %d \n"
         "minimum_datetime_min: %d \n"
         "minimum_datetime_sec: %f \n"
         "maximum_number: %f \n"
         "maximum_formula: %s \n"
         "maximum_datetime_year: %d \n"
         "maximum_datetime_month: %d \n"
         "maximum_datetime_day: %d \n"
         "maximum_datetime_min: %d \n"
         "maximum_datetime_sec: %f \n"
         "input_title: %s \n"
         "input_message: %s \n"
         "error_title: %s \n"
         "error_message: %s ",
	 dv->validate,
	 dv->criteria,
	 dv->ignore_blank,
	 dv->show_input,
	 dv->show_error,
	 dv->error_type,
	 dv->dropdown,
	 dv->value_number,
	 dv->value_formula,
	 dv->value_datetime.year,
	 dv->value_datetime.month,
	 dv->value_datetime.day,
	 dv->value_datetime.min,
	 dv->value_datetime.sec,
	 dv->minimum_number,
	 dv->minimum_formula,
	 dv->minimum_datetime.year,
	 dv->minimum_datetime.month,
	 dv->minimum_datetime.day,
	 dv->minimum_datetime.min,
	 dv->minimum_datetime.sec,
	 dv->maximum_number,
	 dv->maximum_formula,
	 dv->maximum_datetime.year,
	 dv->maximum_datetime.month,
	 dv->maximum_datetime.day,
	 dv->maximum_datetime.min,
	 dv->maximum_datetime.sec,
	 dv->input_title,
	 dv->input_message,
	 dv->error_title,
	 dv->error_message);
}

void debugDataValidations()
{
  int i;
  for (i = 0; i < dataValidationCount; i++)
    {
      lxw_data_validation* v = &dataValidations[i];
      debugDataValidation(v);
    }
}

void debugChartPoints()
{
  int i;
  for (i = 0; i < chartPointCount; i++)
    {
      lxw_chart_fill* a = chartPoints[i].fill;
      printf("============================== %s \n %d",
	     "Processing Chart Point: \n", 
	     i);
      if (chartPoints[i].line)
	{debugChartLine(chartPoints[i].line);}
      if (chartPoints[i].fill)
	{debugChartFill(chartPoints[i].fill);}
      if (chartPoints[i].pattern)
	{debugChartPattern(chartPoints[i].pattern);}
    }
}

void debugChartFill (lxw_chart_fill* c)
{
  if (c->color)
    {
      printf("Fill Color Set To: %lu, (uint32_t) %ld\n", 
	     (uint32_t) c->color,
	     (uint32_t) c->color);
    }
  if (c->none)
    {printf("Fill `None` Value Set To: %u\n", c->none);}
  if (c->transparency)
    {printf("Fill `None` Value Set To: %u\n", c->transparency);}
  printf("===========\n");
}

void debugChartFills()
{
  int i;
  for (i = 0; i < chartFillCount; i++)
    {debugChartFill(&chartFills[i]);}
}

void debugChartLines()
{
  int i;
  for (i = 0; i < chartLineCount; i++)
    {debugChartLine(&chartLines[i]);}
}


void debugChartPatterns()
{
  int i;
  for (i = 0; i < chartPatternCount; i++)
    {debugChartPattern(&chartPatterns[i]);}
}

void setDataValidationIndex(ptrdiff_t ind)
{
  if (ind < dataValidationCount)
    {dataValidationIndex = ind;}
}

void setChartPointIndex(ptrdiff_t ind)
{
  if (ind < maxAllowedChartPoints)
    {chartPointIndex = ind;}
}

void setChartFillIndex(ptrdiff_t ind)
{
  if (ind < chartFillCount)
    {chartFillIndex = ind;}
}

void setChartLineIndex(ptrdiff_t ind)
{
  if (ind < chartLineCount)
    {chartLineIndex = ind;}
}

void setChartPatternIndex(ptrdiff_t ind)
{
  if (ind < chartPatternCount)
    {chartPatternIndex = ind;}
}

void setFormatIndex(ptrdiff_t ind)
{
  if (ind < formatCount)
    {formatIndex = ind;}
}

void setSeriesIndex(ptrdiff_t ind)
{
  if (ind < seriesCount)
    {seriesIndex = ind;}
}

void setChartIndex(ptrdiff_t ind)
{
  if (ind < chartCount)
    {chartIndex = ind;}
}

void setSeriesErrorBarsIndex(ptrdiff_t ind)
{
  if (ind < seriesErrorBarsCount)
    {seriesErrorBarsIndex = ind;}
}

void setChartFontIndex(ptrdiff_t ind)
{
  if (ind < chartFontCount)
    {chartFontIndex = ind;}
}

void addWorksheet(char* worksheetName)
{
  if (workbook)
    {worksheet = workbook_add_worksheet(workbook, worksheetName);}
}

void worksheetWriteString(char* value)
{
  if (worksheet)
    {
      lxw_format *f = NULL;
      if (formats)
	{f = formats[formatIndex];}
      handleLXWError(worksheet_write_string(worksheet, row, col, value, f), 
		     "worksheet-write-string");
    }
}


void worksheetWriteNumber(float value)
{
  if (worksheet)
    {
      lxw_format *f = NULL;
      if (formats)
	{f = formats[formatIndex];}
      handleLXWError(worksheet_write_number(worksheet, row, col, value, f), 
		     "worksheet-write-number");
    }
}

void worksheetWriteFormula(char* formula,
			   uint32_t row,
			   unsigned short col)
{
  if (worksheet)
    {
      lxw_format *f = NULL;
      if (formats)
	{f = formats[formatIndex];}
      handleLXWError(worksheet_write_formula(worksheet, row, col, formula, f), 
		     "worksheet-write-formula");
    }
}


void worksheetWriteArrayFormula(char *formula,
				uint32_t stRow,
				unsigned short stCol,
				uint32_t endRow,
			        unsigned short endCol)
{
  if (worksheet)
    {
      lxw_format *f = NULL;
      if (formats)
	{f = formats[formatIndex];}
      handleLXWError(worksheet_write_array_formula(worksheet, 
						   stRow,
						   stCol,
						   endRow,
						   endCol,
						   formula, 
						   f), 
		     "worksheet-write-formula");
    }
}


void worksheetWriteDatetime(int year,
			    int month,
			    int day,
			    int hour,
			    int mn,
			    double sec,
			    uint32_t row,
			    unsigned short col)
{
  if (worksheet)
    {
      lxw_format *f = NULL;
      if (formats)
	{f = formats[formatIndex];}
      lxw_datetime dt = {year, month, day, hour, mn, sec};
      handleLXWError(worksheet_write_datetime(worksheet, row, col, &dt, f), 
		     "worksheet-write-datetime");
    }
}


void worksheetWriteURL(char* url,
		       uint32_t row,
		       unsigned short col)
{
  if (worksheet)
    {
      lxw_format *f = NULL;
      if (formats)
	{f = formats[formatIndex];}
      handleLXWError(worksheet_write_url(worksheet, row, col, url, f), 
		     "worksheet-write-url");
    }
}

void worksheetWriteBoolean(short value,
			   uint32_t row,
			   unsigned short col)
{
  if (worksheet)
    {
      lxw_format *f = NULL;
      if (formats)
	{f = formats[formatIndex];}
      handleLXWError(worksheet_write_boolean(worksheet, row, col, value, f), 
		     "worksheet-write-boolean");
    }
}

void worksheetWriteBlank(uint32_t row,
			 unsigned short col)
{
  if (worksheet)
    {
      lxw_format *f = NULL;
      if (formats)
	{f = formats[formatIndex];}
      handleLXWError(worksheet_write_blank(worksheet, row, col, f), 
		     "worksheet-write-blank");
    }
}

void worksheetSetRow(double height,
		     uint32_t row)
{
  if (worksheet)
    {
      lxw_format *f = NULL;
      if (formats)
	{f = formats[formatIndex];}
      handleLXWError(worksheet_set_row(worksheet, row, height, f), 
		     "worksheet-set-row");
    }
}


void worksheetSetRowOpt(unsigned char hidden,
			unsigned char level,
			unsigned char collapsed,
			double height,
			uint32_t row)
{
  if (worksheet)
    {
      lxw_format *f = NULL;
      if (formats)
	{f = formats[formatIndex];}
      lxw_row_col_options options = {.hidden = hidden, 
				     .level  = level, 
				     .collapsed = collapsed};
      handleLXWError(worksheet_set_row_opt(worksheet, row, height, f, &options), 
		     "worksheet-set-row-opt");
    }
}


void worksheetSetColumn(double width,
			unsigned short stCol,
			unsigned short endCol)
{
  if (worksheet)
    {
      lxw_format *f = NULL;
      if (formats)
	{f = formats[formatIndex];}
      handleLXWError(worksheet_set_column(worksheet, stCol, endCol, width, f),
		     "worksheet-set-column");
    }
}

void worksheetSetColumnOpt(unsigned char hidden,
			   unsigned char level,
			   unsigned char collapsed,
			   double width,
			   unsigned short stCol,
			   unsigned short endCol)
{
  if (worksheet)
    {
      lxw_format *f = NULL;
      if (formats)
	{f = formats[formatIndex];}
      lxw_row_col_options options = {.hidden = hidden, 
				     .level  = level, 
				     .collapsed = collapsed};
      handleLXWError(worksheet_set_column_opt(worksheet, 
					      stCol, 
					      endCol, 
					      width, 
					      f, 
					      &options),
		     "worksheet-set-column-opt");
    }
}


void worksheetInsertImage(char* filename,
			  uint32_t row,
			  unsigned short col)
{
  if (worksheet)
    {
      handleLXWError(worksheet_insert_image(worksheet, row, col, filename), 
		     "worksheet-insert-image");
    }
}


void worksheetInsertImageOpt(int x_offset,
			     int y_offset,
			     double x_scale,
			     double y_scale,
			     char* filename,
			     uint32_t row,
			     unsigned short col)
{
  if (worksheet)
    {
      lxw_image_options options  = {.x_offset = x_offset,
				    .y_offset = y_offset,
				    .x_scale  = x_scale,
				    .y_scale  = y_scale};
      handleLXWError(worksheet_insert_image_opt(worksheet, 
						row, 
						col, 
						filename,
						&options),
		     "worksheet-insert-image-opt");
    }
}


void worksheetInsertChart(uint32_t row,
			  unsigned short col)
{
  if (worksheet
      && charts)
    {worksheet_insert_chart(worksheet, row, col, charts[chartIndex]);}
}

void worksheetInsertChartOpt(int x_offset,
			     int y_offset,
			     int x_scale,
			     int y_scale,
			     uint32_t row,
			     unsigned short col)
{
  if (worksheet
      && charts)
    {
      lxw_image_options options  = {.x_offset = x_offset,
				    .y_offset = y_offset,
				    .x_scale  = x_scale,
				    .y_scale  = y_scale};
      handleLXWError(worksheet_insert_chart_opt(worksheet, 
						row, 
						col,
						charts[chartIndex],
						&options),
		     "worksheet-insert-image-opt");
    }
}

void worksheetMergeRange(char* val,
			 uint32_t stRow,
			 unsigned short stCol,
			 uint32_t endRow,
			 unsigned short endCol)
			 
{
  if (worksheet && formats)
    {
      handleLXWError(worksheet_merge_range(worksheet,
					   stRow,
					   stCol,
					   endRow,
					   endCol,
					   val,
					   formats[formatIndex]),
		     "worksheet-merge-range");
    }
}

void worksheetAutoFilter(uint32_t stRow,
			 unsigned short stCol,
			 uint32_t endRow,
			 unsigned short endCol)
{
  if (worksheet)
    {
      handleLXWError(worksheet_autofilter(worksheet, stRow, stCol, endRow,  endCol),
		     "worksheet-autofilter");
    }
}

void worksheetDataValidationCell(uint32_t row,
				 unsigned short col)
{
  if (worksheet 
      && (dataValidationIndex < dataValidationCount))
    {
      handleLXWError(worksheet_data_validation_cell(worksheet, row, col, &dataValidations[dataValidationIndex]), 
		     "worksheet-data-validation-cell");
    }
}


void worksheetDataValidationRange(uint32_t stRow,
				  unsigned short stCol,
				  uint32_t endRow,
				  unsigned short endCol)
{
  if (worksheet 
      && (dataValidationIndex < dataValidationCount))
    {
      handleLXWError(worksheet_data_validation_range(worksheet,
						     stRow,
						     stCol,
						     endRow,
						     endCol,
						     &dataValidations[dataValidationIndex]),
		     "worksheet-data-validation-range");
    }
}

void worksheetActivate()
{
  if (worksheet)
    {worksheet_activate(worksheet);}
}

void worksheetSelect()
{
  if (worksheet)
    {worksheet_select(worksheet);}
}

void worksheetHide()
{
  if (worksheet)
    {worksheet_hide(worksheet);}
}

void worksheetSetFirstSheet()
{
  if (worksheet)
    {worksheet_set_first_sheet(worksheet);}
}

void worksheetSplitPanes(double vertical,
			 double horizontal)
{
  if (worksheet)
    {worksheet_split_planes(worksheet, vertical, horizontal);}
}

void worksheetSetSelection(uint32_t stRow,
			   unsigned short stCol,
			   uint32_t endRow,
			   unsigned short endCol)
{
  if (worksheet)
    {worksheet_set_selection(worksheet, stRow, stCol, endRow, endCol);}
}

void worksheetSetLandscape()
{
  if (worksheet)
    {worksheet_set_landscape(worksheet);}
}

void worksheetSetPortrait()
{
  if (worksheet)
    {worksheet_set_portrait(worksheet);}
}

void worksheetSetPageView()
{
  if (worksheet)
    {worksheet_set_page_view(worksheet);}
}

void worksheetSetPaper(unsigned char paper_type)
{
  if (worksheet)
    {worksheet_set_paper(worksheet, paper_type);}
}

void worksheetSetMargins(double left,
			 double right,
			 double top,
			 double bottom)
{
  if (worksheet)
    {worksheet_set_margins(worksheet, left, right, top, bottom);}
}

void worksheetSetHeader(char* headerStr)
{
  if (worksheet)
    {
      handleLXWError(worksheet_set_header(worksheet, headerStr),
		     "worksheet-set-header");
    }
}

void worksheetSetFooter(char* footerStr)
{
  if (worksheet)
    {
      handleLXWError(worksheet_set_footer(worksheet, footerStr),
		     "worksheet-set-footer");
    }
}

void worksheetSetHeaderOpt(char* headerStr,
			   double margin)
{
  if (worksheet)
    {
      lxw_header_footer_options o = (lxw_header_footer_options){.margin = margin};
      handleLXWError(worksheet_set_header_opt(worksheet, headerStr, &o), 
		     "worksheet-set-header-opt");
    }
}

void worksheetSetFooterOpt(char* footerStr,
			   double margin)
{
  if (worksheet)
    {
      lxw_header_footer_options o = (lxw_header_footer_options){.margin = margin};
      handleLXWError(worksheet_set_footer_opt(worksheet, footerStr, &o), 
		     "worksheet-set-footer-opt");
    }
}

void worksheetSetHPageBreaks()
{
  if (worksheet &&
      rowNumbers)
    {worksheet_set_h_pagebreaks(worksheet, colNumbers);}
}

void worksheetSetVPageBreaks()
{
  if (worksheet &&
      colNumbers)
    {worksheet_set_v_pagebreaks(worksheet, rowNumbers);}
}

void createFormat()
{
  //if (workbook)
  //  {format = workbook_add_format(workbook);}
  if (formatCount > maxAllowedFormats)
    {
      printf("trying to add more formats than alotted in init, in add-format.");
      return;
    }
  *(formats + formatCount) = workbook_add_format(workbook);
  formatIndex = formatCount;
  formatCount++;
}

void setCellNumFormat(char* numFormat)
{
  if (formats)
    {format_set_num_format(formats[formatIndex], numFormat);}
}

void setCellNumFormatIndex(uint8_t index)
{
  if (formats)
    {format_set_num_format_index(formats[formatIndex], index);}  
}

void formatHideFormula()
{
  if (formats)
    {format_set_hidden(formats[formatIndex]);}
}

void formatSetUnlocked()
{
  if (formats)
    {format_set_unlocked(formats[formatIndex]);}
}

void formatSetBold()
{
  if (formats)
    {format_set_bold(formats[formatIndex]);}
}

void formatSetItalic()
{
  if (formats)
    {format_set_italic(formats[formatIndex]);}
}

void setUnderlineSingle()
{
  if (formats)
    {format_set_underline(formats[formatIndex], LXW_UNDERLINE_SINGLE);}
}

void setUnderlineDouble()
{
  if (formats)
    {format_set_underline(formats[formatIndex], LXW_UNDERLINE_DOUBLE);}
}

void setUnderlineSingleAcct()
{
  if (formats)
    {format_set_underline(formats[formatIndex], LXW_UNDERLINE_SINGLE_ACCOUNTING);}
}

void setUnderlineDoubleAcct()
{
  if (formats)
    {format_set_underline(formats[formatIndex], LXW_UNDERLINE_DOUBLE_ACCOUNTING);}
}


void setStrikeout()
{
  if (formats)
    {format_set_font_strikeout(formats[formatIndex]);}
}

void setSuperscript()
{
  if (formats)
    {format_set_font_script(formats[formatIndex], LXW_FONT_SUPERSCRIPT);}
}

void setSubscript()
{
  if (formats)
    {format_set_font_script(formats[formatIndex], LXW_FONT_SUBSCRIPT);}
}

void setFontColor(uint32_t color)
{
  if (formats)
    {format_set_font_color(formats[formatIndex], color);}
}

void setFontName(char* name)
{
  if (formats)
    {format_set_font_name(formats[formatIndex], name);}
}

void setRotation(int degrees)
{
  if (formats &&
      (degrees >= -90 && degrees <= 270))
    {format_set_rotation(formats[formatIndex], degrees);}
}

void setIndentation(int level)
{
  if (formats)
    {format_set_indentation(formats[formatIndex], level);}
}

void setBold()
{
  if (formats)
    {format_set_bold(formats[formatIndex]);}
}

void setItalic()
{
  if (formats)
    {format_set_italic(formats[formatIndex]);}
}


void setShrink()
{
  if (formats)
    {format_set_shrink(formats[formatIndex]);}
}

void setPattern(uint8_t pattern)
{
  if (formats) //highest in pattern enum
    {format_set_pattern(formats[formatIndex], pattern);}
}

void setAlign(uint8_t align)
{
  if (formats
      && align <= (uint8_t) LXW_ALIGN_VERTICAL_DISTRIBUTED)//highest in align enum
    {format_set_align(formats[formatIndex], align);}
  else
    {printf("invalid alignment set in set-align.");}
}

void setBorder(uint8_t border)
{
  if (formats
      && border <= (uint8_t) LXW_BORDER_SLANT_DASH_DOT)
    {format_set_border(formats[formatIndex], border);}
  else
    {printf("invalid border type set in set-border.");}
}

void setBorderBottom(uint8_t border)
{
  if (formats 
      && border <= (uint8_t) LXW_BORDER_SLANT_DASH_DOT)
    {format_set_bottom(formats[formatIndex], border);}
  else
    {printf("invalid border type set in set-border-bottom.");}
}


void setBorderTop(uint8_t border)
{
  if (formats 
      && border <= (uint8_t) LXW_BORDER_SLANT_DASH_DOT)
    {format_set_top(formats[formatIndex], border);}
  else
    {printf("invalid border type set in set-border-top.");}
}


void setBorderLeft(uint8_t border)
{
  if (formats
      && border <= (uint8_t) LXW_BORDER_SLANT_DASH_DOT)
    {format_set_left(formats[formatIndex], border);}
  else
    {printf("invalid border type set in set-border-top.");}
}

void setBorderRight(uint8_t border)
{
  if (formats
      && border <= (uint8_t) LXW_BORDER_SLANT_DASH_DOT)
    {format_set_right(formats[formatIndex], border);}
  else
    {printf("invalid border type set in set-border-top.");}
}

void setBGColor(uint32_t color)
{
  if (formats)
    {format_set_bg_color(formats[formatIndex], color);}
}

void setFGColor(uint32_t color)
{
  if (formats)
    {format_set_fg_color(formats[formatIndex], color);}
}

void setBorderColor(uint32_t color)
{
  if (formats)
    {format_set_border_color(formats[formatIndex], color);}
}

void setBorderBottomColor(uint32_t color)
{
  if (formats)
    {format_set_bottom_color(formats[formatIndex], color);}
}

void setBorderTopColor(uint32_t color)
{
  if (formats)
    {format_set_top_color(formats[formatIndex], color);}
}

void setBorderLeftColor(uint32_t color)
{
  if (formats[formatIndex])
    {format_set_left_color(formats[formatIndex], color);}
}

void setBorderRightColor(uint32_t color)
{
  if (formats)
    {format_set_right_color(formats[formatIndex], color);}
}

void turnOnTextWrap()
{
  if (formats)
    {format_set_text_wrap(formats[formatIndex]);}
}

void setDocProperties(char* title,
		      char* subject,
		      char* author,
		      char* manager,
		      char* company,
		      char* category,
		      char* keywords,
		      char* comments,
		      char* status)
{
  if (workbook)
    {
      lxw_doc_properties properties = {
	.title    = title,
	.subject  = subject,
	.author   = author,
	.manager  = manager,
	.company  = company,
	.category = category,
	.keywords = keywords,
	.comments = comments,
	.status   = status 
      };
      handleLXWError(workbook_set_properties(workbook, &properties),
		     "set-doc-properties");
    }
}


void workbookSetCustomPropertyString(char* name,
				     char* value)
{
  if (workbook)
    {
      handleLXWError(workbook_set_custom_property_string(workbook, name, value), 
		     "workbook-set-custom-property-string");
    }
}


void workbookSetCustomPropertyNumber(char* name,
				     double value)
{
  if (workbook)
    {
      handleLXWError(workbook_set_custom_property_number(workbook, name, value), 
		     "workbook-set-custom-property-number");
    }
}

void workbookSetCustomPropertyBoolean(char* name,
				      unsigned char value)
{
  if (workbook)
    {
      handleLXWError(workbook_set_custom_property_boolean(workbook, name, value), 
		     "workbook-set-custom-property-number");
    }
}

void workbookSetCustomPropertyDatetime(char* name,
				       int year,
				       int month,
				       int day,
				       int hour,
				       int min,
				       double sec)
{
  lxw_datetime d = {year, month, day, hour, min, sec};
  if (workbook)
    {
      handleLXWError(workbook_set_custom_property_datetime(workbook, name, &d), 
		     "workbook-set-custom-property-datetime");
    }
}

void workbookDefineName(char* name, char* formula)
{
  if (workbook)
    {
      handleLXWError(workbook_define_name(workbook, name, formula),
		     "workbook-define-name");
    }
}

void workbookGetWorksheetByName(char* name)
{
  if (workbook)
    {worksheet = workbook_get_worksheet_by_name(workbook, name);}
}

void workbookValidateWorksheetName(char* name)
{
  if (workbook)
    {workbook_validate_worksheet_name(workbook, name);}
}

void setCol(unsigned short colNum)
{col = colNum;}

void setRow(uint32_t rowNum)
{row = rowNum;}

void setPos(uint32_t rowNum,
	    unsigned short colNum)
{
  row = rowNum;
  col = colNum;
}





/*chart functions*/

void createChart(unsigned short chartType)
{
  if (chartCount < maxAllowedCharts
      && charts
      && workbook)
    {
      *(charts + chartCount) = workbook_add_chart(workbook, (lxw_chart_type) chartType);
      chartIndex = chartCount;
      chartCount++;
    }
}

void createChartSeries(char* categories, char* vals)
{
  if (seriesCount < maxAllowedSeries &&
      charts)
    {
      *(series + seriesCount) = chart_add_series(&charts[chartIndex], categories, vals);
      seriesIndex = seriesCount;
      seriesCount++;
    }
}

void chartSeriesSetName(char* name)
{
  if (series && charts)
    {chart_series_set_name(series[seriesIndex], name);}
}

void chartSeriesSetLine()
{
  if (series &&
      chartLineIndex < chartLineCount)
    {chart_series_set_line(series[seriesIndex], &chartLines[chartLineIndex]);}
  else
    {printf("Invalid series and/or chart line index passed in (chart-series-set-line)");}
}


void chartSeriesSetFill(int chartFillIndex)
{
  if (series &&
      chartFillIndex < chartFillCount)
    {chart_series_set_fill(series[seriesIndex], &chartFills[chartFillIndex]);}
  else
    {printf("Invalid chart fill index or series does not exist");}
}

void chartSeriesSetPattern(uint8_t patternType,
			   uint32_t fgColor,
			   uint32_t bgColor)
{
  if (series)
    {
      lxw_chart_pattern pattern = {.type = patternType,
				   .fg_color = fgColor,
				   .bg_color = bgColor};
      chart_series_set_pattern(series[seriesIndex], &pattern);
    }
}

void chartSeriesSetMarkerType(unsigned char ty)
{
  if (series)
    {chart_series_set_marker_type(series[seriesIndex], ty);}
}

void chartSeriesSetMarkerSize(unsigned char size)
{
  if (series)
    {chart_series_set_marker_size(series[seriesIndex], size);}
}

void chartSeriesSetMarkerLine()
{
  if (series && chartLineIndex < chartLineCount)
    {chart_series_set_marker_line(series[seriesIndex], &chartLines[chartLineIndex]);}
  else
    {printf("series does not exist, or chartline index is invalid.");}
}

void chartSeriesSetMarkerFill(ptrdiff_t fillIndex)
{
  if (series && fillIndex < chartFillCount)
    {chart_series_set_marker_fill(series[seriesIndex], &chartFills[fillIndex ]);}
  else
    {printf("series does not exist, or chartline index is invalid");}
}

void chartSeriesSetMarkerPattern(uint8_t patternType,
				 uint32_t fgColor,
				 uint32_t bgColor)
{
  if (series)
    {
      lxw_chart_pattern pattern = {.type = patternType,
				   .fg_color = fgColor,
				   .bg_color = bgColor};
      chart_series_set_marker_pattern(series[seriesIndex], &pattern);
    }
}



void chartSeriesInvertIfNegative()
{
  if (series)
    {chart_series_set_invert_if_negative(series[seriesIndex]);}
}


void chartSeriesSetPoints()
{
  if (chartPoints &&
      series)
    {chart_series_set_points(series[seriesIndex], &chartPoints[chartPointIndex]);}
}

void chartSeriesSetSmooth(uint8_t smooth)
{
  if (series)
    {chart_series_set_smooth(series[seriesIndex], smooth);}
}

void chartSeriesSetLabels()
{
  if (series)
    {chart_series_set_labels(series[seriesIndex]);}
}

void chartSeriesSetLabelsOpt(uint8_t showName,
			     uint8_t showCategory,
			     uint8_t showValue)
{
  if (series)
    {
      chart_series_set_labels_options(series[seriesIndex],
				      showName,
				      showCategory,
				      showValue);
    }
}

void chartSeriesSetLabelsSeparator(uint8_t separator)
{
  if (series)
    {
      chart_series_set_labels_separator(series[seriesIndex],
					separator);
    }
}

void chartSeriesSetLabelsPosition(uint8_t position)
{
  if (series)
    {
      chart_series_set_labels_position(series[seriesIndex],
				       position);
    }
}

void chartSeriesSetLabelsLeaderLine()
{
  if (series)
    {chart_series_set_labels_leader_line(series[seriesIndex]);}
}

void chartSeriesSetLabelsLegend()
{
  if (series)
    {chart_series_set_labels_legend(series[seriesIndex]);}
}

void chartSeriesSetLabelsPercentage()
{
  if (series)
    {chart_series_set_labels_percentage(series[seriesIndex]);}
}

void chartSeriesSetLabelsNumFormat(char* numFormat)
{
  if (series)
    {
      chart_series_set_labels_num_format(series[seriesIndex],
					 numFormat);
    }
}

void chartSeriesSetLabelsFont(char* fontName)
{
  if (series)
    {
      chart_series_set_font_name(series[seriesIndex],
				 fontName);
    }
}


void chartSeriesSetTrendline(uint8_t ty,
			     uint8_t val)
{
  if (series)
    {
      chart_series_set_trendline(series[seriesIndex],
				 ty,
				 val);
    }
}

void chartSeriesSetTrendlineForecast(double forward,
				     double backward)
{
  if (series)
    {
      chart_series_set_trendline_forecast(series[seriesIndex],
					  forward,
					  backward);
    }
}

void chartSeriesSetTrendlineEquation()
{
  if (series)
    {chart_series_set_trendline_equation(series[seriesIndex]);}
}

void chartSeriesSetTrendlineRSquared()
{
  if (series)
    {chart_series_set_trendline_r_squared(series[seriesIndex]);}
}

void chartSeriesSetTrendlineIntercept(double intercept)
{
  if (series)
    {
      chart_series_set_trendline_intercept(series[seriesIndex],
					   intercept);
    }
}

void chartSeriesSetTrendlineName(char* name)
{
  if (series)
    {
      chart_series_set_trendline_name(series[seriesIndex],
				      name);
    }
}

void chartSeriesSetTrendlineLine()
{
  if (series
      && chartLines)
    {
      chart_series_set_trendline_line(series[seriesIndex],
				      &chartLines[chartLineIndex]);
    }
}

void chartSeriesGetErrorBars(uint8_t axisType)
{
  if (series)
    {
      *(seriesErrorBars + seriesErrorBarsCount) =
	(lxw_series_error_bars*) chart_series_get_error_bars(series[seriesIndex],
			 				     axisType);
      seriesErrorBarsIndex = seriesErrorBarsCount;
      seriesErrorBarsCount++;
    }
}

void chartSeriesSetErrorBars(uint8_t ty,
			     double val,
			     unsigned short horizontal)
{
  if (series)
    {
      chart_series_set_set_error_bars((horizontal > 0) ?
				      series[seriesIndex]->x_error_bars :
				      series[seriesIndex]->y_error_bars,
				      ty,
				      val);
    }
}

void chartSeriesSetErrorBarsDirection(uint8_t dir)
{
  if (seriesErrorBars)
    {
      set_chart_series_error_bars_direction(seriesErrorBars[seriesErrorBarsIndex],
					    dir);
    }
}

void chartSeriesSetErrorBarsEndCap(uint8_t endCap)
{
  if (seriesErrorBars)
    {
      set_chart_series_error_bars_endcap(seriesErrorBars[seriesErrorBarsIndex],
					 endCap);
    }
}

void chartSeriesSetErrorBarsLine()
{
  if (seriesErrorBars &&
      chartLines)
    {
      set_chart_series_error_bars_endcap(seriesErrorBars[seriesErrorBarsIndex],
					 chartLines[chartLineIndex]);
    }
}

void chartToggleMajorGridLinesXAxis(unsigned short showAxis)
{
  if (charts)
    {
      if (showAxis)
	{chart_axis_major_gridlines_set_visible(charts[chartIndex]->x_axis, LXW_TRUE);}
      else
	{chart_axis_major_gridlines_set_visible(charts[chartIndex]->x_axis, LXW_FALSE);}
    }
}

void chartToggleMajorGridLinesYAxis(unsigned short showAxis)
{
  if (charts)
    {
      if (showAxis)
	{chart_axis_major_gridlines_set_visible(charts[chartIndex]->y_axis, LXW_TRUE);}
      else
	{chart_axis_major_gridlines_set_visible(charts[chartIndex]->y_axis, LXW_FALSE);}
    }
}

void chartToggleMinorGridLinesXAxis(unsigned short showAxis)
{
  if (charts)
    {
      if (showAxis)
	{chart_axis_minor_gridlines_set_visible(charts[chartIndex]->x_axis, LXW_TRUE);}
      else
	{chart_axis_minor_gridlines_set_visible(charts[chartIndex]->x_axis, LXW_FALSE);}
    }
}

void chartToggleMinorGridLinesYAxis(unsigned short showAxis)
{
  if (charts)
    {
      if (showAxis)
	{chart_axis_minor_gridlines_set_visible(charts[chartIndex]->y_axis, LXW_TRUE);}
      else
	{chart_axis_minor_gridlines_set_visible(charts[chartIndex]->y_axis, LXW_FALSE);}
    }
}

void chartXAxisSetName(char* name)
{
  if (charts)
    {chart_axis_set_name(charts[chartIndex]->x_axis, name);}
}

void chartYAxisSetName(char* name)
{
  if (charts)
    {chart_axis_set_name(charts[chartIndex]->y_axis, name);}
}

void chartXAxisSetNameRange(char* sheetName,
			    uint32_t row,
			    unsigned short col)
{
  if (charts)
    {
      chart_axis_set_name_range(charts[chartIndex]->x_axis,
				sheetName,
				row,
				col);
    }
}

void chartYAxisSetNameRange(char* sheetName,
			    uint32_t row,
			    unsigned short col)
{
  if (charts)
    {
      chart_axis_set_name_range(charts[chartIndex]->y_axis,
				sheetName,
				row,
				col);
    }
}

void chartXAxisSetNameFont()
{
  if (charts &&
      chartFonts)
    {chart_axis_set_name_font(charts[chartIndex]->x_axis, &chartFonts[chartFontIndex]);}
}

void chartYAxisSetNameFont()
{
  if (charts &&
      chartFonts)
    {chart_axis_set_name_font(charts[chartIndex]->y_axis, &chartFonts[chartFontIndex]);}
}

void chartXAxisSetNumFont()
{
  if (charts &&
      chartFonts)
    {chart_axis_set_num_font(charts[chartIndex]->x_axis, &chartFonts[chartFontIndex]);}
}

void chartYAxisSetNumFont()
{
  if (charts &&
      chartFonts)
    {chart_axis_set_num_font(charts[chartIndex]->y_axis, &chartFonts[chartFontIndex]);}
}

void chartXAxisSetNumFormat(const char* format)
{
  if (charts &&
      chartFonts)
    {chart_axis_set_num_format(charts[chartIndex]->x_axis, format);}
}

void chartYAxisSetNumFormat(const char* format)
{
  if (charts &&
      chartFonts)
    {chart_axis_set_num_format(charts[chartIndex]->y_axis, format);}
}

void chartXAxisSetLine()
{
  if (charts &&
      chartLines)
    {chart_axis_set_num_format(charts[chartIndex]->x_axis, &chartLines[chartLineIndex]);}
}

void chartYAxisSetLine()
{
  if (charts &&
      chartLines)
    {chart_axis_set_num_format(charts[chartIndex]->y_axis, &chartLines[chartLineIndex]);}
}

void chartXAxisSetFill()
{
  if (charts &&
      chartFills)
    {chart_axis_set_fill(charts[chartIndex]->x_axis, &chartFills[chartFillIndex]);}
}

void chartYAxisSetFill()
{
  if (charts &&
      chartFills)
    {chart_axis_set_fill(charts[chartIndex]->y_axis, &chartFills[chartFillIndex]);}
}

void chartXAxisSetPattern()
{
  if (charts &&
      chartPatterns)
    {chart_axis_set_fill(charts[chartIndex]->x_axis, &chartPatterns[chartPatternIndex]);}
}

void chartYAxisSetPattern()
{
  if (charts &&
      chartPatterns)
    {chart_axis_set_fill(charts[chartIndex]->y_axis, &chartPatterns[chartPatternIndex]);}
}

void chartReverseXAxis()
{
  if (charts)
    {chart_axis_set_reverse(charts[chartIndex]->x_axis);}
}

void chartReverseYAxis()
{
  if (charts)
    {chart_axis_set_reverse(charts[chartIndex]->y_axis);}
}

void chartXAxisSetCrossing(double val)
{
  if (charts)
    {chart_axis_set_crossing(charts[chartIndex]->x_axis, val);}
}

void chartYAxisSetCrossing(double val)
{
  if (charts)
    {chart_axis_set_crossing(charts[chartIndex]->y_axis, val);}
}

void chartXAxisSetCrossingMax()
{
  if (charts)
    {chart_axis_set_crossing_max(charts[chartIndex]->x_axis);}
}

void chartYAxisSetCrossingMax()
{
  if (charts)
    {chart_axis_set_crossing_max(charts[chartIndex]->y_axis);}
}

void chartXAxisOff()
{
  if (charts)
    {chart_axis_off(charts[chartIndex]->x_axis);}
}

void chartYAxisOff()
{
  if (charts)
    {chart_axis_off(charts[chartIndex]->y_axis);}
}

void chartXAxisSetPosition(uint8_t pos)
{
  if (charts)
    {chart_axis_set_position(charts[chartIndex]->x_axis, pos);}
}

void chartYAxisSetPosition(uint8_t pos)
{
  if (charts)
    {chart_axis_set_position(charts[chartIndex]->x_axis, pos);}
}

void chartXAxisSetLabelPosition(uint8_t pos)
{
  if (charts)
    {chart_axis_set_position_label(charts[chartIndex]->x_axis, pos);}
}

void chartYAxisSetLabelPosition(uint8_t pos)
{
  if (charts)
    {chart_axis_set_position_label(charts[chartIndex]->x_axis, pos);}
}

void chartXAxisSetLabelAlign(uint8_t align)
{
  if (charts)
    {chart_axis_set_label_align(charts[chartIndex]->x_axis, align);}
}

void chartYAxisSetLabelAlign(uint8_t align)
{
  if (charts)
    {chart_axis_set_label_align(charts[chartIndex]->y_axis, align);}
}

void chartXAxisSetMin(double min)
{
  if (charts)
    {chart_axis_set_min(charts[chartIndex]->x_axis, min);}
}

void chartXAxisSetMax(double max)
{
  if (charts)
    {chart_axis_set_max(charts[chartIndex]->x_axis, max);}
}

void chartYAxisSetMin(double min)
{
  if (charts)
    {chart_axis_set_min(charts[chartIndex]->y_axis, min);}
}

void chartYAxisSetMax(double max)
{
  if (charts)
    {chart_axis_set_max(charts[chartIndex]->y_axis, max);}
}

void chartXAxisSetLogBase(uint16_t logBase)
{
  if (charts)
    {chart_axis_set_log_base(charts[chartIndex]->x_axis, logBase);}
}

void chartYAxisSetLogBase(uint16_t logBase)
{
  if (charts)
    {chart_axis_set_log_base(charts[chartIndex]->y_axis, logBase);}
}

void chartXAxisSetMajorTickPoint(uint8_t ty)
{
  if (charts)
    {chart_axis_set_major_tick_point(charts[chartIndex]->x_axis, ty);}
}

void chartYAxisSetMajorTickPoint(uint8_t ty)
{
  if (charts)
    {chart_axis_set_major_tick_point(charts[chartIndex]->y_axis, ty);}
}

void chartXAxisSetMinorTickPoint(uint8_t ty)
{
  if (charts)
    {chart_axis_set_minor_tick_point(charts[chartIndex]->x_axis, ty);}
}

void chartYAxisSetMinorTickPoint(uint8_t ty)
{
  if (charts)
    {chart_axis_set_minor_tick_point(charts[chartIndex]->y_axis, ty);}
}

void chartXAxisSetIntervalUnit(uint16_t unit)
{
  if (charts)
    {chart_axis_set_interval_unit(charts[chartIndex]->x_axis, unit);}
}

void chartYAxisSetIntervalUnit(uint16_t unit)
{
  if (charts)
    {chart_axis_set_interval_unit(charts[chartIndex]->y_axis, unit);}
}

void chartXAxisSetIntervalUnitTick(uint16_t unit)
{
  if (charts)
    {chart_axis_set_interval_unit_tick(charts[chartIndex]->x_axis, unit);}
}

void chartYAxisSetIntervalUnitTick(uint16_t unit)
{
  if (charts)
    {chart_axis_set_interval_unit_tick(charts[chartIndex]->y_axis, unit);}
}

void chartXAxisSetMajorUnit(double unit)
{
  if (charts)
    {chart_axis_set_major_unit(charts[chartIndex]->x_axis, unit);}
}

void chartYAxisSetMajorUnit(double unit)
{
  if (charts)
    {chart_axis_set_major_unit(charts[chartIndex]->y_axis, unit);}
}

void chartXAxisSetMinorUnit(double unit)
{
  if (charts)
    {chart_axis_set_minor_unit(charts[chartIndex]->x_axis, unit);}
}

void chartYAxisSetMinorUnit(double unit)
{
  if (charts)
    {chart_axis_set_minor_unit(charts[chartIndex]->y_axis, unit);}
}

void chartXAxisSetDisplayUnits(uint8_t unit)
{
  if (charts)
    {chart_axis_set_display_units(charts[chartIndex]->x_axis, unit);}
}

void chartYAxisSetDisplayUnits(uint8_t unit)
{
  if (charts)
    {chart_axis_set_display_units(charts[chartIndex]->y_axis, unit);}
}


void chartXAxisSetDisplayUnitsVisible(uint8_t visible)
{
  if (charts)
    {chart_axis_set_display_units_visible(charts[chartIndex]->x_axis, visible);}
}

void chartYAxisSetDisplayUnitsVisible(uint8_t visible)
{
  if (charts)
    {chart_axis_set_display_units_visible(charts[chartIndex]->y_axis, visible);}
}

void chartXAxisSetMajorGridlinesVisible(uint8_t visible)
{
  if (charts)
    {chart_axis_set_major_gridlines_visible(charts[chartIndex]->x_axis, visible);}
}

void chartYAxisSetMajorGridlinesVisible(uint8_t visible)
{
  if (charts)
    {chart_axis_set_major_gridlines_visible(charts[chartIndex]->y_axis, visible);}
}

void chartXAxisSetMinorGridlinesVisible(uint8_t visible)
{
  if (charts)
    {chart_axis_set_minor_gridlines_visible(charts[chartIndex]->x_axis, visible);}
}

void chartYAxisSetMinorGridlinesVisible(uint8_t visible)
{
  if (charts)
    {chart_axis_set_minor_gridlines_visible(charts[chartIndex]->y_axis, visible);}
}

void chartXAxisSetMajorGridlinesSetLine()
{
  if (charts
      && chartLines)
    {chart_axis_set_major_gridlines_set_line(charts[chartIndex]->x_axis, &chartLines[chartLineIndex]);}
}

void chartYAxisSetMajorGridlinesSetLine()
{
  if (charts
      && chartLines)
    {chart_axis_set_major_gridlines_set_line(charts[chartIndex]->y_axis, &chartLines[chartLineIndex]);}
}

void chartXAxisSetMinorGridlinesSetLine()
{
  if (charts
      && chartLines)
    {chart_axis_set_minor_gridlines_set_line(charts[chartIndex]->x_axis, &chartLines[chartLineIndex]);}
}

void chartYAxisSetMinorGridlinesSetLine()
{
  if (charts
      && chartLines)
    {chart_axis_set_minor_gridlines_set_line(charts[chartIndex]->y_axis, &chartLines[chartLineIndex]);}
}

void chartSeriesSetNameRange(char* sheetName,
			     uint32_t row,
			     unsigned short col)
{
  if (series)
    {
      chart_series_set_name_range(series[chartIndex],
				  sheetName,
				  row,
				  col);
    }
}

void chartXAxisSetReverse()
{
  if (charts)
    {chart_axis_set_reverse(charts[chartIndex]->x_axis);}
}

void chartYAxisSetReverse()
{
  if (charts)
    {chart_axis_set_reverse(charts[chartIndex]->y_axis);}
}


void chartXAxisSetCrossingOff()
{
  if (charts)
    {chart_axis_set_crossing_off(charts[chartIndex]->x_axis);}
}

void chartYAxisSetCrossingOff()
{
  if (charts)
    {chart_axis_set_crossing_off(charts[chartIndex]->y_axis);}
}






void closeWorkbook()
{
  if (workbook)
    {handleLXWError(workbook_close(workbook), "close-workbook");}
  chartPointsCleanup();
  chartLinesCleanup();
  chartPatternsCleanup();
  chartFillsCleanup();
  chartFontsCleanup();
  formatsCleanup();
  rowNumberCleanup();
  colNumberCleanup();
  dataValidationListsCleanup();
  dataValidationsCleanup();
  seriesCleanup();
  seriesErrorBarsCleanup();
  chartsCleanup();
}

