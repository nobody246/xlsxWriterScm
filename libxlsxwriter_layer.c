// xlsxWriterScm (c) 2019-2025 Alexander Semotan
// BSD2 Software License
void createWorkbook(char* workbookName)
{
  workbook = workbook_new(workbookName);
}


void createWorkbookConstantMemory(char* workbookName)
{
  lxw_workbook_options options = {.constant_memory = LXW_TRUE,
				  .tmpdir = NULL};
  workbook = workbook_new_opt(workbookName, &options);
}

void resizeDataValidations()
{
  lxw_data_validation* v = NULL;
  v = realloc(dataValidations, dataValidationCount * sizeof(lxw_data_validation));
  if (!v)
    {
      printf("error: There was a problem with allocating more data validations in create-data-validation.\n");
      return;
    }
  v[dataValidationIndex] = (lxw_data_validation){
    .value_formula="",
    .minimum_formula="",
    .maximum_formula="",
    .input_title="",
    .input_message="",
    .error_title="",
    .error_message=""};
  dataValidations = v;
  char*** d = NULL;
  //127 total possible list entries, 255 character count which includes commas
  d = realloc(dataValidationList, dataValidationCount * sizeof(char**) * 127);
  if (!d)
    {
      printf("error: There was a problem with allocating more data validation list space in create-data-validation.\n");
      return;
    }
  dataValidationList = d;
  *(dataValidationList + dataValidationCount) = calloc(255 * 127, sizeof(char*) + sizeof(char));
      
  ptrdiff_t* h = NULL;
  h = realloc(dataValidationListCharCount, dataValidationCount * sizeof(ptrdiff_t));
  if (!h)
    {
      printf("error: There was problem allocating data-validation-list-char-count space in create-data-validation.\n");
      return;
    }
  dataValidationListCharCount = h;

  ptrdiff_t* y = NULL;
  y = realloc(dataValidationListStrIndex, dataValidationCount * sizeof(ptrdiff_t));
  if (!y)
    {
      printf("error: There was problem allocating data-validation-list-str-index space in create-data-validation.\n");
      return;
    }
  dataValidationListStrIndex = y;
  maxAllowedDataValidations+=1;

}

void dataValidationsBoundsCheck()
{
  if (!dataValidationCount)
    {initDataValidations(1);}
  if (dataValidationCount >= maxAllowedDataValidations)
    {resizeDataValidations();}
}



void formatsCleanup()
{
  if (formats)
    {
      free(formats);
      formats = NULL;
    }
  maxAllowedFormats = 0;
  formatIndex = 0;
  formatCount = 0;
}


void rowNumberCleanup()
{
  if (rowNumbers)
    {
      free(rowNumbers);
      rowNumbers = NULL;
    }
  maxAllowedRowNumbers = 0;
}

void colNumberCleanup()
{
  if (colNumbers)
    {
      free(colNumbers);
      colNumbers = NULL;
    }
  maxAllowedColNumbers = 0;
}

void dataValidationsCleanup()
{
  if (dataValidations)
    {
      for (int x = 0; x<=dataValStrInd; x++)
	{
	  free(dataValStrs[x]);
	  dataValStrs[x] = NULL;
	}
      free(dataValidations);
      dataValidations = NULL;
      if (dataValStrs)
	{
	  free(dataValStrs);
	  dataValStrs = NULL;
	}
    }
  for (int x = 0; x<dataValidationCount; x++)
    {
      if (dataValidationList[x])
	{free(dataValidationList[x]);}
      dataValidationList[x] = NULL;
    }
  if (dataValidationList)
    {
      free(dataValidationList);
      dataValidationList = NULL;
    }
  maxAllowedDataValidations = 0;
  dataValidationCount = 0;
  dataValidationIndex = 0;
  if (dataValidationListCharCount)
    {
      free(dataValidationListCharCount);
      dataValidationListCharCount = NULL;
    }
  if (dataValidationListStrIndex != NULL)
    {
      free(dataValidationListStrIndex);
      dataValidationListStrIndex = NULL;
    }
  dataValidationListCount = 0;
}

void chartPointsCleanup()
{
  for (int x = 0; x<chartPointCount; x++)
    {
      //free(chartPoints[x]); //segfaults?
      chartPoints[x] = NULL;
    }
  if (chartPoints)
    {
      free(chartPoints);
      chartPoints = NULL;
    }
  if (chartPointHolder)
    {
      free(chartPointHolder);
      chartPointHolder = NULL;
    }
  maxAllowedChartPoints = 0;
  chartPointCount = 0;
  chartPointIndex = 0;
}

void chartFillsCleanup()
{
  if (chartFills)
    {
      free(chartFills);
      chartFills = NULL;
    }
  maxAllowedChartFills = 0;
  chartFillCount = 0;
  chartFillIndex = 0;
}
     
void chartPatternsCleanup()
{
  if (chartPatterns)
    {
      free(chartPatterns);
      chartPatterns = NULL;
    }
  maxAllowedChartPatterns = 0;
  chartPatternCount = 0;
  chartPatternIndex = 0;
}
     
void chartLinesCleanup()
{
  if (chartLines)
    {
      free(chartLines);
      chartLines = NULL;
    }
  maxAllowedChartLines = 0;
  chartLineCount = 0;
  chartLineIndex = 0;
}     

void seriesCleanup()
{
  for (int x = 0; x<seriesCount; x++)
    {
      if (series[x])
	{
	  //free(series[x]); //segfaults?
	  series[x] = NULL;
	}
    }
  if (series)
    {
      free(series);
      series = NULL;
    }
  maxAllowedSeries = 0;
  seriesCount = 0;
  seriesIndex = 0;
}

void chartsCleanup()
{
  for (int x = 0; x<chartCount; x++)
    {
      //free(charts[x]); //segfaults?
      charts[x] = NULL;
    }  
  if (charts)
    {
      free(charts);
      charts = NULL;
    }
  maxAllowedCharts = 0;
  chartCount = 0;
  chartIndex = 0;
}

void seriesErrorBarsCleanup()
{
  if (seriesErrorBars)
    {
      free(seriesErrorBars);
      seriesErrorBars = NULL;
    }
  maxAllowedSeriesErrorBars = 0;
  seriesErrorBarsCount = 0;
  seriesErrorBarsIndex = 0;
}

void chartFontsCleanup()
{
  if (chartFonts)
    {
      free(chartFonts);
      chartFonts = NULL;
    }
  maxAllowedChartFonts = 0;
  chartFontCount = 0;
  chartFontIndex = 0;
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
  dataValidations[0] = (lxw_data_validation){
    .value_formula="",
    .minimum_formula="",
    .maximum_formula="",
    .input_title="",
    .input_message="",
    .error_title="",
    .error_message=""};
  dataValidationList = calloc(allocateN, sizeof(char**) * 127);
  int i;
  int x;
  for (i = 0; i < allocateN; i++)
    {
      //127 total possible list entries, 255 character count which includes commas
      *(dataValidationList + i) = calloc(255 * 127, sizeof(char*) + sizeof(char));
      //for (x = 0; x < 127; x++)
      //{
      //  *(dataValidationList + i + x) = calloc(255, sizeof(char));
      //}
    }
  dataValidationListCharCount = calloc(allocateN, sizeof(ptrdiff_t));
  dataValidationListStrIndex = calloc(allocateN, sizeof(ptrdiff_t));
  maxAllowedDataValidations = allocateN;
  dataValidationListCount = 0;
  dataValidationCount = 1;
  dataValidationIndex = 0;
}

void initDataValidationLists()
{
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
  chartPoints = calloc(allocateN + 1, sizeof(lxw_chart_point*));
  chartPointHolder = calloc(allocateN, sizeof(lxw_chart_point));
  maxAllowedChartPoints = allocateN;
  chartPointCount = 0;
  chartPointIndex = 0;
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
  series = calloc(allocateN, (sizeof(lxw_chart_series*) + sizeof(lxw_chart_series)));
  maxAllowedSeries = allocateN;
}

void initCharts(ptrdiff_t allocateN)
{
  chartsCleanup();
  charts = calloc(allocateN, sizeof(lxw_chart*));
  maxAllowedCharts = allocateN;
}

void initChartFonts(ptrdiff_t allocateN)
{
  chartFontsCleanup();
  chartFonts = calloc(allocateN, sizeof(lxw_chart_font));
  maxAllowedChartFonts = allocateN;
}

void colOptionCleanup()
{
  if (colOptions)
    {free(colOptions);}
  colOptionIndex = 0;
  colOptionCount = 0;
  maxAllowedColOptions = 0;
}


void initColOptions(ptrdiff_t allocateN)
{
  colOptionCleanup();
  colOptions = calloc(allocateN, sizeof(lxw_row_col_options));
  maxAllowedColOptions = allocateN;
}

void createDataValidationListEntry(char* dvl)
{
  dataValidationsBoundsCheck();
  if (dataValidationListCharCount[dataValidationIndex] > 254 ||
      dataValidationListStrIndex[dataValidationIndex] > 127)
    {
      printf("error in create-data-validation-list-entry list entries maxed out.\n");
      return;
    }
  int i = 0;
  while (dvl[i])
    {i++;}
  if ((i > 254) ||
      (dataValidationListCharCount[dataValidationIndex] + i) > 254)
    {
      printf("error in create-data-validation-list-entry list entry is too large.\n");
      return;
    }
  dataValidationListCharCount[dataValidationIndex] += i;
  dataValidationList[dataValidationIndex][dataValidationListStrIndex[dataValidationIndex]] =
    calloc(dataValidationListCharCount[dataValidationIndex], sizeof(char));
  memcpy(dataValidationList[dataValidationIndex][dataValidationListStrIndex[dataValidationIndex]],
	 dvl,
	 dataValidationListCharCount[dataValidationIndex]);
  dataValidationListCharCount[dataValidationIndex]+=1;//inc for comma separator
  dataValidationListStrIndex[dataValidationIndex]+=1;//inc index
}

void setValidationValidate(uint8_t validate)
{
  dataValidationsBoundsCheck();
  dataValidations[dataValidationIndex].validate = validate;
}

void setValidationCriteria(uint8_t criteria)
{
  dataValidationsBoundsCheck();
  dataValidations[dataValidationIndex].criteria = criteria;
}

void setValidationIgnoreBlank(uint8_t ignoreBlank)
{
  dataValidationsBoundsCheck();
  dataValidations[dataValidationIndex].ignore_blank = ignoreBlank;
}

void setValidationShowInput(uint8_t showInput)
{
  dataValidationsBoundsCheck();
  dataValidations[dataValidationIndex].show_input = showInput;
}

void setValidationShowError(uint8_t showError)
{
  if (!dataValidationCount)
    {initDataValidations(1);}
  if (dataValidationCount >= maxAllowedDataValidations)
    {resizeDataValidations();}
  dataValidations[dataValidationIndex].show_error = showError;
}

void setValidationErrorType(uint8_t errorType)
{
  dataValidationsBoundsCheck();
  dataValidations[dataValidationIndex].error_type = errorType;
}

void setValidationDropdown(uint8_t dropdown)
{
  if (!dataValidationCount)
    {initDataValidations(1);}
  if (dataValidationCount >= maxAllowedDataValidations)
    {resizeDataValidations();}
  dataValidations[dataValidationIndex].dropdown = dropdown;
}

void setValidationValueNumber(double num)
{
  dataValidationsBoundsCheck(); 
  dataValidations[dataValidationIndex].value_number = num;
}

void setValidationValueFormula(char* form)
{
  dataValidationsBoundsCheck();
  int i = 1;
  while (form[i])
    {i++;}
  char** d = realloc(dataValStrs, (dataValStrInd + 1) * sizeof(char*));
  d[dataValStrInd] = calloc(i, sizeof(char));
  strcpy(d[dataValStrInd], form);
  dataValidations[dataValidationIndex].value_formula = d[dataValStrInd++];
  dataValStrs = d;
}

void setValidationValueDateTime(int y,
				int m,
				int d,
				int h,
				int mn,
				double s)
{
  dataValidationsBoundsCheck(); 
  lxw_datetime x  =
    (lxw_datetime){.year = y,
		   .month = m,
		   .day = d,
		   .hour = h,
		   .min = mn,
		   .sec = s};
  dataValidations[dataValidationIndex].value_datetime = x;
}

void setValidationMinNumber(double num)
{
  dataValidationsBoundsCheck(); 
  dataValidations[dataValidationIndex].minimum_number = num;
}

void setValidationMinFormula(char* form)
{
  dataValidationsBoundsCheck();
  int i=1;
  while (form[i])
    {i++;}
  char** d = realloc(dataValStrs, (dataValStrInd + 1) * sizeof(char*));
  d[dataValStrInd] = calloc(i, sizeof(char));
  strcpy(d[dataValStrInd], form);
  dataValidations[dataValidationIndex].minimum_formula = d[dataValStrInd++];
  dataValStrs = d;
}

void setValidationMinDateTime(int y,
			      int m,
			      int d,
			      int h,
			      int mn,
			      double s)
{
  dataValidationsBoundsCheck();
  lxw_datetime x  =
    (lxw_datetime){.year = y,
		   .month = m,
		   .day = d,
		   .hour = h,
		   .min = mn,
		   .sec = s};
  dataValidations[dataValidationIndex].minimum_datetime = x;
}

void setValidationMaxNumber(double num)
{
  dataValidationsBoundsCheck();
  dataValidations[dataValidationIndex].maximum_number = num;
}

void setValidationMaxFormula(char* form)
{
  dataValidationsBoundsCheck();
  int i = 1;
  while (form[i])
    {i++;}
  char** d = realloc(dataValStrs, (dataValStrInd + 1) * sizeof(char*));
  d[dataValStrInd] = calloc(i, sizeof(char));
  strcpy(d[dataValStrInd], form);
  dataValidations[dataValidationIndex].minimum_formula = d[dataValStrInd++];
  dataValStrs = d;
}

void setValidationMaxDateTime(int y,
			      int m,
			      int d,
			      int h,
			      int mn,
			      double s)
{
  dataValidationsBoundsCheck();  
  lxw_datetime x  =
    (lxw_datetime){.year = y,
		   .month = m,
		   .day = d,
		   .hour = h,
		   .min = mn,
		   .sec = s};
  dataValidations[dataValidationIndex].maximum_datetime = x;
}

void setValidationInputTitle(char* title)
{
  dataValidationsBoundsCheck();
  int i = 1;
  while (title[i])
    {i++;}
  char** d = realloc(dataValStrs, (dataValStrInd + 1) * sizeof(char*));
  d[dataValStrInd] = calloc(i, sizeof(char));
  strcpy(d[dataValStrInd], title);
  dataValidations[dataValidationIndex].input_title = d[dataValStrInd++];
  dataValStrs = d;
}

void setValidationInputMessage(char* message)
{
  dataValidationsBoundsCheck();
  int i = 1;
  while (message[i])
    {i++;}
  char** d = realloc(dataValStrs, (dataValStrInd + 1) * sizeof(char*));
  d[dataValStrInd] = calloc(i, sizeof(char));
  strcpy(d[dataValStrInd], message);
  dataValidations[dataValidationIndex].input_message = d[dataValStrInd++];
  dataValStrs = d;
}

void setValidationErrorTitle(char* title)
{
  dataValidationsBoundsCheck();
  int i = 1;
  while (title[i])
    {i++;}
  char** d = realloc(dataValStrs, (dataValStrInd + 1) * sizeof(char*));
  d[dataValStrInd] = calloc(i, sizeof(char));
  strcpy(d[dataValStrInd], title);
  dataValidations[dataValidationIndex].error_title = d[dataValStrInd++];
  dataValStrs = d;
}

void setValidationErrorMessage(char* message)
{
  dataValidationsBoundsCheck();
  int i = 1;
  while (message[i])
    {i++;}
  char** d = realloc(dataValStrs, (dataValStrInd + 1) * sizeof(char*));
  d[dataValStrInd] = calloc(i, sizeof(char));
  strcpy(d[dataValStrInd], message);
  dataValidations[dataValidationIndex].error_message = d[dataValStrInd++];
  dataValStrs = d;
}

void setValidationValueList()
{
  dataValidationsBoundsCheck();
  dataValidations[dataValidationIndex].value_list = dataValidationList[dataValidationIndex];
}


void createDataValidation()
{
  if (!worksheet)
    {return;}
  if (!dataValidationCount)
    {
      printf("error from create-data-validation, data Validation is not initialized.");
      return;
    }
  worksheet_data_validation_cell(worksheet, row, col, &dataValidations[dataValidationIndex]);
  dataValidationIndex = dataValidationCount;
  dataValidationCount+=1;
}

void createChartPoint()
{
  if (!chartPoints)
    {
      printf("Error: Chart-points not initialized in create-chart-point");
      return;
    }
  if (!chartLines)
    {initChartLines(1);}
  if (!chartPatterns)
    {initChartPatterns(1);}
  if (!chartFills)
    {initChartFills(1);}
  if (chartPointCount >= maxAllowedChartPoints)
    {
      printf("Trying to create more chart points than allotted in create-chart-point");
      return;
    }
  chartPointIndex = chartPointCount;
  *(chartPointHolder + chartPointIndex)  = (lxw_chart_point){.fill =    &chartFills[chartFillIndex],
							     .line =    &chartLines[chartLineIndex],
							     .pattern = &chartPatterns[chartPatternIndex]};
  *(chartPoints + chartPointIndex) = &chartPointHolder[chartPointCount];
  chartPointCount+=1;
}

void createChartFill(int32_t color,
		     uint8_t none,
		     uint8_t transparency)
{
  if (chartFillCount >= maxAllowedChartFills)
    {
      int x = chartFillCount + 1;
      lxw_chart_fill* d = NULL;// = malloc(x * sizeof(lxw_chart_fill));
      d = realloc(chartFills, x * sizeof(lxw_chart_fill));
      if (!d)
	{
	  printf("error in create-chart-fill, problem allocating more space. \n");
	  return;
	}
      chartFills = d;
      maxAllowedChartFills += 1;
    }
  *(chartFills + chartFillCount) = (lxw_chart_fill) {.color = color, 
						     .none = none, 
						     .transparency = transparency};
  chartFillIndex = chartFillCount;
  chartFillCount+=1;
}

void createChartLine(int32_t color,
		     uint8_t none,
		     float width,
		     uint8_t dashType)
{
  if (!chartLines)
    {initChartLines(1);}
  if (chartLineCount >= maxAllowedChartLines)
    {
      lxw_chart_line* d = NULL;// = malloc(chartLineCount * sizeof(lxw_chart_line));
      d = realloc(chartLines, chartLineCount * sizeof(lxw_chart_line));
      if (!d)
	{
	  printf("error in create-chart-fill, problem allocating more space. \n");
	  return;
	}
      chartLines = d;
      maxAllowedChartLines += 1;
    }
  if (chartLineCount < maxAllowedChartLines)
    {
      *(chartLines + chartLineCount) = (lxw_chart_line){.color     = color,
							.none      = none,
							.width     = width,
							.dash_type = dashType};
      chartLineCount+=1;
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
  if (chartPatternCount >= maxAllowedChartPatterns)
    {
      ptrdiff_t i = chartPatternCount + 1;
      lxw_chart_pattern* d = NULL;// = malloc(i * sizeof(lxw_chart_pattern));
      d = realloc(chartPatterns, i * sizeof(lxw_chart_pattern));
      if (!d)
	{
	  printf("error in create-chart-pattern, problem allocating more space. \n");
	  return;
	}
      chartPatterns = d;
      maxAllowedChartPatterns += 1;
    }
  *(chartPatterns + chartPatternCount) = (lxw_chart_pattern){.fg_color = fgColor,
							     .bg_color = bgColor,
							     .type     = ty};
  chartPatternCount+=1;
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
  if (!chartFonts)
    {initChartFonts(1);}
  if (chartFontCount >= maxAllowedChartFonts)
    {
      lxw_chart_font* d = NULL;// = malloc(chartFontCount * sizeof(lxw_chart_font));
      d = realloc(chartFonts, chartFontCount * sizeof(lxw_chart_font));
      if (!d)
	{
	  printf("error in create-chart-fill, problem allocating more space. \n");
	  return;
	}
      chartFonts = d;
      maxAllowedChartFonts += 1;
    }
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
      chartFontCount+=1;
    }
}
		     

void addToRowNumberList(uint32_t val)
{
  if (!rowNumbers)
    {initRowNumbers(1);}
  if (rowNumberCount >= maxAllowedRowNumbers)
    {
      lxw_row_t* d = NULL;//malloc(rowNumberCount * sizeof(lxw_row_t));
      d = realloc(rowNumbers, rowNumberCount * sizeof(lxw_row_t));
      if (!d)
	{
	  printf("error in add-to-col-number-list with allocating more col numbers.");
	  return;
	}
      rowNumbers = d;
      maxAllowedRowNumbers+=1;
    }
  if (rowNumberCount < maxAllowedRowNumbers)
    {*(rowNumbers + rowNumberCount) = (lxw_row_t) val;}
  rowNumberCount += 1;
}

void addToColNumberList(unsigned short val)
{
  if (!colNumbers)
    {initColNumbers(1);}
  if (colNumberCount >= maxAllowedColNumbers)
    {
      lxw_col_t* d = NULL; //malloc(colNumberCount * sizeof(lxw_col_t));
      d = realloc(colNumbers, colNumberCount * sizeof(lxw_col_t));
      if (!d)
	{
	  printf("error in add-to-col-number-list with allocating more col numbers.");
	  return;
	}
      colNumbers = d;
      maxAllowedColNumbers+=1;
    }
  if (colNumberCount < maxAllowedColNumbers)
    {*(colNumbers + colNumberCount) = (lxw_col_t) val;}
  colNumberCount += 1;
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

void setColOptionIndex(ptrdiff_t ind)
{
  if (ind < colOptionCount)
    {colOptionIndex = ind;}
}

void addWorksheet(char* worksheetName)
{
  if (workbook)
    {
      int i = 0;
      while (worksheetName[i])
	{i++;}
      if (i == 0)
	{worksheet = workbook_add_worksheet(workbook, NULL);}
      else
	{worksheet = workbook_add_worksheet(workbook, worksheetName);}

    }
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


void worksheetWriteFormula(char* formula)
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
		     "worksheet-write-array-formula");
    }
}


void worksheetWriteDatetime(int year,
			    int month,
			    int day,
			    int hour,
			    int mn,
			    double sec)
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


void worksheetWriteUrl(char* url)
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

void worksheetWriteBoolean(short value)
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

void worksheetWriteBlank()
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

void worksheetSetRowOpt(double height,
			uint32_t row)
{
  if (!colOptionCount)
    {
      printf("error in worksheet-set-row-opt Please init col row opts.");
      return;
    }
  if (!worksheet)
    {return;}
  lxw_format *f = NULL;
  if (formats)
    {f = formats[formatIndex];}
  handleLXWError(worksheet_set_row_opt(worksheet, row, height, f, &colOptions[colOptionIndex]), 
		 "worksheet-set-row-opt");
}


void worksheetSetColumn(double width,
			uint16_t stCol,
			uint16_t endCol)
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

void worksheetSetColumnOpt(double width,
			   unsigned short stCol,
			   unsigned short endCol)
{
  if (!colOptionCount)
    {
      printf("error in worksheet-set-column-opt Please init col row opts.");
      return;
    }
  if (!worksheet)
    {return;}
  lxw_format *f = NULL;
  if (formats)
    {f = formats[formatIndex];}
  handleLXWError(worksheet_set_column_opt(worksheet, 
					  stCol, 
					  endCol, 
					  width, 
					  f, 
					  &colOptions[colOptionIndex]),
		 "worksheet-set-column-opt");
}


void worksheetInsertImage(char* filename)
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
			     char* filename)
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


void worksheetInsertChart()
{
  if (worksheet
      && charts)
    {
      worksheet_insert_chart(worksheet, row, col, charts[chartIndex]);
    }
}

void worksheetInsertChartOpt(int x_offset,
			     int y_offset,
			     int x_scale,
			     int y_scale)
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
  if (worksheet)
    {
      lxw_format* f = NULL;
      if (formats)
	{f = formats[formatIndex];}
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

void worksheetDataValidationCell()
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

void worksheetFreezePanes(uint32_t row,
			  uint16_t col)
{
  if (worksheet)
    {worksheet_freeze_panes(worksheet, row, col);}
}

void worksheetSplitPanes(double vertical,
			 double horizontal)
{
  if (worksheet)
    {worksheet_split_panes(worksheet, vertical, horizontal);}
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
    {worksheet_set_h_pagebreaks(worksheet, rowNumbers);}
  rowNumberCleanup();
}

void worksheetSetVPageBreaks()
{
  if (worksheet &&
      colNumbers)
    {worksheet_set_v_pagebreaks(worksheet, colNumbers);}
  colNumberCleanup();
}

void createFormat()
{
  if (!workbook)
    {return;}
  if (formatCount > maxAllowedFormats)
    {
      printf("trying to add more formats than alotted in init, in add-format.\n");
      return;
    }
  *(formats + formatCount) = workbook_add_format(workbook);
  formatIndex = formatCount;
  formatCount+=1;
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

void formatSetHidden()
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

void setFontColor(int32_t color)
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

void setIndent(int level)
{
  if (formats)
    {format_set_indent(formats[formatIndex], level);}
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

void setBgColor(int32_t color)
{
  if (formats)
    {format_set_bg_color(formats[formatIndex], color);}
}

void setFgColor(int32_t color)
{
  if (formats)
    {format_set_fg_color(formats[formatIndex], color);}
}

void setBorderColor(int32_t color)
{
  if (formats)
    {format_set_border_color(formats[formatIndex], color);}
}

void setBorderBottomColor(int32_t color)
{
  if (formats)
    {format_set_bottom_color(formats[formatIndex], color);}
}

void setBorderTopColor(int32_t color)
{
  if (formats)
    {format_set_top_color(formats[formatIndex], color);}
}

void setBorderLeftColor(int32_t color)
{
  if (formats[formatIndex])
    {format_set_left_color(formats[formatIndex], color);}
}

void setBorderRightColor(int32_t color)
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

void workbookSetCustomPropertyDateTime(char* name,
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

void workbookDefineName(char* name, char* val)
{
  if (!workbook)
    {return;}
  handleLXWError(workbook_define_name(workbook, name, val),
		 "workbook-define-name");
}

void workbookGetWorksheetByName(char* name)
{
  if (workbook)
    {worksheet = workbook_get_worksheet_by_name(workbook, name);}
}

void workbookValidateWorksheetName(char* name)
{
  if (workbook)
    {workbook_validate_sheet_name(workbook, name);}
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
  if (!charts)
    {initCharts(1);}
  if (chartCount >= maxAllowedCharts)
    {
      size_t s = chartCount * (sizeof(lxw_chart*) + sizeof(lxw_chart));
      lxw_chart** d = NULL;
      d = realloc(charts, s);
      if (!d)
	{
	  printf("error: problem allocating more charts in create-chart");
	  return;
	}
      charts = d;
      maxAllowedCharts+=1;
    }
  if (chartCount < maxAllowedCharts
      && charts
      && workbook)
    {
      *(charts + chartCount) =
	workbook_add_chart(workbook,
			   (lxw_chart_type) chartType);
      chartIndex = chartCount;
      chartCount+=1;
    }
}

void createChartSeries(char* categories, char* vals)
{
  if (!series)
    {initSeries(1);}
  if (seriesCount >= maxAllowedSeries)
    {
      lxw_chart_series** d = NULL;
      d = realloc(series, seriesCount * (sizeof(lxw_chart_series*) + sizeof(lxw_chart_series)));
      if (!d)
	{
	  printf("error problem allocating more space for chart-series in create-chart-series.");
	  return;
	}
      series = d;
      maxAllowedSeries+=1;
    }
  char* c = NULL;
  int i=0;
  while (categories[i])
    {i++;}
  if (i)
    {c = categories;}
  *(series + seriesCount) = chart_add_series(charts[chartIndex], c, vals);
  seriesIndex = seriesCount;
  seriesCount+=1;
}

void createRowColOpt(unsigned short level,
		     unsigned short hidden,
		     unsigned short collapsed)
{
  if (!colOptionCount)
    {initColOptions(1);}
  if (colOptionCount >= maxAllowedColOptions)
    {
      lxw_row_col_options* d = NULL; 
      d = realloc(colOptions, colOptionCount * sizeof(lxw_row_col_options));
      if (!d)
	{
	  printf("error in worksheet-set-row-opt, problem allocating more memory for col-options ");
	  return;
	}
      colOptions = d;
      maxAllowedColOptions+=1;
    }
  colOptions[colOptionIndex] = (lxw_row_col_options)
    {.hidden = hidden,
     .level = level,
     .collapsed = collapsed};
  colOptionIndex = colOptionCount;
  colOptionCount += 1;
}



void chartSeriesSetName(char* name)
{
  if (series && charts)
    {
      if (!series[seriesIndex])
	{
	  printf("error from chart-series-set-name, please create a series first.");
	  return;
	}
      chart_series_set_name(series[seriesIndex], name);
    }
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
			   int32_t fgColor,
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
				 int32_t fgColor,
				 int32_t bgColor)
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
  if (series &&
      chartPoints)
    {
      chart_series_set_points(series[seriesIndex], chartPoints);
    }
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
      lxw_chart_font font = {.name=fontName, .color=LXW_COLOR_BLACK};
      chart_series_set_labels_font(series[seriesIndex],
				   &font);
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
      seriesErrorBarsCount+=1;
    }
}

void chartSeriesSetErrorBars(uint8_t ty,
			     double val,
			     unsigned short horizontal)
{
  if (series)
    {
      chart_series_set_error_bars((horizontal > 0) ?
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
      chart_series_set_error_bars_direction(seriesErrorBars[seriesErrorBarsIndex],
					    dir);
    }
}

void chartSeriesSetErrorBarsEndCap(uint8_t endCap)
{
  if (seriesErrorBars)
    {
      chart_series_set_error_bars_endcap(seriesErrorBars[seriesErrorBarsIndex],
					 endCap);
    }
}

void chartSeriesSetErrorBarsLine()
{
  if (seriesErrorBars &&
      chartLines)
    {
      chart_series_set_error_bars_line(seriesErrorBars[seriesErrorBarsIndex],
				       &chartLines[chartLineIndex]);
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

void chartXAxisSetNameRange(char* sheetName)
{
  if (charts)
    {
      chart_axis_set_name_range(charts[chartIndex]->x_axis,
				sheetName,
				row,
				col);
    }
}

void chartYAxisSetNameRange(char* sheetName)
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
    {chart_axis_set_line(charts[chartIndex]->x_axis, &chartLines[chartLineIndex]);}
}

void chartYAxisSetLine()
{
  if (charts &&
      chartLines)
    {chart_axis_set_line(charts[chartIndex]->y_axis, &chartLines[chartLineIndex]);}
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
    {chart_axis_set_pattern(charts[chartIndex]->x_axis, &chartPatterns[chartPatternIndex]);}
}

void chartYAxisSetPattern()
{
  if (charts &&
      chartPatterns)
    {chart_axis_set_pattern(charts[chartIndex]->y_axis, &chartPatterns[chartPatternIndex]);}
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
    {chart_axis_set_label_position(charts[chartIndex]->x_axis, pos);}
}

void chartYAxisSetLabelPosition(uint8_t pos)
{
  if (charts)
    {chart_axis_set_label_position(charts[chartIndex]->x_axis, pos);}
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

void chartYAxisSetMin(double min)
{
  if (charts)
    {chart_axis_set_min(charts[chartIndex]->y_axis, min);}
}

void chartXAxisSetMax(double max)
{
  if (charts)
    {chart_axis_set_max(charts[chartIndex]->x_axis, max);}
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

void chartXAxisSetMajorTickMark(uint8_t ty)
{
  if (charts)
    {chart_axis_set_major_tick_mark(charts[chartIndex]->x_axis, ty);}
}

void chartYAxisSetMajorTickMark(uint8_t ty)
{
  if (charts)
    {chart_axis_set_major_tick_mark(charts[chartIndex]->y_axis, ty);}
}

void chartXAxisSetMinorTickMark(uint8_t ty)
{
  if (charts)
    {chart_axis_set_minor_tick_mark(charts[chartIndex]->x_axis, ty);}
}

void chartYAxisSetMinorTickMark(uint8_t ty)
{
  if (charts)
    {chart_axis_set_minor_tick_mark(charts[chartIndex]->y_axis, ty);}
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

void chartXAxisSetIntervalTick(uint16_t unit)
{
  if (charts)
    {chart_axis_set_interval_tick(charts[chartIndex]->x_axis, unit);}
}

void chartYAxisSetIntervalTick(uint16_t unit)
{
  if (charts)
    {chart_axis_set_interval_tick(charts[chartIndex]->y_axis, unit);}
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
    {chart_axis_major_gridlines_set_visible(charts[chartIndex]->x_axis, visible);}
}

void chartYAxisSetMajorGridlinesVisible(uint8_t visible)
{
  if (charts)
    {chart_axis_major_gridlines_set_visible(charts[chartIndex]->y_axis, visible);}
}

void chartXAxisSetMinorGridlinesVisible(uint8_t visible)
{
  if (charts)
    {chart_axis_minor_gridlines_set_visible(charts[chartIndex]->x_axis, visible);}
}

void chartYAxisSetMinorGridlinesVisible(uint8_t visible)
{
  if (charts)
    {chart_axis_minor_gridlines_set_visible(charts[chartIndex]->y_axis, visible);}
}

void chartXAxisSetMajorGridlinesSetLine()
{
  if (charts
      && chartLines)
    {chart_axis_major_gridlines_set_line(charts[chartIndex]->x_axis, &chartLines[chartLineIndex]);}
}

void chartYAxisSetMajorGridlinesSetLine()
{
  if (charts
      && chartLines)
    {chart_axis_major_gridlines_set_line(charts[chartIndex]->y_axis, &chartLines[chartLineIndex]);}
}

void chartXAxisSetMinorGridlinesSetLine()
{
  if (charts
      && chartLines)
    {chart_axis_minor_gridlines_set_line(charts[chartIndex]->x_axis, &chartLines[chartLineIndex]);}
}

void chartYAxisSetMinorGridlinesSetLine()
{
  if (charts
      && chartLines)
    {chart_axis_minor_gridlines_set_line(charts[chartIndex]->y_axis, &chartLines[chartLineIndex]);}
}

void chartTitleSetName(char* name)
{
  if (charts)
    {chart_title_set_name(charts[chartIndex], name);}
}

void chartTitleSetNameRange(char* sheetName)
{
  if (charts)
    {
      chart_title_set_name_range(charts[chartIndex],
				 sheetName,
				 row,
				 col);
    }
}

void chartTitleSetNameFont()
{
  if (charts &&
      chartFonts)
    {chart_title_set_name_font(charts[chartIndex], &chartFonts[chartFontIndex]);}
}

void chartTitleOff()
{
  if (charts)
    {chart_title_off(charts[chartIndex]);}
}

void chartLegendSetPosition(uint8_t pos)
{
  if (charts)
    {chart_legend_set_position(charts[chartIndex], pos);}
}

void chartLegendSetFont()
{
  if (charts &&
      chartFonts)
    {chart_legend_set_font(charts[chartIndex], &chartFonts[chartFontIndex]);}
}

void addChartLegendSeriesToDelete(uint16_t series)
{
  if (!deleteSeries)
    {deleteSeries = calloc(100, sizeof(uint16_t));}
  if (deleteSeriesCount < 100)
    {
      *(deleteSeries + deleteSeriesCount) = series;
      deleteSeriesCount += 1;
    }
  else
    {chartLegendDeleteSeries();}
}
 
void chartLegendDeleteSeries()
{
  if (charts &&
      deleteSeries)
    {
      chart_legend_delete_series(charts[chartIndex], deleteSeries);
      free(deleteSeries);
    }
}

void chartAreaSetLine()
{
  if (charts &&
      chartLines)
    {chart_chartarea_set_line(charts[chartIndex], &chartLines[chartLineIndex]);}
}

void chartAreaSetFill()
{
  if (charts &&
      chartFills)
    {chart_chartarea_set_fill(charts[chartIndex], &chartFills[chartFillIndex]);}
}

void chartAreaSetPattern()
{
  if (charts &&
      chartPatterns)
    {chart_chartarea_set_pattern(charts[chartIndex], &chartPatterns[chartPatternIndex]);}
}

void chartPlotAreaSetLine()
{
  if (charts &&
      chartLines)
    {chart_plotarea_set_line(charts[chartIndex], &chartLines[chartLineIndex]);}
}

void chartPlotAreaSetFill()
{
  if (charts &&
      chartFills)
    {chart_plotarea_set_fill(charts[chartIndex], &chartFills[chartFillIndex]);}
}

void chartPlotAreaSetPattern()
{
  if (charts &&
      chartPatterns)
    {chart_plotarea_set_pattern(charts[chartIndex], &chartPatterns[chartPatternIndex]);}
}


void chartSetStyle(uint8_t styleID)
{
  if (charts)
    {chart_set_style(charts[chartIndex],styleID);}
}

void chartSetTable()
{
  if (charts)
    {chart_set_table(charts[chartIndex]);}
}

void chartSetTableGrid(uint8_t horizontal,
		       uint8_t vertical,
		       uint8_t outline,
		       uint8_t legendKeys)
{
  if (charts)
    {
      chart_set_table_grid(charts[chartIndex],
			   horizontal,
			   vertical,
			   outline,
			   legendKeys);
    }
}

void chartSetUpDownBars()
{
  if (charts)
    {chart_set_up_down_bars(charts[chartIndex]);}
}

void chartSetUpBarsLineAndFill()
{
  if (chartLines &&
      chartFills)
    {
      downBarLine = &chartLines[chartLineIndex];
      downBarFill = &chartFills[chartFillIndex];
      if (upBarLine &&
	  upBarFill)
	{
	  chart_set_up_down_bars_format(charts[chartIndex],
					upBarLine,
					upBarFill,
					downBarLine,
					downBarFill);
	}
    }
  
}


void chartSetDownBarsLineAndFill()
{
  if (chartLines &&
      chartFills)
    {
      upBarLine = &chartLines[chartLineIndex];
      upBarFill = &chartFills[chartFillIndex];
      if (downBarLine &&
	  downBarFill)
	{
	  chart_set_up_down_bars_format(charts[chartIndex],
					upBarLine,
					upBarFill,
					downBarLine,
					downBarFill);
	}
    }
}

void chartSetDropLines()
{
  if (charts &&
      chartLines)
    {chart_set_drop_lines(charts[chartIndex], &chartLines[chartLineIndex]);}
}

void chartSetHighLowLines()
{
  if (charts &&
      chartLines)
    {chart_set_high_low_lines(charts[chartIndex], &chartLines[chartLineIndex]);}
}

void chartSetSeriesOverlap(uint8_t overlap)
{
  if (charts)
    {chart_set_series_overlap(charts[chartIndex], overlap);}
}

void chartSetSeriesGap(uint16_t gap)
{
  if (charts)
    {chart_set_series_gap(charts[chartIndex], gap);}
}

void chartShowBlanksAs(uint8_t option)
{
  if (charts)
    {chart_show_blanks_as(charts[chartIndex], option);}
}

void chartShowHiddenData()
{
  if (charts)
    {chart_show_hidden_data(charts[chartIndex]);}
}

void chartSetRotation(uint16_t rotation)
{
  if (charts)
    {chart_set_rotation(charts[chartIndex], rotation);}
}

void chartSetHoleSize(uint8_t size)
{
  if (charts)
    {chart_set_hole_size(charts[chartIndex], size);}
}

void chartSeriesSetNameRange(char* sheetName)
{
  if (series)
    {
      if (!series[seriesIndex])
	{
	  printf("error from chart-series-set-name-range, please create a series first.\n");
	  return;
	}
      chart_series_set_name_range(series[seriesIndex],
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


void createChartsheet(char* sheetName)
{
  if (workbook)
    {
      const char* s = sheetName;
      chartsheet = workbook_add_chartsheet(workbook, s);
    }
}

void chartsheetSetChart()
{
  if (charts &&
      chartsheet)
    {chartsheet_set_chart(chartsheet, charts[chartIndex]);}
}


void chartsheetActivate()
{
  if (chartsheet)
    {chartsheet_activate(chartsheet);}
}

void chartsheetSelect()
{
  if (chartsheet)
    {chartsheet_select(chartsheet);}
}

void chartsheetHide()
{
  if (chartsheet)
    {chartsheet_hide(chartsheet);}
}

void chartsheetSetFirstSheet()
{
  if (chartsheet)
    {chartsheet_set_first_sheet(chartsheet);}
}

void chartsheetSetTabColor(int32_t color)
{
  if (chartsheet)
    {chartsheet_set_tab_color(chartsheet, color);}
}

void worksheetSetDefaultRow(double height,
			    uint8_t hideUnusedRows)
{
  if (!worksheet)
    {return;}
  worksheet_set_default_row(worksheet, height, hideUnusedRows);
}

void worksheetProtect(char* pswd,
		      uint8_t 	no_select_locked_cells,
		      uint8_t 	no_select_unlocked_cells,
		      uint8_t 	format_cells,
		      uint8_t 	format_columns,
		      uint8_t 	format_rows,
		      uint8_t 	insert_columns,
		      uint8_t 	insert_rows,
		      uint8_t 	insert_hyperlinks,
		      uint8_t 	delete_columns,
		      uint8_t 	delete_rows,
		      uint8_t 	sort,
		      uint8_t 	autofilter,
		      uint8_t 	pivot_tables,
		      uint8_t 	scenarios,
		      uint8_t 	objects,
		      uint8_t 	no_content,
		      uint8_t 	no_objects)
{
  if (worksheet)
    {
      lxw_protection o = (lxw_protection)
	{.no_select_locked_cells = no_select_locked_cells,
	 .no_select_unlocked_cells = no_select_unlocked_cells,
	 .format_cells = format_cells,
	 .format_columns = format_columns,
	 .format_rows = format_rows,
	 .insert_columns = insert_columns,
	 .insert_rows = insert_rows,
	 .insert_hyperlinks = insert_hyperlinks,
	 .delete_columns = delete_columns,
	 .delete_rows = delete_rows,
	 .sort = sort,
	 .autofilter = autofilter,
	 .pivot_tables = pivot_tables,
	 .scenarios = scenarios,
	 .objects = objects,
	 .no_content = no_content,
	 .no_objects = no_objects};
      worksheet_protect(worksheet, pswd ? pswd : "", &o);
    }
}

void chartsheetProtect(char* pswd,
		       uint8_t 	no_select_locked_cells,
 		       uint8_t 	no_select_unlocked_cells,
 		       uint8_t 	format_cells,
 		       uint8_t 	format_columns,
 		       uint8_t 	format_rows,
 		       uint8_t 	insert_columns,
 		       uint8_t 	insert_rows,
 		       uint8_t 	insert_hyperlinks,
 		       uint8_t 	delete_columns,
 		       uint8_t 	delete_rows,
 		       uint8_t 	sort,
 		       uint8_t 	autofilter,
 		       uint8_t 	pivot_tables,
 		       uint8_t 	scenarios,
 		       uint8_t 	objects,
 		       uint8_t 	no_content,
 		       uint8_t 	no_objects)
{
  if (chartsheet)
    {
      lxw_protection o = (lxw_protection)
	{.no_select_locked_cells = no_select_locked_cells,
	 .no_select_unlocked_cells = no_select_unlocked_cells,
	 .format_cells = format_cells,
	 .format_columns = format_columns,
	 .format_rows = format_rows,
	 .insert_columns = insert_columns,
	 .insert_rows = insert_rows,
	 .insert_hyperlinks = insert_hyperlinks,
	 .delete_columns = delete_columns,
	 .delete_rows = delete_rows,
	 .sort = sort,
	 .autofilter = autofilter,
	 .pivot_tables = pivot_tables,
	 .scenarios = scenarios,
	 .objects = objects,
	 .no_content = no_content,
	 .no_objects = no_objects};
      chartsheet_protect(chartsheet, pswd ? pswd : "", &o);
    }
}

void chartsheetSetZoom(uint16_t scale)
{
  if (chartsheet)
    {chartsheet_set_zoom(chartsheet, scale);}
}

void chartsheetSetLandscape()
{
  if (chartsheet)
    {chartsheet_set_landscape(chartsheet);}
}

void chartsheetSetPortrait()
{
  if (chartsheet)
    {chartsheet_set_portrait(chartsheet);}
}

void chartsheetSetPaper(uint8_t paperType)
{
  if (chartsheet)
    {chartsheet_set_paper(chartsheet, paperType);}
}

void chartsheetSetMargins(double left,
			  double right,
			  double top,
			  double bottom)
{
  if (chartsheet)
    {
      chartsheet_set_margins(chartsheet,
			     left,
			     right,
			     top,
			     bottom);
    }
}

void chartsheetSetHeader(char* headerStr)
{
  if (chartsheet)
    {
      const char* s = headerStr;
      handleLXWError(chartsheet_set_header(chartsheet, s),
		     "chartsheet-set-header");
    }
}

void chartsheetSetFooter(char* footerStr)
{
  if (chartsheet)
    {
      const char* s = footerStr;
      handleLXWError(chartsheet_set_footer(chartsheet, s),
		     "chartsheet-set-header");
    }
}


void chartsheetSetHeaderOpt(char* headerStr, double margin)
{
  if (chartsheet)
    {
      lxw_header_footer_options o = (lxw_header_footer_options){.margin = margin};
      const char* s = headerStr;
      handleLXWError(chartsheet_set_header_opt(chartsheet, s, &o),
		     "chartsheet-set-header-opt");      
    }
}

void chartsheetSetFooterOpt(char* footerStr, double margin)
{
  if (chartsheet)
    {
      lxw_header_footer_options o = (lxw_header_footer_options){.margin = margin};
      const char* s = footerStr;
      handleLXWError(chartsheet_set_footer_opt(chartsheet, s, &o),
		     "chartsheet-set-footer-opt");      
    }
}

void initRichStringList(ptrdiff_t allocateN)
{
  maxAllowedRichStrings = allocateN;
  richStringList = calloc(allocateN, sizeof(lxw_rich_string_tuple**));
  richStringCount = 0;
}

void initRichString(ptrdiff_t allocateN)
{
  richString = calloc(allocateN, sizeof(lxw_rich_string_tuple*));
  richStringFragments = calloc(allocateN, sizeof(lxw_rich_string_tuple));
  rsFragmentStrings = calloc(allocateN , sizeof(char*));
  maxAllowedRichStringFragments = allocateN;
  richStringFragmentCount = 0;
}

void createRichStringFragment(char* stringChunk)
{
  if (!richStringFragmentCount)
    {richStringFragmentCount = 0;}
  if (richStringFragments &&
      richStringFragmentCount < maxAllowedRichStringFragments)
    {
      int i = 1;
      while (stringChunk[i])
	{i++;}
      rsFragmentStrings[richStringFragmentCount] = calloc(i,sizeof(char));
      strcpy(rsFragmentStrings[richStringFragmentCount], stringChunk);
      lxw_format* f = NULL;
      if (formats)
	{f = formats[formatIndex];}
      *(richStringFragments + richStringFragmentCount) =
	(lxw_rich_string_tuple)
	{.format = f,
	 .string = rsFragmentStrings[richStringFragmentCount]};
      richStringFragmentCount+=1;
    }
}

void richStringCleanup()
{
  if (richString)
    {free(richString);}
  if (richStringFragments)
    {free(richStringFragments);}
  richStringFragmentCount = 0;
  maxAllowedRichStringFragments = 0;
}

void richStringListCleanup()
{
  if (richStringList)
    {
      free(richStringList);
      richStringList = NULL;
    }
  for (int i = 0; i<richStringFragmentCount; i++)
    {
      if(rsFragmentStrings[i])
	{
	  free(rsFragmentStrings[i]);
	  rsFragmentStrings[i] = NULL;
	}
    }
  if (rsFragmentStrings)
    {
      free(rsFragmentStrings);
      rsFragmentStrings = NULL;
    }
  richStringCount = 0;
  maxAllowedRichStrings = 0;
}

void worksheetSetTabColor(int32_t color)
{
  if (worksheet)
    {worksheet_set_tab_color(worksheet, color);}
}

void worksheetWriteRichStringFragments()
{
  if (!worksheet)
    {return;}
  if (!richStringCount)
    {initRichStringList(1);}  
  if (richStringFragmentCount < 2)
    {
      printf("worksheet-write-rich-string error, rich string needs more than 1 fragment.\n");
      return;
    }
  if (richStringCount >= maxAllowedRichStrings)
    {
      int c = richStringCount + 1;
      lxw_rich_string_tuple*** r = NULL; 
      r = realloc(richStringList, c * sizeof(lxw_rich_string_tuple**));
      if (!r)
	{
	  printf("error, problem reallocating memory in worksheet-write-rich-string-fragments\n");
	  return;
	}
      richStringList = r;
      maxAllowedRichStrings+=1;
    }
  lxw_format* f = NULL;
  if (formats)
    {f = formats[formatIndex];}
  int i;
  for (i = 0; i < richStringFragmentCount; i++)
    {richString[i] = &richStringFragments[i];}
  *(richStringList + richStringCount) = richString;
  worksheet_write_rich_string(worksheet, row, col, richStringList[richStringCount], f);
  richStringCount+=1;
}

void allWorksheetsSetColumn(double width,
			    unsigned short stCol,
			    unsigned short endCol)
{
  lxw_format* f = NULL;
  if (formats)
    {f = formats[formatIndex];}  
  LXW_FOREACH_WORKSHEET(worksheet, workbook){
    worksheet_set_column(worksheet, stCol, endCol, width, f);
  }
}

void allWorksheetsSetRow(double height,
			 uint32_t row)
{
  lxw_format* f = NULL;
  if (formats)
    {f = formats[formatIndex];}  
  LXW_FOREACH_WORKSHEET(worksheet, workbook){
    worksheet_set_row(worksheet, row, height, f);
  }
}


void allWorksheetsWriteString(char* val)
{
  lxw_format* f = NULL;
  if (formats)
    {f = formats[formatIndex];}  
  LXW_FOREACH_WORKSHEET(worksheet, workbook){
    worksheet_write_string(worksheet, row, col, val, f);
  }
}

void allWorksheetsWriteFormula(char* val)
{
  lxw_format* f = NULL;
  if (formats)
    {f = formats[formatIndex];}  
  LXW_FOREACH_WORKSHEET(worksheet, workbook){
    worksheet_write_formula(worksheet, row, col, val, f);
  }
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
  dataValidationsCleanup();
  seriesCleanup();
  seriesErrorBarsCleanup();
  chartsCleanup();
  richStringCleanup();
  richStringListCleanup();
  colOptionCleanup();
}


