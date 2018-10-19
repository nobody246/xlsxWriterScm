lxw_workbook*         workbook;
lxw_worksheet*        worksheet;
lxw_chartsheet*       chartsheet;
uint32_t              row = 0;
unsigned short        col = 0;


lxw_chart**            charts;
ptrdiff_t maxAllowedCharts = 0;
ptrdiff_t chartCount = 0;
ptrdiff_t chartIndex = 0;

lxw_chart_series**     series;
ptrdiff_t maxAllowedSeries = 0;
ptrdiff_t seriesCount = 0;
ptrdiff_t seriesIndex = 0;

lxw_chart_font*        chartFonts;
ptrdiff_t maxAllowedChartFonts;
ptrdiff_t chartFontCount = 0;
ptrdiff_t chartFontIndex = 0;

lxw_format**          formats;
ptrdiff_t maxAllowedFormats = 0;
ptrdiff_t formatCount = 0;
ptrdiff_t formatIndex = 0;

lxw_row_t* rowNumbers;
ptrdiff_t maxAllowedRowNumbers = 0;
ptrdiff_t rowNumberCount = 0;            
 
lxw_col_t* colNumbers;
ptrdiff_t maxAllowedColNumbers = 0;
ptrdiff_t colNumberCount = 0;            

lxw_chart_line* chartLines;
ptrdiff_t maxAllowedChartLines = 0;
ptrdiff_t chartLineCount = 0;
ptrdiff_t chartLineIndex = 0;
 
lxw_chart_point* chartPoints;
ptrdiff_t maxAllowedChartPoints = 0;
ptrdiff_t chartPointCount = 0;
ptrdiff_t chartPointIndex = 0;

lxw_chart_fill* chartFills;
ptrdiff_t maxAllowedChartFills = 0;
ptrdiff_t chartFillCount = 0;
ptrdiff_t chartFillIndex = 0;

lxw_chart_pattern* chartPatterns;
ptrdiff_t maxAllowedChartPatterns = 0;
ptrdiff_t chartPatternCount = 0;
ptrdiff_t chartPatternIndex = 0;

lxw_data_validation* dataValidations;
ptrdiff_t maxAllowedDataValidations = 0;
ptrdiff_t dataValidationCount = 0;
ptrdiff_t dataValidationIndex = 0;

lxw_series_error_bars** seriesErrorBars;
ptrdiff_t maxAllowedSeriesErrorBars = 0;
ptrdiff_t seriesErrorBarsCount = 0;
ptrdiff_t seriesErrorBarsIndex = 0;


char** dataValidationList;

char* errorStr;

unsigned short gridLineXAxisShow = 0;
unsigned short gridLineYAxisShow = 0;

uint16_t* deleteSeries;
ptrdiff_t deleteSeriesCount;

lxw_chart_line*  upBarLine;
lxw_chart_fill*  upBarFill;
lxw_chart_line*  downBarLine;
lxw_chart_fill*  downBarFill;

lxw_rich_string_tuple*** richStringList;
ptrdiff_t richStringCount;
ptrdiff_t maxAllowedRichStrings;
lxw_rich_string_tuple** richString;
lxw_rich_string_tuple* richStringFragments;
char** rsFragmentStrings;
ptrdiff_t richStringFragmentCount;
ptrdiff_t maxAllowedRichStringFragments;

