// (c) xlsxWriterScm 2019-2025 Alexander Semotan
// Released under the BSD 2 Software License 
void handleLXWError(lxw_error error, char* calledFromFunc)
{
  if (error)
    {
      printf("Error in (%s), num: %d = %s\n", 
	     calledFromFunc ,
	     error, 
	     lxw_strerror(error));
    }
}
