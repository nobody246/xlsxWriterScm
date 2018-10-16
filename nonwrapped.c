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
