# xlsxwriter-scm
chicken-scheme libxlsxwriter bindings

uses easyffi, s, srfi-69 eggs, build file uses posix egg

utilizes libz, libxslxwriter c libraries.

run build.sh, bin folder will contain compiled files, 
(use xlsxwriterscm)

Very basic functionality is good to go and tested, but I will work on documenting and testing everything,
In the mean time check source files (particularly libxlsxwriter_layer.c) and https://libxlsxwriter.github.io/
and expect bugs.
