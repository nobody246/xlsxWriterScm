# xlsxwriter-scm
chicken-scheme libxlsxwriter bindings

# Requirements
build file uses posix and s eggs

built chicken scheme source file uses easyffi, data-structures eggs


The c dependencies are: libz, libxslxwriter c libraries.


# Usage
1) with successfully running build.sh after installing all required dependencies, ./bin folder will contain compiled files. 

2) to include in chicken-scheme: (use xlsxwriterscm)

Basic functionality seems good to go but I still find things to fix continuously, For reference check the /tests folder for examples as well as the documentation of libxlsxwriter https://libxlsxwriter.github.io/.

This is released under same license as libxlsxwriter (BSD).


# FYI

When I updated Libxlsxwriter and rebuilt xlsxWriterScm I had to run ldconfig to get everything to run. You may or may not run into this issue, I'm not sure why it happened.
