LD_LIBRARY_PATH=$LD_LIBRARY_PATH:/usr/local/lib/
export LD_LIBRARY_PATH
chicken-wrap libxlsxwriter_layer.c;
csi processBuildFile.scm;
csc -s xlsxwriterscm.scm -emit-import-library xlsxwriterscm -lxlsxwriter -lz;
cp *.so tests
cp *.import.scm tests
mv *.so bin
mv *.import.scm bin
