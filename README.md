# Galton-Watson process simulation 
  
This program was created as a part of Probability and Statistics class to Simulates Galton Watson process.  
  
Outputs a .xlsx file, with data and few charts. Name of the file can by specified by `-o` argument.  
Default output file consist of input parameters. 
  
Use `--help` to get all optional arguments. 
  
Because Excel can't handle more then 256 series in one chart, if initial entities (`-n`) exceeds this number,  
chart won't be complete. However, chart will be created using mathplotlib in separate .png file.  
But this process can take a long time. At this time, program can be terminated safely, without losing .xlsx file. 
