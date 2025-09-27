# Correlations tester, a python script
A helper python script that creates an Excel with data colored with different colors (orange/red for stronger correlations >0.5 and light/yellow for moderate values >0.3). It takes SPSS as input. It's something similar to what MADs is doing.

If we are having issues with data quality we can run it on our side and test different groups, like different providers, and identify better sample (which shows stronger correlations), or worse, more like randomly-generated sample, where there will not be any strong correlations.

For now,it tests multi-punch variables only (I believe other types are easy to do too but I did not need it and did not add possibility to calculate it yet).

What you need is that python .py file, but you can also grab the BAT file to launch it easier and adjust for multiple subgroups faster.

Input parameters are:  
--inpfile the path to your SPSS file  
--outfile the desired file name for the resulting excel (can be omitted)  
--pattern (required)  the regex pattern to select variables. For example ,if you need to look at M5-M9 which are multi-punch within a loop with brands, and variable names in spss are M5\_(3-digit attribute code)\_(3-digit brand code), and you only need to check brand 001, the pattern would be `--pattern "^M(?:5|6|7|8|9)_\d{3}_001$"`  
--filter data filter expression. For example, `--filter "DV\_LinkType==2"` or `--filter "DV\_LinkType!=2"`

## How to use:

Download files from the latest Release page:
[Releases](../../releases/latest)

Then edit parameters in the bat file - add as many calls as you need. Update paths to spss, data filters, output file name, etc... Just follow the same example as what you see in that bat file. Currently there are 10 calls for amazon 2401543

Enjoy.

If some python packages are missing (the script has minimum requirements but it needs pandas, numpy, openpyxl, scipy, pathlib, arglib, etc..), just type
`python -m pip install xxx`
where xxx is a missing package.

