# Correlations tester, a python script
A helper python script to calculate correlations.

The result is a beautiful Excel, formatted similar way what MADs does. The script creates an Excel with data highlighted: orange/red for stronger correlations >0.5 and light/yellow for moderate values >0.3.

Input can be SPSS or MDD.

As parameters, you provide 1. file name, 2. variables to choose, through "mask" or regular expression, 3. (optional) case data filter, to analyze only within link type, provider, or certain wgt group, 4. (optional) output file name - if not provided, a file with "correlations" added will be created at the same location as input data file.

This can be used:
- to inspect data quality (good data will show strong correlations in attributes, low-quality bot-generated or dummy data will not show any correlations)
- to inspect why weighting groups are not working, if there is no other obvious reason

The tool can only inspect multi-punch variables. Technically, anuthing that is a number is working (all categoricals in spss).

What you need is that python .py file, but you can also grab the BAT file to launch it easier and adjust for multiple subgroups faster.

Input parameters are:  
`--inpfile` the path to your SPSS file  
`--outfile` the desired file name for the resulting excel (can be omitted)  
`--pattern_regex` or `--pattern_mask` (one of those 2 is required)  
`--pattern_regex` the regex pattern to select variables. For example, if you need to look at M5-M9 which are multi-punch within a loop with brands, and variable names in spss are M5\_(3-digit attribute code)\_(3-digit brand code), and you only need to check brand 001, the pattern would be `--pattern_regex "^M(?:5|6|7|8|9)_\d{3}_001$"`  
`--pattern_mask` the wildcarded pattern to select variables. "`?`" stands for any one single character, "`*`" stands for any sequence (including empty) of characters. This approach is for those non-technical-savvy who struggle adjusting regular expressions. But this is very limited in its functionality. For example, you need to test all variables in M5, M6, M7, M8, M9 batteries. What can you do? Filter `M*_001`? This way you also test M1, M2, M3, M4... Or, use `M5*_001` pattern? Then you are only testing attributes within M5 and testing it against each other but not against attributes from M6, M7, M8, M9.  
`--filter` data filter expression. For example, `--filter "DV_LinkType==2"` or `--filter "DV_LinkType!=2"`

## How to use:

Download files from the latest Release page:
[Releases](../../releases/latest)

Then edit parameters in the bat file - add as many calls as you need. Update paths to spss, data filters, output file name, etc... Just follow the same example as what you see in that bat file. Currently there are 10 calls for amazon 2401543

Enjoy.

If some python packages are missing (the script has minimum requirements but it needs pandas, numpy, openpyxl, scipy, pathlib, arglib, etc..), just type
`python -m pip install xxx`
where xxx is a missing package.

