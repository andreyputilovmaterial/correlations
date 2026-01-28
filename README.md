# Correlations tester, a python script
A helper python script to calculate correlations.

The result is a beautiful Excel, formatted similar way what MADs does. The script creates an Excel with data highlighted: orange/red for stronger correlations >0.5 and light/yellow for moderate values >0.3.

Input can be SPSS or MDD.

As parameters, you provide 1. file name, 2. variables to choose, through "mask" or regular expression, 3. (optional) case data filter, to analyze only within link type, provider, or certain wgt group, 4. (optional) output file name - if not provided, a file with "correlations" added will be created at the same location as input data file.

This can be used:
- to inspect data quality (good data will show strong correlations in attributes, low-quality bot-generated or dummy data will not show any correlations)
- to inspect why weighting groups are not working, if there is no other obvious reason

The tool treats everything as a multi-punch variable. Single-punch or text variables can also be used and are categorized. Everything is converted to 1/0 flag, same as stubs of multi-punch in spss. If you have single-punch variables tested, responses within one variables will also be checked for correlations against each other (as each response is 1/0 flag), but you can just ignore those. Some single-punch variables, like WGT_Device, show phi=1 (100%) correlation. Like Mobile and Non-Mobile correlate. Some codeframe vartiables can show phi=0.06 correlation - weak but still positive. Just ignore this, if compared categorories are from the same single-punch variable. Sometimes Mobile and Non-Mobile categories do not show correlation with phi=1 but with phi=0.99876, because of Yates correction. This is normal, but the correction should be off, so expected phi=1. Comparing responses from multi-punch variables normally makes sense - we can compare selected attributes within one single var, so not everything should be ignored. Use common sense.

What you need is that python .py file, but you can also grab the BAT file to launch it easier and adjust for multiple subgroups faster.

Input parameters are:  
- **`--inpfile`** the path to your MDD or SPSS file  
- **`--outfile`** the desired file name for the resulting excel (can be omitted)  
- **`--pattern_regex`** or **`--pattern_mask`** (one of those 2 is required)  
    - **`--pattern_regex`** is the regex pattern to select variables. Advanced way to provide flexible syntax to select variables. For example, if you need to look at M5-M9 which are multi-punch within a loop with brands, and variable names in spss are M5\_(3-digit attribute code)\_(3-digit brand code), and you only need to check brand 001 - the pattern would be `--pattern_regex "^M(?:5|6|7|8|9)_\d{3}_001$"`. Note: if providing this from BAT file or command prompt, some characters have special meaning. For example, `|` and `^`. Use carefully. 
    - **`--pattern_mask`** the wildcarded pattern to select variables, with variables separated with comma. For example, `DV_AgeGender,DV_BannerWGT,SampleProvider,BrandLoop[{Tesla}].KidGender`. "`?`" stands for any one single character, "`*`" stands for any sequence (including empty) of characters.  
- **`--filter`** data filter expression.<br>In MDD, just pass the normal expression, as you do. For example, "`DV_LinkType.ContainsAny({Adult})`"<br>In SPSS, as an example, the expression could be `--filter "DV_LinkType==2"` or `--filter "DV_LinkType!=2"`. Syntax used is what is used for pandas.query(), or pandas.eval(): "==" for checking if something is equal, "!=" for not equal, "&" for "and" and "&" for "or". Besides that, normal arithmetic, boolean, comparison operators can be used, common functions - see https://pandas.pydata.org/docs/reference/api/pandas.eval.html#pandas.eval
- **`--chi2_contingency_correction`** is an optional flag to set Yates correction to true or false. If not passed, "false" is used, as it's better for 1/0  flags. However, as scipy tools are defined, the default is correction=true. But we are using correction=false here, if not overriden with this param here

## How to use:

Download files from the latest Release page:
[Releases](../../releases/latest)

See the BAT files provided as an example.

Enjoy.

If some python packages are missing (the script has minimum requirements but it needs pandas, numpy, openpyxl, scipy, pathlib, arglib, etc..), just type
`python -m pip install -r requirements.txt`


