@ECHO OFF


@REM :: REM run this if script dependencies are not installed

@REM python -m pip install -r requirements.txt



@REM Other examples with spss
python "%~dp0correlations.py" ^
    --inpfile "..\Tables\Data\R2401543.sav" ^
    --pattern_regex "^M(?:5|6|7|8|9)_\d{3}_001" ^
    --filter "DV_LinkType==2" ^
    --outfile "R2401543_test_clientsample_amazonrelay.xlsx"


python "%~dp0correlations.py" ^
    --inpfile "..\Tables\Data\R2401543.sav" ^
    --pattern_regex "^M(?:5|6|7|8|9)_\d{3}_002" ^
    --filter "DV_LinkType==2" ^
    --outfile "R2401543_test_clientsample_truckstop.xlsx"


python "%~dp0correlations.py" ^
    --inpfile "..\Tables\Data\R2401543.sav" ^
    --pattern_regex "^M(?:5|6|7|8|9)_\d{3}_003" ^
    --filter "DV_LinkType==2" ^
    --outfile "R2401543_test_clientsample_datloadboard.xlsx"


python "%~dp0correlations.py" ^
    --inpfile "..\Tables\Data\R2401543.sav" ^
    --pattern_regex "^M(?:5|6|7|8|9)_\d{3}_004" ^
    --filter "DV_LinkType==2" ^
    --outfile "R2401543_test_clientsample_123loadboard.xlsx"


python "%~dp0correlations.py" ^
    --inpfile "..\Tables\Data\R2401543.sav" ^
    --pattern_regex "^M(?:5|6|7|8|9)_\d{3}_005" ^
    --filter "DV_LinkType==2" ^
    --outfile "R2401543_test_clientsample_rxodrive.xlsx"


python "%~dp0correlations.py" ^
    --inpfile "..\Tables\Data\R2401543.sav" ^
    --pattern_regex "^M(?:5|6|7|8|9)_\d{3}_001" ^
    --filter "DV_LinkType!=2" ^
    --outfile "test_panelsample_amazonrelay.xlsx"


python "%~dp0correlations.py" ^
    --inpfile "..\Tables\Data\R2401543.sav" ^
    --pattern_regex "^M(?:5|6|7|8|9)_\d{3}_002" ^
    --filter "DV_LinkType!=2" ^
    --outfile "test_panelsample_truckstop.xlsx"


python "%~dp0correlations.py" ^
    --inpfile "..\Tables\Data\R2401543.sav" ^
    --pattern_regex "^M(?:5|6|7|8|9)_\d{3}_003" ^
    --filter "DV_LinkType!=2" ^
    --outfile "test_panelsample_datloadboard.xlsx"


python "%~dp0correlations.py" ^
    --inpfile "..\Tables\Data\R2401543.sav" ^
    --pattern_regex "^M(?:5|6|7|8|9)_\d{3}_004" ^
    --filter "DV_LinkType!=2" ^
    --outfile "test_panelsample_123loadboard.xlsx"


python "%~dp0correlations.py" ^
    --inpfile "..\Tables\Data\R2401543.sav" ^
    --pattern_regex "^M(?:5|6|7|8|9)_\d{3}_005" ^
    --filter "DV_LinkType!=2" ^
    --outfile "test_panelsample_rxodrive.xlsx" ^
