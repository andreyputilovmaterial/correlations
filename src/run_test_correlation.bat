@ECHO OFF

python test_correlation.py --inpfile ..\Tables\Data\R2401543.sav --outfile test_clientsample_amazonrelay.xlsx --pattern "^M(?:5|6|7|8|9)_\d{3}_001" --filter "DV_LinkType==2"
python test_correlation.py --inpfile ..\Tables\Data\R2401543.sav --outfile test_clientsample_truckstop.xlsx --pattern "^M(?:5|6|7|8|9)_\d{3}_002" --filter "DV_LinkType==2"
python test_correlation.py --inpfile ..\Tables\Data\R2401543.sav --outfile test_clientsample_datloadboard.xlsx --pattern "^M(?:5|6|7|8|9)_\d{3}_003" --filter "DV_LinkType==2"
python test_correlation.py --inpfile ..\Tables\Data\R2401543.sav --outfile test_clientsample_123loadboard.xlsx --pattern "^M(?:5|6|7|8|9)_\d{3}_004" --filter "DV_LinkType==2"
python test_correlation.py --inpfile ..\Tables\Data\R2401543.sav --outfile test_clientsample_rxodrive.xlsx --pattern "^M(?:5|6|7|8|9)_\d{3}_005" --filter "DV_LinkType==2"
python test_correlation.py --inpfile ..\Tables\Data\R2401543.sav --outfile test_panelsample_amazonrelay.xlsx --pattern "^M(?:5|6|7|8|9)_\d{3}_001" --filter "DV_LinkType!=2"
python test_correlation.py --inpfile ..\Tables\Data\R2401543.sav --outfile test_panelsample_truckstop.xlsx --pattern "^M(?:5|6|7|8|9)_\d{3}_002" --filter "DV_LinkType!=2"
python test_correlation.py --inpfile ..\Tables\Data\R2401543.sav --outfile test_panelsample_datloadboard.xlsx --pattern "^M(?:5|6|7|8|9)_\d{3}_003" --filter "DV_LinkType!=2"
python test_correlation.py --inpfile ..\Tables\Data\R2401543.sav --outfile test_panelsample_123loadboard.xlsx --pattern "^M(?:5|6|7|8|9)_\d{3}_004" --filter "DV_LinkType!=2"
python test_correlation.py --inpfile ..\Tables\Data\R2401543.sav --outfile test_panelsample_rxodrive.xlsx --pattern "^M(?:5|6|7|8|9)_\d{3}_005" --filter "DV_LinkType!=2"
