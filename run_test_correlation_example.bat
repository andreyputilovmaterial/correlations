@ECHO OFF


@REM :: REM run this if script dependencies are not installed

@REM python -m pip install -r requirements.txt



@REM Hasbro WGT
python "%~dp0correlations.py" ^
    --inpfile "S:\Trackers\Hasbro Centralized Trackers\Hasbro MTG\Wave_2025Q3_2501152\Tables\Data\R2501152_FILTERCOMPLETED.mdd" ^
    --pattern_regex "\b(?:DV_US_IncomeNet|DV_US_EthnicityNet|DV_US_Region|ParentStatus|DV_FinalAgeGender|WGT_Device|DV_EngagementTiers|DV_PokemonEngageTiers)\b" ^
    --filter "WeightingStatus.ContainsAny({Completes}) AND DV_TargetVsNon_MTG.ContainsAny({Adult}) AND (DV_NestAugmentLink_MTG.ContainsAny({AdultGenpop}) OR (DV_NestAugmentLink_MTG.ContainsAny({ParentGenpop}) AND DV_FinalAgeGender.ContainsAny({Boys812, Boys1317, Girls812, Girls1317}))) and DV_WeightBanner.ContainsAny({UsAdult})" ^
    --outfile "R2501152_USAdult_correlations.xlsx"


