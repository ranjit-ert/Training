::********************************Modify the attrbiutes below to meet you requirements**************************************

:: 3 things you must tell the tool...
:: Set the 'debugFolder' so that this tool can find all the *.feature files within it recursively
:: Set the features dll that the features will be run against
:: Tell the tool where the Libraries folder is in relation to where this tool is executed

set debugFolder=TestTeamAutomationFrameNew\bin\Debug;
set featuresDll=TestTeamAutomationFrameNew\bin\Debug\TestTeamAutomationFrameNew.dll
set librariesFolder=Lib

::********************************No need to modify anything below here*****************************************************

:: Get all the feature files in the Debug folder
SETLOCAL EnableDelayedExpansion
set featureFiles=
for /r %debugFolder% %%X in (*.Feature) do ( SET featureFiles=!featureFiles!;%%X)
::Remove the leading semi-colon
SET featureFiles=%featureFiles:*;=% 

:: Starting Feature Creation, provide the dll and feature files, if more than one dir for feature files then seperate with ';' and make sure the entire feature files string is enclosed within quotes (as ; is a special symbol), e.g. "bin\Debug\*.feature;bin\Debug\MyOtherFeatures*.feature"
call %librariesFolder%\NBehaveEIT\ExecuteUserStories.bat %librariesFolder% %featuresDll% "%featureFiles%"
:: To view the output as html you need to set up a virtual directory on a folder called 'FeatureOutput' in the same dir as this bat file.
:: so change the URL as appropriate
call %librariesFolder%\XmlTransformCmdLineTool\TransformXml.exe -i "FeatureOutput\Features.xml" -s "%librariesFolder%\NBehaveEIT\NBehaveResultsTemplate.xsl" -o "FeatureOutput\Output.html"
:: start file://%CD%/FeatureOutput/Output.html