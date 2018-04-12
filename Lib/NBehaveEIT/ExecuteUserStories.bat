rmdir FeatureOutput /s /q
mkdir FeatureOutput
%1\NBehave\NBehave-Console.exe %2 /sf=%3 /storyOutput=FeatureOutput\features.txt /xml=FeatureOutput\Features.xml