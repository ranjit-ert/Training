cd %ANDROID_HOME%\plat*
adb connect %1
adb shell dumpsys battery > C:\temp\Battery-%1-%2.txt