@ECHO OFF
ECHO Deploying publish file from 
ECHO "..\..\Publish\PowerPoint Warrior" 
ECHO to Azure blob storage

PAUSE

"c:\Program Files (x86)\Microsoft SDKs\Azure\AzCopy\AzCopy.exe" "..\..\Publish\PowerPoint Warrior" https://ppwarrior.blob.core.windows.net/install /DestKey:bjBXjyOJHOXamFjRAjn5YGLr9xfvpZv4fxVb8SZ+JY5GFvmjWFKIGlQD4rowpTrbopsiev4v320F8N6osuBk0A== /S

PAUSE