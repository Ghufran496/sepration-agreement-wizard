@echo off
echo Stopping any running instances of the API...
taskkill /IM dotnet.exe /F

echo Starting API in HTTP mode (port 5254)...
cd DocumentGenerationApi
start "DocumentGenerationApi-HTTP" cmd /c "dotnet run --urls=http://localhost:5254"

echo API should be starting now. Check browser at:
echo http://localhost:5254/api/document/ping

echo.
echo If you have issues, ensure ports 5254 are not in use by other applications.
pause 