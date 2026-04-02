@echo off
echo ========================================
echo    Registry Expert - Build Script
echo ========================================
echo.

:: Check if dotnet is available
where dotnet >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo ERROR: .NET SDK not found. Please install .NET 8 SDK.
    echo Download from: https://dotnet.microsoft.com/download/dotnet/8.0
    pause
    exit /b 1
)

echo [1/2] Restoring packages...
dotnet restore RegistryExpert.Wpf\RegistryExpert.Wpf.csproj
if %ERRORLEVEL% neq 0 (
    echo ERROR: Package restore failed
    pause
    exit /b 1
)

echo.
echo [2/2] Building single-file executable...
dotnet publish RegistryExpert.Wpf\RegistryExpert.Wpf.csproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -o .\publish-wpf
if %ERRORLEVEL% neq 0 (
    echo ERROR: Build failed
    pause
    exit /b 1
)

echo.
echo ========================================
echo    Build Complete!
echo ========================================
echo.
echo The executable is located at:
echo   .\publish-wpf\RegistryExpert.exe
echo.
echo This is a single portable .exe file.
echo.
pause
