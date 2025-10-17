REM Step 1: Autobuild solution in Release-mode
set VSVER=[17.0^,18.0^)

::Edit path if VS 2022 is installed on other path
call "C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvarsall.bat" x64

rmdir /s /q docxGenerator4DataBin

::Build VinnyLibConverter
devenv docxGenerator4Data.sln /Build "Release|Any CPU"

::For net8.0
xcopy .\README.md "bin\Release\net8.0" /Y /I
xcopy .\LICENSE.md "bin\Release\net8.0" /Y /I

xcopy bin\Release\*.* docxGenerator4DataBin /Y /I /E

::ZIP release
del "docxGenerator4DataBin.zip"
"C:\Program Files\7-Zip\7z" a -tzip "docxGenerator4DataBin.zip" docxGenerator4DataBin
rmdir /s /q docxGenerator4DataBin

::@endlocal
::@exit /B 1
