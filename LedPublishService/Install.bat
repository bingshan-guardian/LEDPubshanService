%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\installutil.exe C:\LEDPublishService\LedPublishService.exe
Net Start ServiceEFNETSYS
sc config ServiceEFNETSYS start= auto
pause