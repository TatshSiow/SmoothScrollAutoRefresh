# SmoothScrollAutoRefresh (SSAR)
This is a simple script for automatically refresh SmoothScroll App License Validation

## How does it work?
This VBS script set a 1200 Seconds Timer when executed and will restart SmoothScroll Service repeatedly
![image](https://github.com/TatshSiow/SmoothScrollAutoRefresh/assets/100989709/16cdae27-9edd-4df7-a8da-8e68f957b45f)
### Yes: Run AutoRefresh Script
### No: Stop AutoRefresh Service
### Cancel : Exit the VBS

## Important Notes
### Don't simply think this code will work immediately!
### Setup your SmoothScroll Location before starting the service
### Or else, it won't be able to restart SmoothScroll

![image](https://github.com/TatshSiow/SmoothScrollAutoRefresh/assets/100989709/2b4ddcc4-c9a2-4f44-a02d-f1f82abd0acb)
#### Under Batch Script Code, locate your _SmoothScroll.exe_ 

## Another Important note for CMD users
### When you kill AutoRefresh Service, it will kill CMD.exe and openconsole.exe
### Make sure you are not using CMD while stopping the AutoRefresh Service
