# SmoothScrollAutoRefresh (SSAR)
A script that silently auto-restart SmoothScroll.exe to eliminate annoying License Validation Check Prompt.\
![image](https://github.com/TatshSiow/SmoothScrollAutoRefresh/assets/100989709/d55e3615-bdf1-4aaf-b5f5-832f6722f534)


## How does it work?
This script will automatically restarts SmoothScroll periodically, endlessly and silently.\
It won't prompt anything on each execution intervals, letting you focus on your foreground activities.


## So so so, how to use this?
![image](https://github.com/TatshSiow/SmoothScrollAutoRefresh/assets/100989709/58cfd9f1-83b0-4ace-a6e7-3b64cf607ee4)

After launching the **SSAR_(1.x).vbs**, it will shows like what you see above.\
Yes: Run AutoRefresh Service (it will prompt "Service Activated")\
No: Stop AutoRefresh Service (it will prompt "Service Deactivated")\
Cancel : Exit the VBS (hmmm, to be more specific, just a do nothing button)

## Important Notes
**SmoothScroll.exe** needs to be **running** before launching the service.\
Or else the service won't be able to find the origin location of SmoothScroll.exe.

**Save your CMD works** before deactivate service!\
It will kill all **cmd.exe** and **conhost.exe** to avoid/eliminate memory leaks.

Feedbacks are appreciated.

## Disclaimer

* By using this script, you acknowledge and agree to the automated restart task.

* Scheduled Restarts: This automated script is executed at specific intervals to refresh the application. These restarts are performed seamlessly and with minimal (nearly zero) impact on your experience.

* User Impact: While I have taken measures to minimize any inconvenience, Users will lose Command Prompt and Console Host sessions when executing task deactivation. Save your work to avoid potential disruptions.

* I'M NOT RESPONSIBLE FOR YOUR DATA LOSS FOR NOT SAVING THE COMMAND PROMPT LINES.

## Legal Disclaimer

* I shall not be held liable for any damages, losses, or legal action arising from the implementation of this automated restart script. 
  
* Users are advised to review and comply with the terms of service of this application.
  
* In the event of any disagreement or potential legal action, users and I agree to resolve disputes through arbitration or other alternative dispute resolution mechanisms.
  
* Thank you for your understanding.
