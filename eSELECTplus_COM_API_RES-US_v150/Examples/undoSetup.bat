@echo off
Echo MONERIS MPG COM API
Echo ===================
Echo This will unregister all Moneris MPG DLLs.
Echo The process will pop up a dialog box for each DLL unregistered.
Echo Click okay in each of these dialog boxes.
echo ------------------------------------------------------------
Echo To ABORT hit ctrl-c at this time and answer y
echo ------------------------------------------------------------
@echo on

@pause

regsvr32 -u MonerisUSCOMAPI.dll

for %%x in (Requests\*.dll) do regsvr32 -u %%x

@echo off
echo ------------------------------------------------------------
echo uninstall complete
echo ------------------------------------------------------------
@pause
notepad doc.txt