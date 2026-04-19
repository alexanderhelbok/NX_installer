@echo on
 
set JAVA_HOME=C:\Apps\java\zulu11
set NX_HOME=C:\Apps\NX
set Path=%JAVA_HOME%\bin;%path%
set JRE_HOME=%JAVA_HOME%
set JRE64_HOME=%JAVA_HOME%
set FMS_HOME=%NX_HOME%\UGMANAGER\tccs
set UGII_BASE_DIR=%NX_HOME%
rem set UGII_ROOT_DIR=%UGII_BASE_DIR%\NXBIN
set UGII_UGMGR_COMMUNICATION=HTTP
set UGII_LANG=english
set UGII_UGMGR_HTTP_URL=http://10.10.220.12:3000/tc
set SPLM_LICENSE_SERVER=28000@10.10.220.13
rem set UGS_LICENSE_BUNDLE=ACD10
rem set UGS_SESSION_BUNDLE=ACD10
set UGII_UGMGR_FMS_PARTIAL_FILE=DISABLED
:: Run Teamcenter Integration with NX
:: with Teamcenter as a background process
::------------------------------------------------------------------
start "" "%UGII_BASE_DIR%\ugii\ugraf.exe" -pim=yes -NX
