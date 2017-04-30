Set OpSysSet = GetObject("winmgmts:{(Shutdown)}//./root/cimv2").ExecQuery("select *
from Win32_OperatingSystem where Primary=true")

for each OpSys in OpSysSet
     OpSys.ShutDown()
next