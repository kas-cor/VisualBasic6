@echo off
copy mswinsck.ocx c:\windows\system32
regsvr32 c:\windows\system32\mswinsck.ocx