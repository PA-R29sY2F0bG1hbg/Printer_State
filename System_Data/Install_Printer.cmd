pnputil.exe -a "path_driver"

cd /d "C:\Windows\System32\Printing_Admin_Scripts\fr-FR"

# bind ip to created lpr port
Cscript prnport.vbs -a -r "Printer Local IP" -o "lpr" -h Printer Local IP
Cscript prnport.vbs -a -r "Printer Local IP" -o "lpr" -h Printer Local IP
Cscript prnport.vbs -a -r "Printer Local IP" -o "lpr" -h Printer Local IP
Cscript prnport.vbs -a -r "Printer Local IP" -o "lpr" -h Printer Local IP
Cscript prnport.vbs -a -r "Printer Local IP" -o "lpr" -h Printer Local IP
Cscript prnport.vbs -a -r "Printer Local IP" -o "lpr" -h Printer Local IP

# Bin port and create printer
rundll32 printui.dll,PrintUIEntry /if /b "Printer Name" /f "C:\Driver_Konica\KOAXWJ__.inf" /r "Printer Local IP" /m "KONICA MINOLTA C658SeriesPCL"
rundll32 printui.dll,PrintUIEntry /if /b "Printer Name" /f "C:\Driver_Konica\KOAXWJ__.inf" /r "Printer Local IP" /m "KONICA MINOLTA C658SeriesPCL"
rundll32 printui.dll,PrintUIEntry /if /b "Printer Name" /f "C:\Driver_Konica\KOAXWJ__.inf" /r "Printer Local IP" /m "KONICA MINOLTA C658SeriesPCL"
rundll32 printui.dll,PrintUIEntry /if /b "Printer Name" /f "C:\Driver_Konica\KOAXWJ__.inf" /r "Printer Local IP" /m "KONICA MINOLTA C658SeriesPCL"
rundll32 printui.dll,PrintUIEntry /if /b "Printer Name" /f "C:\Driver_Konica\KOAXWJ__.inf" /r "Printer Local IP" /m "KONICA MINOLTA C658SeriesPCL"
rundll32 printui.dll,PrintUIEntry /if /b "Printer Name" /f "C:\Driver_Konica\KOAXWJ__.inf" /r "Printer Local IP" /m "KONICA MINOLTA C658SeriesPCL"

