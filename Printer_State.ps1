#+-------------------------------------------------------------------+  
#| = : = : = : = : = : = : = : = : = : = : = : = : = : = : = : = : = |  
#|{>/-------------------------------------------------------------\<}|           
#|: | Author:  Philippe-Alexandre Munch                           | :|           
#| :| Email:   ----------------------------------                 |: |
#|: | Purpose: Printer_State :)  in GUI Version                   | :|
#| :|                                                             |: |  
#|: |                      						                  | :|  
#| :|                                                             |: |      
#|: |         		Date:05-Jan-2018                              | :|  
#|: |                                                             |: |  
#| :| 	/^(o.o)^\    Version: 3.0            	                  | :|
#|{>\-------------------------------------------------------------/<}|
#| = : = : = : = : = : = : = : = : = : = : = : = : = : = : = : = : = |
#+-------------------------------------------------------------------+


# LOAD ASSEMBLY
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[void][reflection.assembly]::LoadWithPartialName("System.Windows.Forms")
[reflection.assembly]::LoadWithPartialName( "System.Windows.Forms")
[reflection.assembly]::LoadWithPartialName( "System.Drawing")
[Windows.Forms.Application]::EnableVisualStyles()
$menuMain = New-Object System.Windows.Forms.MenuStrip
$menuAbout = New-Object System.Windows.Forms.ToolStripMenuItem
$MenuContact = New-Object System.Windows.Forms.ToolStripMenuItem
$MenuDrivers = New-Object System.Windows.Forms.ToolStripMenuItem
$DriversKonica = New-Object System.Windows.Forms.ToolStripMenuItem

# Visual DATA
$file = (get-item $env:USERPROFILE\Documents\Visual_Data\Image2.png)
$img = [System.Drawing.Image]::Fromfile($file);
$file2 = (get-item $env:USERPROFILE\Documents\Visual_Data\wallpaper.jpg) 
$img2 = [System.Drawing.Image]::Fromfile($file2);
    
# GraphiC Unite Interface
[System.Windows.Forms.Application]::EnableVisualStyles(); 
$form = new-object system.Windows.Forms.Form
$form.Text = "Printer State"
$form.ForeColor = "White"
$form.AutoSizeMode = "GrowOn"
$form.WindowState="Maximized"
$form.BackColor = "Green"
$form.BackgroundImage = $img2
$form.BackgroundImageLayout = "none" 

# Main Menu Bar
[void]$Form.Controls.Add($menuMain)

# about
$menuAbout.Text = "&about"
[void]$menuMain.Items.Add($menuAbout)
$menuAbout.add_Click({about})

# contacte
$MenuContact.Text = "&Contact"
[void]$menuMain.Items.Add($MenuContact)
$MenuContact.add_Click({Contact})

# Menue Drivers
[void]$menuMain.Items.Add($MenuDrivers)
$MenuDrivers.Text = "Drivers"

# Menue Drivers Knoica
[void]$MenuDrivers.DropDownItems.Add($DriversKonica)
$DriversKonica.text = "Printer Driver"
$DriversKonica.add_Click{KonicaDriver}

# Function Get, Store and injection Printer Driver
function KonicaDriver {
    New-Item -ItemType Directory -Path "C:\" -Name Drivers_Konica
    Copy-Item -Path 'source path' -Recurse -Destination "Destination path"
    start "Batch file for install driver"
    Add-PrinterDriver -Name "Driver Name"
    if (Test-Path -Path "Driver Source Path") {
        $wshell = New-Object -ComObject Wscript.Shell    
        $wshell.Popup("Driver Have Been Copied", 0, "Done", 0x1)
    }
}

# funtion contact
function Contact {

    $ol = New-Object -comObject Outlook.Application

    #Create the new email
    $mail = $ol.CreateItem(0)

    # 
    $mail.To = "Recipient@domain.com"
    
    #Optional, set the subject
    $mail.Subject = "Printer_State"
    
    #Get the new email object
    $inspector = $mail.GetInspector
    
    #Bring the message window to the front
    $inspector.Activate()
    
}

# funtion about 

function about {
    $apropos = New-Object -ComObject Wscript.Shell    
    $apropos.Popup("
         
Author:  Philippe-Alexandre Munch                                    
Email:   --------------------------------                  
Purpose: Printer_State :)  in GUI Version                   
                                                                 
Date:05-Jan-2018                                                            
Version: 3.0                    


** Welcomle to Printer_State **  


**************************************
************************
********************
*****************
**************



", 0, "about", 0x1)
    
}

# LOGO BAR
$Icon = New-Object system.drawing.icon ("$env:USERPROFILE\Documents\Visual_Data\logo_Printer_State.ico")
$form.Icon = $Icon


# LOGO PRINTER
$pictureBox = new-object Windows.Forms.PictureBox
$pictureBox.Width =  $img.Size.Width
$pictureBox.Height =  $img.Size.Height
$pictureBox.Location = new-object System.Drawing.Size(10,35)
$pictureBox.Image = $img;
$pictureBox.BackColor = "Transparent"
$form.controls.add($pictureBox)

# LOGO PRINTER 2 
$pictureBox2 = new-object Windows.Forms.PictureBox
$pictureBox2.Width =  $img.Size.Width
$pictureBox2.Height =  $img.Size.Height
$pictureBox2.Location = new-object System.Drawing.Size(10,85)
$pictureBox2.Image = $img;
$pictureBox2.BackColor = "Transparent"
$form.controls.add($pictureBox2)

# LOGO PRINTER 3 
$pictureBox3 = new-object Windows.Forms.PictureBox
$pictureBox3.Width =  $img.Size.Width
$pictureBox3.Height =  $img.Size.Height
$pictureBox3.Location = new-object System.Drawing.Size(10,140)
$pictureBox3.Image = $img;
$pictureBox3.BackColor = "Transparent"
$form.controls.add($pictureBox3)

# LOGO PRINTER 4 
$pictureBox4 = new-object Windows.Forms.PictureBox
$pictureBox4.Width =  $img.Size.Width
$pictureBox4.Height =  $img.Size.Height
$pictureBox4.Location = new-object System.Drawing.Size(10,200)
$pictureBox4.Image = $img;
$pictureBox4.BackColor = "Transparent"
$form.controls.add($pictureBox4)

# LOGO PRINTER 5 
$pictureBox5 = new-object Windows.Forms.PictureBox
$pictureBox5.Width =  $img.Size.Width
$pictureBox5.Height =  $img.Size.Height
$pictureBox5.Location = new-object System.Drawing.Size(10,260)
$pictureBox5.Image = $img;
$pictureBox5.BackColor = "Transparent"
$form.controls.add($pictureBox5)

# IP ADRESSE PRINTER & Physic Location 
$objip = New-Object System.Windows.Forms.Label
$objip.Location = New-Object System.Drawing.Size(60,40) 
$objip.Size = New-Object System.Drawing.Size(300,20) 
$objip.Text = "Local IP And Physic Location" 
$objip.BackColor = "Transparent"
$form.Controls.Add($objip) 

# IP ADRESSE PRINTER & Physic Location 2
$objip2 = New-Object System.Windows.Forms.Label
$objip2.Location = New-Object System.Drawing.Size(60,100) 
$objip2.Size = New-Object System.Drawing.Size(300,20) 
$objip2.Text = "Local IP And Physic Location"
$objip2.BackColor = "Transparent"
$form.Controls.Add($objip2) 

# IP ADRESSE PRINTER & Physic Location 3
$objip3 = New-Object System.Windows.Forms.Label
$objip3.Location = New-Object System.Drawing.Size(60,160) 
$objip3.Size = New-Object System.Drawing.Size(325,20) 
$objip3.Text = "Local IP And Physic Location"
$objip3.BackColor = "Transparent"
$form.Controls.Add($objip3) 

# IP ADRESSE PRINTER & Physic Location 4
$objip4 = New-Object System.Windows.Forms.Label
$objip4.Location = New-Object System.Drawing.Size(60,220) 
$objip4.Size = New-Object System.Drawing.Size(300,20) 
$objip4.Text =  "Local IP And Physic Location"
$objip4.BackColor = "Transparent"
$form.Controls.Add($objip4) 

# IP ADRESSE PRINTER & Physic Location 5
$objip5 = New-Object System.Windows.Forms.Label
$objip5.Location = New-Object System.Drawing.Size(60,280) 
$objip5.Size = New-Object System.Drawing.Size(300,20) 
$objip5.Text = "Local IP And Physic Location"
$objip5.BackColor = "Transparent"
$form.Controls.Add($objip5) 

# LINK TO PRINTER WEB SITE 
$LinkLabel = New-Object System.Windows.Forms.LinkLabel 
$LinkLabel.Location = New-Object System.Drawing.Size(400,40) 
$LinkLabel.Size = New-Object System.Drawing.Size(150,20) 
$LinkLabel.LinkColor = "Cyan" 
$LinkLabel.ActiveLinkColor = "Cyan" 
$LinkLabel.Text = "Link To Printer WebPage" 
$LinkLabel.add_Click({[system.Diagnostics.Process]::start("website addr")})
$LinkLabel.BackColor = "Transparent" 
$form.Controls.Add($LinkLabel) 

# LINK TO PRINTER WEB SITE 2    
$LinkLabel2 = New-Object System.Windows.Forms.LinkLabel 
$LinkLabel2.Location = New-Object System.Drawing.Size(400,100) 
$LinkLabel2.Size = New-Object System.Drawing.Size(150,20) 
$LinkLabel2.LinkColor = "Cyan" 
$LinkLabel2.ActiveLinkColor = "Cyan" 
$LinkLabel2.Text = "Link To Printer WebPage" 
$LinkLabel2.BackColor = "Transparent"
$LinkLabel2.add_Click({[system.Diagnostics.Process]::start("website addr")}) 
$form.Controls.Add($LinkLabel2) 

# LINK TO PRINTER WEB SITE 3
$LinkLabel3 = New-Object System.Windows.Forms.LinkLabel 
$LinkLabel3.Location = New-Object System.Drawing.Size(400,160) 
$LinkLabel3.Size = New-Object System.Drawing.Size(150,20) 
$LinkLabel3.LinkColor = "Cyan" 
$LinkLabel3.ActiveLinkColor = "Cyan" 
$LinkLabel3.Text = "Link To Printer WebPage" 
$LinkLabel3.BackColor = "Transparent"
$LinkLabel3.add_Click({[system.Diagnostics.Process]::start("website addr")}) 
$form.Controls.Add($LinkLabel3) 

# LINK TO PRINTER WEB SITE 4
$LinkLabel4 = New-Object System.Windows.Forms.LinkLabel 
$LinkLabel4.Location = New-Object System.Drawing.Size(400,220) 
$LinkLabel4.Size = New-Object System.Drawing.Size(150,20) 
$LinkLabel4.LinkColor = "Cyan" 
$LinkLabel4.ActiveLinkColor = "Cyan" 
$LinkLabel4.Text = "Link To Printer WebPage" 
$LinkLabel4.BackColor = "Transparent"
$LinkLabel4.add_Click({[system.Diagnostics.Process]::start("website addr")}) 
$form.Controls.Add($LinkLabel4) 

# LINK TO PRINTER WEB SITE 5
$LinkLabel5 = New-Object System.Windows.Forms.LinkLabel 
$LinkLabel5.Location = New-Object System.Drawing.Size(400,280) 
$LinkLabel5.Size = New-Object System.Drawing.Size(150,20) 
$LinkLabel5.LinkColor = "Cyan" 
$LinkLabel5.ActiveLinkColor = "Cyan" 
$LinkLabel5.Text = "Link To Printer WebPagee" 
$LinkLabel5.BackColor = "Transparent"
$LinkLabel5.add_Click({[system.Diagnostics.Process]::start("website addr")}) 
$form.Controls.Add($LinkLabel5) 

###########################
# Install Printer Boutton  #
Function StartProgressBar
{
	if($i -le 0){
	    $pbrTest.Value = $i
	    $script:i += 1
	}
	
	}
$pbrTest = New-Object System.Windows.Forms.ProgressBar
$pbrTest.Maximum = 100
$pbrTest.Minimum = 0
$pbrTest.Location = new-object System.Drawing.Size(705,45)
$pbrTest.size = new-object System.Drawing.Size(100,15)
$i = 0	

$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Size(555,40)
$OKButton.Size = New-Object System.Drawing.Size(100,23)
$OKButton.ForeColor = "DarkBlue"
$OKButton.BackColor = "White"
$OKButton.Text = "Install"
$OKButton.Add_Click({
Add-PrinterPort -Name "Local_Printer_IP" -LprHostAddress "Local_Printer_IP" -LprQueueName "LPR" 
Add-Printer -Name "Printer_Name" -PortName "IP_PORT" -DriverName "Driver Name"

    
    While ($i -le 100) {
        $pbrTest.Value = $i
        Start-Sleep -m 1
        "VALLUE EQ"
        $i
        $i += 1
    }
})

$Form.Controls.Add($pbrTest)
$form.Controls.Add($OKButton)

#Install Printer 2 
Function StartProgressBar2
{
	if($i2 -le 0){
	    $pbrTest2.Value = $i2
	    $script2:i += 1
	}
	
	}
$pbrTest2 = New-Object System.Windows.Forms.ProgressBar
$pbrTest2.Maximum = 100
$pbrTest2.Minimum = 0
$pbrTest2.Location = new-object System.Drawing.Size(705,105)
$pbrTest2.size = new-object System.Drawing.Size(100,15)
$i2 = 0	

$OKButton2 = New-Object System.Windows.Forms.Button
$OKButton2.Location = New-Object System.Drawing.Size(555,100)
$OKButton2.Size = New-Object System.Drawing.Size(100,23)
$OKButton2.ForeColor = "DarkBlue"
$OKButton2.BackColor = "White"
$OKButton2.Text = "Install"
$OKButton2.Add_Click({
Add-PrinterPort -Name "Local_Printer_IP" -LprHostAddress "Local_Printer_IP" -LprQueueName "LPR" 
Add-Printer -Name "Printer_Name" -PortName "IP_PORT" -DriverName "Driver Name"
    

        While ($i2 -le 100) {
            $pbrTest2.Value = $i2
            Start-Sleep -m 1
            "VALLUE EQ"
            $i2
            $i2 += 1
        }
    })

$Form.Controls.Add($pbrTest2)
$form.Controls.Add($OKButton2)

#Install Printer 3
Function StartProgressBar3
{
	if($i3 -le 0){
	    $pbrTest3.Value = $i3
	    $script3:i += 1
	}
	
	}
$pbrTest3 = New-Object System.Windows.Forms.ProgressBar
$pbrTest3.Maximum = 100
$pbrTest3.Minimum = 0
$pbrTest3.Location = new-object System.Drawing.Size(705,163)
$pbrTest3.size = new-object System.Drawing.Size(100,15)
$i3 = 0	

$OKButton3 = New-Object System.Windows.Forms.Button
$OKButton3.Location = New-Object System.Drawing.Size(555,160)
$OKButton3.Size = New-Object System.Drawing.Size(100,23)
$OKButton3.ForeColor = "DarkBlue"
$OKButton3.BackColor = "White"
$OKButton3.Text = "Install"
$OKButton3.Add_Click({
Add-PrinterPort -Name "Local_Printer_IP" -LprHostAddress "Local_Printer_IP" -LprQueueName "LPR" 
Add-Printer -Name "Printer_Name" -PortName "IP_PORT" -DriverName "Driver Name"


    While ($i3 -le 100) {
        $pbrTest3.Value = $i3
        Start-Sleep -m 1
        "VALLUE EQ"
        $i3
        $i3 += 1
    }
})

$Form.Controls.Add($pbrTest3)
$form.Controls.Add($OKButton3)

#Install Printer 4
Function StartProgressBar4
{
	if($i4 -le 0){
	    $pbrTest4.Value = $i4
	    $script4:i += 1
	}
	
	}
$pbrTest4 = New-Object System.Windows.Forms.ProgressBar
$pbrTest4.Maximum = 100
$pbrTest4.Minimum = 0
$pbrTest4.Location = new-object System.Drawing.Size(705,223)
$pbrTest4.size = new-object System.Drawing.Size(100,15)
$i4 = 0	

$OKButton4 = New-Object System.Windows.Forms.Button
$OKButton4.Location = New-Object System.Drawing.Size(555,220)
$OKButton4.Size = New-Object System.Drawing.Size(100,23)
$OKButton4.ForeColor = "DarkBlue"
$OKButton4.BackColor = "White"
$OKButton4.Text = "Install"
$OKButton4.Add_Click({
Add-PrinterPort -Name "Local_Printer_IP" -LprHostAddress "Local_Printer_IP" -LprQueueName "LPR" 
Add-Printer -Name "Printer_Name" -PortName "IP_PORT" -DriverName "Driver Name"


    While ($i4 -le 100) {
        $pbrTest4.Value = $i4
        Start-Sleep -m 1
        "VALLUE EQ"
        $i4
        $i4 += 1
    }
})

$Form.Controls.Add($pbrTest4)
$form.Controls.Add($OKButton4)

#Install Printer 5
Function StartProgressBar5
{
	if($i5 -le 0){
	    $pbrTest5.Value = $i5
	    $script5:i += 1
	}
	
	}
$pbrTest5 = New-Object System.Windows.Forms.ProgressBar
$pbrTest5.Maximum = 100
$pbrTest5.Minimum = 0
$pbrTest5.Location = new-object System.Drawing.Size(705,283)
$pbrTest5.size = new-object System.Drawing.Size(100,15)
$i5 = 0	

$OKButton5 = New-Object System.Windows.Forms.Button
$OKButton5.Location = New-Object System.Drawing.Size(555,280)
$OKButton5.Size = New-Object System.Drawing.Size(100,23)
$OKButton5.ForeColor = "DarkBlue"
$OKButton5.BackColor = "White"
$OKButton5.Text = "Install"
$OKButton5.Add_Click({
Add-PrinterPort -Name "Local_Printer_IP" -LprHostAddress "Local_Printer_IP" -LprQueueName "LPR" 
Add-Printer -Name "Printer_Name" -PortName "IP_PORT" -DriverName "Driver Name"


    While ($i5 -le 100) {
        $pbrTest5.Value = $i5
        Start-Sleep -m 1
        "VALLUE EQ"
        $i5
        $i5 += 1
    }
})

$Form.Controls.Add($pbrTest5)
$form.Controls.Add($OKButton5)

#============================
# Uninstall Printer 
#===========================
Function StartProgressBar9
{
	if($i9 -le 0){
	    $pbrTest.Value = $i9
	    $script9:i += 1
	}
	
	}
$pbrTest9 = New-Object System.Windows.Forms.ProgressBar
$pbrTest9.Maximum = 100
$pbrTest9.Minimum = 0
$pbrTest9.Location = new-object System.Drawing.Size(1000,45)
$pbrTest9.size = new-object System.Drawing.Size(100,15)
$i9 = 0	

$OKButton9 = New-Object System.Windows.Forms.Button
$OKButton9.Location = New-Object System.Drawing.Size(850,40)
$OKButton9.Size = New-Object System.Drawing.Size(100,23)
$OKButton9.ForeColor = "DarkBlue"
$OKButton9.BackColor = "White"
$OKButton9.Text = "Uninstall"
$OKButton9.Add_Click({Remove-Printer "Printer_Name"
Remove-PrinterPort -Name "Port_Name"

    
    While ($i9 -le 100) {
        $pbrTest9.Value = $i9
        Start-Sleep -m 1
        "VALLUE EQ"
        $i9
        $i9 += 1
    }
})

$Form.Controls.Add($pbrTest9)
$form.Controls.Add($OKButton9)

#============================
# Unistall Printer 2
#===========================

Function StartProgressBar2
{
	if($i10 -le 0){
	    $pbrTest.Value = $i10
	    $script10:i += 1
	}
	
	}
$pbrTest10 = New-Object System.Windows.Forms.ProgressBar
$pbrTest10.Maximum = 100
$pbrTest10.Minimum = 0
$pbrTest10.Location = new-object System.Drawing.Size(1000,103)
$pbrTest10.size = new-object System.Drawing.Size(100,15)
$i10 = 0	

$OKButton10 = New-Object System.Windows.Forms.Button
$OKButton10.Location = New-Object System.Drawing.Size(850,100)
$OKButton10.Size = New-Object System.Drawing.Size(100,23)
$OKButton10.ForeColor = "DarkBlue"
$OKButton10.BackColor = "White"
$OKButton10.Text = "Uninstall"
$OKButton10.Add_Click({Remove-Printer "Printer_Name"
Remove-PrinterPort -Name "Port_Name"
    
    While ($i10 -le 100) {
        $pbrTest10.Value = $i10
        Start-Sleep -m 1
        "VALLUE EQ"
        $i10
        $i10 += 1
    }
})

$Form.Controls.Add($pbrTest10)
$form.Controls.Add($OKButton10)


#============================
# uninstall printer 3 
#===========================

Function StartProgressBar3
{
	if($i11 -le 0){
	    $pbrTest.Value = $i11
	    $script11:i += 1
	}
	
	}
$pbrTest11 = New-Object System.Windows.Forms.ProgressBar
$pbrTest11.Maximum = 100
$pbrTest11.Minimum = 0
$pbrTest11.Location = new-object System.Drawing.Size(1000,163)
$pbrTest11.size = new-object System.Drawing.Size(100,15)
$i11 = 0	

$OKButton11 = New-Object System.Windows.Forms.Button
$OKButton11.Location = New-Object System.Drawing.Size(850,160)
$OKButton11.Size = New-Object System.Drawing.Size(100,23)
$OKButton11.ForeColor = "DarkBlue"
$OKButton11.BackColor = "White"
$OKButton11.Text = "Uninstall"
$OKButton11.Add_Click({Remove-Printer "Printer_Name"
Remove-PrinterPort -Name "Port_Name"

    
    While ($i11 -le 100) {
        $pbrTest11.Value = $i11
        Start-Sleep -m 1
        "VALLUE EQ"
        $i11
        $i11 += 1
    }
})

$Form.Controls.Add($pbrTest11)
$form.Controls.Add($OKButton11)


#============================
# Uninstall Printer 4 
#============================

Function StartProgressBar2
{
	if($i12 -le 0){
	    $pbrTest.Value = $i12
	    $script12:i += 1
	}
	
	}
$pbrTest12 = New-Object System.Windows.Forms.ProgressBar
$pbrTest12.Maximum = 100
$pbrTest12.Minimum = 0
$pbrTest12.Location = new-object System.Drawing.Size(1000,223)
$pbrTest12.size = new-object System.Drawing.Size(100,15)
$i12 = 0	

$OKButton12 = New-Object System.Windows.Forms.Button
$OKButton12.Location = New-Object System.Drawing.Size(850,220)
$OKButton12.Size = New-Object System.Drawing.Size(100,23)
$OKButton12.ForeColor = "DarkBlue"
$OKButton12.BackColor = "White"
$OKButton12.Text = "Uninstall"
$OKButton12.Add_Click({Remove-Printer "Printer_Name"
Remove-PrinterPort -Name "Port_Name"

    
    While ($i12 -le 100) {
        $pbrTest12.Value = $i12
        Start-Sleep -m 1
        "VALLUE EQ"
        $i12
        $i12 += 1
    }
})

$Form.Controls.Add($pbrTest12)
$form.Controls.Add($OKButton12)

#============================
# Uninstall Printer 5 
#============================

Function StartProgressBar2
{
	if($i13 -le 0){
	    $pbrTest.Value = $i13
	    $script13:i += 1
	}
	
	}
$pbrTest13 = New-Object System.Windows.Forms.ProgressBar
$pbrTest13.Maximum = 110
$pbrTest13.Minimum = 0
$pbrTest13.Location = new-object System.Drawing.Size(1000,283)
$pbrTest13.size = new-object System.Drawing.Size(100,15)
$i13 = 0	

$OKButton13 = New-Object System.Windows.Forms.Button
$OKButton13.Location = New-Object System.Drawing.Size(850,280)
$OKButton13.Size = New-Object System.Drawing.Size(100,23)
$OKButton13.ForeColor = "DarkBlue"
$OKButton13.BackColor = "White"
$OKButton13.Text = "Uninstall"
$OKButton13.Add_Click({Remove-Printer "Printer_Name"
Remove-PrinterPort -Name "Port_Name"
    
    While ($i13 -le 110) {
        $pbrTest13.Value = $i13
        Start-Sleep -m 1
        "VALLUE EQ"
        $i13
        $i13 += 1
    }
})

$Form.Controls.Add($pbrTest13)
$form.Controls.Add($OKButton13)

# ping Printer  & freidnly output
function pingInfo1 {

$pingResult1 = ping "PRINTER_IP_ADDRESS" -n 3  | fl | Out-String; 
if ($pingResult1 -match "TTL"){$outputBox1.text = "*** Online ***"
$outputBox1.ForeColor = "Blue"}
else {$outputBox1.Text = "*** Offline ***"
$outputBox1.ForeColor = "Red"} 
}     
$Button1 = New-Object System.Windows.Forms.Button 
$Button1.Location = New-Object System.Drawing.Size(1140,40) 
$Button1.Size = New-Object System.Drawing.Size(102,23) 
$Button1.ForeColor = "BLUE"
$Button1.BackColor = "WHITE"
$Button1.Text = "Check Connexion" 
$Button1.Add_Click({pingInfo1}) 
$Form.Controls.Add($Button1) 
$outputBox1 = New-Object System.Windows.Forms.TextBox 
$outputBox1.Location = New-Object System.Drawing.Size(1250,35) 
$outputBox1.Size = New-Object System.Drawing.Size(200,35) 
$outputBox1.MultiLine = $True 
$outputBox1.ForeColor = "Blue"

$outputBox1.ScrollBars = "Vertical" 
$Form.Controls.Add($outputBox1) 

# ping Printer  & freidnly output 2 
function pingInfo2 {
$pingResult2 = ping "PRINTER_IP_ADDRESS" -n 3  | fl | Out-String; 
if ($pingResult2 -match "TTL"){$outputBox2.text = "*** Online ***"
$outputBox2.ForeColor = "Blue"}
else {$outputBox2.Text = "*** Offline ***"
$outputBox2.ForeColor = "Red"} 
}                                  
$Button2 = New-Object System.Windows.Forms.Button 
$Button2.Location = New-Object System.Drawing.Size(1140,100) 
$Button2.Size = New-Object System.Drawing.Size(102,23) 
$Button2.ForeColor = "BLUE"
$Button2.BackColor = "WHITE"
$Button2.Text = "Check Connexion" 
$Button2.Add_Click({pingInfo2}) 
$Form.Controls.Add($Button2) 
$outputBox2 = New-Object System.Windows.Forms.TextBox 
$outputBox2.Location = New-Object System.Drawing.Size(1250,95) 
$outputBox2.Size = New-Object System.Drawing.Size(200,35) 
$outputBox2.MultiLine = $True 
$outputBox2.ForeColor = "Blue"

$outputBox2.ScrollBars = "Vertical" 
$Form.Controls.Add($outputBox2) 

# ping Printer  & freidnly output 3
function pingInfo3 {
$pingResult3 = ping "PRINTER_IP_ADDRESS" -n 3  | Format-List | Out-String; 
if ($pingResult3 -match "TTL"){$outputBox3.text = "*** Online ***"
$outputBox3.ForeColor = "Blue"}
else {$outputBox3.Text = "*** Offline ***"
$outputBox3.ForeColor = "Red"}  
}  
$Button3 = New-Object System.Windows.Forms.Button 
$Button3.Location = New-Object System.Drawing.Size(1140,160) 
$Button3.Size = New-Object System.Drawing.Size(102,23) 
$Button3.ForeColor = "BLUE"
$Button3.BackColor = "WHITE"
$Button3.Text = "Check Connexion" 
$Button3.Add_Click({pingInfo3}) 
$Form.Controls.Add($Button3) 
$outputBox3 = New-Object System.Windows.Forms.TextBox 
$outputBox3.Location = New-Object System.Drawing.Size(1250,155) 
$outputBox3.Size = New-Object System.Drawing.Size(200,35) 
$outputBox3.MultiLine = $True 
$outputBox3.ForeColor = "Blue"
$outputBox3.ScrollBars = "Vertical" 
$Form.Controls.Add($outputBox3) 

# ping Printer  & freidnly output 4
function pingInfo4 {

$pingResult4 = ping "PRINTER_IP_ADDRESS" -n 3  | fl | Out-String; 
if ($pingResult4 -match "TTL"){$outputBox4.text = "*** Online ***"
$outputBox4.ForeColor = "Blue"}
else {$outputBox4.Text = "*** Offline ***"
$outputBox4.ForeColor = "Red"} 
}                                  
$Button4 = New-Object System.Windows.Forms.Button 
$Button4.Location = New-Object System.Drawing.Size(1140,220) 
$Button4.Size = New-Object System.Drawing.Size(102,23) 
$Button4.ForeColor = "BLUE"
$Button4.BackColor = "WHITE"
$Button4.Text = "Check Connexion" 
$Button4.Add_Click({pingInfo4}) 
$Form.Controls.Add($Button4) 
$outputBox4 = New-Object System.Windows.Forms.TextBox 
$outputBox4.Location = New-Object System.Drawing.Size(1250,215) 
$outputBox4.Size = New-Object System.Drawing.Size(200,35) 
$outputBox4.MultiLine = $True 
$outputBox4.ForeColor = "Blue"

$outputBox4.ScrollBars = "Vertical" 
$Form.Controls.Add($outputBox4) 
 
# ping Printer  & freidnly output 5
function pingInfo5 {

$pingResult5 = ping "PRINTER_IP_ADDRESS" -n 3  | fl | Out-String; 
if ($pingResult5 -match "TTL"){$outputBox5.text = "*** Online ***"
$outputBox5.ForeColor = "Blue"}
else {$outputBox5.Text = "*** Offline ***"
$outputBox5.ForeColor = "Red"} 
}                                  
$Button5 = New-Object System.Windows.Forms.Button 
$Button5.Location = New-Object System.Drawing.Size(1140,280) 
$Button5.Size = New-Object System.Drawing.Size(102,23) 
$Button5.ForeColor = "BLUE"
$Button5.BackColor = "WHITE"
$Button5.Text = "Check Connexion" 
$Button5.Add_Click({pingInfo5}) 
$Form.Controls.Add($Button5) 

$outputBox5 = New-Object System.Windows.Forms.TextBox 
$outputBox5.Location = New-Object System.Drawing.Size(1250,275) 
$outputBox5.Size = New-Object System.Drawing.Size(200,35) 
$outputBox5.MultiLine = $True 
$outputBox5.ForeColor = "Blue"

$outputBox5.ScrollBars = "Vertical" 
$Form.Controls.Add($outputBox5) 

# Signature 
$Sinature = New-Object System.Windows.Forms.Label
$Sinature.Location = New-Object System.Drawing.Size(0,960) 
$Sinature.Size = New-Object System.Drawing.Size(300,50) 
$Sinature.Text = "
PrinterState 3.0 - 
Created by Philippe-Alexandre Munch"
$Sinature.BackColor = "Transparent"
$form.Controls.Add($Sinature) 

 # Show result 
 $form.Add_Shown( { $form.Activate() } )
 $form.ShowDialog()