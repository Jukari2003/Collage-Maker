################################################################################
#                                 Collage Maker                                #
#                      Written By: MSgt Anthony Brechtel                       #
#                                    Ver 2.2                                   #
#                                                                              #
################################################################################
clear-host
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
################################################################################
######Load Assemblies###########################################################
[System.Windows.Forms.Application]::EnableVisualStyles();


################################################################################
######Global Variables##########################################################
$Global:outFile = ""
#$Global:restricted_zones = @{};
$Global:settings_file = "$scriptPath\Settings.txt"

$Global:default_input_path = $scriptPath
$Global:default_output_path = "$scriptPath\Output"
$global:background_color = "Black"
$global:border_color = "Black"
[int]$Global:default_pic_number = 150
[int]$Global:default_min_image_size = 400
[int]$Global:default_max_image_size = 1520
[int]$global:border_width = 2;
[int]$global:overlap = 33;

$screen = [System.Windows.Forms.SystemInformation]::PrimaryMonitorSize
#[string]$screen = [string]$screen.width + "x" + [string]$screen.Height

if($screen.width)
{
    [int]$global:imageWidth = $screen.width
}
else
{
    [int]$global:imageWidth = 1024
}
if($screen.height)
{
    [int]$global:imageHeight = $screen.height
}
else
{
    [int]$global:imageHeight = 768
}

#################################################################################
#####Loading Settings############################################################
function load_settings
{
    if(!(Test-path -literalpath "$scriptPath\Output"))
    {
        New-Item  -ItemType directory -Path "$scriptPath\Output"
    }
    if(test-path -literalpath $Global:settings_file)
    {
        $reader = [System.IO.File]::OpenText($settings_file)
        while($null -ne ($line = $reader.ReadLine())) 
        {

            $line_split = $line -split ':::';
            ###################
            if($line_split[0] -eq "OUTPUT_DIR")
            {
                if(Test-Path -literalpath $line_split[1])
                {
                    $Global:default_output_path = $line_split[1];
                }
            }
            ###################
            if($line_split[0] -eq "INPUT_DIR")
            {
                if(Test-Path -literalpath $line_split[1])
                {
                    $global:default_input_path = $line_split[1];
                }
            }
            ###################
            if($line_split[0] -eq "PIC_NUMBER")
            {
                $global:default_pic_number = $line_split[1];
            }
            ###################
            if($line_split[0] -eq "WALLPAPER_WIDTH")
            {
                $global:imageWidth = $line_split[1];
            }
            ###################
            if($line_split[0] -eq "WALLPAPER_HEIGHT")
            {
                $global:imageHeight = $line_split[1];
            }
            ###################
            if($line_split[0] -eq "MIN_SIZE")
            {
                $global:default_min_image_size = $line_split[1];
            }
            ###################
            if($line_split[0] -eq "MAX_SIZE")
            {
                $global:default_max_image_size = $line_split[1];
            }
            ###################
            if($line_split[0] -eq "OVERLAP")
            {
                $global:overlap = $line_split[1];
            }
            ###################
            if($line_split[0] -eq "BACKGROUND_COLOR")
            {
                $global:background_color = $line_split[1];
            }
            ###################
            if($line_split[0] -eq "BORDER_COLOR")
            {
                $global:border_color = $line_split[1];
            }
            ###################
            if($line_split[0] -eq "BORDER_WIDTH")
            {
                $global:border_width = $line_split[1];
            } 
        }
        $reader.Close()
    }
}
#################################################################################
#####Save Settings###############################################################
function save_settings
{
    if(test-path -literalpath $settings_file)
    {
        remove-item -literalpath $settings_file
    }
    Add-Content $settings_file "OUTPUT_DIR:::$Global:default_output_path"
    Add-Content $settings_file "INPUT_DIR:::$Global:default_input_path"
    Add-Content $settings_file "PIC_NUMBER:::$global:default_pic_number"
    Add-Content $settings_file "WALLPAPER_WIDTH:::$global:imageWidth"
    Add-Content $settings_file "WALLPAPER_HEIGHT:::$global:imageHeight"
    Add-Content $settings_file "MIN_SIZE:::$global:default_min_image_size"
    Add-Content $settings_file "MAX_SIZE:::$global:default_max_image_size"
    Add-Content $settings_file "OVERLAP:::$global:overlap"
    Add-Content $settings_file "BACKGROUND_COLOR:::$global:background_color"
    Add-Content $settings_file "BORDER_WIDTH:::$global:border_width"
    Add-Content $settings_file "BORDER_COLOR:::$global:border_color"
}
#################################################################################
#####Main########################################################################
function main
{
    ##################################################################################
    ###########Main Form
    $mainForm = New-Object System.Windows.Forms.Form
    $mainForm.Location = "200, 200"
    $mainForm.Font = "Copperplate Gothic,8.1"
    $mainForm.FormBorderStyle = "FixedDialog"
    $mainForm.ForeColor = "Black"
    $mainForm.BackColor = "#434343"
    $mainForm.Text = "  Collage Maker"
    $mainForm.Width = 530 #1245
    $mainForm.Height = 450
    #$mainForm.ClientSize = "530,450"

    ##################################################################################
    ###########Progress Bar
    $progressBar1 = New-Object System.Windows.Forms.ProgressBar
    $progressBar1.Name = 'progressBar1'
    $progressBar1.Value = 0
    $progressBar1.Style="Continuous"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 0
    $System_Drawing_Size.Height = 0
    $progressBar1.Size = $System_Drawing_Size
    $progressBar1.Location = "15, 425"
    $mainForm.Controls.Add($progressBar1)

    ##################################################################################
    ###########Path Input
    $textBoxIn = New-Object System.Windows.Forms.TextBox
    $textBoxIn.text = $Global:default_input_path
    $textBoxIn.Location = "15, 30"
    $textBoxIn.Size = "420, 20"
    $textBoxIn.Add_TextChanged({
        if($textBoxIn.text -and (test-path -literalpath $textBoxIn.text))
        {
            $Global:default_input_path = $textBoxIn.text;
            save_settings
        }
        else
        {
            $Global:default_input_path = "";
        }
    })
    $mainForm.Controls.Add($textBoxIn)

    ##################################################################################
    ###########Input Label
    $ProcessLabel = New-Object System.Windows.Forms.Label
    $ProcessLabel.Location = "15, 12"
    $ProcessLabel.Size = "300, 23"
    $ProcessLabel.ForeColor = "White" 
    $ProcessLabel.Text = "Input Images Folder"
    $mainForm.Controls.Add($ProcessLabel)

    ##################################################################################
    ###########Browse Button
    $buttonBrowse = New-Object System.Windows.Forms.Button
    $buttonBrowse.Location = "440, 28"
    $buttonBrowse.Size = "75, 23"
    $buttonBrowse.ForeColor = "White"
    $buttonBrowse.Backcolor = "#606060"
    $buttonBrowse.Text = "Browse"
    $buttonBrowse.add_Click({selectFolderIn})
    $mainForm.Controls.Add($buttonBrowse)

    ##################################################################################
    ###########Picture Number Trackbar
    $numberTrackBar = New-Object System.Windows.Forms.TrackBar
    $numberTrackBar.Location = "230, 60"
    $numberTrackBar.Orientation = "Horizontal"
    $numberTrackBar.Width = 290
    $numberTrackBar.Height = 40
    $numberTrackBar.TickFrequency = 100
    $numberTrackBar.TickStyle = "TopLeft"
    $numberTrackBar.SetRange(1, 2000)
    $numberTrackBar.Value = $Global:default_pic_number
    $numberTrackBarValue = $Global:default_pic_number
    $numberTrackBar.add_ValueChanged({
        $numberTrackBarValue = $numberTrackBar.Value
        $numberTrackBarLabel.Text = "Number of Images ($numberTrackBarValue)"
        $Global:default_pic_number = $numberTrackBarValue
        save_settings
    })
    $mainForm.Controls.add($numberTrackBar)

    ##################################################################################
    ###########Picture Number Label
    $numberTrackBarLabel = New-Object System.Windows.Forms.Label 
    $numberTrackBarLabel.Location = "15,70"
    $numberTrackBarLabel.Size = "190,23"
    $numberTrackBarLabel.ForeColor = "White"
    $numberTrackBarLabel.Text = "Number of Images ($numberTrackBarValue)"
    $mainForm.Controls.Add($numberTrackBarLabel)

    ##################################################################################
    ###########Minimum Number of Images Trackbar
    $min_size_TrackBar = New-Object System.Windows.Forms.TrackBar
    $min_size_TrackBar.Location = "230,110"
    $min_size_TrackBar.Orientation = "Horizontal"
    $min_size_TrackBar.Width = 290
    $min_size_TrackBar.Height = 10
    $min_size_TrackBar.TickFrequency = 100
    $min_size_TrackBar.TickStyle = "TopLeft"
    $min_size_TrackBar.SetRange(100, 1000)
    $min_size_TrackBar.Value = $Global:default_min_image_size
    $min_size_TrackBarValue = $Global:default_min_image_size
    $min_size_TrackBar.add_ValueChanged({
        $min_size_TrackBarValue = $min_size_TrackBar.Value
        $min_size_TrackBarLabel.Text = "Minimum Image Size ($min_size_TrackBarValue)"
        $Global:default_min_image_size = $min_size_TrackBarValue
        save_settings
    })
    $mainForm.Controls.add($min_size_TrackBar)

    ##################################################################################
    ###########Minimum Number of Images Label
    $min_size_TrackBarLabel = New-Object System.Windows.Forms.Label 
    $min_size_TrackBarLabel.Location = "15,120"
    $min_size_TrackBarLabel.Size = "210,23"
    $min_size_TrackBarLabel.ForeColor = "White"
    $min_size_TrackBarLabel.Text = "Minimum Image Size ($min_size_TrackBarValue)"
    $mainForm.Controls.Add($min_size_TrackBarLabel)


    ##################################################################################
    ###########Max Image Size Trackbar
    $max_size_TrackBar = New-Object System.Windows.Forms.TrackBar
    $max_size_TrackBar.Location = "230,160"
    $max_size_TrackBar.Orientation = "Horizontal"
    $max_size_TrackBar.Width = 290
    $max_size_TrackBar.Height = 10
    $max_size_TrackBar.LargeChange = 20
    $max_size_TrackBar.SmallChange = 1
    $max_size_TrackBar.TickFrequency = 200
    $max_size_TrackBar.TickStyle = "TopLeft"
    $max_size_TrackBar.SetRange(300, 3000)
    $max_size_TrackBar.Value = $Global:default_max_image_size
    $max_size_TrackBarValue = $Global:default_max_image_size
    $max_size_TrackBar.add_ValueChanged({
        $max_size_TrackBarValue = $max_size_TrackBar.Value
        $max_size_TrackBarLabel.Text = "Maximum Image Size ($max_size_TrackBarValue)"
        $Global:default_max_image_size = $max_size_TrackBarValue
        save_settings
    })
    $mainForm.Controls.add($max_size_TrackBar)

    ##################################################################################
    ###########Max Image Size Label
    $max_size_TrackBarLabel = New-Object System.Windows.Forms.Label 
    $max_size_TrackBarLabel.Location = "15,175"
    $max_size_TrackBarLabel.Size = "210,23"
    $max_size_TrackBarLabel.ForeColor = "White"
    $max_size_TrackBarLabel.Text = "Maximum Image Size ($max_size_TrackBarValue)"
    $mainForm.Controls.Add($max_size_TrackBarLabel)

    ##################################################################################
    ###########Overlap Trackbar
    $global:overlap_TrackBar = New-Object System.Windows.Forms.TrackBar
    $global:overlap_TrackBar.Location = "230,210"
    $global:overlap_TrackBar.Orientation = "Horizontal"
    $global:overlap_TrackBar.Width = 290
    $global:overlap_TrackBar.Height = 10
    $global:overlap_TrackBar.TickFrequency = 10
    $global:overlap_TrackBar.TickStyle = "TopLeft"
    $global:overlap_TrackBar.SetRange(0, 100)
    $global:overlap_TrackBar.Value = $global:overlap
    $global:overlap_TrackBarValue = $global:overlap
    $global:overlap_TrackBar.add_ValueChanged({
        $global:overlap_TrackBarValue = $global:overlap_TrackBar.Value
        $global:overlap_TrackBarLabel.Text = "Overlap ($global:overlap_TrackBarValue)"
        $global:overlap = $global:overlap_TrackBarValue
        save_settings
    })
    $mainForm.Controls.add($global:overlap_TrackBar)

    ##################################################################################
    ###########Overlap Label
    $global:overlap_TrackBarLabel = New-Object System.Windows.Forms.Label 
    $global:overlap_TrackBarLabel.Location = "15,220"
    $global:overlap_TrackBarLabel.Size = "210,23"
    $global:overlap_TrackBarLabel.ForeColor = "White"
    $global:overlap_TrackBarLabel.Text = "Overlap ($global:overlap_TrackBarValue%)"
    $mainForm.Controls.Add($global:overlap_TrackBarLabel)

    ##################################################################################
    ###########Background Color ComboBox
    $backgroundComboBox = New-Object System.Windows.Forms.ComboBox
    $backgroundComboBox.Size = "110,20"
    #$backgroundComboBox.Location = "410,270"
    $backgroundComboBox.Location = "170,270"
    $backgroundComboBox.ForeColor = "Indigo"
    $backgroundComboBox.BackColor = "White"
    [void]$backgroundComboBox.items.add("Black")
    [void]$backgroundComboBox.items.add("White")
    [void]$backgroundComboBox.items.add("Random")
    [void]$backgroundComboBox.items.add("Transparent")
    [void]$backgroundComboBox.items.add("Select Color")
    if(!($backgroundComboBox.items.Contains($global:background_color)))
    {
        [void]$backgroundComboBox.items.add("$global:background_color")
    }
    $backgroundComboBox.SelectedIndex = $backgroundComboBox.FindStringExact("$global:background_color")
    $backgroundComboBox.Add_SelectedValueChanged({
        if($this.text -eq "Select Color")
            {
                $new_color = color_picker
                write-host $new_color[1]
                $this.items.add($new_color[1])
                $backgroundComboBox.text = $new_color[1]
                $global:background_color = $new_color[1]
                save_settings

            }

        $global:background_color = $backgroundComboBox.text;
        save_settings

    })
    $mainForm.Controls.Add($backgroundComboBox)

    ##################################################################################
    ###########Background Color Label
    $backgroundComboBoxLabel = New-Object System.Windows.Forms.Label 
    #$backgroundComboBoxLabel.Location = "300,272"
    $backgroundComboBoxLabel.Location = "15,272"
    $backgroundComboBoxLabel.Size = "180,20"
    $backgroundComboBoxLabel.ForeColor = "White"
    $backgroundComboBoxLabel.Text = "Background Color"
    $mainForm.Controls.Add($backgroundComboBoxLabel)


    ##################################################################################
    ###########Border Color Combo
    $borderColorComboBox = New-Object System.Windows.Forms.ComboBox
    $borderColorComboBox.Location = "410,270"
    $borderColorComboBox.Size = "110,20"
    $borderColorComboBox.ForeColor = "Indigo"
    $borderColorComboBox.BackColor = "White"
    [void]$borderColorComboBox.items.add("Black")
    [void]$borderColorComboBox.items.add("White")
    [void]$borderColorComboBox.items.add("Random")
    [void]$borderColorComboBox.items.add("Select Color")
    if(!($borderColorComboBox.items.Contains($global:border_color)))
    {
        [void]$borderColorComboBox.items.add("$global:border_color")
    }        
    $borderColorComboBox.SelectedIndex = 0
    $borderColorComboBox.SelectedIndex = $borderColorComboBox.FindStringExact("$global:border_color")
    $borderColorComboBox.Add_SelectedValueChanged({
        if($this.text -eq "Select Color")
        {
            $new_color = color_picker
            write-host $new_color[1]
            $this.items.add($new_color[1])
            $borderColorComboBox.text = $new_color[1]
            $global:border_color = $new_color[1]
            save_settings

        }
        $global:border_color = $borderColorComboBox.text;
        save_settings

    })
    $mainForm.Controls.Add($borderColorComboBox)

    ##################################################################################
    ###########Border Color Label
    $borderColorComboBoxLabel = New-Object System.Windows.Forms.Label 
    $borderColorComboBoxLabel.Location = "300,272"
    $borderColorComboBoxLabel.Size = "180,20"
    $borderColorComboBoxLabel.ForeColor = "White"
    $borderColorComboBoxLabel.Text = "Border Color"
    $mainForm.Controls.Add($borderColorComboBoxLabel)

    ##################################################################################
    ###########Screen Size Combo
    $screen = [System.Windows.Forms.SystemInformation]::PrimaryMonitorSize
    [string]$screen = [string]$screen.width + "x" + [string]$screen.Height
    $imageComboBox = New-Object System.Windows.Forms.ComboBox
    $imageComboBox.Location = "170,300"
    $imageComboBox.Size = "110,20"
    $imageComboBox.ForeColor = "Indigo"
    $imageComboBox.BackColor = "White"
    [void]$imageComboBox.items.add("$screen") 
    [void]$imageComboBox.items.add("800 x600")
    [void]$imageComboBox.items.add("1024x768")
    [void]$imageComboBox.items.add("1280x1024")
    [void]$imageComboBox.items.add("1600x1200") 
    [void]$imageComboBox.items.add("1366x768") 
    [void]$imageComboBox.items.add("1920x1200")
    [void]$imageComboBox.items.add("1920x1080")
    [void]$imageComboBox.items.add("5120x1440")
    [void]$imageComboBox.items.add("1200x1200")   
    $imageComboBox.SelectedIndex = $imageComboBox.FindStringExact("$global:imageWidth" + "x" + "$global:imageHeight")
    $imageComboBox.Add_SelectedValueChanged({
        $global:imageWidth  = [int]$imageComboBox.Text.Split("x")[0]
        $global:imageHeight = [int]$imageComboBox.Text.Split("x")[1]
        save_settings
    })
    $mainForm.Controls.Add($imageComboBox)

    ##################################################################################
    ###########Screen Size Label
    $imageComboBoxLabel = New-Object System.Windows.Forms.Label 
    $imageComboBoxLabel.Location = "15,300"
    $imageComboBoxLabel.Size = "150,20"
    $imageComboBoxLabel.ForeColor = "White"
    $imageComboBoxLabel.Text = "Output Size"
    $mainForm.Controls.Add($imageComboBoxLabel)

    ##################################################################################
    ###########Border Color Combo
    $borderComboBox = New-Object System.Windows.Forms.ComboBox
    $borderComboBox.Location = "410,300"
    $borderComboBox.Size = "110,20"
    $borderComboBox.ForeColor = "Indigo"
    $borderComboBox.BackColor = "White"
    For ($i=0; $i -le 10; $i++) {
        [void]$borderComboBox.items.add($i)
    }
    $borderComboBox.SelectedIndex = $borderComboBox.FindStringExact("$global:border_width")
    $borderComboBox.Add_SelectedValueChanged({

        $global:border_width = $borderComboBox.text;
        save_settings

    })
    $mainForm.Controls.Add($borderComboBox)

    ##################################################################################
    ###########Border Width Label
    $borderComboBoxLabel = New-Object System.Windows.Forms.Label 
    $borderComboBoxLabel.Location = "300,300"
    $borderComboBoxLabel.Size = "200,20"
    $borderComboBoxLabel.ForeColor = "White"
    $borderComboBoxLabel.Text = "Border Width"
    $mainForm.Controls.Add($borderComboBoxLabel)

    ##################################################################################
    ###########Output Label
    $outputLabel = New-Object System.Windows.Forms.Label
    $outputLabel.Location = "15, 330"
    $outputLabel.Size = "150, 20"
    $outputLabel.ForeColor = "White"
    $outputLabel.Text = "Output Folder"
    $mainForm.Controls.Add($outputLabel)

    ##################################################################################
    ###########Output Text Box
    $textBoxOut = New-Object System.Windows.Forms.TextBox
    $textBoxOut.text = $default_output_path
    $textBoxOut.Location = "15, 353"
    $textBoxOut.Size = "420, 20"
    $textBoxOut.Add_TextChanged({
        if($textBoxOut.text -and (test-path -literalpath $textBoxOut.text))
        {
            $Global:default_output_path = $textBoxOut.text;
            save_settings
        }
        else
        {
            $Global:default_output_path = "";
        }
    })
    $mainForm.Controls.Add($textBoxOut)

    ##################################################################################
    ###########Browse Button
    $outputBrowse = New-Object System.Windows.Forms.Button
    $outputBrowse.Location = "440, 350"
    $outputBrowse.Size = "75, 23"
    $outputBrowse.ForeColor = "White"
    $outputBrowse.Backcolor = "#606060" 
    $outputBrowse.Text = "Browse"
    $outputBrowse.add_Click({selectFolderOut})
    $mainForm.Controls.Add($outputBrowse)

    ##################################################################################
    ###########Complete Label
    $completeLabel = New-Object System.Windows.Forms.Label
    $completeLabel.Location = "140, 270"
    $completeLabel.Size = "60, 18"
    $completeLabel.ForeColor = "Green"
    $completeLabel.Text = ""
    $mainForm.Controls.Add($completeLabel)

    ##################################################################################
    ###########Process Button
    $buttonProcess = New-Object System.Windows.Forms.Button
    $buttonProcess.Location = "105,390"
    $buttonProcess.Size = "100, 23"
    $buttonProcess.ForeColor = "White"
    $buttonProcess.BackColor = "#606060"
    $buttonProcess.Text = "Process"
    $buttonProcess.add_Click({
    
        if($this.text -eq "Process")
        {
            $this.text = "Stop"
            ProcessCollage
        }
        else
        {
            Stop-Job -job $Global:job
            Remove-Job -job $Global:job
            $this.text = "Process"
            $progressBar1.Value = "100"
        }
    
    
    })
    $mainForm.Controls.Add($buttonProcess)

    ##################################################################################
    ###########Display Button
    $buttonDisplay = New-Object System.Windows.Forms.Button
    $buttonDisplay.Location = "220,390"
    $buttonDisplay.Enabled = $false
    $buttonDisplay.Size = "100, 23"
    $buttonDisplay.ForeColor = "White"
    $buttonDisplay.BackColor = "#606060" 
    $buttonDisplay.Text = "Launch"
    $buttonDisplay.add_Click({DisplayImage})
    $mainForm.Controls.Add($buttonDisplay)

    ##################################################################################
    ###########Exit Button
    $buttonExit = New-Object System.Windows.Forms.Button
    $buttonExit.Location = "335,390"
    $buttonExit.Size = "100, 23"
    $buttonExit.ForeColor = "White"
    $buttonExit.BackColor = "#606060" 
    $buttonExit.Text = "Exit"
    $buttonExit.add_Click({$mainForm.close()})
    $mainForm.Controls.Add($buttonExit)

    [void] $mainForm.ShowDialog()
}
#################################################################################
#####Process Collage#############################################################
Function ProcessCollage 
{
    clear-host;
    ##################################################################################
    ###########Check Validity
    [Array]$errors;
    $files = @()
    if($Global:default_input_path -and (test-path -literalpath $Global:default_input_path))
    {
        if($Global:default_output_path -and (test-path -literalpath $Global:default_output_path))
        {
            $files = Get-ChildItem -literalpath $Global:default_input_path -Recurse | where {$_.extension -in ".png",".jpg",".jpeg",".bmp",".tiff",".gif",".webp"}
        }
        else
        {
            $errors += "Invalid Output Directory"
        }
    }
    else
    {
        $errors += "Invalid Input Directory"
    }
    if($files.count -ne 0)
    {
        $aspect_ratio = $global:imageHeight / 400
        $wheight =  $global:imageHeight / $aspect_ratio;
        $wwidth =  $global:imageWidth / $aspect_ratio;
        
        $mainForm.Width = 530 + $wwidth + 20
        $mainForm.height = 500
        #$mainForm.ClientSize = "1000,500";
        $progressBar1.value = 0;
        $System_Drawing_Size = New-Object System.Drawing.Size
        $System_Drawing_Size.Width = $mainForm.Width - 30
        $System_Drawing_Size.Height = 20
        $progressBar1.Size = $System_Drawing_Size
        $formGraphics = $mainForm.createGraphics() 
        
        $super_name = Split-Path $Global:default_input_path -Leaf


        $Global:outFile = $Global:default_output_path + "\"  + "$super_name" + "_Collage_" + $global:imageWidth +"x" + $global:imageHeight + "_" + (Get-Date -UFormat %Y%m%d_%H%M%S) + ".bmp"
        ##################################################################################
        ###########Build Collage Job
        $build_collage = {

            $save_timer = Get-Date
            Add-Type -Assembly System.Windows.Forms
            Add-Type -Assembly System.Drawing
            $numberTrackBar = $using:numberTrackBar
            $files = $using:files
            $imageWidth = $using:imageWidth
            $imageHeight = $using:imageHeight
            $default_max_image_size = $using:default_max_image_size
            $restricted_zones = @{};
            $default_output_path =  $using:default_output_path
            $outFile = $using:outFile
            $default_pic_number = $using:default_pic_number
            $backgroundComboBox = $using:backgroundComboBox
            $borderColorComboBox = $using:borderColorComboBox
            $overlap = $using:overlap
            $border_width = $using:border_width
            $max_size_TrackBar = $using:max_size_TrackBar
            $min_size_TrackBar = $using:min_size_TrackBar
            ##################################################################################
            ###########Initialize Interface

            ##Setup Function Variables
        
           

            # Default Random Color
            $red   = (Get-Random -minimum 0 -maximum 255)
            $green = (Get-Random -minimum 0 -maximum 255)
            $blue  = (Get-Random -minimum 0 -maximum 255)

            # Create Image
            $bitmap = New-Object System.Drawing.Bitmap($imageWidth,$imageHeight)
            $bitmapGraphics = [System.Drawing.Graphics]::FromImage($bitmap)
            $brushBorder = New-Object System.Drawing.SolidBrush("Black")

            # Image Background Color
            If ($backgroundComboBox.Text -eq "Transparent") {
                $bitmap.MakeTransparent()
            } Else {
                
                If ($backgroundComboBox.Text -eq "Random") {
                    $red   = (Get-Random -minimum 0 -maximum 255)
                    $green = (Get-Random -minimum 0 -maximum 255)
                    $blue  = (Get-Random -minimum 0 -maximum 255)
                    $backColor = [System.Drawing.Color]::FromArgb($red, $green, $blue)
                    $bitmapGraphics.Clear($backColor)
                } Else {
                    $bitmapGraphics.Clear($backgroundComboBox.Text)
                }
            }

            # Image Border Color
            If ($borderColorComboBox.Text -eq "Random") {
                $red   = (Get-Random -minimum 0 -maximum 255)
                $green = (Get-Random -minimum 0 -maximum 255)
                $blue  = (Get-Random -minimum 0 -maximum 255)
                $borderColor = [System.Drawing.Color]::FromArgb($red, $green, $blue)
                $brushBorder = New-Object System.Drawing.SolidBrush($borderColor)
            } Else {
                $brushBorder = New-Object System.Drawing.SolidBrush($borderColorComboBox.Text)
            }
                      
            $i = 0;
            While ($i -le $numberTrackBar.Value) 
            { 
                $pointX = 0
                $pointY = 0
                $width = 0
                $height = 0;
                $i++
                $file_number = Get-Random -minimum 0 -maximum $files.count
                $fileName = $files[$file_number].FullName

                
                Try
                {
                    $imageFile = [System.Drawing.Image]::FromFile($fileName)
                
                    #Write-Output $imageFile.GetPropertyItem(274)
                    $value = $imageFile.GetPropertyItem(274).Value[0]
                    
                    
                    if($value -eq 6)
                    {
                        $imageFile.rotateflip("Rotate90FlipNone")
                    }
                    if($value -eq 8)
                    {
                        $imageFile.rotateflip("Rotate270FlipNone")
                    }
                    
                } catch {}

                $tries = 0;
                $go = 0;
                $pointX = (Get-Random -minimum -250 -maximum ([int]$imageWidth + 250))
                $pointY = (Get-Random -minimum -250 -maximum ([int]$imageHeight + 250))

                

                while($tries -le 150)
                { 
                    $height = 0;
                    $width = 0;
                    $tries++;
                    ##See if Wider or Taller
                    if($imageFile.Height -ge $imageFile.Width)
                    {
                        if($imageFile.Height -ge $max_size_TrackBar.Value)
                        {
                            $height = $default_max_image_size
                        }
                        ##Randomize Size/Fix Aspect Ratio
                        [int]$height = Get-Random -minimum ([int]$min_size_TrackBar.Value) -maximum ([int]$max_size_TrackBar.Value)
                        $aspect_ratio = $imageFile.Height / $height
                        [int]$width = $imageFile.Width / $aspect_ratio;
                    }
                    else
                    {
                        if($imageFile.Width -ge $max_size_TrackBar.Value)
                        {
                            $width = $default_max_image_size
                        }
                        ##Randomize Size/Fix Aspect Ratio
                        [int]$width = Get-Random -minimum ($min_size_TrackBar.Value) -maximum ([int]$max_size_TrackBar.Value)
                        $aspect_ratio = $imageFile.width / $width
                        [int]$height = $imageFile.Height / $aspect_ratio;
                    }

                    ##Get Image Location Areas
                    $c1x  = $pointX
                    $c1y  = $pointY
                    $c2x  = $pointX + $width
                    $c2y  = $pointY
                    $c3x  = $pointX
                    $c3y  = $pointY + $height
                    $c4x  = $pointX + $width
                    $c4y  = $pointY + $height

                    $failed = 0;
                    foreach ($key in $restricted_zones.keys) 
                    {
                        ([int]$x1,[int]$x2,[int]$y1,[int]$y2) = $key -split " - "
                        if((($c1x -in $x1..$x2) -or ($x1 -in $c1x..$c4x) -or ($x2 -in $c1x..$c4x)) -and (($c1y -in $y1..$y2) -or ($y1 -in $c1y..$c4y) -or ($y2 -in $c1y..$c4y)))
                        {
                            ##Placement Overlaps too much try and find a place nearby
                            $failed++;
                            $pointX = $pointX + (Get-Random -minimum -100 -maximum 100)
                            $pointY = $pointY + (Get-Random -minimum -100 -maximum 100)
                        }
                    }
                    
                    if($failed -eq 0)
                    {
                        $go = 1;
                        break
                    }
                }
                ##Prepare for First Picture
                if($restricted_zones.count -eq 0)
                {
                    $pointX = -10
                    $pointY = -10
                    $go = 1;
                }

                ##Execute Placment
                if($go -eq 1)
                {
                    ##Find Image Placement for Zoning
                    [int]$zone_x1 = ($pointX + (($overlap / 2) * ($width / 100)))
                    [int]$zone_x2 = (($pointX + $width) - (($overlap / 2) * ($width / 100)))
                    [int]$zone_y1 = ($pointY + (($overlap / 2) * ($height / 100)))
                    [int]$zone_y2 = (($pointY + $height) - (($overlap / 2) * ($height / 100)))
                    if(($zone_x1 -ne $zone_x2) -or ($zone_y1 -ne $zone_y2))
                    {
                        #Write-Output "$zone_x1 - $zone_x2 - $zone_y1 - $zone_y2  = 1"
                        if(!($restricted_zones.containskey("$zone_x1 - $zone_x2 - $zone_y1 - $zone_y2")))
                        {
                            #Write-Output "$zone_x1 - $zone_x2 - $zone_y1 - $zone_y2  = 2"
                            $restricted_zones.add("$zone_x1 - $zone_x2 - $zone_y1 - $zone_y2",0)
                        }
                    }

                    ##Draw Border
                    If ($border_width -gt 0) 
                    {
	                    $bitmapGraphics.FillRectangle($brushBorder, `
                        $pointX-$border_width, `
                        $pointY-$border_width, `
                        $width +($border_width*2), `
                        $height+($border_width*2))
                    }

                    ##Place Bitmap on Canvas
                    $bitmapGraphics.DrawImage($imageFile, $pointX, $pointY, $width, $height)

                    $duration = (Get-Date) - $save_timer
                    if(($duration.TotalSeconds -gt 4 -or ($i -eq 1)))
                    {    
                        $save_timer = Get-Date
                        $bitmap.Save($outFile)
                        write-Output "Updated"
                    }
                    else
                    {
                        ##Update Progress Bar
                        [int]$status = (($i/ $Global:default_pic_number) * 100)
                        if($status -gt 100){$status = 100;}
                        write-output $status
                    }



                }
                else
                {
                    #write-output Skipped
                }
                
            }
            $bitmap.Save($outFile)
            $bitmap.Dispose()
            $bitmapGraphics.Dispose()
        }
        ##################################################################################
        ###########Start Job and Display Updates
        $script:load_image_timer = Get-Date
        $first = 1;
        $Global:job = Start-Job -ScriptBlock  $build_collage
        Do {[System.Windows.Forms.Application]::DoEvents() 
            $status = $Global:job.ChildJobs.Output | Select-Object -Last 1
            if($status -match "\d")
            {
                $progressBar1.Value = $status
            }
            else
            {   
                $duration = (Get-Date) - $script:load_image_timer
                if(($duration.TotalSeconds -gt 3) -or ($first -eq 1))
                {
                    $first = 2;
                    #write-host $Global:job.ChildJobs.Output
                    $script:load_image_timer = Get-Date
                    if(Test-Path $Global:outFile)
                    {
                        try
                        {
                            $bitmap = [System.Drawing.Image]::Fromfile($Global:outFile);

                            if($bitmap)
                            {
                                $formGraphics.DrawImage($bitmap, 530, 15, $wwidth, $wheight)
                                $bitmap.Dispose()
                            }
                            else
                            {
                                write-host Failed
                            }
                        }
                        catch
                        {
                            write-host Failed to Update Preview
                        }
                    }
                }
            }
        } Until (($Global:job.State -ne "Running"))
        
        $bitmap = [System.Drawing.Image]::Fromfile($Global:outFile);
        $formGraphics.DrawImage($bitmap, 530, 15, $wwidth, $wheight)
        $progressBar1.Value = 100
        $buttonProcess.Text = "Process"
        #$message = "Complete`n`n"
        #$yesno = [System.Windows.Forms.MessageBox]::Show("$message","Complete", "Ok" , "Information")
        
        ##################################################################################
        ###########Finalize Work
        


    }#Files Count
    if($errors.count -gt 0)
    {      
        errors($errors)
    }
    $buttonDisplay.Enabled = $true
 

}
################################################################################
######Color Picker##############################################################
function color_picker
{
    $colorDialog = new-object System.Windows.Forms.ColorDialog
    $colorDialog.AllowFullOpen = $true
    $colorDialog.FullOpen = $true
    $colorDialog.color = "Black"
    $colorDialog.ShowDialog()
    [string]$colors = (" " + ([System.Drawing.Color] | gm -Static -MemberType Properties).name + " ")
    if($colors -match $colorDialog.Color.Name)
    {
        return $colorDialog.Color.Name
    }
    else
    {
        $red = [System.Convert]::ToString($colordialog.color.R,16)
        $green = [System.Convert]::ToString($colordialog.color.G,16)
        $blue = [System.Convert]::ToString($colordialog.color.B,16)
        if($red.length -eq 1){$red = "0" + "$red"}
        if($green.length -eq 1){$green = "0" + "$green"}
        if($blue.length -eq 1){$blue = "0" + "$blue"}
        $hex = "$red" + "$green" + "$blue"
        return "#$hex"
    }     
}
#################################################################################
#####Select Folder###############################################################
function selectFolderIn 
{
	$selectForm = New-Object System.Windows.Forms.FolderBrowserDialog
    $selectForm.SelectedPath = $Global:default_input_path
	$getKey = $selectForm.ShowDialog()
	If ($getKey -eq "OK") 
    {
            $Global:default_input_path = $selectForm.SelectedPath
        	$textBoxIn.Text = $selectForm.SelectedPath
            save_settings
	}
}
function selectFolderOut 
{
	$selectForm = New-Object System.Windows.Forms.FolderBrowserDialog
    $selectForm.SelectedPath = $Global:default_output_path
	$getKey = $selectForm.ShowDialog()
	If ($getKey -eq "OK") 
    {
            $Global:default_output_path = $selectForm.SelectedPath
        	$textBoxOut.Text = $selectForm.SelectedPath
            save_settings
	}
}
#################################################################################
#####Display Image###############################################################
Function DisplayImage 
{
    If ($Global:outFile.Length -gt 0) 
    {
        Invoke-Item $Global:outFile
    }
}
#################################################################################
#################################################################################
function errors($errors)
{
    $message = "Please fix the following errors:`n`n"
    $counter = 0;
    foreach($error in $errors)
    {
        if($error -ne "")
        {
            $counter++;
            $message = $message + "$counter - $error`n"
        } 
    }
    [System.Windows.MessageBox]::Show($message,"Error",'Ok','Error')
}
load_settings
main

