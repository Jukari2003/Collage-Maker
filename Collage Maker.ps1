################################################################################
#                                 Collage Maker                                #
#                         Written By: Anthony Brechtel                         #
#                                    Ver 3.1                                   #
#                                                                              #
################################################################################
#  Version History

#  Ver 1.0 - 17 May 2021 
        #Working Proof of Concept 

#  Ver 2.0 - 28 Nov 2021 
        #Improved User Interface
        #Moved Collage process to Child-Job
        #Reduced Crashes 

#  Ver 3.0 - 13 Jan 2023 
        #Added Support for Video Collages
        #Added Error Checking
        #Switched to FFmpeg Image placement (More Stable)
        #Improved Interface

# Ver 3.1 - 17 Jan 2023
        #Fixed issue with network drives

################################################################################
######Global Variables##########################################################
clear-host
$script:dir                         = split-path -parent $MyInvocation.MyCommand.Definition
$script:outFile                     = ""
$script:settings_file               = "$script:dir\Settings.txt"
$script:default_input_path          = $script:dir
$script:default_output_path         = "$script:dir\Output"
$script:background_color            = "Black"
$script:collage_type                = "Image"
$script:border_color                = "Black"
[int]$script:default_pic_number     = 100
[int]$script:default_min_image_size = 400
[int]$script:default_max_image_size = 600
$script:border_width                = 2;
$script:overlap                     = 50;
$script:canvas_duration             = 5
$script:off_screen_pixels           = 150
$script:status_image                = "$script:dir" + "\Buffer\Status.bmp"
$script:status_video                = "$script:dir" + "\Buffer\Status.mp4"


################################################################################
######Load Assemblies###########################################################
[System.Windows.Forms.Application]::EnableVisualStyles();
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing


################################################################################
######Get Screen Variables #####################################################
$screen = [System.Windows.Forms.SystemInformation]::PrimaryMonitorSize
if($screen.width)
{
    [int]$script:canvas_width = $screen.width
}
else
{
    [int]$script:canvas_width = 1024
}
if($screen.height)
{
    [int]$script:canvas_height = $screen.height
}
else
{
    [int]$script:canvas_height = 768
}


#################################################################################
#####Loading Settings############################################################
function load_settings
{
    if(!(Test-path -literalpath "$script:dir\Output"))
    {
        New-Item  -ItemType directory -Path "$script:dir\Output"
    }
    if(!(Test-path -literalpath "$script:dir\Buffer"))
    {
        New-Item  -ItemType directory -Path "$script:dir\Buffer"
    }
    if(test-path -literalpath $script:settings_file)
    {
        $reader = [System.IO.File]::OpenText($settings_file)
        while($null -ne ($line = $reader.ReadLine())) 
        {

            $line_split = $line -split ':::';
            ###################
            if($line_split[0] -eq "COLLAGE_TYPE")
            {
                $script:collage_type = $line_split[1];
            }
            ###################
            if($line_split[0] -eq "FFMPEG")
            {
                if(($line_split[1].length -gt 3) -and (Test-Path -literalpath $line_split[1]))
                {
                    $script:ffmpeg = $line_split[1];
                }
            }
            ###################
            if($line_split[0] -eq "OUTPUT_DIR")
            {
                if(Test-Path -literalpath $line_split[1])
                {
                    $script:default_output_path = $line_split[1];
                }
            }
            ###################
            if($line_split[0] -eq "INPUT_DIR")
            {
                if(Test-Path -literalpath $line_split[1])
                {
                    $script:default_input_path = $line_split[1];
                }
            }
            ###################
            if($line_split[0] -eq "PIC_NUMBER")
            {
                $script:default_pic_number = $line_split[1];
            }
            ###################
            if($line_split[0] -eq "WALLPAPER_WIDTH")
            {
                $script:canvas_width = $line_split[1];
            }
            ###################
            if($line_split[0] -eq "WALLPAPER_HEIGHT")
            {
                $script:canvas_height = $line_split[1];
            }
            ###################
            if($line_split[0] -eq "MIN_SIZE")
            {
                $script:default_min_image_size = $line_split[1];
            }
            ###################
            if($line_split[0] -eq "MAX_SIZE")
            {
                $script:default_max_image_size = $line_split[1];
            }
            ###################
            if($line_split[0] -eq "OVERLAP")
            {
                $script:overlap = $line_split[1];
            }
            ###################
            if($line_split[0] -eq "BACKGROUND_COLOR")
            {
                $script:background_color = $line_split[1];
            }
            ###################
            if($line_split[0] -eq "BORDER_COLOR")
            {
                $script:border_color = $line_split[1];
            }
            ###################
            if($line_split[0] -eq "BORDER_WIDTH")
            {
                $script:border_width = $line_split[1];
            } 
            ###################
            if($line_split[0] -eq "DURATION")
            {
                $script:canvas_duration = $line_split[1];
            }
            ###################
            if($line_split[0] -eq "OFF_SCREEN")
            {
                $script:off_screen_pixels = $line_split[1];
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
    Add-Content $settings_file "COLLAGE_TYPE:::$script:collage_type"
    Add-Content $settings_file "DURATION:::$script:canvas_duration"
    Add-Content $settings_file "OFF_SCREEN:::$script:off_screen_pixels"
    Add-Content $settings_file "FFMPEG:::$script:ffmpeg"
    Add-Content $settings_file "OUTPUT_DIR:::$script:default_output_path"
    Add-Content $settings_file "INPUT_DIR:::$script:default_input_path"
    Add-Content $settings_file "PIC_NUMBER:::$script:default_pic_number"
    Add-Content $settings_file "WALLPAPER_WIDTH:::$script:canvas_width"
    Add-Content $settings_file "WALLPAPER_HEIGHT:::$script:canvas_height"
    Add-Content $settings_file "MIN_SIZE:::$script:default_min_image_size"
    Add-Content $settings_file "MAX_SIZE:::$script:default_max_image_size"
    Add-Content $settings_file "OVERLAP:::$script:overlap"
    Add-Content $settings_file "BACKGROUND_COLOR:::$script:background_color"
    Add-Content $settings_file "BORDER_WIDTH:::$script:border_width"
    Add-Content $settings_file "BORDER_COLOR:::$script:border_color"
}
#################################################################################
#####Main########################################################################
function main
{
    ##################################################################################
    ###########Main Form
    $Form = New-Object System.Windows.Forms.Form
    $Form.Location = "200, 200"
    $Form.Font = "Copperplate Gothic,8.1"
    $Form.FormBorderStyle = "FixedDialog"
    $Form.ForeColor = "Black"
    $Form.BackColor = "#434343"
    $Form.Text = "  Collage Maker"
    $Form.Width = 530 #1245
    $Form.Height = 600


    ##################################################################################
    ###########Collage Type Label
    $y_pos = 12;
    $collage_type_label = New-Object System.Windows.Forms.Label
    $collage_type_label.Size = "115, 18"
    $collage_type_label.Location = New-Object System.Drawing.Point ((($Form.width / 2) - $collage_type_label.width), ($y_pos + 2))  
    $collage_type_label.ForeColor = "White"
    #$collage_type_label.backColor = "Green"
    $collage_type_label.Text = "Collage  Type"
    $Form.Controls.Add($collage_type_label)


    ##################################################################################
    ###########Collage Type Combo
    $collage_type_combo_box = New-Object System.Windows.Forms.ComboBox
    $collage_type_combo_box.Location = New-Object System.Drawing.Point ((($Form.width / 2)), $y_pos)
    $collage_type_combo_box.Size = "140,20"
    $collage_type_combo_box.ForeColor = "Indigo"
    $collage_type_combo_box.BackColor = "White"
    [void]$collage_type_combo_box.items.add("Image")
    [void]$collage_type_combo_box.items.add("Video")
    [void]$collage_type_combo_box.items.add("Image & Video")
    if(!($collage_type_combo_box.items.Contains($script:collage_type)))
    {
        [void]$collage_type_combo_box.items.add("$script:collage_type")
    }        
    $collage_type_combo_box.SelectedIndex = 0
    $collage_type_combo_box.SelectedIndex = $collage_type_combo_box.FindStringExact("$script:collage_type")
    $collage_type_combo_box.Add_SelectedValueChanged({
        $script:collage_type = $collage_type_combo_box.text;

        if($script:collage_type -eq "Image")
        {
            $canvas_duration_input.text = 1
            $canvas_duration_input.enabled = $false
        }
        else
        {
            $canvas_duration_input.enabled = $true
        }

        save_settings
    })
    $Form.Controls.Add($collage_type_combo_box)


    ##################################################################################
    ###########FFmpeg Location Label
    $y_pos = $y_pos + 30
    $ffmpeg_location_label = New-Object System.Windows.Forms.Label 
    $ffmpeg_location_label.Location = New-Object System.Drawing.Point(15,($y_pos))
    $ffmpeg_location_label.Size = "300, 18"
    $ffmpeg_location_label.ForeColor = "White"
    $ffmpeg_location_label.Text = "FFmpeg Location"
    $Form.Controls.Add($ffmpeg_location_label)


    ##################################################################################
    ###########FFmpeg Input
    $y_pos = $y_pos + 20
    $ffmpeg_box = New-Object System.Windows.Forms.TextBox
    $ffmpeg_box.Location = New-Object System.Drawing.Point(15, $y_pos)
    $ffmpeg_box.Size = "420, 20"
    if(($script:ffmpeg -eq "") -or ($script:ffmpeg -eq $null) -or (!(Test-Path -literalpath $script:ffmpeg)))
    {
        $ffmpeg_box.text = "Browse or Enter a file path for FFmpeg.exe"
    }
    else
    {
        $ffmpeg_box.text = $script:ffmpeg
    }
    $ffmpeg_box.Add_Click({
        if($ffmpeg_box.Text -eq "Browse or Enter a file path for FFmpeg.exe")
        {
            $ffmpeg_box.Text = ""
            $ffmpeg_box.Text = ""
        }
    })
    $ffmpeg_box.Add_TextChanged({
        
            if(($this.text -ne $Null) -and ($this.text -ne "") -and (Test-Path -literalpath $this.text) -and ($this.text -match ".exe$"))
            {
                $script:ffmpeg = $ffmpeg_box.Text
                $ffmpeg_box.Text = $script:ffmpeg
                save_settings
            }
            else
            {
                $script:ffmpeg = "";
                save_settings
            }
    })
    $ffmpeg_box.Add_lostFocus({

        if(($script:ffmpeg -eq "") -or ($script:ffmpeg -eq $null) -or (!(Test-Path -literalpath $script:ffmpeg)))
        {
            $this.text = "Browse or Enter a file path for FFmpeg.exe"
            $ffmpeg_box.Text = "Browse or Enter a file path for FFmpeg.exe"
        }
    })
    $Form.Controls.Add($ffmpeg_box)


    ##################################################################################
    ###########FFmpeg Browse
    $ffmpeg_browse_button = New-Object System.Windows.Forms.Button
    $ffmpeg_browse_button.Location= New-Object System.Drawing.Size(440, $y_pos)
    $ffmpeg_browse_button.BackColor = "#606060"
    $ffmpeg_browse_button.ForeColor = "White"
    $ffmpeg_browse_button.Size = "75, 23"
    $ffmpeg_browse_button.Text='Browse'
    $ffmpeg_browse_button.Add_Click(
    {    
        $return = prompt_for_file_exe
        if($return.length -ge 3)
        {
            $ffmpeg_box.text = $return
        }
    })
    $Form.Controls.Add($ffmpeg_browse_button)


    ##################################################################################
    ###########Input Label
    $y_pos = $y_pos + 40;
    $input_images_label = New-Object System.Windows.Forms.Label
    $input_images_label.Location = "15, $y_pos"
    $input_images_label.Size = "300, 18"
    $input_images_label.ForeColor = "White" 
    $input_images_label.Text = "Input Images Folder"
    $Form.Controls.Add($input_images_label)


    ##################################################################################
    ###########Path Input
    $y_pos = $y_pos + 20
    $textBoxIn = New-Object System.Windows.Forms.TextBox
    $textBoxIn.text = $script:default_input_path
    $textBoxIn.Location = "15, $y_pos"
    $textBoxIn.Size = "420, 20"
    $textBoxIn.Add_TextChanged({
        if($textBoxIn.text -and (test-path -literalpath $textBoxIn.text))
        {
            $script:default_input_path = $textBoxIn.text;
            save_settings
        }
        else
        {
            $script:default_input_path = "";
        }
    })
    $Form.Controls.Add($textBoxIn)
   

    ##################################################################################
    ###########Browse Button
    $y_pos = $y_pos - 2
    $buttonBrowse = New-Object System.Windows.Forms.Button
    $buttonBrowse.Location = "440, $y_pos"
    $buttonBrowse.Size = "75, 23"
    $buttonBrowse.ForeColor = "White"
    $buttonBrowse.Backcolor = "#606060"
    $buttonBrowse.Text = "Browse"
    $buttonBrowse.add_Click({selectFolderIn})
    $Form.Controls.Add($buttonBrowse)


    ##################################################################################
    ###########Picture Number Trackbar
    $y_pos = $y_pos + 30
    $numberTrackBar = New-Object System.Windows.Forms.TrackBar
    $numberTrackBar.Location = "230, $y_pos"
    $numberTrackBar.Orientation = "Horizontal"
    $numberTrackBar.Width = 290
    $numberTrackBar.Height = 40
    $numberTrackBar.TickFrequency = 100
    $numberTrackBar.TickStyle = "TopLeft"
    $numberTrackBar.SetRange(1, 2000)
    $numberTrackBar.Value = $script:default_pic_number
    $numberTrackBarValue = $script:default_pic_number
    $numberTrackBar.add_ValueChanged({
        $numberTrackBarValue = $numberTrackBar.Value
        $numberTrackBarLabel.Text = "Number of Images ($numberTrackBarValue)"
        $script:default_pic_number = $numberTrackBarValue
        save_settings
    })
    $Form.Controls.add($numberTrackBar)


    ##################################################################################
    ###########Picture Number Label
    $y_pos = $y_pos + 10
    $numberTrackBarLabel = New-Object System.Windows.Forms.Label 
    $numberTrackBarLabel.Location = "15,$y_pos"
    $numberTrackBarLabel.Size = "215,23"
    $numberTrackBarLabel.ForeColor = "White"
    $numberTrackBarLabel.Text = "Number of Images ($numberTrackBarValue)"
    $Form.Controls.Add($numberTrackBarLabel)


    ##################################################################################
    ###########Minimum Number of Images Trackbar
    $y_pos = $y_pos + 40
    $min_size_TrackBar = New-Object System.Windows.Forms.TrackBar
    $min_size_TrackBar.Location = "230,$y_pos"
    $min_size_TrackBar.Orientation = "Horizontal"
    $min_size_TrackBar.Width = 290
    $min_size_TrackBar.Height = 10
    $min_size_TrackBar.TickFrequency = 100
    $min_size_TrackBar.TickStyle = "TopLeft"
    $min_size_TrackBar.SetRange(100, 1000)
    $min_size_TrackBar.Value = $script:default_min_image_size
    $min_size_TrackBarValue = $script:default_min_image_size
    $min_size_TrackBar.add_ValueChanged({
        $min_size_TrackBarValue = $min_size_TrackBar.Value
        $min_size_TrackBarLabel.Text = "Minimum Image Size ($min_size_TrackBarValue)"
        $script:default_min_image_size = $min_size_TrackBarValue
        save_settings
    })
    $Form.Controls.add($min_size_TrackBar)


    ##################################################################################
    ###########Minimum Number of Images Label
    $y_pos = $y_pos + 10
    $min_size_TrackBarLabel = New-Object System.Windows.Forms.Label 
    $min_size_TrackBarLabel.Location = "15,$y_pos"
    $min_size_TrackBarLabel.Size = "215,23"
    $min_size_TrackBarLabel.ForeColor = "White"
    $min_size_TrackBarLabel.Text = "Minimum Image Size ($min_size_TrackBarValue)"
    $Form.Controls.Add($min_size_TrackBarLabel)


    ##################################################################################
    ###########Max Image Size Trackbar
    $y_pos = $y_pos + 40
    $max_size_TrackBar = New-Object System.Windows.Forms.TrackBar
    $max_size_TrackBar.Location = "230,$y_pos"
    $max_size_TrackBar.Orientation = "Horizontal"
    $max_size_TrackBar.Width = 290
    $max_size_TrackBar.Height = 10
    $max_size_TrackBar.LargeChange = 20
    $max_size_TrackBar.SmallChange = 1
    $max_size_TrackBar.TickFrequency = 200
    $max_size_TrackBar.TickStyle = "TopLeft"
    $max_size_TrackBar.SetRange(300, 3000)
    $max_size_TrackBar.Value = $script:default_max_image_size
    $max_size_TrackBarValue = $script:default_max_image_size
    $max_size_TrackBar.add_ValueChanged({
        $max_size_TrackBarValue = $max_size_TrackBar.Value
        $max_size_TrackBarLabel.Text = "Maximum Image Size ($max_size_TrackBarValue)"
        $script:default_max_image_size = $max_size_TrackBarValue
        save_settings
    })
    $Form.Controls.add($max_size_TrackBar)


    ##################################################################################
    ###########Max Image Size Label
    $y_pos = $y_pos + 10
    $max_size_TrackBarLabel = New-Object System.Windows.Forms.Label 
    $max_size_TrackBarLabel.Location = "15,$y_pos"
    $max_size_TrackBarLabel.Size = "215,23"
    $max_size_TrackBarLabel.ForeColor = "White"
    $max_size_TrackBarLabel.Text = "Maximum Image Size ($max_size_TrackBarValue)"
    $Form.Controls.Add($max_size_TrackBarLabel)


    ##################################################################################
    ###########Overlap Trackbar
    $y_pos = $y_pos + 40
    $script:overlap_TrackBar = New-Object System.Windows.Forms.TrackBar
    $script:overlap_TrackBar.Location = "230,$y_pos"
    $script:overlap_TrackBar.Orientation = "Horizontal"
    $script:overlap_TrackBar.Width = 290
    $script:overlap_TrackBar.Height = 10
    $script:overlap_TrackBar.TickFrequency = 10
    $script:overlap_TrackBar.TickStyle = "TopLeft"
    $script:overlap_TrackBar.SetRange(0, 100)
    $script:overlap_TrackBar.Value = $script:overlap
    $script:overlap_TrackBarValue = $script:overlap
    $script:overlap_TrackBar.add_ValueChanged({
        $script:overlap_TrackBarValue = $script:overlap_TrackBar.Value
        $script:overlap_TrackBarLabel.Text = "Overlap ($script:overlap_TrackBarValue)"
        $script:overlap = $script:overlap_TrackBarValue
        save_settings
    })
    $Form.Controls.add($script:overlap_TrackBar)


    ##################################################################################
    ###########Overlap Label
    $y_pos = $y_pos + 10
    $script:overlap_TrackBarLabel = New-Object System.Windows.Forms.Label 
    $script:overlap_TrackBarLabel.Location = "15,$y_pos"
    $script:overlap_TrackBarLabel.Size = "215,23"
    $script:overlap_TrackBarLabel.ForeColor = "White"
    $script:overlap_TrackBarLabel.Text = "Overlap ($script:overlap_TrackBarValue%)"
    $Form.Controls.Add($script:overlap_TrackBarLabel)


    ##################################################################################
    ###########Background Color ComboBox
    $y_pos = $y_pos + 45
    $backgroundComboBox = New-Object System.Windows.Forms.ComboBox
    $backgroundComboBox.Location = New-Object System.Drawing.Point (165,$y_pos)
    $backgroundComboBox.size = "110,20"
    $backgroundComboBox.ForeColor = "Indigo"
    $backgroundComboBox.BackColor = "White"
    [void]$backgroundComboBox.items.add("Black")
    [void]$backgroundComboBox.items.add("White")
    [void]$backgroundComboBox.items.add("Random")
    [void]$backgroundComboBox.items.add("Select Color")
    if(!($backgroundComboBox.items.Contains($script:background_color)))
    {
        [void]$backgroundComboBox.items.add("$script:background_color")
    }
    $backgroundComboBox.SelectedIndex = $backgroundComboBox.FindStringExact("$script:background_color")
    $backgroundComboBox.Add_SelectedValueChanged({
        if($this.text -eq "Select Color")
            {
                $new_color = color_picker
                write-host $new_color[1]
                $this.items.add($new_color[1])
                $backgroundComboBox.text = $new_color[1]
                $script:background_color = $new_color[1]
                save_settings

            }

        $script:background_color = $backgroundComboBox.text;
        save_settings

    })
    $Form.Controls.Add($backgroundComboBox)
    

    ##################################################################################
    ###########Background Color Label    
    $backgroundComboBoxLabel = New-Object System.Windows.Forms.Label 
    $backgroundComboBoxLabel.Location = New-Object System.Drawing.Point (15,($y_pos + 2))
    $backgroundComboBoxLabel.Size = "160,20"
    $backgroundComboBoxLabel.ForeColor = "White"
    $backgroundComboBoxLabel.Text = "Background Color"
    $Form.Controls.Add($backgroundComboBoxLabel)
    

    ##################################################################################
    ###########Border Color Combo
    $borderColorComboBox = New-Object System.Windows.Forms.ComboBox
    $borderColorComboBox.Location = "410,$y_pos"
    $borderColorComboBox.Size = "110,20"
    $borderColorComboBox.ForeColor = "Indigo"
    $borderColorComboBox.BackColor = "White"
    [void]$borderColorComboBox.items.add("Black")
    [void]$borderColorComboBox.items.add("White")
    [void]$borderColorComboBox.items.add("Random")
    [void]$borderColorComboBox.items.add("Random All")
    [void]$borderColorComboBox.items.add("Select Color")
    if(!($borderColorComboBox.items.Contains($script:border_color)))
    {
        [void]$borderColorComboBox.items.add("$script:border_color")
    }        
    $borderColorComboBox.SelectedIndex = 0
    $borderColorComboBox.SelectedIndex = $borderColorComboBox.FindStringExact("$script:border_color")
    $borderColorComboBox.Add_SelectedValueChanged({
        if($this.text -eq "Select Color")
        {
            $new_color = color_picker
            write-host $new_color[1]
            $this.items.add($new_color[1])
            $borderColorComboBox.text = $new_color[1]
            $script:border_color = $new_color[1]
            save_settings

        }
        $script:border_color = $borderColorComboBox.text;
        save_settings

    })
    $Form.Controls.Add($borderColorComboBox)


    ##################################################################################
    ###########Border Color Label
    $borderColorComboBoxLabel = New-Object System.Windows.Forms.Label 
    $borderColorComboBoxLabel.Location = New-Object System.Drawing.Point (300,($y_pos + 2))
    $borderColorComboBoxLabel.Size = "180,20"
    $borderColorComboBoxLabel.ForeColor = "White"
    $borderColorComboBoxLabel.Text = "Border Color"
    $Form.Controls.Add($borderColorComboBoxLabel)


    ##################################################################################
    ###########Off Screen Pixels Input
    $y_pos = $y_pos + 30  
    $off_screen_pixel_input = New-Object System.Windows.Forms.TextBox
    $off_screen_pixel_input.text = $script:off_screen_pixels
    $off_screen_pixel_input.Location = "165, $y_pos"
    $off_screen_pixel_input.Size = "110,20"
    $off_screen_pixel_input.Add_TextChanged({
        if($off_screen_pixel_input.text -and ($off_screen_pixel_input.text -match "^\d+$"))
        {
            [int]$script:off_screen_pixels = $off_screen_pixel_input.text;
            save_settings
        }
        else
        {
            $off_screen_pixel_input.text = 100;
            $script:off_screen_pixels = 100;
            save_settings
        }
    })
    $Form.Controls.Add($off_screen_pixel_input)
    

    ##################################################################################
    ###########Off Screen Pixels Label
    $off_screen_pixel_label = New-Object System.Windows.Forms.Label 
    $off_screen_pixel_label.Location = New-Object System.Drawing.Point (15,($y_pos + 2))
    $off_screen_pixel_label.Size = "150,20"
    $off_screen_pixel_label.ForeColor = "White"
    $off_screen_pixel_label.Text = "Off-Screen Pixels"
    $Form.Controls.Add($off_screen_pixel_label)


    ##################################################################################
    ###########Border Width Combo
    $borderComboBox = New-Object System.Windows.Forms.ComboBox
    $borderComboBox.Location = "410,$y_pos"
    $borderComboBox.Size = "110,20"
    $borderComboBox.ForeColor = "Indigo"
    $borderComboBox.BackColor = "White"
    [void]$borderComboBox.items.add("Random")
    [void]$borderComboBox.items.add("Random All")
    For ($i=0; $i -le 20; $i++) {
        [void]$borderComboBox.items.add($i)
    }
    $borderComboBox.SelectedIndex = $borderComboBox.FindStringExact("$script:border_width")
    $borderComboBox.Add_SelectedValueChanged({

        $script:border_width = $borderComboBox.text;
        save_settings

    })
    $Form.Controls.Add($borderComboBox)


    ##################################################################################
    ###########Border Width Label
    $borderComboBoxLabel = New-Object System.Windows.Forms.Label 
    $borderComboBoxLabel.Location = New-Object System.Drawing.Point (300,($y_pos + 2))
    $borderComboBoxLabel.Size = "200,20"
    $borderComboBoxLabel.ForeColor = "White"
    $borderComboBoxLabel.Text = "Border Width"
    $Form.Controls.Add($borderComboBoxLabel)


    ##################################################################################
    ###########Duration Input
    $y_pos = $y_pos + 30  
    $canvas_duration_input = New-Object System.Windows.Forms.TextBox
    $canvas_duration_input.text = $script:canvas_duration
    $canvas_duration_input.Location = "165, $y_pos"
    $canvas_duration_input.Size = "110,20"
    $canvas_duration_input.Add_TextChanged({
        if($canvas_duration_input.text -and ($canvas_duration_input.text -match "^\d+$"))
        {
            [int]$script:canvas_duration = $canvas_duration_input.text;
            if(($script:canvas_duration -gt 10) -and ($script:warned -ne 1))
            {
                $script:warned = 1;
                $message = "WARNING: A long duration will substantially increase processing time!"
                [System.Windows.MessageBox]::Show($message,"!!!WARNING!!!",'Ok')
            }
            save_settings
        }
        else
        {
            $canvas_duration_input.text = 5;
            $script:canvas_duration = 5;
            save_settings
        }
    })
    $Form.Controls.Add($canvas_duration_input)


    ##################################################################################
    ###########Duration Label  
    $canvas_duration_label = New-Object System.Windows.Forms.Label 
    $canvas_duration_label.Location = New-Object System.Drawing.Point (15,($y_pos + 2))
    $canvas_duration_label.Size = "160,20"
    $canvas_duration_label.ForeColor = "White"
    $canvas_duration_label.Text = "Duration Seconds"
    $Form.Controls.Add($canvas_duration_label)


    ##################################################################################
    ###########Screen Size Combo
    $screen = [System.Windows.Forms.SystemInformation]::PrimaryMonitorSize
    [string]$screen = [string]$screen.width + "x" + [string]$screen.Height
    $imageComboBox = New-Object System.Windows.Forms.ComboBox
    $imageComboBox.Location = "410,$y_pos"
    $imageComboBox.Size = "110,20"
    $imageComboBox.ForeColor = "Indigo"
    $imageComboBox.BackColor = "White"
    [void]$imageComboBox.items.add("$screen") 
    [void]$imageComboBox.items.add("800x600")
    [void]$imageComboBox.items.add("1024x768")
    [void]$imageComboBox.items.add("1200x1200")
    [void]$imageComboBox.items.add("1280x800")
    [void]$imageComboBox.items.add("1280x1024")
    [void]$imageComboBox.items.add("1366x768")
    [void]$imageComboBox.items.add("1440x900")
    [void]$imageComboBox.items.add("1600x900")
    [void]$imageComboBox.items.add("1680x1050")
    [void]$imageComboBox.items.add("1600x1200") 
    [void]$imageComboBox.items.add("1920x1200")
    [void]$imageComboBox.items.add("1920x1080")
    [void]$imageComboBox.items.add("2560x1440")
    [void]$imageComboBox.items.add("3840x2160")
    [void]$imageComboBox.items.add("5120x1440")   
    $imageComboBox.SelectedIndex = $imageComboBox.FindStringExact("$script:canvas_width" + "x" + "$script:canvas_height")
    $imageComboBox.Add_SelectedValueChanged({
        $script:canvas_width  = [int]$imageComboBox.Text.Split("x")[0]
        $script:canvas_height = [int]$imageComboBox.Text.Split("x")[1]
        save_settings
    })
    $Form.Controls.Add($imageComboBox)


    ##################################################################################
    ###########Screen Size Label
    $imageComboBoxLabel = New-Object System.Windows.Forms.Label 
    $imageComboBoxLabel.Location = New-Object System.Drawing.Point (300,($y_pos + 2))
    $imageComboBoxLabel.Size = "150,20"
    $imageComboBoxLabel.ForeColor = "White"
    $imageComboBoxLabel.Text = "Output Size"
    $Form.Controls.Add($imageComboBoxLabel)


    ##################################################################################
    ###########Output Label
    $y_pos = $y_pos + 40
    $outputLabel = New-Object System.Windows.Forms.Label
    $outputLabel.Location = "15, $y_pos"
    $outputLabel.Size = "150, 20"
    $outputLabel.ForeColor = "White"
    $outputLabel.Text = "Output Folder"
    $Form.Controls.Add($outputLabel)


    ##################################################################################
    ###########Output Text Box
    $y_pos = $y_pos + 20
    $textBoxOut = New-Object System.Windows.Forms.TextBox
    $textBoxOut.text = $default_output_path
    $textBoxOut.Location = "15, $y_pos"
    $textBoxOut.Size = "420, 20"
    $textBoxOut.Add_TextChanged({
        if($textBoxOut.text -and (test-path -literalpath $textBoxOut.text))
        {
            $script:default_output_path = $textBoxOut.text;
            save_settings
        }
        else
        {
            $script:default_output_path = "";
        }
    })
    $Form.Controls.Add($textBoxOut)


    ##################################################################################
    ###########Browse Button
    $y_pos = $y_pos - 2
    $outputBrowse = New-Object System.Windows.Forms.Button
    $outputBrowse.Location = "440, $y_pos"
    $outputBrowse.Size = "75, 23"
    $outputBrowse.ForeColor = "White"
    $outputBrowse.Backcolor = "#606060" 
    $outputBrowse.Text = "Browse"
    $outputBrowse.add_Click({selectFolderOut})
    $Form.Controls.Add($outputBrowse)

    
    ##################################################################################
    ###########Process Button
    $y_pos = $y_pos + 40
    $buttonProcess = New-Object System.Windows.Forms.Button
    $buttonProcess.Size = "100, 23"
    $buttonProcess.Location = New-Object System.Drawing.Point ((($Form.width / 2) - 170), $y_pos)
    $buttonProcess.ForeColor = "White"
    $buttonProcess.BackColor = "#606060"
    $buttonProcess.Text = "Process"
    $buttonProcess.add_Click({
    
        if($this.text -eq "Process")
        {
            $status = check_settings
            if($status -ne 1)
            {
                if(Test-Path -LiteralPath "$script:dir\Buffer")
                {
                    $buffer_files = Get-ChildItem "$script:dir\Buffer"
                    foreach($file in $buffer_files)
                    {
                        Remove-Item -LiteralPath $file.FullName
                    }
                }
                $completeLabel.Text = ""
                $script:progress_bar.visible = $true
                $script:progress_bar.Value = 0
                $this.text = "Stop"
                $completeLabel.ForeColor = "White"
                ProcessCollage
                $script:progress_bar.Value = 100
                $buttonProcess.Text = "Process"
                if($script:early_stop -ne 1)
                {
                    $completeLabel.Text = "Complete"
                    $completeLabel.ForeColor = "Green"
                    $completeLabel.Location = New-Object System.Drawing.Point ((($Form.width / 2) - ($completeLabel.width / 2)), $y_pos)
                }
            }
        }
        else
        {
            $this.text = "Stopping..."
            $this.Enabled = $false 
            $completeLabel.ForeColor = "Yellow"
            $completeLabel.Text = "Stopping..."

            Add-Content "$script:dir\Buffer\Exit.txt" ""
            $start_trying = (Get-Date)
            while((Test-Path -literalpath "$script:dir\Buffer\Exit.txt") -and ($start_trying.TotalSeconds -lt 15))
            {
                sleep 1;
            }
            $this.Enabled = $true 
            $this.text = "Process"
            $script:progress_bar.Value = 100
            #$script:progress_bar.visible = $false
            $completeLabel.Text = "Stopped By User"
            $completeLabel.ForeColor = "Red"
            $completeLabel.Location = New-Object System.Drawing.Point ((($Form.width / 2) - ($completeLabel.width / 2)), $y_pos)
        }
    
    
    })
    $Form.Controls.Add($buttonProcess)


    ##################################################################################
    ###########Display Button
    $buttonDisplay = New-Object System.Windows.Forms.Button
    $buttonDisplay.Size = "100, 23"
    $buttonDisplay.Location = New-Object System.Drawing.Point ((($Form.width / 2) - ($buttonDisplay.width / 2)), $y_pos)
    $buttonDisplay.Enabled = $false  
    $buttonDisplay.ForeColor = "White"
    $buttonDisplay.BackColor = "#606060" 
    $buttonDisplay.Text = "Display"
    $buttonDisplay.add_Click({DisplayImage})
    $Form.Controls.Add($buttonDisplay)


    ##################################################################################
    ###########Exit Button
    $buttonExit = New-Object System.Windows.Forms.Button
    $buttonExit.Size = "100, 23"
    $buttonExit.Location = New-Object System.Drawing.Point ((($Form.width / 2) + 70), $y_pos)
    $buttonExit.ForeColor = "White"
    $buttonExit.BackColor = "#606060" 
    $buttonExit.Text = "Exit"
    $buttonExit.add_Click({
        try{$Form.close();}
        catch {exit;}
    })
    $Form.Controls.Add($buttonExit)


    ##################################################################################
    ###########Progress Bar
    $y_pos = $y_pos + 40
    $script:progress_bar = New-Object System.Windows.Forms.ProgressBar
    $script:progress_bar.Location = "15, $y_pos"
    $script:progress_bar.Name = 'progressBar1'
    $script:progress_bar.Value = 0
    $script:progress_bar.Style="Continuous"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 0
    $System_Drawing_Size.Height = 0
    $script:progress_bar.Size = $System_Drawing_Size
    $Form.Controls.Add($script:progress_bar)


    ##################################################################################
    ###########Complete Label
    $y_pos = $y_pos + 30
    $completeLabel = New-Object System.Windows.Forms.Label
    $completeLabel.Size = "530, 18"
    $completeLabel.Location = New-Object System.Drawing.Point ((($Form.width / 2) - ($completeLabel.width / 2)), $y_pos)
    $completeLabel.ForeColor = "Green"
    $completeLabel.Text = "sdfsdfsf"
    $completeLabel.TextAlign = "MiddleCenter"
    $Form.Controls.Add($completeLabel)


    [void] $Form.ShowDialog()
}
#################################################################################
#####Check Settings #############################################################
function check_settings
{
    $status = 0;
    [array]$errors = "";
    if(($script:ffmpeg.Length -lt 5) -or (!(Test-Path -literalpath "$script:ffmpeg" -PathType Leaf)))
    {
        $errors += "Invalid FFmpeg Location"

    }
    if(($script:default_output_path.length -le 3) -or (!(Test-Path -literalpath "$script:default_output_path" -PathType Container)))
    {
        $errors += "Invalid Output Directory"
    }
    if(($script:default_input_path.length -le 3) -or (!(Test-Path -literalpath "$script:default_input_path"  -PathType Container)))
    {
        $errors += "Invalid Input Directory"
    }
    if($errors.count -ne 1)
    {
        $status = 1;
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
    return $status
}
#################################################################################
#####Process Collage#############################################################
Function ProcessCollage 
{
    clear-host;
###################################################################################################################################################################################################################################################
######Build Collage Job Start #####################################################################################################################################################################################################################
$build_collage = {

    #################################################################################
    ######Function Check For Exit ###################################################
    function check_exit
    {
        if(Test-Path -LiteralPath "$script:dir\Buffer\Exit.txt")
        {
            Write-Output "STOP-"
            Write-Output "Stopped by User!"
            if(Test-Path -LiteralPath "$script:dir\Buffer")
            {
                $buffer_files = Get-ChildItem "$script:dir\Buffer"
                foreach($file in $buffer_files)
                {
                    Remove-Item -LiteralPath $file.FullName -Force
                }
            }
            exit;
        }
    }


    #################################################################################
    ######Job Transfered Variables###################################################
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    $script:dir                     = $using:dir
    $script:ffmpeg                  = $using:ffmpeg
    $script:ffprobe                 = $script:ffmpeg -replace "ffmpeg.exe","ffprobe.exe"
    $script:collage_type            = $using:collage_type
    $script:max_time_seconds        = $using:canvas_duration
    $script:default_input_path      = $using:default_input_path 
    $script:default_output_path     = $using:default_output_path
    $script:default_pic_number      = $using:default_pic_number
    $script:default_max_image_size  = $using:default_max_image_size
    $script:default_min_image_size  = $using:default_min_image_size
    $script:canvas_width            = $using:canvas_width
    $script:canvas_height           = $using:canvas_height
    $script:overlap                 = $using:overlap
    $script:status_image            = $using:status_image
    $script:status_video            = $using:status_video
    $script:off_screen_pixels       = $using:off_screen_pixels
     
    $script:background_color        = $using:background_color
        #Random

    $script:border_color            = $using:border_color
        #Random
        #Random All

    $script:border_width           = $using:border_width
        #Random
        #Random All


    $script:background_color = $script:background_color -replace "#","0x"
    $script:border_color = $script:border_color -replace "#","0x" 
    #################################################################################
    ######Child Variables############################################################
    $script:Shell = New-Object -ComObject Shell.Application
    $script:save_timer = Get-Date
    $script:restricted_zones = @{};
    $script:previous_files = @()
    

    $script:watch = 2;
        #1 = Watch FFmpeg Process
        #2 = Silent
    #################################################################################
    ######Var Print #################################################################
    Write-Output "----------------------------------------------------------------------------------------------------------"
    Write-Output "Script Dir:    $script:dir"
    Write-Output "FFmpeg:        $script:ffmpeg"
    Write-Output "FFprobe:       $script:ffprobe"
    Write-Output "Collage Type:  $script:collage_type"
    Write-Output "Max Seconds:   $script:max_time_seconds"
    Write-Output "Input Path:    $script:default_input_path"
    Write-Output "Output Path:   $script:default_output_path"
    Write-Output "Pic Number:    $script:default_pic_number"
    Write-Output "Max Size:      $script:default_max_image_size"
    Write-Output "Min Size:      $script:default_min_image_size"
    Write-Output "Canvas Width:  $script:canvas_width"
    Write-Output "Canvas Height: $script:canvas_height"
    Write-Output "Overlap:       $script:overlap"
    Write-Output "BackGrd Color: $script:background_color"
    Write-Output "Border Color:  $script:border_color"
    Write-Output "Border Width:  $script:border_width"
    Write-Output "Status Image:  $script:status_image"
    Write-Output "Status Video:  $script:status_video"
    Write-Output "Off Screen:    $script:off_screen_pixels"
    Write-Output "----------------------------------------------------------------------------------------------------------"
    Write-Output " "


    #################################################################################
    ######Flush Buffers #############################################################
    check_exit
    if(Test-Path -LiteralPath "$script:dir\Buffer")
    {
        $buffer_files = Get-ChildItem "$script:dir\Buffer"
        foreach($file in $buffer_files)
        {
            Remove-Item -LiteralPath $file.FullName
        }
    }


    #################################################################################
    ######Set Output File Names #####################################################
    $super_name = Split-Path $script:default_input_path -Leaf
    $date = (Get-Date -UFormat %Y%m%d_%H%M%S)
    $script:output_mp4 = $script:dir + "\Buffer\"          + "$super_name" + "_Collage_" + $script:canvas_width +"x" + $script:canvas_height + "_" + $date + ".mp4"
    $script:buffer_mp4 = $script:dir + "\Buffer\"          + "$super_name" + "_Collage_" + $script:canvas_width +"x" + $script:canvas_height + "_" + $date + "_buffer.mp4"
    $script:final_file = "$super_name" + "_Collage_" + $script:canvas_width +"x" + $script:canvas_height + "_" + $date

    write-Output "END-$script:final_file"

    #################################################################################
    ######Gather Files ##############################################################
    check_exit
    Write-Output "STAT-Gathering Files..."
    $files = @();
    if($script:collage_type -eq "Image & Video")
    {
        #Image & Video
        $script:files = Get-ChildItem -literalpath $script:default_input_path -Recurse | where {$_.extension -in ".png",".jpg",".jpeg",".bmp",".tiff",".gif",".webp",".avi",".mp4",".mpeg",".mkv",".rm",".mpg",".m4v",".flv",".wmv",".ogm",".m2ts",".vob",".ts"}
    }
    elseif($script:collage_type -eq "Video")
    {
        #Video
        $script:files = Get-ChildItem -literalpath $script:default_input_path -Recurse | where {$_.extension -in ".gif",".avi",".mp4",".mpeg",".mkv",".rm",".mpg",".m4v",".flv",".wmv",".ogm",".m2ts",".vob",".ts"}
    }
    else
    {
        #Images
        $script:files = Get-ChildItem -literalpath $script:default_input_path -Recurse | where {$_.extension -in ".png",".jpg",".jpeg",".bmp",".tiff",".gif",".webp"}
    }
    check_exit
    if($files.count -eq 0)
    {
        Write-Output "STAT-FAILED: No Valid Files"
        Write-Output "FAILED: NO VALID FILES!"
        sleep 1
        exit;
    }


    #################################################################################
    ######Build Default Canvas ######################################################
    $backcolor = $script:background_color
    if($script:background_color -eq "Random")
    {
        $red   = (Get-Random -minimum 0 -maximum 255).ToString("X").PadLeft(2,"0")
        $green = (Get-Random -minimum 0 -maximum 255).ToString("X").PadLeft(2,"0")
        $blue  = (Get-Random -minimum 0 -maximum 255).ToString("X").PadLeft(2,"0")
        $backcolor = "0x$red$green$blue"
    }
    $cmd = "color=" + "$backcolor" + ":size=" + "$script:canvas_width" + "x" + "$script:canvas_height" + ":duration=" + "$script:max_time_seconds"
    $console = & cmd /u /c  "$script:ffmpeg -f lavfi -i $cmd -c:v libx264 -x264-params `"nal-hrd=cbr`" -b:v 81920k -minrate 81920k -maxrate 81920k -bufsize 81920k -movflags +faststart -hide_banner $script:show `"$script:output_mp4`" -y"
    

    #################################################################################
    ######Process Files #############################################################
    $count = 0;
    while($count -lt $script:default_pic_number)
    {
        check_exit
        #################################################################################
        ######File Loop Vars ############################################################
        $count++ 

        ########################
        ####Show FFMpeg Output
        if($script:watch -eq 2)
        {
            $script:show = "-loglevel error"
        }
        else
        {
            $script:show = "";
        }

        #################################################################################
        ######Generate a Random File ####################################################
        $file_number = Get-Random -minimum 0 -maximum $script:files.Count
        $tries = 0;
        while(($file_number -in $script:previous_files) -and ($tries -lt 100))
        {
            $tries++;
            $file_number = Get-Random -minimum 0 -maximum $script:files.Count
        
        }
        if($tries -eq 100)
        {
            #write-host Threshold Was Met
            $script:previous_files = @()
        }
        $script:previous_files += $file_number
    

        #################################################################################
        ######Generate a Random File ####################################################
        check_exit
        $file_name    = Get-ChildItem -literalpath $files[$file_number].FullName
        $simple_name  = $file_name.Name
        $full_path    = $file_name.FullName
        $shell_folder = $Shell.Namespace($file_name.DirectoryName)
        $shell_file   = $shell_folder.ParseName($file_name.Name)
        $file_width   = $shell_file.ExtendedProperty("System.Video.FrameWidth")
        $file_height  = $shell_file.ExtendedProperty("System.Video.FrameHeight")
        if($file_width -eq $null)
        {
            if($full_path -match ".gif$")
            {
                [string]$execute = & cmd /u /c  "$script:ffprobe -i `"$full_path`" -show_streams -select_streams a 2>&1"
                $script:duration = $execute.Substring(($execute.IndexOf("Duration:") + 10),8)
                ([int]$hours,[int]$minutes,[int]$seconds) = $script:duration -split ":"
                [int]$script:duration_seconds = (($seconds + ($minutes * 60) + ($hours * 60 * 60)))
                [int]$script:duration_offset = [Math]::Log10($script:duration_seconds) * ([Math]::Sqrt($script:duration_seconds) * 2) * ($script:duration_threshold / 100)
                $start_point = Get-Random -Minimum 0 -Maximum $duration_seconds
            }
            else
            {
                $duration_seconds = 0;
                $start_point = 0;
            }
          
            $script:image = New-Object System.Drawing.Bitmap $file_name.FullName
            $file_width  = $image.Width
            $file_height = $image.Height
            
            
        }
        else
        {
            #write-host -ForegroundColor Cyan "Video File"
            $duration    = $shell_folder.GetDetailsOf($shell_file, 27)
            ([int]$hours,[int]$minutes,[int]$seconds) = $script:duration -split ":"
            [int]$duration_seconds = (($seconds + ($minutes * 60) + ($hours * 60 * 60)))
            $start_point = Get-Random -Minimum 0 -Maximum $duration_seconds
        }

    
        Write-Output "Working on #$count - $simple_name"
        Write-Output "STAT-Working on #$count - $simple_name"
        #write-output "File:$file_name Dur:$duration_seconds Dim:$file_width x $file_height Start:$start_point"
        if(($file_width -eq 0) -or ($file_width -eq "") -or ($file_width -eq $null) -or ($file_height -eq 0) -or ($file_height -eq "") -or ($file_height -eq $null))
        {
            write-output -ForegroundColor Red "Error: Invalid File $file_name"
            continue
        }


        #################################################################################
        ######Find a Random Canvas Opening ##############################################
        check_exit
        $tries = 0;
        $go = 0;
        while($tries -lt 1000)
        { 
            $tries++;
            $height = 0;
            $width = 0;    
            #################################################################################
            ######Determine New Object Sizes ################################################
            if($file_height -ge $file_width)
            {
                ##Randomize Size/Fix Aspect Ratio
                [int]$height = Get-Random -minimum ([int]$script:default_min_image_size) -maximum ([int]$script:default_max_image_size)
                $aspect_ratio = $file_height / $height
                [int]$width = $file_width / $aspect_ratio;
            }
            else
            {
                ##Randomize Size/Fix Aspect Ratio
                [int]$width = Get-Random -minimum ([int]$script:default_min_image_size) -maximum ([int]$script:default_max_image_size)
                $aspect_ratio = $file_width / $width
                [int]$height = $file_height / $aspect_ratio;
            }

            $pointX = (Get-Random -minimum -$script:off_screen_pixels  -maximum (([int]$script:canvas_width + $script:off_screen_pixels) - $width))
            $pointY = (Get-Random -minimum -$script:off_screen_pixels  -maximum (([int]$script:canvas_height + $script:off_screen_pixels) - $height))

            #################################################################################
            ######Calculate Object Placement #################################################
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
                    $failed++
                }
            }
                    
            if($failed -eq 0)
            {
                $go = 1;
                break
            }
        }
        

        #################################################################################
        ######Execute Placement #########################################################
        check_exit
        if($go -eq 1)
        {
            #################################################################################
            ######Create a Restricted Zone ##################################################
            [int]$zone_x1 = ($pointX + (($script:overlap / 2) * ($width / 100)))
            [int]$zone_x2 = (($pointX + $width) - (($script:overlap / 2) * ($width / 100)))
            [int]$zone_y1 = ($pointY + (($script:overlap / 2) * ($height / 100)))
            [int]$zone_y2 = (($pointY + $height) - (($script:overlap / 2) * ($height / 100)))
            if(($zone_x1 -ne $zone_x2) -or ($zone_y1 -ne $zone_y2))
            {
                #Write-Output "$zone_x1 - $zone_x2 - $zone_y1 - $zone_y2  = 1"
                if(!($restricted_zones.containskey("$zone_x1 - $zone_x2 - $zone_y1 - $zone_y2")))
                {
                    #Write-Output "$zone_x1 - $zone_x2 - $zone_y1 - $zone_y2  = 2"
                    $restricted_zones.add("$zone_x1 - $zone_x2 - $zone_y1 - $zone_y2",0)
                }
            }

            #################################################################################
            ######Place Object On Canvas ## #################################################
            check_exit
            $fc = "-filter_complex `"" + "[1]scale=" + $width +":" + $height +",setsar=1[fg];[0][fg]overlay=x=" + $pointX + ":y=" + $pointY + ",format=yuv420p;`""
            if($start_point -ne "0")
            {
                #write-output "Video"
                $console = & cmd /u /c "$script:ffmpeg -i `"$script:output_mp4`" -stream_loop -1 -ss $start_point -i `"$file_name`" -an -hide_banner $script:show $fc -c:a copy -movflags +faststart -t $script:max_time_seconds `"$script:buffer_mp4`" -y"
            }
            else
            {
                #write-output "Image"
                #write-output "$script:ffmpeg -i `"$script:output_mp4`" -stream_loop -1 -i `"$file_name`" -an -hide_banner $script:show $fc -c:a copy -movflags +faststart -t $script:max_time_seconds `"$script:buffer_mp4`" -y"
                $console = & cmd /u /c "$script:ffmpeg -i `"$script:output_mp4`" -i `"$file_name`" -an -hide_banner $script:show $fc -c:a copy -movflags +faststart -t $script:max_time_seconds `"$script:buffer_mp4`" -y"
            }
            if(Test-Path -LiteralPath "$script:buffer_mp4")
            {
                Remove-Item -LiteralPath "$script:output_mp4"
                Rename-Item -LiteralPath "$script:buffer_mp4" "$script:output_mp4"
            }


            #################################################################################
            ######Draw Object Borders #######################################################
            check_exit
            if($script:border_width -ne 0)
            {
                $bordercolor = $script:border_color
                if($script:border_color -match "Random")
                {
                    $red   = (Get-Random -minimum 0 -maximum 255).ToString("X").PadLeft(2,"0")
                    $green = (Get-Random -minimum 0 -maximum 255).ToString("X").PadLeft(2,"0")
                    $blue  = (Get-Random -minimum 0 -maximum 255).ToString("X").PadLeft(2,"0")
                    $bordercolor = "0x$red$green$blue"
                    ##Randomize Only Once
                    if($script:border_color -ne "Random All")
                    {
                        $script:border_color = "$bordercolor"
                    }
                }
                $borderwidth = $script:border_width
                if($script:border_width -match "Random")
                {
                    $borderwidth = Get-Random -minimum 0 -maximum 20

                    ##Randomize Only Once
                    if($script:border_width -ne "Random All")
                    {
                        $script:border_width = "$borderwidth"
                    }
                }
                $pointX = $pointX - 1 #Calibrate
                $pointY = $pointY - 1 #Calibrate
                $width  = $width  + 1 #Calibrate
                $height = $height + 1 #calibrate

                $box = "-filter_complex `"drawbox=" + "$pointX" + ":" + "$pointY" + ":" + "$width" + ":" + "$height" + ":" + "$bordercolor" +"@1" + ":t=" + $borderwidth + ";`""
                $console = & cmd /u /c "$script:ffmpeg -i `"$script:output_mp4`" -hide_banner $script:show $box `"$script:buffer_mp4`" -y"
                if(Test-Path -LiteralPath "$script:buffer_mp4")
                {
                   Remove-Item -LiteralPath "$script:output_mp4"
                   Rename-Item -LiteralPath "$script:buffer_mp4" "$script:output_mp4"
                }
            }
            #################################################################################
            ######Draw Object Borders #######################################################
            check_exit
            $duration = (Get-Date) - $script:save_timer
            if(($duration.TotalSeconds -gt 5) -or ($count -eq 1) -or ($count -eq $script:default_pic_number))
            {   
                #write-output "Running Update Start $duration"
                $script:save_timer = Get-Date
                $snap = $script:max_time_seconds - 1
                $console = & cmd /u /c "$script:ffmpeg -ss $snap -i `"$script:output_mp4`" -hide_banner $script:show `"$script:status_image`" -y"
                #write-output "Running Update End $duration"
            }
            #################################################################################
            ######Update Progress ###########################################################
            [int]$status = (($count / $script:default_pic_number) * 100)
            if($status -gt 100){$status = 100;}
            write-output "PB-$status"
            Copy-Item $script:output_mp4 $script:status_video -Force
            Write-Output "     Placement Successful: $tries Tries"
        }
        else
        {
            Write-Output "     Placement Failed: $tries Tries"
        }
        Write-Output " "
    }#File Count Loop
    write-Output "FIN-"
    write-Output "Finished!"
    sleep 1
    
            
}
###################################################################################################################################################################################################################################################
######Build Collage Job End #####################################################################################################################################################################################################################
##################################################################################


    ################################################################################
    #####Start Job & Display Output ################################################
    $script:job = Start-Job -ScriptBlock  $build_collage
    $script:status_counter = 0;
    $script:load_image_timer = Get-Date
    $script:first = 0;
    $script:final_file = "";
    $script:early_stop = 0;
    while(($script:job.State -eq [System.Management.Automation.JobState]::Running) -or ($script:job.ChildJobs.Output.Count -ne $script:status_counter))
    {  
        [System.Windows.Forms.Application]::DoEvents()
        if(($script:job.ChildJobs.Output.count -ge 2) -and ($script:job.ChildJobs.Output[$script:status_counter] -ne $null))
        {
            $output = $script:job.ChildJobs.Output[$script:status_counter]
            $script:status_counter++
            if(($output -ne $null) -and ($output -ne ""))
            {
                if($output -match "^PB-") #Update Progress Bar Value
                {
                    $script:progress_bar.Value = [int]$output.substring(3,[string]$output.length -3);
                    if($script:progress_bar.Value -eq 100)
                    {
                        $completeLabel.Text = "Complete"
                    } 
                }
                elseif($output -match "^STOP-") #Update Progress Bar Value
                {
                    $script:early_stop = 1;
                }
                elseif($output -match "^STAT-") #Update Progress Bar Value
                {
                    $completeLabel.Text = $output.substring(5,[string]$output.length -5);
                }
                elseif($output -match "^END-") #Final File
                {
                    $script:final_file = $output.substring(4,[string]$output.length -4);
                }
                elseif($output -match "^FIN-") #Finished
                {
                    load_status_image
                }
                else
                {
                    write-host  $output
                }
            }#Not Null                            
        }#If Gt 2
        ##################################################
        $duration = (Get-Date) - $script:load_image_timer
        if($duration.TotalSeconds -gt 3)
        {
            $script:load_image_timer = Get-Date
            [string]$returned = load_status_image
            if(Test-Path -LiteralPath $final_file)
            {
                $script:outFile = $final_file
                $buttonDisplay.Enabled = $true
            }
            else
            {
                if($script:collage_type -eq "Image")
                {
                    $script:outFile = $script:status_image
                    $buttonDisplay.Enabled = $true
                }
                else
                {
                    $script:outFile = $script:status_video
                    $buttonDisplay.Enabled = $true
                }
            }
            
        }
    }#While Status Running
    ################################################################################
    #####Finalize - Force Image GUI Update #########################################
    $returned = "";
    $start_trying = (Get-Date)
    while(($returned -ne "Success") -and ($start_trying.TotalSeconds -gt 5))
    {
        [string]$returned = load_status_image
        sleep 1
    }
    $buttonDisplay.Enabled = $true
    
    
    ################################################################################
    #####Finalize - Copy Files #####################################################
    if($script:early_stop -eq 0)
    {
        if($script:collage_type -eq "Image")
        {
            $final_file = $script:default_output_path + "\" + $final_file + ".bmp"
            Copy-Item $script:status_image $final_file -Force
            $script:outFile = $final_file
        }
        else
        {
            $final_file = $script:default_output_path + "\" + $final_file + ".mp4"
            write-host $final_file
            Copy-Item $script:status_video $final_file -Force
            $script:outFile = $final_file
        }
    }
}
################################################################################
######Load Status Image ########################################################
function load_status_image
{
    $return = ""
    if(Test-Path -Literalpath $script:status_image)
    {
        $aspect_ratio = $script:canvas_height / 550
        $wheight =  $script:canvas_height / $aspect_ratio;
        $wwidth =  $script:canvas_width / $aspect_ratio;
        if($script:first -eq 0)
        {
            $script:first = 1;
            $Form.Width = 530 + $wwidth + 20
            $Form.height = 650
            $script:progress_bar.value = 0;
            $completeLabel.Location = New-Object System.Drawing.Point ((($Form.width / 2) - ($completeLabel.width / 2)), $y_pos)
            $System_Drawing_Size = New-Object System.Drawing.Size
            $System_Drawing_Size.Width = $Form.Width - 30
            $System_Drawing_Size.Height = 20
            $script:progress_bar.Size = $System_Drawing_Size
            $script:formGraphics = $Form.createGraphics()
        }
        try
        {
            $script:bitmap = [System.Drawing.Image]::Fromfile($script:status_image);

            if($script:bitmap)
            {
                $script:formGraphics.DrawImage($script:bitmap, 530, 15, $wwidth, $wheight)
                $script:bitmap.Dispose()
                $return = "Success"
            }
            else
            {
                #write-host Failed
                $return = "Failed"
            }
        }
        catch
        {
            #write-host Failed to Update Preview
            $return = "Failed"
        }
    }
    return $return
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
    $selectForm.SelectedPath = $script:default_input_path
	$getKey = $selectForm.ShowDialog()
	If ($getKey -eq "OK") 
    {
            $script:default_input_path = $selectForm.SelectedPath
        	$textBoxIn.Text = $selectForm.SelectedPath
            save_settings
	}
}
function selectFolderOut 
{
	$selectForm = New-Object System.Windows.Forms.FolderBrowserDialog
    $selectForm.SelectedPath = $script:default_output_path
	$getKey = $selectForm.ShowDialog()
	If ($getKey -eq "OK") 
    {
            $script:default_output_path = $selectForm.SelectedPath
        	$textBoxOut.Text = $selectForm.SelectedPath
            save_settings
	}
}
################################################################################
######Prompt for File Exe#######################################################
function prompt_for_file_exe()
{  
 [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")|Out-Null

 $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
 $OpenFileDialog.initialDirectory = "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
 #$OpenFileDialog.filter = "All files (*.*)| *.*"
 $OpenFileDialog.filter = "FFMpeg (*.exe)|*.exe;"
 $OpenFileDialog.ShowDialog() | Out-Null
 $OpenFileDialog.filename
}
#################################################################################
#####Display Image###############################################################
Function DisplayImage 
{
    if(($script:outFile -gt 3) -and (Test-Path -LiteralPath $script:outFile)) 
    {
        if($script:outFile -match "Status")
        {
            $display_file = (Get-Item -LiteralPath $script:outFile)
            $display_file = "Working Display Only" + $display_file.Extension
            $display_file = "$script:dir" + "\Buffer\" + $display_file
            if(Test-Path -LiteralPath $display_file)
            {
                Remove-Item -LiteralPath $display_file
            }
            Copy-Item $script:outFile $display_file

            Invoke-Item $display_file
        }
        else
        {
            Invoke-Item $script:outFile
        }
    }
}
#################################################################################
#################################################################################
load_settings
main

