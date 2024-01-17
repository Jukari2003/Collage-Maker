Creates Desktop Collages 

Interface Example:
![alt text](https://github.com/Jukari2003/Collage-Maker/blob/main/Preview.png?raw=true)

Video Collage Output Example:
![alt text](https://github.com/Jukari2003/Collage-Maker/blob/main/Batman%20Preview.gif?raw=true)


Install Instructions:<br />
&emsp;&emsp;&emsp;- Collage-Maker.zip from GitHub<br />
&emsp;&emsp;&emsp;&emsp;&emsp;- Top Right Hand Corner Click "Code"<br />
&emsp;&emsp;&emsp;&emsp;&emsp;- Select "Download Zip"<br />
&emsp;&emsp;&emsp;- Extract Files to a desired location<br />
&emsp;&emsp;&emsp;- Right Click on BB.ps1<br />
&emsp;&emsp;&emsp;- Click "Edit"     (This should open up Collage-Maker in Powershell ISE)<br />
&emsp;&emsp;&emsp;- Once PowerShell ISE is opened. Click the Green Play Arrow.<br />
&emsp;&emsp;&emsp;- Success!<br />

--------------------------------------------------------------------------------------------------------------------------------
Settings Information:
- Collage Type
  - Determine what type of Collage you want to make Image, Video, or Both.
- FFmpeg location
  - Script requires FFmpeg to process you can download it free at: https://ffmpeg.org/download.html
- Input Images Folder
  - Directory to pull images from
- Number of Images 
  - Number of images the system will try to place, depending on your setting some images will not place becaus there is not enough screen space remaining
- Minimum Image Size
  - Determines the minimal images size for placement (Images are scaled randomly)
- Maximum Image Size
  - Determines the maximum image size for placement (Images are scaled randomly)
- Overlap
  - Determines the amount of overlap you are willing to accept 
    - 0% = No overlap
    - 100% = Overlaps everything
- Background Color
  - Determines the main background behind the images
  - "Random" will make produce a Random Background Color
- Off-Screen Pixils
  - Will determine how much images can expand off the canvas into non-visible space
- Border Color
  - Determines the Border Color around pictures
  - "Random" will produce a random border color but keep it consistent for all borders
  - "Random All" will produce a new random color for each image/video border
- Output Size
  - This will determine how big you want the Collage to be. Default is Primary monitor size
- Border Width
  - Determines how big you want the borders around pictures to be
  - "Random" will create a random border size but keep it condsisten for all borders
  - "Random All" will generate a new thickness size for each image/video 
- Duration Seconds
  - Applies to video collages only, will determine how long videos will run in seconds. It is important to note that the longer the time, the longer it will take to generate a new collage 
- Output Folder
  - This is the save location of the Collage

For best results:
  - Use a small batch of images first, and increase as desired.


