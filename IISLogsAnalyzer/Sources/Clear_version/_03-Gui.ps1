<#
Library Name:	_03-Gui.ps1
Authors:		Fabian Caignard
Date:			March 7th 2019
Version:		1.0
Description:    This Powershell library contains Code of items used by IISLogsAnalyser.ps1
                script to provide a Graphical User Interface (assemblies, custom classes
                and all Windows Forms objects and controls).
#>

#-------------------------------------------------------------------------------
#region: DPI custom Class (Thanks to Peter Hinchley at https://bit.ly/2XOizdA)
Add-Type @'
  using System;
  using System.Runtime.InteropServices;
  using System.Drawing;

  public class DPI {
    [DllImport("gdi32.dll")]
    static extern int GetDeviceCaps(IntPtr hdc, int nIndex);

    public enum DeviceCap {
      VERTRES = 10,
      DESKTOPVERTRES = 117
    }

    public static float scaling() {
      Graphics g = Graphics.FromHwnd(IntPtr.Zero);
      IntPtr desktop = g.GetHdc();
      int LogicalScreenHeight = GetDeviceCaps(desktop, (int)DeviceCap.VERTRES);
      int PhysicalScreenHeight = GetDeviceCaps(desktop, (int)DeviceCap.DESKTOPVERTRES);

      return (float)PhysicalScreenHeight / (float)LogicalScreenHeight;
    }
  }
'@ -ReferencedAssemblies 'System.Drawing.dll'
$curDPIrate = [Math]::round([DPI]::scaling(), 2) * 100
#endregion
#-------------------------------------------------------------------------------
#region: IconExtractor custom Class (Thanks to Thomas Levesque at http://bit.ly/1KmLgyN)
$icoexcode = @"
using System;
using System.Drawing;
using System.Runtime.InteropServices;

namespace System
{
	public class IconExtractor
	{

	 public static Icon Extract(string file, int number, bool largeIcon)
	 {
	  IntPtr large;
	  IntPtr small;
	  ExtractIconEx(file, number, out large, out small, 1);
	  try
	  {
	   return Icon.FromHandle(largeIcon ? large : small);
	  }
	  catch
	  {
	   return null;
	  }

	 }
	 [DllImport("Shell32.dll", EntryPoint = "ExtractIconExW", CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
	 private static extern int ExtractIconEx(string sFile, int iIndex, out IntPtr piLargeVersion, out IntPtr piSmallVersion, int amountIcons);

	}
}
"@
Add-Type -TypeDefinition $icoexcode -ReferencedAssemblies System.Drawing
#endregion
#-------------------------------------------------------------------------------
#region: Load/Use Assemblies
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms, System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
#endregion
#-------------------------------------------------------------------------------
#region: SmallImgList and LargeImgList ImageList objects
$SmallImgList = [System.Windows.Forms.ImageList]::new()
$LargeImgList = [System.Windows.Forms.ImageList]::new()
$SmallImgList.Images.Add([System.IconExtractor]::Extract("imageres.dll", 146, $false))   #  0 - Remote Computers icon
$LargeImgList.Images.Add([System.IconExtractor]::Extract("imageres.dll", 146, $true))    #  0 - Remote Computers icon
$SmallImgList.Images.Add([System.IconExtractor]::Extract("imageres.dll", 104, $false))   #  1 - Unchecked Computer icon
$LargeImgList.Images.Add([System.IconExtractor]::Extract("imageres.dll", 104, $true))    #  1 - Unchecked Computer icon
$SmallImgList.Images.Add([System.IconExtractor]::Extract("imageres.dll", 143, $false))   #  2 - Checked Computer icon
$LargeImgList.Images.Add([System.IconExtractor]::Extract("imageres.dll", 143, $true))    #  2 - Checked Computer icon
$SmallImgList.Images.Add([System.IconExtractor]::Extract("imageres.dll", 77, $false))    #  3 - Key icon
$LargeImgList.Images.Add([System.IconExtractor]::Extract("imageres.dll", 77, $true))     #  3 - Key icon
$SmallImgList.Images.Add([System.IconExtractor]::Extract("imageres.dll", 54, $false))    #  4 - Golden Lock icon
$LargeImgList.Images.Add([System.IconExtractor]::Extract("imageres.dll", 54, $true))     #  4 - Golden Lock icon
$SmallImgList.Images.Add([System.IconExtractor]::Extract("imageres.dll", 101, $false))   #  5 - Checked Users icon
$LargeImgList.Images.Add([System.IconExtractor]::Extract("imageres.dll", 101, $true))    #  5 - Checked Users icon
$SmallImgList.Images.Add([System.IconExtractor]::Extract("shell32.dll", 266, $false))    #  6 - Directory 1 icon
$LargeImgList.Images.Add([System.IconExtractor]::Extract("shell32.dll", 266, $true))     #  6 - Directory 1 icon
$SmallImgList.Images.Add([System.IconExtractor]::Extract("imageres.dll", 27, $false))    #  7 - Local Disk icon
$LargeImgList.Images.Add([System.IconExtractor]::Extract("imageres.dll", 27, $true))     #  7 - Local Disk icon
$SmallImgList.Images.Add([System.IconExtractor]::Extract("imageres.dll", 28, $false))    #  8 - Remote Disk icon
$LargeImgList.Images.Add([System.IconExtractor]::Extract("imageres.dll", 28, $true))     #  8 - Remote Disk icon
$SmallImgList.Images.Add([System.IconExtractor]::Extract("imageres.dll", 156, $false))   #  9 - Checks App Window icon
$LargeImgList.Images.Add([System.IconExtractor]::Extract("imageres.dll", 156, $true))    #  9 - Checks App Window icon
$SmallImgList.Images.Add([System.IconExtractor]::Extract("shell32.dll", 258, $false))    # 10 - Save Floppy icon
$LargeImgList.Images.Add([System.IconExtractor]::Extract("shell32.dll", 258, $true))     # 10 - Save Floppy icon
$SmallImgList.Images.Add([System.IconExtractor]::Extract("imageres.dll", 50, $false))    # 11 - Recycle Bin icon
$LargeImgList.Images.Add([System.IconExtractor]::Extract("imageres.dll", 50, $true))     # 11 - Recycle Bin icon
$SmallImgList.Images.Add([System.IconExtractor]::Extract("shell32.dll", 261, $false))    # 12 - Play App Window icon
$LargeImgList.Images.Add([System.IconExtractor]::Extract("shell32.dll", 261, $true))     # 12 - Play App Window icon
$SmallImgList.Images.Add([System.IconExtractor]::Extract("mmcndmgr.dll", 12, $false))    # 13 - Network user icon
$LargeImgList.Images.Add([System.IconExtractor]::Extract("mmcndmgr.dll", 12, $true))     # 13 - Network user icon
$SmallImgList.Images.Add([System.IconExtractor]::Extract("mmcndmgr.dll", 58, $false))    # 14 - Local user icon
$LargeImgList.Images.Add([System.IconExtractor]::Extract("mmcndmgr.dll", 58, $true))     # 14 - Local user icon
$SmallImgList.Images.Add([System.IconExtractor]::Extract("mmcndmgr.dll", 43, $false))    # 15 - Server user icon
$LargeImgList.Images.Add([System.IconExtractor]::Extract("mmcndmgr.dll", 43, $true))     # 15 - Server user icon
$SmallImgList.Images.Add([System.IconExtractor]::Extract("mmcndmgr.dll", 75, $false))    # 16 - Process engines icon
$LargeImgList.Images.Add([System.IconExtractor]::Extract("mmcndmgr.dll", 75, $true))     # 16 - Process engines icon
$SmallImgList.Images.Add([System.IconExtractor]::Extract("mmcndmgr.dll", 5, $false))     # 17 - Keys 1 icon
$LargeImgList.Images.Add([System.IconExtractor]::Extract("mmcndmgr.dll", 5, $true))      # 17 - Keys 1 icon
$SmallImgList.Images.Add([System.IconExtractor]::Extract("mmcndmgr.dll", 0, $false))     # 18 - Directory 2 icon
$LargeImgList.Images.Add([System.IconExtractor]::Extract("mmcndmgr.dll", 0, $true))      # 18 - Directory 2 icon
$SmallImgList.Images.Add([System.IconExtractor]::Extract("mmcndmgr.dll", 9, $false))     # 19 - Red cross icon
$LargeImgList.Images.Add([System.IconExtractor]::Extract("mmcndmgr.dll", 9, $true))      # 19 - Red cross icon
$SmallImgList.Images.Add([System.IconExtractor]::Extract("shell32.dll", 104, $false))    # 20 - Keys 2 icon
$LargeImgList.Images.Add([System.IconExtractor]::Extract("shell32.dll", 104, $true))     # 20 - Keys 2 icon
#endregion
#-------------------------------------------------------------------------------
#region: Get-EmbededImage function (returns a System.Drawing.Image Object from Base64 String)
function Get-EmbededImage {
    param (
        # Embedded Base64 Image string
        [Parameter(Mandatory = $true)]
        [string]$ImgBase64
    )
    $tmpStream = [System.IO.MemoryStream][System.Convert]::FromBase64String($ImgBase64)
    $DecodedImage = [System.Drawing.Bitmap][System.Drawing.Image]::FromStream($tmpStream)
    return $DecodedImage
}
#endregion
#-------------------------------------------------------------------------------
#region: Show-WelcomeForm function (displays the script Welcome splash screen and returns "OK" if next button is clicked)
function Show-WelcomeForm {
    # Creation of Welcome_Frm Form Object
    $Welcome_Frm = [System.Windows.Forms.Form]::new()
    # With Welcome_Frm...
    $Welcome_Frm | ForEach-Object {
        # Definition of Form Object's properties
        $_.Name = "Welcome_Frm"
        $_.TopMost = $true
        $_.ClientSize = '635,290'
        $_.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
        $_.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
        #$_.ControlBox = $false
        $_.MaximizeBox = $false
        $_.MinimizeBox = $false
        # Get Form Icon from SmallImgList ImageList Object
        $_.Icon = [System.Drawing.Icon]::FromHandle([System.Drawing.Bitmap]::new($SmallImgList.Images[9]).GetHicon())
        $_.Text = "{0} - Welcome" -f $MyScript.Name
        # Creation of SplashImg_PBx PictureBox Object (picture on left of the welcoming splash screen)
        $SplashImg_PBx = [System.Windows.Forms.PictureBox]::new()
        # With SplashImg_PBx...
        $SplashImg_PBx | ForEach-Object {
            # Definition of PictureBox Object's properties
            $_.Name = "SplashImg_Pbx"
            $_.Size = '150,290'
            $_.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::StretchImage
            $_.TabStop = $false
            $_.Location = '0,0'
            # Get Image from WelcomeSlpash_Img Base64 String
            $_.Image = Get-EmbededImage $WelcomeSplash_Img
        }
        # Creation of WelcomeTitle_Lbl Label Object (title of the welcoming splash screen)
        $WelcomeTitle_Lbl = [System.Windows.Forms.Label]::new()
        # With WelcomeTitle_Lbl...
        $WelcomeTitle_Lbl | ForEach-Object {
            # Definition of Label Object's properties
            $_.Name = "WelcomeTitle_Lbl"
            $_.Size = '470,30'
            $_.TabStop = $false
            $_.Location = '160,0'
            $_.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
            $_.Font = [string]$MyFonts.MainTitle
            # Get Script Name and Script Version from MyScript PSObject and add them in Title text
            $_.Text = "Welcome in {0} v{1}!" -f $MyScript.Name, $MyScript.Version
        }
        # Creation of WelcomeDescription_Lbl Label Object (text to describe script in the welcoming splash screen)
        $WelcomeDescription_Lbl = [System.Windows.Forms.Label]::new()
        # With WelcomeDescription_Lbl...
        $WelcomeDescription_Lbl | ForEach-Object {
            # Definition of Label Object's properties
            $_.Name = "WelcomeDescription_Lbl"
            $_.Size = '470,225'
            $_.TabStop = $false
            $_.Location = '160,30'
            $_.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
            $_.Font = [string]$MyFonts.Normal
            # Get Script description from MyScript PSObject and add it as Description text
            $_.Text = $MyScript.Description
        }
        # Creation of WelcomeNext_Btn Button Object (button to close the splash screen with "OK" dialog result - Set as "Accept Button" of the Form)
        $WelcomeNext_Btn = [System.Windows.Forms.Button]::new()
        # With WelcomeNext_Btn...
        $WelcomeNext_Btn | ForEach-Object {
            # Definition of Button Object's properties
            $_.Name = "WelcomeNext_Btn"
            $_.Size = '80,30'
            $_.TabStop = $true
            $_.Location = '550,255'
            $_.TabIndex = 1
            $_.UseVisualStyleBackColor = $true
            $_.Font = [string]$MyFonts.Normal
            $_.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
            #// Tips & Tricks: Adding a "&" in the Text property String allow use keyboard shortcut by pressing Alt+[Character after &].
            $_.Text = "&Next >"
            $_.DialogResult = "OK"
            # Adding of Click event handler statement
            $_.Add_Click( { $($this.FindForm()).Close() })
        }
        # Creation of WelcomeEsc_Btn Button Object (button to close the splash screen with "Cancel" dialog result - Set as "Cancel Button" of the Form)
        $WelcomeEsc_Btn = [System.Windows.Forms.Button]::new()
        # With WelcomeEsc_Btn...
        $WelcomeEsc_Btn | ForEach-Object {
            # Definition of Button Object's properties
            $_.Name = "WelcomeEsc_Btn"
            $_.Size = '80,30'
            $_.TabStop = $false
            # Button is placed under splash screen PictureBox to be hidden but usable
            $_.Location = '0,255'
            $_.TabIndex = 2
            $_.UseVisualStyleBackColor = $true
            $_.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
            $_.Text = "Cancel"
            $_.DialogResult = "Cancel"
            # Adding of Click envent handler statement
            $_.Add_Click( { $($this.FindForm()).Close() })
        }
        # Adding Controls to Form Object
        $_.Controls.AddRange(@($SplashImg_PBx, $WelcomeTitle_Lbl, $WelcomeDescription_Lbl, $WelcomeNext_Btn, $WelcomeEsc_Btn))
        # Definition of Next and Cancel buttons as, respectively, AcceptButton and CancelButton of Form Object
        $_.AcceptButton = $WelcomeNext_Btn
        $_.CancelButton = $WelcomeEsc_Btn
        # Displaying the Form
        $_.ShowDialog()
    } # End With Welcome_Frm
} # End of Show-Welcome function
#endregion
#-------------------------------------------------------------------------------
#region: Show-MainForm function (displays the script's main Form)
function Show-MainForm {
    # Creation of Main_Frm Form Object
    $Main_Frm = [System.Windows.Forms.Form]::new()
    # With Main_Frm Form Object...
    $Main_Frm | ForEach-Object {
        # Form Object properties definition
        $_.Name = "Main_Frm"
        $_.Text = "{0} - Target selection" -f $MyScript.Name
        $_.ClientSize = '1000,600'
        $_.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
        $_.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
        #$_.ControlBox = $false
        $_.MaximizeBox = $false
        $_.MinimizeBox = $false
        # Get Form Icon from SmallImgList ImageList Object
        $_.Icon = [System.Drawing.Icon]::FromHandle([System.Drawing.Bitmap]::new($SmallImgList.Images[9]).GetHicon())
        # Creation of HPanels_Spl SplitContainer Object
        $HPanels_Spl = [System.Windows.Forms.SplitContainer]::new()
        # With HPanels_Spl SplitContainer Object...
        $HPanels_Spl | ForEach-Object {
            # SplitContainer properties definition
            $_.Name = "HPanels_Spl"
            $_.Location = '0,0'
            $_.Size = '1000,600'
            $_.Orientation = "Horizontal"
            $_.Dock = "None"
            $_.BorderStyle = "None"
            $_.IsSplitterFixed = $false
            $_.SplitterDistance = 175
            $_.SplitterWidth = 1
            #region: SplitterPanel Panel1
            $OnWDirs_Grp = [System.Windows.Forms.GroupBox]::new()
            # With OnWDirs_Grp Control...
            $OnWDirs_Grp | ForEach-Object {
                $_.Name = "OnWDirs_Grp"
                $_.Size = '400,160'
                $_.Location = '10,8'
                $_.Text = "Output and working directories..."
                $_.Font = [string]$MyFonts.Strong
                # Creation of OutUnique_CBx CheckBox Object
                $OutUnique_CBx = [System.Windows.Forms.CheckBox]::new()
                # With OutUnique_CBx Control...
                $OutUnique_CBx | ForEach-Object {
                    # Control properties definition
                    $_.Name = 'OutUnique_CBx'
                    $_.Text = "Returns results in a &unique file instead of in one per computer."
                    $_.Size = '380,20'
                    $_.Location = '10,20'
                    $_.Font = [string]$MyFonts.Normal
                    $_.CheckAlign = "MiddleLeft"
                    $_.UseVisualStyleBackColor = $true
                    $_.TextAlign = "MiddleLeft"
                    $_.TabStop = $true
                    $_.TabIndex = 0
                    $_.Enabled = $true
                    $_.Visible = $true
                    $_.Add_CheckedChanged( { $script:OutSingle = $this.Checked } )
                }
                # Creation of OutDirSel_Lbl Label Object
                $OutDirSel_Lbl = [System.Windows.Forms.Label]::new()
                # With OutDirSel_Lbl Control...
                $OutDirSel_Lbl | ForEach-Object {
                    # Control properties definition
                    $_.Name = 'OutDirSel_Lbl'
                    $_.Text = "&Output directory:"
                    $_.AutoSize = $false
                    $_.Size = '150,20'
                    $_.Location = '10,45'
                    $_.Font = [string]$MyFonts.Normal
                    $_.TextAlign = "MiddleLeft"
                    $_.TabStop = $true
                    $_.TabIndex = 1
                    $_.Enabled = $true
                    $_.Visible = $true
                }
                # Creation of OutDirSel_Txt TextBox Object
                $OutDirSel_Txt = [System.Windows.Forms.TextBox]::new()
                # With OutDirSel_Txt Control...
                $OutDirSel_Txt | ForEach-Object {
                    # Control properties definition
                    $_.Name = 'OutDirSel_Txt'
                    $_.Text = $script:OutputPath
                    $_.AutoSize = $false
                    $_.Location = '10,65'
                    $_.Size = '355,20'
                    $_.Font = [string]$MyFonts.Normal
                    $_.ReadOnly = $true
                    $_.TextAlign = "Left"
                    $_.TabStop = $true
                    $_.TabIndex = 2
                    $_.Enabled = $true
                    $_.Visible = $true

                }
                # Creation of OutDirSel_Btn Button Object
                $OutDirSel_Btn = [System.Windows.Forms.Button]::new()
                # With OutDirSel_Btn Control...
                $OutDirSel_Btn | ForEach-Object {
                    # Control properties definition
                    $_.Name = 'OutDirSel_Btn'
                    $_.Text = "..."
                    $_.AutoSize = $false
                    $_.Size = '27,22'
                    $_.Location = '363,64'
                    $_.Font = [string]$MyFonts.Normal
                    $_.TextAlign = "MiddleCenter"
                    $_.UseVisualStyleBackColor = $true
                    $_.TabStop = $true
                    $_.TabIndex = 3
                    $_.Enabled = $true
                    $_.Visible = $true
                    $_.Add_Click( {
                            $path = $OutDirSel_Txt.Text
                            If (!(Test-Path $path -PathType Container)) { $path = "C:\Temp" }
                            # Creation of OutDirBrowse FolderBrowserDialog Object
                            $OutDirBrowse = [System.Windows.Forms.FolderBrowserDialog]::new()
                            # With OutDirBrowse Control...
                            $OutDirBrowse | ForEach-Object {
                                # Control properties definition
                                $_.Description = "Select an Output directory"
                                $_.SelectedPath = $OutDirSel_Txt.Text
                            }
                            if ($OutDirBrowse.ShowDialog($this.FindForm()) -eq "OK") {
                                $script:OutputPath = $OutDirSel_Txt.Text = $OutDirBrowse.SelectedPath
                            }
                        } )
                }
                # Creation of UseWrkDir_CBx CheckBox Object
                $UseWrkDir_CBx = [System.Windows.Forms.CheckBox]::new()
                # With UseWrkDir_CBx Control...
                $UseWrkDir_CBx | ForEach-Object {
                    # Control properties definition
                    $_.Name = 'UseWrkDir_CBx'
                    $_.Text = "Use a different directory for &temporary files."
                    $_.Size = '380,20'
                    $_.Location = '10,95'
                    $_.Font = [string]$MyFonts.Normal
                    $_.CheckAlign = "MiddleLeft"
                    $_.UseVisualStyleBackColor = $true
                    $_.TextAlign = "MiddleLeft"
                    $_.TabStop = $true
                    $_.TabIndex = 4
                    #$_.Enabled = $true
                    $_.Enabled = $false
                    $_.Visible = $true
                    # Adding a onCheckedChanged event handler to UseWrkDir_CBx control.
                    $_.Add_CheckedChanged( {
                            if ($this.Checked) {
                                $WrkDirSel_Txt.Enabled = $true
                                $WrkDirSel_Btn.Enabled = $true
                                $script:WorkDir = $WrkDirSel_Txt.Text
                            }
                            else {
                                $WrkDirSel_Txt.Enabled = $false
                                $WrkDirSel_Btn.Enabled = $false
                                $script:WorkDir = $script:OutputPath
                            }
                            Test-OutDirSettings
                        } )
                }
                # Creation of WrkDirSel_Lbl Label Object
                $WrkDirSel_Lbl = [System.Windows.Forms.Label]::new()
                # With WrkDirSel_Lbl Control...
                $WrkDirSel_Lbl | ForEach-Object {
                    # Control properties definition
                    $_.Name = 'WrkDirSel_Lbl'
                    $_.Text = "(by default temporary files are written in the output directory)"
                    $_.AutoSize = $false
                    $_.Size = '360,20'
                    $_.Location = '30,110'
                    $_.Font = [string][GuiFont]::ToItalic($MyFonts.Normal)
                    $_.TextAlign = "MiddleLeft"
                    $_.TabStop = $true
                    $_.TabIndex = 5
                    #$_.Enabled = $true
                    $_.Enabled = $false
                    $_.Visible = $true
                }
                # Creation of WrkDirSel_Txt TextBox Object
                $WrkDirSel_Txt = [System.Windows.Forms.TextBox]::new()
                # With WrkDirSel_Txt Control...
                $WrkDirSel_Txt | ForEach-Object {
                    # Control properties definition
                    $_.Name = 'WrkDirSel_Txt'
                    $_.Text = $script:OutputPath
                    $_.AutoSize = $false
                    $_.Location = '10,130'
                    $_.Size = '355,20'
                    $_.Font = [string]$MyFonts.Normal
                    $_.BorderStyle = "Fixed3D"
                    $_.ReadOnly = $true
                    $_.Multiline = $false
                    $_.TextAlign = "Left"
                    $_.TabStop = $true
                    $_.TabIndex = 6
                    $_.Enabled = $false
                    $_.Visible = $true
                }
                # Creation of WrkDirSel_Btn Button Object
                $WrkDirSel_Btn = [System.Windows.Forms.Button]::new()
                # With WrkDirSel_Btn Control...
                $WrkDirSel_Btn | ForEach-Object {
                    # Control properties definition
                    $_.Name = 'WrkDirSel_Btn'
                    $_.Text = "..."
                    $_.AutoSize = $false
                    $_.Size = '27,22'
                    $_.Location = '363,129'
                    $_.Font = [string]$MyFonts.Normal
                    $_.TextAlign = "MiddleCenter"
                    $_.AutoEllipsis = $false
                    $_.UseVisualStyleBackColor = $true
                    $_.TabStop = $true
                    $_.TabIndex = 7
                    $_.Enabled = $false
                    $_.Visible = $true
                    # Adding Click event handler to WrkDirSel_Btn button
                    $_.Add_Click( {
                            $path = $WrkDirSel_Txt.Text
                            If (!(Test-Path $path -PathType Container)) { $path = "C:\Temp" }
                            # Creation of WrkDirBrowse FolderBrowserDialog Object
                            $WrkDirBrowse = [System.Windows.Forms.FolderBrowserDialog]::new()
                            # With WrkDirBrowse Control...
                            $WrkDirBrowse | ForEach-Object {
                                # Control properties definition
                                $_.Description = "Select a Working directory"
                                $_.SelectedPath = $WrkDirSel_Txt.Text
                            }
                            # If User selected a directory, set and store path of selected directory.
                            if ($WrkDirBrowse.ShowDialog($this.FindForm()) -eq "OK") {
                                $script:WorkDir = $WrkDirSel_Txt.Text = $WrkDirBrowse.SelectedPath
                                Test-OutDirSettings
                            }
                        } )
                }
                # Adding controls to OnWDirs_Grp GroupBox Object (group of Output n Working Directories)
                $_.Controls.AddRange(@($OutUnique_CBx, $OutDirSel_Lbl, $OutDirSel_Txt, $OutDirSel_Btn, $UseWrkDir_CBx, $WrkDirSel_Lbl, $WrkDirSel_Txt, $WrkDirSel_Btn))
            }
            # Creation of SharedCreds_Grp GroupBox Object
            $SharedCreds_Grp = [System.Windows.Forms.GroupBox]::new()
            # With SharedCreds_Grp Control...
            $SharedCreds_Grp | ForEach-Object {
                # Control properties definition
                $_.Name = 'SharedCreds_Grp'
                $_.Text = "Shared credentials..."
                $_.AutoSize = $false
                $_.Size = '260,160'
                $_.Location = '420,8'
                $_.Font = [string]$MyFonts.Strong
                $_.TabStop = $true
                $_.TabIndex = 8
                $_.Enabled = $true
                $_.Visible = $true
                # Creation of UseSharedCreds_CBx CheckBox Object
                $UseSharedCreds_CBx = [System.Windows.Forms.CheckBox]::new()
                # With UseSharedCreds_CBx Control...
                $UseSharedCreds_CBx | ForEach-Object {
                    # Control properties definition
                    $_.Name = 'UseSharedCreds_CBx'
                    $_.Text = "&Force use of common credentials for all remote computers."
                    $_.Size = '240,33'
                    $_.Location = '10,20'
                    $_.Font = [string]$MyFonts.Normal
                    $_.CheckAlign = "TopLeft"
                    $_.UseVisualStyleBackColor = $true
                    $_.TextAlign = "MiddleLeft"
                    $_.TabStop = $true
                    $_.TabIndex = 9
                    $_.Enabled = $true
                    $_.Visible = $true
                    # Adding a onCheckedChanged event handler to UseSharedCreds_CBx control
                    $_.Add_CheckedChanged( {
                            Update-SharedCredsControls
                            $script:UseSharedCreds = $this.Checked
                        })
                }
                # Creation of SharedCredsExplain_Lbl Label Object
                $SharedCredsExplain_Lbl = [System.Windows.Forms.Label]::new()
                # With SharedCredsExplain_Lbl Control...
                $SharedCredsExplain_Lbl | ForEach-Object {
                    # Control properties definition
                    $_.Name = 'SharedCredsExplain_Lbl'
                    $_.Text = "Shared credentials must be saved before add any remote computer."
                    $_.AutoSize = $false
                    $_.Size = '165,33'
                    $_.Location = '10,117'
                    $_.ForeColor = "Red"
                    $_.Font = $("{0},8,style=Italic" -f $MyFonts.DefaultFamily)
                    $_.TextAlign = "MiddleLeft"
                    $_.TabStop = $false
                    $_.TabIndex = 10
                    $_.Enabled = $true
                    $_.Visible = $true
                }
                # Creation of SharedCredsDetails_Lbl Label Object
                $SharedCredsDetails_Lbl = [System.Windows.Forms.Label]::new()
                # With SharedCredsDetails_Lbl Control...
                $SharedCredsDetails_Lbl | ForEach-Object {
                    # Control properties definition
                    $_.Name = 'SharedCredsDetails_Lbl'
                    $_.Text = ""
                    $_.AutoSize = $false
                    $_.Size = '240,50'
                    $_.Location = '10,60'
                    $_.ForeColor = "Green"
                    $_.Font = $("{0},10,style=Italic" -f $MyFonts.DefaultFamily)
                    $_.TextAlign = "MiddleCenter"
                    $_.TabStop = $false
                    $_.TabIndex = 10
                    $_.Enabled = $false
                    $_.Visible = $false
                }
                # Creation of SCUsername_Lbl Label Object
                $SCUsername_Lbl = [System.Windows.Forms.Label]::new()
                # With SCUsername_Lbl Control...
                $SCUsername_Lbl | ForEach-Object {
                    # Control properties definition
                    $_.Name = 'SCUsername_Lbl'
                    $_.Text = "Username:"
                    $_.Size = '80,20'
                    $_.Location = '10,60'
                    $_.Margin = '3,0,3,0'
                    $_.Padding = '0,0,0,0'
                    $_.BorderStyle = "None"
                    $_.BackColor = "Control"
                    $_.ForeColor = "ControlText"
                    $_.Font = [string]$MyFonts.Normal
                    $_.TextAlign = "MiddleLeft"
                    $_.AutoEllipsis = $false
                    $_.TabStop = $false
                    $_.TabIndex = 11
                    $_.Enabled = $false
                    $_.Visible = $true
                }
                # Creation of SCUsername_Txt TextBox Object
                $SCUsername_Txt = [System.Windows.Forms.TextBox]::new()
                # With SCUsername_Txt Control...
                $SCUsername_Txt | ForEach-Object {
                    # Control properties definition
                    $_.Name = 'SCUsername_Txt'
                    $_.Text = ""
                    $_.Location = '100,60'
                    $_.Size = '150,20'
                    $_.Font = [string]$MyFonts.Normal
                    $_.AcceptsReturn = $false
                    $_.AcceptsTab = $false
                    $_.ScrollBars = "None"
                    $_.WordWrap = $false
                    $_.Multiline = $false
                    $_.TextAlign = "Left"
                    $_.TabStop = $true
                    $_.TabIndex = 12
                    $_.Visible = $true
                    $_.Enabled = $false
                    # Adding a KeyDown event handler (click on Save button if Enter Key is pressed while the control has focus)
                    $_.Add_KeyDown( { if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) { $SCSave_Btn.PerformClick() } })
                }
                # Creation of SCPassword_Lbl Label Object
                $SCPassword_Lbl = [System.Windows.Forms.Label]::new()
                # With SCPassword_Lbl Control...
                $SCPassword_Lbl | ForEach-Object {
                    # Control properties definition
                    $_.Name = 'SCPassword_Lbl'
                    $_.Text = "Password:"
                    $_.Size = '80,20'
                    $_.Location = '10,90'
                    $_.Font = [string]$MyFonts.Normal
                    $_.TextAlign = "MiddleLeft"
                    $_.AutoEllipsis = $false
                    $_.TabStop = $false
                    $_.TabIndex = 13
                    $_.Enabled = $false
                    $_.Visible = $true
                }
                # Creation of SCPassword_Txt TextBox Object
                $SCPassword_Txt = [System.Windows.Forms.TextBox]::new()
                # With SCPassword_Txt Control...
                $SCPassword_Txt | ForEach-Object {
                    # Control properties definition
                    $_.Name = 'SCPassword_Txt'
                    $_.Text = ""
                    $_.Location = '100,90'
                    $_.Size = '150,20'
                    $_.Font = [string]$MyFonts.Normal
                    $_.UseSystemPasswordChar = $true
                    $_.TextAlign = "Left"
                    $_.TabStop = $true
                    $_.TabIndex = 14
                    $_.Visible = $true
                    $_.Enabled = $false
                    # Adding a KeyDown event handler (click on Save button if Enter Key is pressed while the control has focus)
                    $_.Add_KeyDown( { if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) { $SCSave_Btn.PerformClick() } })
                }
                # Creation of SCSave_Btn Button Object
                $SCSave_Btn = [System.Windows.Forms.Button]::new()
                # With SCSave_Btn Control...
                $SCSave_Btn | ForEach-Object {
                    # Control properties definition
                    $_.Name = 'SCSave_Btn'
                    $_.Text = "Sa&ve"
                    $_.Size = '75,30'
                    $_.Location = '175,120'
                    $_.Font = [string]$MyFonts.Normal
                    $_.TextAlign = "MiddleCenter"
                    $_.TextImageRelation = "ImageBeforeText"
                    $_.ImageAlign = "MiddleRight"
                    $_.ImageList = $SmallImgList
                    $_.ImageIndex = 10
                    $_.UseVisualStyleBackColor = $true
                    $_.TabStop = $true
                    $_.TabIndex = 15
                    $_.Enabled = $false
                    $_.Visible = $true
                    # Adding a Click event handler to SCSave_Btn control.
                    $_.Add_Click( {
                            $this.FindForm().Enabled = $false
                            If (!([string]::IsNullOrEmpty($SCUsername_Txt.Text))) {
                                # If a password was provided then...
                                if (!([string]::IsNullOrEmpty($SCPassword_Txt.Text))) {
                                    # Convert plaintext clear password to a secure string
                                    $secpassword = ConvertTo-SecureString -String $SCPassword_Txt.Text -AsPlainText -Force
                                    # Temporary Save Shared credentials
                                    $TmpCred = [System.Management.Automation.PSCredential]::new($SCUsername_Txt.Text, $secpassword)
                                }
                                else {
                                    # Else temporary Save Shared credentials with empty password
                                    $TmpCred = [System.Management.Automation.PSCredential]::new($SCUsername_Txt.Text, [System.Security.SecureString]::new())
                                }
                                if (Test-Credentials $TmpCred) {
                                    $this.Enabled = $false
                                    $this.Visible = $false
                                    $script:SharedCredentials = $TmpCred
                                    $SCUsername_Lbl.Visible = $false
                                    $SCUsername_Txt.Visible = $false
                                    $SCPassword_Lbl.Visible = $false
                                    $SCPassword_Txt.Visible = $false
                                    $SCClear_Btn.Enabled = $true
                                    $SCClear_Btn.Visible = $true
                                    $SharedCredsDetails_Lbl.Text = "{0}`nis currently saved as shared credential." -f $SCUsername_Txt.Text
                                    $SharedCredsDetails_Lbl.Enabled = $true
                                    $SharedCredsDetails_Lbl.Visible = $true
                                    $this.FindForm().Enabled = $true
                                }
                                else { $this.FindForm().Enabled = $true ; return }
                            }
                            else {
                                Show-Warning "The script cannot save empty Username and Password!" "Oups!"
                                $this.FindForm().Enabled = $true
                                return
                            }
                        } )
                }
                # Creation of SCClear_Btn Button Object
                $SCClear_Btn = [System.Windows.Forms.Button]::new()
                # With SCClear_Btn Control...
                $SCClear_Btn | ForEach-Object {
                    # Control properties definition
                    $_.Name = 'SCClear_Btn'
                    $_.Text = "&Clear"
                    $_.Size = '75,30'
                    $_.Location = '175,120'
                    $_.Font = [string]$MyFonts.Normal
                    $_.TextAlign = "MiddleCenter"
                    $_.TextImageRelation = "ImageBeforeText"
                    $_.ImageAlign = "MiddleRight"
                    $_.ImageList = $SmallImgList
                    $_.ImageIndex = 11
                    $_.UseVisualStyleBackColor = $true
                    $_.TabStop = $true
                    $_.TabIndex = 16
                    $_.Enabled = $false
                    $_.Visible = $false
                    # Adding a Click event handler to SCClear_Btn control.
                    $_.Add_Click( {
                            $this.Enabled = $false
                            $this.Visible = $false
                            $script:SharedCredentials = [System.Management.Automation.PSCredential]::Empty
                            $SCUsername_Lbl.Visible = $true
                            $SCUsername_Txt.Visible = $true
                            $SCUsername_Txt.Text = ""
                            $SCPassword_Lbl.Visible = $true
                            $SCPassword_Txt.Visible = $true
                            $SCPassword_Txt.Text = ""
                            $SCSave_Btn.Enabled = $true
                            $SCSave_Btn.Visible = $true
                            $SharedCredsDetails_Lbl.Text = ""
                            $SharedCredsDetails_Lbl.Enabled = $false
                            $SharedCredsDetails_Lbl.Visible = $false
                        } )
                }
                # Adding controls to SharedCreds_Grp GroupBox Object (group of Shared Credentials)
                $_.Controls.AddRange(@($UseSharedCreds_CBx, $SharedCredsExplain_Lbl, $SCUsername_Lbl, $SCUsername_Txt, $SCPassword_Lbl, $SCPassword_Txt, $SharedCredsDetails_Lbl, $SCSave_Btn, $SCClear_Btn))
            }
            # Creation of W3CFields_Grp GroupBox Object
            $W3CFields_Grp = [System.Windows.Forms.GroupBox]::new()
            # With W3CFields_Grp Control...
            $W3CFields_Grp | ForEach-Object {
                # Control properties definition
                $_.Name = 'W3CFields_Grp'
                $_.Text = "W3C logs fields selection..."
                $_.Size = '300,160'
                $_.Location = '690,8'
                $_.Font = [string]$MyFonts.Strong
                $_.TabStop = $true
                $_.TabIndex = 17
                $_.Enabled = $true
                $_.Visible = $true
                # Creation of LogFieldsExplain_Lbl Label Object
                $LogFieldsExplain_Lbl = [System.Windows.Forms.Label]::new()
                # With LogFieldsExplain_Lbl Control...
                $LogFieldsExplain_Lbl | ForEach-Object {
                    # Control properties definition
                    $_.Name = 'LogFieldsExplain_Lbl'
                    $_.Text = "If these fields are found in log files, the script will export them to its results.`n(analysis will still based on date, url, user and IP)"
                    $_.Size = '280,55'
                    $_.Location = '10,20'
                    $_.Padding = '5,5,5,5'
                    $_.BackColor = "Info"
                    $_.ForeColor = "DarkGoldenrod"
                    $_.Font = [string][GuiFont]::ToItalic($MyFonts.Normal)
                    $_.TextAlign = "MiddleLeft"
                    $_.TabStop = $false
                    $_.TabIndex = 18
                    $_.Enabled = $true
                    $_.Visible = $true
                    # Adding a Paint event handler to LogFieldsExplain_Lbl control to force "DarkGoldenrod" colored border.
                    $_.Add_Paint( { [System.Windows.Forms.ControlPaint]::DrawBorder($_.Graphics, $_.ClipRectangle, "DarkGoldenrod", "Solid") })
                }
                # Creation of LogFieldsSel_Lbl Label Object
                $LogFieldsSel_Lbl = [System.Windows.Forms.Label]::new()
                # With LogFieldsSel_Lbl Control...
                $LogFieldsSel_Lbl | ForEach-Object {
                    # Control properties definition
                    $_.Name = 'LogFieldsSel_Lbl'
                    # Get and display list of selected LogFields for LogFieldList array
                    $_.Text = $($LogFieldList.where( { $_.Selected })).W3Cname -join ', '
                    $_.AutoSize = $false
                    $_.Location = '10,80'
                    $_.Size = '280,40'
                    $_.Font = "Microsoft Sans Serif,9,style=Regular"
                    $_.BorderStyle = "Fixed3D"
                    $_.TextAlign = "MiddleCenter"
                    $_.TabStop = $false
                    $_.TabIndex = 19
                    $_.Visible = $true
                    $_.Enabled = $true
                }
                # Creation of LogFieldsSel_Btn Button Object
                $LogFieldsSel_Btn = [System.Windows.Forms.Button]::new()
                # With LogFieldsSel_Btn Control...
                $LogFieldsSel_Btn | ForEach-Object {
                    # Control properties definition
                    $_.Name = 'LogFieldsSel_Btn'
                    $_.Text = "&Select/unselect log fields..."
                    $_.AutoSize = $false
                    $_.Size = '280,25'
                    $_.Location = '10,125'
                    $_.Font = [string]$MyFonts.Normal
                    $_.TextAlign = "MiddleCenter"
                    $_.UseVisualStyleBackColor = $true
                    $_.TabStop = $true
                    $_.TabIndex = 20
                    $_.Enabled = $true
                    $_.Visible = $true
                    # Adding a Click event handler to LogFieldsSel_Btn control.
                    $_.Add_Click( {
                            # If Log Fields selection form dialog result equals to "OK", refresh selected fields list
                            if (Show-LogFieldsSelection -eq "OK") { $LogFieldsSel_Lbl.Text = $($LogFieldList.where( { $_.Selected })).W3Cname -join ', ' }
                        } )
                }
                # Adding controls to W3CFields_Grp GroupBox Object (group of W3C Log Fields selection)
                $_.Controls.AddRange(@($LogFieldsExplain_Lbl, $LogFieldsSel_Lbl, $LogFieldsSel_Btn))
            }
            # Adding controls (groupboxes only) to Panel1 of SplitContainer Object (Upper panel in the Form)
            $_.Panel1.Controls.AddRange(@($OnWDirs_Grp, $SharedCreds_Grp, $W3CFields_Grp))
            #endregion
            #region: SplitterPanel Panel2
            # Creation of WorkingArea_Pnl Panel Object
            $WorkingArea_Pnl = [System.Windows.Forms.Panel]::new()
            # With WorkingArea_Pnl Control...
            $WorkingArea_Pnl | ForEach-Object {
                # Control properties definition
                $_.Name = 'WorkingArea_Pnl'
                $_.Size = '1000,395'
                $_.Location = '0,0'
                $_.BorderStyle = "Fixed3D"
                $_.TabStop = $false
                $_.TabIndex = 21
                $_.Enabled = $true
                $_.Visible = $true
                # Creation of LogDirectories_LVw ListView Object
                $LogDirectories_LVw = [System.Windows.Forms.ListView]::new()
                # With LogDirectories_LVw Control...
                $LogDirectories_LVw | ForEach-Object {
                    # Control properties definition
                    $_.Name = 'LogDirectories_LVw'
                    $_.AutoArrange = $true
                    $_.Location = '0,0'
                    $_.Size = '700,340'
                    $_.BorderStyle = "Fixed3D"
                    $_.CheckBoxes = $true
                    $_.Font = [string]$MyFonts.Normal
                    $_.View = "Details"
                    $_.HeaderStyle = "Clickable"
                    $_.ShowGroups = $true
                    $_.GridLines = $false
                    $_.Alignment = "Top"
                    $_.AllowColumnReorder = $false
                    $_.FullRowSelect = $true
                    $_.HideSelection = $false
                    $_.HoverSelection = $false
                    $_.MultiSelect = $false
                    $_.Scrollable = $true
                    $_.Sorting = "None"
                    $_.LabelEdit = $false
                    $_.LabelWrap = $true
                    $_.TabStop = $true
                    $_.TabIndex = 22
                    $_.Enabled = $true
                    $_.Visible = $true
                    $_.LargeImageList = $LargeImgList
                    $_.SmallImageList = $SmallImgList
                    # Creation of LDLDirectory_LVG ColumnHeader object (Directory column for Log directories Listview)
                    $LDLDirectory_LVG = [System.Windows.Forms.ColumnHeader]::new()
                    # With LDLDirectory_LVG ColumnHeader object...
                    $LDLDirectory_LVG | ForEach-Object {
                        $_.Text = "Directory"
                        $_.TextAlign = 0
                        $_.Width = 350
                    }
                    # Creation of LDLLogFiles_LVG ColumnHeader object (Log files column for Log directories Listview)
                    $LDLLogFiles_LVG = [System.Windows.Forms.ColumnHeader]::new()
                    # With LDLLogFiles_LVG ColumnHeader object...
                    $LDLLogFiles_LVG | ForEach-Object {
                        $_.Text = "Log files"
                        $_.TextAlign = 2
                        $_.Width = 70
                    }
                    # Creation of LDLOldest_LVG ColumnHeader object (First file column for Log directories Listview)
                    $LDLOldest_LVG = [System.Windows.Forms.ColumnHeader]::new()
                    # With LDLOldest_LVG ColumnHeader object...
                    $LDLOldest_LVG | ForEach-Object {
                        $_.Text = "First file"
                        $_.TextAlign = 2
                        $_.Width = 70
                    }
                    # Creation of LDLLatest_LVG ColumnHeader object (Last file column for Log directories Listview)
                    $LDLLatest_LVG = [System.Windows.Forms.ColumnHeader]::new()
                    # With LDLLatest_LVG ColumnHeader object...
                    $LDLLatest_LVG | ForEach-Object {
                        $_.Text = "Last file"
                        $_.TextAlign = 2
                        $_.Width = 70
                    }
                    # Creation of LDLStarting_LVG ColumnHeader object (Starts at column for Log directories Listview)
                    $LDLStarting_LVG = [System.Windows.Forms.ColumnHeader]::new()
                    # With LDLStarting_LVG ColumnHeader object...
                    $LDLStarting_LVG | ForEach-Object {
                        $_.Text = "Starts at"
                        $_.TextAlign = 2
                        $_.Width = 70
                    }
                    # Creation of LDLEnding_LVG ColumnHeader object (Ends at column for Log directories Listview)
                    $LDLEnding_LVG = [System.Windows.Forms.ColumnHeader]::new()
                    # With LDLEnding_LVG ColumnHeader object...
                    $LDLEnding_LVG | ForEach-Object {
                        $_.Text = "Ends at"
                        $_.TextAlign = 2
                        $_.Width = 70
                    }
                    # Adding Columns to the ListView
                    $_.Columns.AddRange(@($LDLDirectory_LVG, $LDLLogFiles_LVG, $LDLOldest_LVG, $LDLLatest_LVG, $LDLStarting_LVG, $LDLEnding_LVG))
                    # Adding an ItemSelectionChanged event handler to Listview (to update details for selected directory in rightside panel)
                    $_.Add_ItemSelectionChanged( {
                            # If there is no log file for selected directory...
                            if ([int]$_.Item.Subitems[1].Text -eq 0) {
                                # Unselect and uncheck the directory in the list then return to Form
                                $_.Item.Selected = $false
                                $_.Item.Checked = $false
                            }
                            # Update details in right side panel
                            Update-SelDirDetails $this $_.Item
                        } )
                    # Adding an ItemChecked event handler to Listview (to update log directories list directories selection)
                    $_.Add_ItemChecked( {
                            if ($_.Item.Tag) {
                                # If there is no log file for selected directory...
                                if ([int]$_.Item.Subitems[1].Text -eq 0) {
                                    # Unselect and uncheck the directory in the list then return to Form
                                    $_.Item.Selected = $false
                                    $_.Item.Checked = $false
                                }
                                # Update Directory Object's Selected property to Item checked status.
                                $_.Item.Tag.Selected = $_.Item.Checked
                            }
                            # If no directory item is checked, then disable Proceed button, else enable it.
                            if ($this.CheckedItems.Count -eq 0) { $Proceed_Btn.Enabled = $false } else { $Proceed_Btn.Enabled = $true }
                            # Refresh status bar scope text
                            Update-StatusBarScope
                        } )
                    # First filling of listview with already existing log directories in Collection
                    $script:LogDirectoryList | Add-DirectoryToLV
                }
                # Creation of ManageComputers_Btn Button Object
                $ManageComputers_Btn = [System.Windows.Forms.Button]::new()
                # With ManageComputers_Btn Control...
                $ManageComputers_Btn | ForEach-Object {
                    # Control properties definition
                    $_.Name = 'ManageComputers_Btn'
                    $_.Text = "&Manage computer list..."
                    $_.Size = '180,30'
                    $_.Location = '10,350'
                    $_.Font = [string]$MyFonts.Normal
                    $_.TextAlign = "MiddleCenter"
                    $_.AutoEllipsis = $false
                    $_.TextImageRelation = "ImageBeforeText"
                    $_.ImageAlign = "MiddleRight"
                    $_.ImageList = $SmallImgList
                    $_.ImageIndex = 0
                    $_.UseVisualStyleBackColor = $true
                    $_.TabStop = $true
                    $_.TabIndex = 23
                    $_.Enabled = $true
                    $_.Visible = $true
                    # Adding a Click event handler to ManageComputers_Btn control.
                    $_.Add_Click( {
                            # Display Computer management form, and get IDs of newly added computers.
                            $newCompList = Show-ComputerManagement
                            # If there is any remote computer in the collection...
                            if ($script:ComputerList.Where( { $_.Type -EQ [ComputerType]::Remote })) {
                                # Disable cotrols of "Shared credentials panel"
                                $UseSharedCreds_CBx.Enabled = $false
                                $SCUsername_Txt.Enabled = $false
                                $SCPassword_Txt.Enabled = $false
                                $SCSave_Btn.Enabled = $false
                                $SCClear_Btn.Enabled = $false
                            }
                            else {
                                # Else (no remote computer in the collection), enable the shared credentials checkbox and update status of other controls depending of checking status of the checkbox.
                                $UseSharedCreds_CBx.Enabled = $true
                                Update-SharedCredsControls
                            }
                            # If there is at least 1 newly added computer...
                            if ($newCompList) {
                                # If "Shared credentials" settings are activated...
                                if ($script:UseSharedCreds) {
                                    # Mount PSDrives for each computer with shared credentials and store PSDrives list in a variable
                                    $NewPsDrives = $newCompList | Connect-Computer -Credential $script:SharedCredentials
                                }
                                else {
                                    # Else (not using shared credentials), mount PSDrives for each computer with current user or custom credentials and store PSDrives list in a variable
                                    $NewPsDrives = $newCompList | Connect-Computer
                                }
                                # Get and store log directories (and their log files lists) for newly added computers.
                                Get-RemoteDirectories $NewPsDrives
                            }
                            # Clear all Log Directory listview items in main from
                            $LogDirectories_LVw.SelectedItems.Clear()
                            $LogDirectories_LVw.Items.Clear()
                            # Reload log Directories form collection to Listview in main form
                            $script:LogDirectoryList | Add-DirectoryToLV
                            # Refresh status bar scope text
                            Update-StatusBarScope
                        } )
                }
                # Creation of AddDirectory_Btn Button Object
                $AddDirectory_Btn = [System.Windows.Forms.Button]::new()
                # With AddDirectory_Btn Control...
                $AddDirectory_Btn | ForEach-Object {
                    # Control properties definition
                    $_.Name = 'AddDirectory_Btn'
                    $_.Text = "Add a log &directory..."
                    $_.Size = '180,30'
                    $_.Location = '200,350'
                    $_.FlatStyle = "Standard"
                    $_.Font = [string]$MyFonts.Normal
                    $_.TextAlign = "MiddleCenter"
                    $_.AutoEllipsis = $false
                    $_.TextImageRelation = "ImageBeforeText"
                    $_.ImageAlign = "MiddleRight"
                    $_.ImageList = $SmallImgList
                    $_.ImageIndex = 18
                    $_.UseVisualStyleBackColor = $true
                    $_.TabStop = $true
                    $_.TabIndex = 24
                    $_.Enabled = $true
                    $_.Visible = $true
                    # Adding a Click event handler to AddDirectory_Btn control.
                    $_.Add_Click( {
                            # If Log directory addition form dialog result equals to "OK", refresh log directories listview
                            if (Show-AddLogDirForm -eq "OK") {
                                # Clear all Log Directory listview items in main from
                                $LogDirectories_LVw.SelectedItems.Clear()
                                $LogDirectories_LVw.Items.Clear()
                                # Reload log Directories form collection to Listview in main form
                                $script:LogDirectoryList | Add-DirectoryToLV
                                # Refresh status bar scope text
                                Update-StatusBarScope
                            }
                        })
                }
                # Creation of SelectedDirDetails_Lbl Label Object
                $SelectedDirDetails_Lbl = [System.Windows.Forms.Label]::new()
                # With SelectedDirDetails_Lbl Control...
                $SelectedDirDetails_Lbl | ForEach-Object {
                    # Control properties definition
                    $_.Name = 'SelectedDirDetails_Lbl'
                    $_.Text = "Details for..."
                    $_.Size = '280,25'
                    $_.Location = '705,5'
                    $_.Font = $("{0},11,style=Bold" -f $MyFonts.DefaultFamily)
                    $_.TextAlign = "TopLeft"
                    $_.TabStop = $false
                    $_.TabIndex = 25
                    $_.Enabled = $true
                    $_.Visible = $true
                }
                # Creation of SelectedLogDir_Lbl Label Object
                $SelectedLogDir_Lbl = [System.Windows.Forms.Label]::new()
                # With SelectedLogDir_Lbl Control...
                $SelectedLogDir_Lbl | ForEach-Object {
                    # Control properties definition
                    $_.Name = 'SelectedLogDir_Lbl'
                    $_.Text = ""
                    $_.Size = '280,20'
                    $_.Location = '710,30'
                    $_.BorderStyle = "Fixed3D"
                    $_.Font = [string]$MyFonts.Normal
                    $_.TextAlign = "TopLeft"
                    $_.AutoEllipsis = $true
                    $_.TabStop = $false
                    $_.TabIndex = 26
                    $_.Enabled = $true
                    $_.Visible = $false
                }
                # Creation of LogsAmountInDirL_Lbl Label Object
                $LogsAmountInDirL_Lbl = [System.Windows.Forms.Label]::new()
                # With LogsAmountInDirL_Lbl Control...
                $LogsAmountInDirL_Lbl | ForEach-Object {
                    # Control properties definition
                    $_.Name = 'LogsAmountInDirL_Lbl'
                    $_.Text = "Amount of found log files:"
                    $_.Size = '150,20'
                    $_.Location = '710,55'
                    $_.Font = [string]$MyFonts.Normal
                    $_.TextAlign = "TopLeft"
                    $_.TabStop = $false
                    $_.TabIndex = 27
                    $_.Enabled = $true
                    $_.Visible = $false
                }
                # Creation of LogsAmountInDirV_Lbl Label Object
                $LogsAmountInDirV_Lbl = [System.Windows.Forms.Label]::new()
                # With LogsAmountInDirV_Lbl Control...
                $LogsAmountInDirV_Lbl | ForEach-Object {
                    # Control properties definition
                    $_.Name = 'LogsAmountInDirV_Lbl'
                    $_.Text = "0"
                    $_.Size = '100,20'
                    $_.Location = '870,55'
                    $_.Font = [string]$MyFonts.Normal
                    $_.TextAlign = "TopLeft"
                    $_.TabStop = $false
                    $_.TabIndex = 28
                    $_.Enabled = $true
                    $_.Visible = $false
                }
                # Creation of SDFirstFileDateL_Lbl Label Object
                $SDFirstFileDateL_Lbl = [System.Windows.Forms.Label]::new()
                # With SDFirstFileDateL_Lbl Control...
                $SDFirstFileDateL_Lbl | ForEach-Object {
                    # Control properties definition
                    $_.Name = 'SDFirstFileDateL_Lbl'
                    $_.Text = "First log file date:"
                    $_.Size = '150,20'
                    $_.Location = '710,80'
                    $_.Font = [string]$MyFonts.Normal
                    $_.TextAlign = "TopLeft"
                    $_.TabStop = $false
                    $_.TabIndex = 29
                    $_.Enabled = $true
                    $_.Visible = $false
                }
                # Creation of SDFirstFileDateV_Lbl Label Object
                $SDFirstFileDateV_Lbl = [System.Windows.Forms.Label]::new()
                # With SDFirstFileDateV_Lbl Control...
                $SDFirstFileDateV_Lbl | ForEach-Object {
                    # Control properties definition
                    $_.Name = 'SDFirstFileDateV_Lbl'
                    $_.Text = "08/09/2020"
                    $_.Size = '110,20'
                    $_.Location = '870,80'
                    $_.Font = [string]$MyFonts.Normal
                    $_.TextAlign = "TopRight"
                    $_.TabStop = $false
                    $_.TabIndex = 30
                    $_.Enabled = $true
                    $_.Visible = $false
                }
                # Creation of SDLastFileDateL_Lbl Label Object
                $SDLastFileDateL_Lbl = [System.Windows.Forms.Label]::new()
                # With SDLastFileDateL_Lbl Control...
                $SDLastFileDateL_Lbl | ForEach-Object {
                    # Control properties definition
                    $_.Name = 'SDLastFileDateL_Lbl'
                    $_.Text = "Last log file date:"
                    $_.Size = '150,20'
                    $_.Location = '710,105'
                    $_.Font = [string]$MyFonts.Normal
                    $_.TextAlign = "TopLeft"
                    $_.TabStop = $false
                    $_.TabIndex = 31
                    $_.Enabled = $true
                    $_.Visible = $false
                }
                # Creation of SDLastFileDateV_Lbl Label Object
                $SDLastFileDateV_Lbl = [System.Windows.Forms.Label]::new()
                # With SDLastFileDateV_Lbl Control...
                $SDLastFileDateV_Lbl | ForEach-Object {
                    # Control properties definition
                    $_.Name = 'SDLastFileDateV_Lbl'
                    $_.Text = "08/09/2020"
                    $_.Size = '110,20'
                    $_.Location = '870,105'
                    $_.Font = [string]$MyFonts.Normal
                    $_.TextAlign = "TopRight"
                    $_.TabStop = $false
                    $_.TabIndex = 32
                    $_.Enabled = $true
                    $_.Visible = $false
                }
                # Creation of LogsFiltering_Grp GroupBox Object
                $LogsFiltering_Grp = [System.Windows.Forms.GroupBox]::new()
                # With LogsFiltering_Grp Control...
                $LogsFiltering_Grp | ForEach-Object {
                    # Control properties definition
                    $_.Name = 'LogsFiltering_Grp'
                    $_.Text = "Log files filtering..."
                    $_.Size = '280,85'
                    $_.Location = '710,130'
                    $_.Font = [string]$MyFonts.Normal
                    $_.TabStop = $true
                    $_.TabIndex = 33
                    $_.Enabled = $true
                    $_.Visible = $false
                    # Creation of OldestDate_Lbl Label Object
                    $OldestDate_Lbl = [System.Windows.Forms.Label]::new()
                    # With OldestDate_Lbl Control...
                    $OldestDate_Lbl | ForEach-Object {
                        # Control properties definition
                        $_.Name = 'OldestDate_Lbl'
                        $_.Text = "Oldest date:"
                        $_.Size = '70,20'
                        $_.Location = '10,20'
                        $_.Font = [string]$MyFonts.Normal
                        $_.TextAlign = "BottomLeft"
                        $_.AutoEllipsis = $false
                        $_.TabStop = $false
                        $_.TabIndex = 34
                        $_.Enabled = $true
                        $_.Visible = $true
                    }
                    # Creation of OldestDate_DTP DateTimePicker Object
                    $OldestDate_DTP = [System.Windows.Forms.DateTimePicker]::new()
                    # With OldestDate_DTP Control...
                    $OldestDate_DTP | ForEach-Object {
                        # Control properties definition
                        $_.Name = 'OldestDate_DTP'
                        $_.Size = '100,20'
                        $_.Location = '90,20'
                        $_.Font = [string]$MyFonts.Normal
                        $_.CalendarFont = "SegoeUI,9,style=Regular"
                        $_.ShowCheckBox = $false
                        $_.ShowUpDown = $false
                        $_.Format = "Short"
                        #$_.MaxDate = [[datetime]]
                        #$_.MinDate = [[datetime]]
                        $_.Value = Get-Date
                        #$_.Text = ""
                        $_.TabStop = $true
                        $_.TabIndex = 35
                        $_.Enabled = $true
                        $_.Visible = $true
                    }
                    # Creation of LatestDate_Lbl Label Object
                    $LatestDate_Lbl = [System.Windows.Forms.Label]::new()
                    # With LatestDate_Lbl Control...
                    $LatestDate_Lbl | ForEach-Object {
                        # Control properties definition
                        $_.Name = 'LatestDate_Lbl'
                        $_.Text = "Latest date:"
                        $_.Size = '70,20'
                        $_.Location = '10,50'
                        $_.Font = [string]$MyFonts.Normal
                        $_.TextAlign = "BottomLeft"
                        $_.TabStop = $false
                        $_.TabIndex = 36
                        $_.Enabled = $true
                        $_.Visible = $true
                    }
                    # Creation of LatestDate_DTP DateTimePicker Object
                    $LatestDate_DTP = [System.Windows.Forms.DateTimePicker]::new()
                    # With LatestDate_DTP Control...
                    $LatestDate_DTP | ForEach-Object {
                        # Control properties definition
                        $_.Name = 'LatestDate_DTP'
                        $_.Size = '100,20'
                        $_.Location = '90,50'
                        $_.Font = [string]$MyFonts.Normal
                        $_.CalendarFont = "SegoeUI,9,style=Regular"
                        $_.DropDownAlign = "Left"
                        $_.ShowCheckBox = $false
                        $_.ShowUpDown = $false
                        $_.Format = "Short"
                        #$_.MaxDate = [[datetime]]
                        #$_.MinDate = [[datetime]]
                        $_.Value = Get-Date
                        #$_.Text = ""
                        $_.TabStop = $true
                        $_.TabIndex = 37
                        $_.Enabled = $true
                        $_.Visible = $true
                    }
                    # Creation of ApplyFilter_Btn Button Object
                    $ApplyFilter_Btn = [System.Windows.Forms.Button]::new()
                    # With ApplyFilter_Btn Control...
                    $ApplyFilter_Btn | ForEach-Object {
                        # Control properties definition
                        $_.Name = 'ApplyFilter_Btn'
                        $_.Text = "&Apply"
                        $_.Size = '70,55'
                        $_.Location = '200,19'
                        $_.Font = [string]$MyFonts.Normal
                        $_.TextAlign = "MiddleCenter"
                        $_.UseVisualStyleBackColor = $true
                        $_.TabStop = $true
                        $_.TabIndex = 38
                        $_.Enabled = $true
                        $_.Visible = $true
                        # Adding a Click event handler to ApplyFilter_Btn button.
                        $_.Add_Click( {
                                if ($OldestDate_DTP.Value -gt $LatestDate_DTP.Value) {
                                    Show-Warning "Latest date cannot be earlier than Oldest date!" "Bad filtering dates"
                                    return
                                }
                                # Update filtering dates with dates defined in DateTimePicker controls
                                Update-FilterDates $OldestDate_DTP.Value $LatestDate_DTP.Value
                                # Refresh and update filtering fields and controls.
                                Update-SelDirFilter $LogDirectories_LVw.SelectedItems[0]
                                # Refresh status bar scope text
                                Update-StatusBarScope
                            })
                    }
                    # Creation of ClearFilter_Btn Button Object
                    $ClearFilter_Btn = [System.Windows.Forms.Button]::new()
                    # With ClearFilter_Btn Control...
                    $ClearFilter_Btn | ForEach-Object {
                        # Control properties definition
                        $_.Name = 'ClearFilter_Btn'
                        $_.Text = "Clea&r"
                        $_.Size = '70,55'
                        $_.Location = '200,19'
                        $_.Font = [string]$MyFonts.Normal
                        $_.TextAlign = "MiddleCenter"
                        $_.UseVisualStyleBackColor = $true
                        $_.TabStop = $true
                        $_.TabIndex = 39
                        $_.Enabled = $true
                        $_.Visible = $true
                        # Adding a Click event handler to ClearFilter_Btn button.
                        $_.Add_Click( {
                                # Update filtering dates with no dates (then clear them)
                                Update-FilterDates
                                # Refresh and update filtering fields and controls.
                                Update-SelDirFilter $LogDirectories_LVw.SelectedItems[0]
                                # Refresh status bar scope text
                                Update-StatusBarScope
                            })
                    }
                    # Adding controls to LogsFiltering_Grp GroupBox Object (group of Log files filtering by dates range)
                    $_.Controls.AddRange(@($OldestDate_Lbl, $OldestDate_DTP, $LatestDate_Lbl, $LatestDate_DTP, $ApplyFilter_Btn, $ClearFilter_Btn))
                }
                # Creation of DirAnalysisScope_Grp GroupBox Object
                $DirAnalysisScope_Grp = [System.Windows.Forms.GroupBox]::new()
                # With DirAnalysisScope_Grp Control...
                $DirAnalysisScope_Grp | ForEach-Object {
                    # Control properties definition
                    $_.Name = 'DirAnalysisScope_Grp'
                    $_.Text = "Scope of the analysis"
                    $_.Size = '280,120'
                    $_.Location = '710,220'
                    $_.Font = [string]$MyFonts.Normal
                    $_.TabStop = $false
                    $_.TabIndex = 40
                    $_.Enabled = $true
                    $_.Visible = $false
                    # Creation of ScopeStartL_Lbl Label Object
                    $ScopeStartL_Lbl = [System.Windows.Forms.Label]::new()
                    # With ScopeStartL_Lbl Control...
                    $ScopeStartL_Lbl | ForEach-Object {
                        # Control properties definition
                        $_.Name = 'ScopeStartL_Lbl'
                        $_.Text = "Starting date:"
                        $_.Size = '150,25'
                        $_.Location = '10,25'
                        $_.Font = [string][GuiFont]::Size($MyFonts.Normal, 11)
                        $_.TextAlign = "TopLeft"
                        $_.TabStop = $false
                        $_.TabIndex = 41
                        $_.Enabled = $true
                        $_.Visible = $true
                    }
                    # Creation of ScopeStartV_Lbl Label Object
                    $ScopeStartV_Lbl = [System.Windows.Forms.Label]::new()
                    # With ScopeStartV_Lbl Control...
                    $ScopeStartV_Lbl | ForEach-Object {
                        # Control properties definition
                        $_.Name = 'ScopeStartV_Lbl'
                        $_.Text = "30/09/2020"
                        $_.Size = '110,25'
                        $_.Location = '160,25'
                        $_.Font = [string][GuiFont]::Size($MyFonts.Normal, 11)
                        $_.TextAlign = "TopRight"
                        $_.TabStop = $false
                        $_.TabIndex = 42
                        $_.Enabled = $true
                        $_.Visible = $true
                    }
                    # Creation of ScopeEndL_Lbl Label Object
                    $ScopeEndL_Lbl = [System.Windows.Forms.Label]::new()
                    # With ScopeEndL_Lbl Control...
                    $ScopeEndL_Lbl | ForEach-Object {
                        # Control properties definition
                        $_.Name = 'ScopeEndL_Lbl'
                        $_.Text = "Ending date:"
                        $_.Size = '150,25'
                        $_.Location = '10,55'
                        $_.Font = [string][GuiFont]::Size($MyFonts.Normal, 11)
                        $_.TextAlign = "TopLeft"
                        $_.TabStop = $false
                        $_.TabIndex = 43
                        $_.Enabled = $true
                        $_.Visible = $true
                    }
                    # Creation of ScopeEndV_Lbl Label Object
                    $ScopeEndV_Lbl = [System.Windows.Forms.Label]::new()
                    # With ScopeEndV_Lbl Control...
                    $ScopeEndV_Lbl | ForEach-Object {
                        # Control properties definition
                        $_.Name = 'ScopeEndV_Lbl'
                        $_.Text = "30/09/2020"
                        $_.Size = '110,25'
                        $_.Location = '160,55'
                        $_.Font = [string][GuiFont]::Size($MyFonts.Normal, 11)
                        $_.TextAlign = "TopRight"
                        $_.TabStop = $false
                        $_.TabIndex = 44
                        $_.Enabled = $true
                        $_.Visible = $true
                    }
                    # Creation of LogsAmountInScopeL_Lbl Label Object
                    $LogsAmountInScopeL_Lbl = [System.Windows.Forms.Label]::new()
                    # With LogsAmountInScopeL_Lbl Control...
                    $LogsAmountInScopeL_Lbl | ForEach-Object {
                        # Control properties definition
                        $_.Name = 'LogsAmountInScopeL_Lbl'
                        $_.Text = "Amount of log files:"
                        $_.Size = '150,25'
                        $_.Location = '10,85'
                        $_.Font = [string][GuiFont]::Size($MyFonts.Normal, 11)
                        $_.TextAlign = "TopLeft"
                        $_.TabStop = $false
                        $_.TabIndex = 45
                        $_.Enabled = $true
                        $_.Visible = $true
                    }
                    # Creation of LogsAmountInScopeV_Lbl Label Object
                    $LogsAmountInScopeV_Lbl = [System.Windows.Forms.Label]::new()
                    # With LogsAmountInScopeV_Lbl Control...
                    $LogsAmountInScopeV_Lbl | ForEach-Object {
                        # Control properties definition
                        $_.Name = 'LogsAmountInScopeV_Lbl'
                        $_.Text = "0"
                        $_.Size = '100,25'
                        $_.Location = '160,85'
                        $_.Font = [string][GuiFont]::Size($MyFonts.Normal, 11)
                        $_.TextAlign = "TopLeft"
                        $_.TabStop = $false
                        $_.TabIndex = 46
                        $_.Enabled = $true
                        $_.Visible = $true
                    }
                    # Adding controls to DirAnalysisScope_Grp GroupBox Object (group of analysis scope details)
                    $_.Controls.AddRange(@($ScopeStartL_Lbl, $ScopeStartV_Lbl, $ScopeEndL_Lbl, $ScopeEndV_Lbl, $LogsAmountInScopeL_Lbl, $LogsAmountInScopeV_Lbl))
                }
                # Creation of Proceed_Btn Button Object
                $Proceed_Btn = [System.Windows.Forms.Button]::new()
                # With Proceed_Btn Control...
                $Proceed_Btn | ForEach-Object {
                    # Control properties definition
                    $_.Name = 'Proceed_Btn'
                    $_.Text = "&Proceed to analysis"
                    $_.Size = '310,30'
                    $_.Location = '390,350'
                    $_.Font = [string]$MyFonts.Normal
                    $_.TextAlign = "MiddleCenter"
                    $_.TextImageRelation = "ImageBeforeText"
                    $_.ImageAlign = "MiddleRight"
                    $_.ImageList = $SmallImgList
                    $_.ImageIndex = 12
                    $_.UseVisualStyleBackColor = $true
                    $_.TabStop = $true
                    $_.TabIndex = 47
                    $_.Enabled = $false
                    $_.Visible = $true
                    # Adding a Click event handler to Proceed_Btn button.
                    $_.Add_Click( {
                            # Disable most of the form's controls
                            $this.Enabled = $false
                            $ManageComputers_Btn.Enabled = $false
                            $AddDirectory_Btn.Enabled = $false
                            $HPanels_Spl.Panel1.Enabled = $false
                            # Make visible and enable the progressbar in status bar
                            $MainStatusProgress1_TSP.Visible = $true
                            $MainStatusProgress1_TSP.Enabled = $true
                            $MainStatusLabel2_TSL.Visible = $true
                            $MainStatusLabel2_TSL.Enabled = $true
                            # Update status bar text
                            $MainStatusLabel1_TSL.Text = "Starting of the analysis..."
                            Start-Sleep -Millisecond 1
                            # Set Output timestamp string
                            $script:OutTimeStamp = [datetime]::Now.ToString("yyyyMMdd-HHmm_ss")
                            # Initialize the Log Fields column mapping array:
                            # For each selceted LogField object in log fields list...
                            $script:LogFieldList.Where( { $_.Selected }) | ForEach-Object {
                                # Add a custom object with W3Cname, mandatory state and negative column ID of log field.
                                [PSCustomObject[]]$script:LogFieldsMapping += [PSCustomObject]@{Name = $_.W3Cname; Mandatory = $_.Mandatory; ColumnID = -1 }
                            }
                            # Create a filtering array of strings and add fields/columns which are selected and mandatory to it
                            [string[]]$filterarray = "Server", "Instance"
                            $filterarray += ($script:LogFieldsMapping.Where( { $_.Mandatory })).Name
                            # Update status bar text
                            $MainStatusLabel1_TSL.Text = "Analysis in progress..."
                            Start-Sleep -Millisecond 1
                            # 1. Get all log files based on selections in collections (Get-SelectedLogFiles function)
                            # 2. Extract all data lines from log files (Invoke-LogDataFromFile function - and passing it amount of log files in the scope, and progressbar and label objects from status bar)
                            # 3. Remove duplicates (Invoke-DedupData function - and passing it a buffer size and deduplication filter)
                            Get-SelectedLogFiles | Invoke-LogDataFromFile -TotalCount $script:ScopeFilesCount -ProgressBarObj $MainStatusProgress1_TSP -ProgressTextObj $MainStatusLabel1_TSL | ForEach-Object {
                                # Set and increase a counter of total amount of data lines from selected log files
                                if ($InTotalLines -eq $null) { $InTotalLines = 1 } else { $InTotalLines++ }
                                # Initialise a counter to 0 (for progression text update)
                                if ($lcounter -eq $null) { $lcounter = 0 }
                                # If a progression status bar label is provided in PorgressTextObj parameter and the counter equals to 0...
                                if ($MainStatusLabel2_TSL -and $lcounter -eq 0) {
                                    # If text length in label is less than 15, add a 'O', else reset it to empty string.
                                    if ($MainStatusLabel2_TSL.Text.Length -lt 15) { $MainStatusLabel2_TSL.Text += 'o'; Start-Sleep -Millisecond 1 } else { $MainStatusLabel2_TSL.Text = "" }
                                }
                                # If counter is less than 5, increase it, else reset it to 0
                                if ($lcounter -lt 5) { $lcounter++ } else { $lcounter = 0 }
                                # Output the received object in pipeline
                                $_
                            } | Invoke-DedupData -Buffer 20 -Filter $filterarray -GroupBy "Server", "Instance" | ForEach-Object {
                                if ($_.Count -gt 0) {
                                    # Set and increase a counter total amount of data lines there are finally in script's results
                                    if ($OutTotalLines -eq $null) { $OutTotalLines = 0 }
                                    $OutTotalLines += $_.Count
                                    # Update status bar texts
                                    $MainStatusLabel1_TSL.Text = ""
                                    $MainStatusLabel2_TSL.Text = ""
                                    Start-Sleep -Millisecond 1
                                    # If script is defined to output all its results in a unique file...
                                    if ($script:OutSingle) {
                                        # Build fullpath for output file
                                        $OutFilePath = $script:OutputPath + '\ILA-Report_' + $script:OutTimeStamp + '.txt'
                                        # Update status bar text
                                        $MainStatusLabel1_TSL.Text = 'Output results in ' + $OutFilePath
                                        Start-Sleep -Millisecond 1
                                        [string[]]$proplist = "Server", "Instance"
                                        $proplist += ($script:LogFieldsMapping).Name
                                        # Write results in the defined output file
                                        $_ | Select-Object -Property $proplist | ConvertTo-Csv -NoTypeInformation | Out-File -FilePath $OutFilePath -Append
                                        # Else...
                                    }
                                    else {
                                        # Set a scope string to Server and "Instance"
                                        $scopeStr = $_[0].Server + '-' + $_[0].Instance + '_'
                                        # Build fullpath for output file
                                        $OutFilePath = $script:OutputPath + '\ILA-Report_' + $scopeStr + $script:OutTimeStamp + '.txt'
                                        # Update status bar text
                                        $MainStatusLabel1_TSL.Text = 'Output results in ' + $OutFilePath
                                        Start-Sleep -Millisecond 1
                                        # Write results in the defined output file
                                        $_ | Select-Object -Property ($script:LogFieldsMapping).Name | ConvertTo-Csv -NoTypeInformation | Out-File -FilePath $OutFilePath -Append
                                    }

                                }
                            }
                            # Update status bar texts
                            $MainStatusLabel1_TSL.Text = "Analysis finished !"
                            $MainStatusLabel2_TSL.Text = $("{0} lines of logs found and only {1} lines output: IIS Log Analyser simplified the logs by {2} lines!" -f $InTotalLines, $OutTotalLines, $($InTotalLines - $OutTotalLines))
                            Start-Sleep -Millisecond 1
                            # Make form colsing button visible and enable it
                            $OpenOutput_Btn.Enabled = $true
                            $OpenOutput_Btn.Visible = $true
                            # Hide this button (proceed button)
                            $this.Visible = $false
                        })
                }
                # Creation of OpenOutput_Btn Button Object
                $OpenOutput_Btn = [System.Windows.Forms.Button]::new()
                # With OpenOutput_Btn Control...
                $OpenOutput_Btn | ForEach-Object {
                    # Control properties definition
                    $_.Name = 'OpenOutput_Btn'
                    $_.Text = "E&xit and open output directory"
                    $_.Size = '310,30'
                    $_.Location = '390,350'
                    $_.Font = [string]$MyFonts.Strong
                    $_.TextAlign = "MiddleCenter"
                    $_.UseVisualStyleBackColor = $true
                    $_.TabStop = $true
                    $_.TabIndex = 48
                    $_.Enabled = $false
                    $_.Visible = $false
                    # Adding a Click event handler to OpenOutput_Btn button.
                    $_.Add_Click( { Invoke-Item $script:OutputPath ; $this.FindForm().Close() })
                }
                # Adding controls to WorkingArea_Pnl Panel Object (Panel of lower part of the Form)
                $_.Controls.AddRange(@($ManageComputers_Btn, $AddDirectory_Btn, $LogDirectories_LVw, $SelectedDirDetails_Lbl, $SelectedLogDir_Lbl, $LogsAmountInDirL_Lbl, $LogsAmountInDirV_Lbl, $SDFirstFileDateL_Lbl, $SDFirstFileDateV_Lbl, $SDLastFileDateL_Lbl, $SDLastFileDateV_Lbl, $LogsFiltering_Grp, $DirAnalysisScope_Grp, $Proceed_Btn, $OpenOutput_Btn))
            }
            # Create a new StatusStrip object (in fact a status bar)
            $MainStatus_STS = [System.Windows.Forms.StatusStrip]::New()
            # With StatusStrip object...
            $MainStatus_STS | ForEach-Object {
                $_.Name = "MainStatus_STS"
                $_.Dock = "Bottom"
                $_.Font = [string]$MyFonts.Normal
                # Create MainStatusProgress1_TSP status bar progress bar
                $MainStatusProgress1_TSP = [System.Windows.Forms.ToolStripProgressBar]::new()
                # With MainStatusProgress1_TSP object...
                $MainStatusProgress1_TSP | ForEach-Object {
                    $_.Name = "MainStatusProgress1_TSP"
                    $_.Text = "..."
                    $_.Visible = $false
                    $_.Step = 1
                    $_.Value = 0
                }
                # Adding the progress bar to Status bar
                [void]$_.Items.Add($MainStatusProgress1_TSP)
                # Create MainStatusLabel1_TSL status bar label
                $MainStatusLabel1_TSL = [System.Windows.Forms.ToolStripStatusLabel]::new()
                # With MainStatusLabel1_TSL object...
                $MainStatusLabel1_TSL | ForEach-Object {
                    $_.Name = "MainStatusLabel1_TSL"
                    $_.Text = "..."
                }
                # Adding the label to status bar
                [void]$_.Items.Add($MainStatusLabel1_TSL)
                # Create MainStatusLabel2_TSL status bar label
                $MainStatusLabel2_TSL = [System.Windows.Forms.ToolStripStatusLabel]::new()
                # With MainStatusLabel2_TSL object...
                $MainStatusLabel2_TSL | ForEach-Object {
                    $_.Name = "MainStatusLabel2_TSL"
                    $_.Visible = $false
                    $_.Text = ""
                    $_.Font = [string]$MyFonts.Strong
                }
                # Adding the label to status bar
                [void]$_.Items.Add($MainStatusLabel2_TSL)
            }
            # Adding controls (1 Panel and the status bar only) to Panel2 of SplitContainer Object (Lower panel in the Form)
            $_.Panel2.Controls.AddRange(@($WorkingArea_Pnl, $MainStatus_STS))
            #endregion
        }
        # Adding Controls to Form Object (only the Splitcontainer)
        $_.Controls.AddRange(@($HPanels_Spl))
        # Creation of MainToolTip_TTp ToolTip Object
        $MainToolTip_TTp = [System.Windows.Forms.ToolTip]::new()
        # With MainToolTip_TTp Object...
        $MainToolTip_TTp | ForEach-Object {
            $_.isBalloon = $true
            $_.AutoPopDelay = 5000
            $_.InitialDelay = 500
            $_.ReshowDelay = 100
            $_.ToolTipIcon = "Info"
            # Set help messages to display in tooltips for each control of the form.
            $_.SetToolTip($OnWDirs_Grp, "Define here location for output files, the working directory location and output files options.`nBy default:`n`t- temporary files and output files are saved in the script directory`n`t- the script create 1 output file per computer")
            $_.SetToolTip($OutUnique_CBx, "Allow the script to concatenate all of its results in 1 unique file instead of 1 file per computer.")
            $OutDirSel_TTT = "Output directory: location for output files. The default location is the script directory.`nBy default temporary files are also temporary saved here."
            $_.SetToolTip($OutDirSel_Lbl, $OutDirSel_TTT)
            $_.SetToolTip($OutDirSel_Txt, $OutDirSel_TTT)
            $_.SetToolTip($OutDirSel_Btn, "Open the Output directory selection dialog.")
            $_.SetToolTip($UseWrkDir_CBx, "Allow script to temporary save its temporary files in a different directory than the Output directory.`nUsefull when Output directory is set to a location which should not get any temporary files for any automation reason.")
            $_.SetToolTip($WrkDirSel_Txt, "Working directory: location for temporary files. Ths script needs to write data during the analysis process in files. These files are then deleted if analysis process succeed.")
            $_.SetToolTip($WrkDirSel_Btn, "Open the Working directory selection dialog.")
            $_.SetToolTip($SharedCreds_Grp, "Allow you to define and use common credentials to connect on all remote computers.`nScript will always run under current user context for Local computer.`n`nTo be able to use ""Shared credentials"", they must be defined and saved before adding any remote computer!")
            $_.SetToolTip($UseSharedCreds_CBx, "Activate the usage of common credentials for all remote computers.`n`nCredentials must then be saved to be effective: if credentials are not saved and this option is checked, then script will use the current user context for all computers (even if specific credentials are provided for each remote computer)!")
            $SCUsername_TTT = "Username: user login name to use as ""Shared credentials"".`nDomain name must be provided for domain users. Accepted formats are:`nMyDomain\JohnDoe or JohnDoe@Mydomain.local"
            $_.SetToolTip($SCUsername_Lbl, $SCUsername_TTT)
            $_.SetToolTip($SCUsername_Txt, $SCUsername_TTT)
            $SCPassword_TTT = "Password: password string for user login name provided in ""Username""."
            $_.SetToolTip($SCPassword_Lbl, $SCPassword_TTT)
            $_.SetToolTip($SCPassword_Txt, $SCPassword_TTT)
            $_.SetToolTip($SCSave_Btn, "Check the provided credentials, then save them to be used by the script as common credentials for all remote computers.`n`n""Shared credentials"" must be saved before any remote computer is added to be effective!")
            $_.SetToolTip($SCClear_Btn, "Clear saved ""Shared credentials"".`nIf no credential is saved and ""Force use of common credentials..."" option is checked, then script will use current user context to connect on remote computers.")
            $_.SetToolTip($W3CFields_Grp, "W3C format log fields to export in script's results from log files.`nIt is possible here to select or unselect fields from a standard list of W3C IIS log fields. Some of them are considered by script as madatory and cannot be unselected.")
            $_.SetToolTip($LogFieldsSel_Lbl, "Currently selected W3C log fields. If some of them are not found in log files, script will export fields with empty value.")
            $_.SetToolTip($LogFieldsSel_Btn, "Open the W3C log field selection dialog.")
            $_.SetToolTip($ManageComputers_Btn, "Open the dialog of list of remote computers. Click here to add or remove remote computers and define specific credentials if neccessary.`nIf you want to use common credentials for all remote computers, please define and save ""Shared credentials"" first.")
            $_.SetToolTip($AddDirectory_Btn, "Open a custom folder selection dialog to add more log directories in the list.`n`nBy default, the script is looking for default IIS logs locations on computers ""C:"" drives (mounted by the script for each remote computer). Default locations are ""W3SVC*"" and ""*FTPSVC*"" directories in ""C:\inetpub\logs\LogFiles"" or ""C:\Windows\system32\LogFiles"".`nIf no log directory is found, you can specifie their location here.`nFor remote computer, the script is only allow to browse the ""C:"" drive.")
            $_.SetToolTip($SelectedDirDetails_Lbl, "Displays bellow details about the selected directory in the Leftside list.")
            $_.SetToolTip($LogsFiltering_Grp, "Allow to limit script analysis on selected directory to a specific dates range.")
            $OldestDate_TTT = "Specify the first log file date to include in the script analysis.`nThe date here cannot be earlier than date of first found log file, and later than last found log file and ""Latest date"" date."
            $_.SetToolTip($OldestDate_Lbl, $OldestDate_TTT)
            $_.SetToolTip($OldestDate_DTP, $OldestDate_TTT)
            $LatestDate_TTT = "Specify the last log file date to include in the script analysis.`nThe date here cannot be earlier than date of first found log file and ""Oldest date"" date, and later than last found log file."
            $_.SetToolTip($LatestDate_Lbl, $LatestDate_TTT)
            $_.SetToolTip($LatestDate_DTP, $LatestDate_TTT)
            $_.SetToolTip($ApplyFilter_Btn, "Apply and save filtering settings for selected directory.")
            $_.SetToolTip($ClearFilter_Btn, "Clear filtering settings for selected directory, then script will analyze all found log files in this directory.")
            $_.SetToolTip($Proceed_Btn, "Run the analysis of all log files in selected directories in the list (according to defined filtering settings).")
        }
        # Refresh status bar scope text
        Update-StatusBarScope
        # Displaying the Form
        [void]$_.ShowDialog()
    }
}
#endregion
#-------------------------------------------------------------------------------
#region: Show-LogFieldsSelection function (display the log fileds selection dialog)
function Show-LogFieldsSelection {
    # Creation of LogFieldsSel_Frm Form Object
    $LogFieldsSel_Frm = [System.Windows.Forms.Form]::new()
    # With LogFieldsSel_Frm Form Object...
    $LogFieldsSel_Frm | ForEach-Object {
        # Form Object properties definition
        $_.Name = "LogFieldsSel_Frm"
        $_.Text = "{0} - Log fields selection" -f $MyScript.Name
        $_.ClientSize = '590,280'
        $_.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
        $_.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
        #$_.ControlBox = $false
        $_.MaximizeBox = $false
        $_.MinimizeBox = $false
        # Get Form Icon from SmallImgList ImageList Object
        $_.Icon = [System.Drawing.Icon]::FromHandle([System.Drawing.Bitmap]::new($SmallImgList.Images[9]).GetHicon())
        # Creation of LogFieldsList_DGV DataGridView Object
        $LogFieldsList_DGV = [System.Windows.Forms.DataGridView]::new()
        # With LogFieldsList_DGV Control...
        $LogFieldsList_DGV | ForEach-Object {
            # Control properties definition
            $_.Name = 'LogFieldsList_DGV'
            $_.AutoSize = $false
            $_.Size = '570,220'
            $_.Location = '10,10'
            $_.Font = [string]$MyFonts.Normal
            $_.AutoSizeColumnsMode = "None"
            $_.AutoSizeRowsMode = 6
            $_.RowHeadersVisible = $false
            $_.ColumnHeadersVisible = $true
            $_.AllowUserToAddRows = $false
            $_.AllowUserToDeleteRows = $false
            $_.SelectionMode = "FullRowSelect"
            $_.MultiSelect = $false
            $_.AllowUserToOrderColumns = $false
            $_.EnableHeadersVisualStyles = $true
            $_.AllowUserToResizeColumns = $false
            $_.RowsDefaultCellStyle.BackColor = "White"
            $_.AlternatingRowsDefaultCellStyle.BackColor = "GradientInactiveCaption"
            $_.TabStop = $true
            $_.TabIndex = 0
            $_.Enabled = $true
            $_.Visible = $true
            #region: Adding DataGridView columns...
            $tmpColObj = [System.Windows.Forms.DataGridViewCheckBoxColumn]::New()
            $tmpColObj | ForEach-Object {
                $_.Name = "fieldMandatory"
                $_.FalseValue = $false
                $_.TrueValue = $true
                $_.Resizable = "False"
                $_.Visible = $false
                $_.ReadOnly = $true
            }
            [void]$_.Columns.Add($tmpColObj)
            $tmpColObj = [System.Windows.Forms.DataGridViewCheckBoxColumn]::New()
            $tmpColObj | ForEach-Object {
                $_.HeaderText = ""
                $_.ToolTipText = "List of standard W3C log fields for Microsoft IIS.`nSelect here fields that the script will have to export in its results."
                $_.Name = "fieldSelection"
                $_.FalseValue = $false
                $_.TrueValue = $true
                $_.Resizable = "False"
                $_.Width = 30
                $_.Visible = $true
                $_.ReadOnly = $false
            }
            [void]$_.Columns.Add($tmpColObj)
            $tmpColObj = [System.Windows.Forms.DataGridViewTextBoxColumn]::New()
            $tmpColObj | ForEach-Object {
                $_.HeaderText = ""
                $_.Name = "fieldMandatoryIf"
                $_.Resizable = "False"
                $_.Visible = $false
                $_.ReadOnly = $true
            }
            [void]$_.Columns.Add($tmpColObj)
            $tmpColObj = [System.Windows.Forms.DataGridViewTextBoxColumn]::New()
            $tmpColObj | ForEach-Object {
                $_.HeaderText = "Name"
                $_.Name = "fieldName"
                $_.Resizable = "True"
                $_.Width = 195
                $_.DefaultCellStyle.WrapMode = "True"
                $_.Visible = $true
                $_.ReadOnly = $true
            }
            [void]$_.Columns.Add($tmpColObj)
            $tmpColObj = [System.Windows.Forms.DataGridViewTextBoxColumn]::New()
            $tmpColObj | ForEach-Object {
                $_.HeaderText = "Description"
                $_.Name = "fieldDescription"
                $_.Resizable = "True"
                $_.Width = 325
                $_.DefaultCellStyle.WrapMode = "True"
                $_.Visible = $true
                $_.ReadOnly = $true
            }
            [void]$_.Columns.Add($tmpColObj)
            #endregion
            #region: Adding DataGridView lines...
            # For each LogField object in List...
            $LogFieldList | ForEach-Object {
                # Set string of field name to display
                $tmpNameStr = "{0} ({1})" -f $_.DisplayName, $_.W3Cname
                # Set array of data for gridview row
                $tmpRow = @($_.Mandatory, $_.Selected, $_.FMD, $tmpNameStr, $_.Description)
                # Add new row to gridview, get its index, then get back row
                $tmpRow = $LogFieldsList_DGV.Rows[$LogFieldsList_DGV.Rows.Add($tmpRow)]
                # If corresponding field is set to mandatory...
                if ($_.Mandatory) {
                    # If there is no mandatory dependency field...
                    if ([string]::IsNullOrEmpty($_.FMD)) {
                        # Make row read only
                        $tmpRow.ReadOnly = $true
                        # Change row text and background colors to gray colors
                        $tmpRow.DefaultCellStyle.BackColor = "LightGray"
                        $tmpRow.DefaultCellStyle.ForeColor = "Gray"
                        # Else (there is dependency field), only set row text in gray color
                    }
                    else { $tmpRow.DefaultCellStyle.ForeColor = "Gray" }
                }
            }
            #endregion
            # Adding CellContentClick event handler to DataGridView (to immediatelly commit change on click in a Cell)
            $_.Add_CellContentClick( { $this.CommitEdit(512) })
            # Adding CellValueChanged event handler to DataGridView (to check mandatory dependencies and ensure that at least 1 of the dependents fields is checked)
            $_.Add_CellValueChanged( {
                    # Get current changed row
                    $tmpRow = $this.Rows[$_.RowIndex]
                    # If the current Field from the row is mandatory, not selected (checked in fact) and have any field dependency
                    if ($tmpRow.Cells[0].Value -and $tmpRow.Cells[2].Value -ne "" -and !$tmpRow.Cells[1].Value) {
                        # Build search string with current W3C field name
                        $searchTxt = "*({0})" -f $tmpRow.Cells[2].Value
                        # Search dependent field in DataGridView, then select (check) it.
                        $this.Rows.Where( { $_.Cells[3].Value -like $searchTxt }).Cells[1].Value = $true
                    }
                })
            $_.Add_SelectionChanged( {
                    # If currently selected field is mandatory and does not have dependency field...
                    $this.SelectedRows.Where( { $_.Cells[0].Value -and $_.Cells[2].Value -eq "" }) | ForEach-Object {
                        # Select (give focus) next filed in the list.
                        $LogFieldsList_DGV.Rows[$($_.Index + 1)].Selected = $true
                    }
                })
        }
        # Creation of Cancel_Btn Button Object
        $Cancel_Btn = [System.Windows.Forms.Button]::new()
        # With Cancel_Btn Control...
        $Cancel_Btn | ForEach-Object {
            # Control properties definition
            $_.Name = 'Cancel_Btn'
            $_.Text = "Cancel"
            $_.AutoSize = $false
            $_.Size = '70,30'
            $_.Location = '10,240'
            $_.Font = [string]$MyFonts.Normal
            $_.TextAlign = "MiddleCenter"
            $_.DialogResult = "Cancel"
            $_.UseVisualStyleBackColor = $true
            $_.TabStop = $true
            $_.TabIndex = 3
            $_.Enabled = $true
            $_.Visible = $true
        }
        # Creation of Reset_Btn Button Object
        $Reset_Btn = [System.Windows.Forms.Button]::new()
        # With Reset_Btn Control...
        $Reset_Btn | ForEach-Object {
            # Control properties definition
            $_.Name = 'Reset_Btn'
            $_.Text = "Reset to default"
            $_.AutoSize = $false
            $_.Size = '100,30'
            $_.Location = '245,240'
            $_.Font = [string]$MyFonts.Normal
            $_.TextAlign = "MiddleCenter"
            $_.UseVisualStyleBackColor = $true
            $_.TabStop = $true
            $_.TabIndex = 2
            $_.Enabled = $true
            $_.Visible = $true
            # Adding Click event handler to button
            $_.Add_Click( {
                    $LogFieldsList_DGV | ForEach-Object {
                        $_.Rows[0].Cells[1].Value = $true
                        $_.Rows[1].Cells[1].Value = $false
                        $_.Rows[2].Cells[1].Value = $false
                        $_.Rows[3].Cells[1].Value = $false
                        $_.Rows[4].Cells[1].Value = $false
                        $_.Rows[5].Cells[1].Value = $false
                        $_.Rows[6].Cells[1].Value = $true
                        $_.Rows[7].Cells[1].Value = $false
                        $_.Rows[8].Cells[1].Value = $false
                        $_.Rows[9].Cells[1].Value = $true
                        $_.Rows[10].Cells[1].Value = $true
                        $_.Rows[11].Cells[1].Value = $false
                        $_.Rows[12].Cells[1].Value = $false
                        $_.Rows[13].Cells[1].Value = $false
                        $_.Rows[14].Cells[1].Value = $false
                        $_.Rows[15].Cells[1].Value = $false
                        $_.Rows[16].Cells[1].Value = $false
                        $_.Rows[17].Cells[1].Value = $false
                        $_.Rows[18].Cells[1].Value = $false
                        $_.Rows[19].Cells[1].Value = $false
                        $_.Rows[20].Cells[1].Value = $false
                        $_.Rows[21].Cells[1].Value = $false
                    }
                })
        }
        # Creation of OK_Btn Button Object
        $OK_Btn = [System.Windows.Forms.Button]::new()
        # With OK_Btn Control...
        $OK_Btn | ForEach-Object {
            # Control properties definition
            $_.Name = 'OK_Btn'
            $_.Text = "OK"
            $_.AutoSize = $false
            $_.Size = '70,30'
            $_.Location = '510,240'
            $_.Font = [string]$MyFonts.Normal
            $_.TextAlign = "MiddleCenter"
            $_.DialogResult = "OK"
            $_.UseVisualStyleBackColor = $true
            $_.TabStop = $true
            $_.TabIndex = 1
            $_.Enabled = $true
            $_.Visible = $true
            # Adding Click event handler to button
            $_.Add_Click( {
                    # For each row of the DataGridView...
                    $LogFieldsList_DGV.Rows | ForEach-Object {
                        # Get W3Cname from DataGridView row of Field item
                        $tmpFieldName = [string]$_.Cells[3].Value | ForEach-Object { $start = $_.IndexOf('(') + 1; $end = $_.LastIndexOf(')') - $start; $_.Substring($start, $end) }
                        # Set new value of Selected property of corresponding LogField object according to selection (checkbox status) in DataGridView.
                        $($script:LogFieldList.Where( { $_.W3Cname -eq $tmpFieldName })).Selected = $_.Cells[1].Value
                    }
                })
        }
        # Define OK and Cancel buttons as acceptButton and CancelButton for the form.
        $_.AcceptButton = $OK_Btn
        $_.CancelButton = $Cancel_Btn
        # Adding Controls to Form Object
        $_.Controls.AddRange(@($LogFieldsList_DGV, $Cancel_Btn, $Reset_Btn, $OK_Btn))
        # Creation of LogFieldsSel_TTp ToolTip Object
        $LogFieldsSel_TTp = [System.Windows.Forms.ToolTip]::new()
        # With LogFieldsSel_TTp Object...
        $LogFieldsSel_TTp | ForEach-Object {
            $_.isBalloon = $true
            $_.AutoPopDelay = 5000
            $_.InitialDelay = 500
            $_.ReshowDelay = 100
            $_.ToolTipIcon = "Info"
            # Set help messages to display in tooltips for each control of the form.
            $_.SetToolTip($Reset_Btn, "Click on this button to restore default selected log fields.`nTo apply these changes click on OK button.")
        }
        $_.ShowDialog()
    }
}
#endregion
#-------------------------------------------------------------------------------
#region: Show-ComputerManagement function (Display a dialog form to add remote computers in the computerlist)
function Show-ComputerManagement {
    $script:RecentlyAdded = @()
    # Creation of RCMgmt_Frm Form Object
    $RCMgmt_Frm = [System.Windows.Forms.Form]::new()
    # With RCMgmt_Frm Form Object...
    $RCMgmt_Frm | ForEach-Object {
        # Form Object properties definition
        $_.Name = "RCMgmt_Frm"
        $_.Text = "{0} - Remote computers management" -f $MyScript.Name
        $_.ClientSize = '520,440'
        $_.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
        $_.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
        #$_.ControlBox = $false
        $_.MaximizeBox = $false
        $_.MinimizeBox = $false
        # Get Form Icon from SmallImgList ImageList Object
        $_.Icon = [System.Drawing.Icon]::FromHandle([System.Drawing.Bitmap]::new($SmallImgList.Images[0]).GetHicon())
        # Creation of ComputerMgmtExplain_Lbl Label Object
        $ComputerMgmtExplain_Lbl = [System.Windows.Forms.Label]::new()
        # With ComputerMgmtExplain_Lbl Control...
        $ComputerMgmtExplain_Lbl | ForEach-Object {
            # Control properties definition
            $_.Name = 'ComputerMgmtExplain_Lbl'
            $_.Text = "This dialog allow you to add or remove remote computers from the script analysis scope.`nBy default, local computer is always part of the analysis scope.`n`nIf no ""Shared Credentials"" is defined, you can specify custom credentials per computer.`nPlease note that adding a remote computer locks the ""Shared credentials"" settings."
            $_.AutoSize = $false
            $_.Size = '500,90'
            $_.Location = '10,10'
            $_.Padding = '5,5,5,5'
            $_.BackColor = "Info"
            $_.ForeColor = "DarkGoldenrod"
            $_.Font = [string]$MyFonts.Normal
            $_.TextAlign = "MiddleLeft"
            $_.TabStop = $false
            $_.TabIndex = 0
            $_.Enabled = $true
            $_.Visible = $true
            # Adding a Paint event handler to LogFieldsExplain_Lbl control to force "DarkGoldenrod" colored border.
            $_.Add_Paint( { [System.Windows.Forms.ControlPaint]::DrawBorder($_.Graphics, $_.ClipRectangle, "DarkGoldenrod", "Solid") })
        }
        # Creation of Computers_LVw ListView Object
        $Computers_LVw = [System.Windows.Forms.ListView]::new()
        # With Computers_LVw Control...
        $Computers_LVw | ForEach-Object {
            # Control properties definition
            $_.Name = 'Computers_LVw'
            $_.AutoArrange = $true
            $_.Location = '10,110'
            $_.Size = '440,200'
            $_.Font = [string]$MyFonts.Normal
            $_.ForeColor = "WindowText"
            $_.View = "Details"
            $_.HeaderStyle = "Clickable"
            $_.Alignment = "Top"
            $_.AllowColumnReorder = $false
            $_.FullRowSelect = $true
            $_.HideSelection = $true
            $_.MultiSelect = $true
            $_.Scrollable = $true
            $_.Sorting = "None"
            $_.LabelEdit = $false
            $_.LabelWrap = $true
            $_.Text = "Test de texte pour voir"
            $_.TabStop = $true
            $_.TabIndex = 1
            $_.Enabled = $true
            $_.Visible = $true
            $_.LargeImageList = $LargeImgList
            $_.SmallImageList = $SmallImgList
            # Creation of Computer_LVG ColumnHeader object (Directory column for Log directories Listview)
            $Computer_LVG = [System.Windows.Forms.ColumnHeader]::new()
            # With Computer_LVG ColumnHeader object...
            $Computer_LVG | ForEach-Object {
                $_.Text = "Computer"
                $_.TextAlign = 0
                $_.Width = 150
                $_.ImageIndex = 0
            }
            # Creation of CurrentUser_LVG ColumnHeader object (Directory column for Log directories Listview)
            $CurrentUser_LVG = [System.Windows.Forms.ColumnHeader]::new()
            # With CurrentUser_LVG ColumnHeader object...
            $CurrentUser_LVG | ForEach-Object {
                $_.Text = ""
                $_.TextAlign = 0
                $_.Width = 35
                $_.ImageIndex = 14
            }
            # Creation of CustomCreds_LVG ColumnHeader object (Directory column for Log directories Listview)
            $CustomCreds_LVG = [System.Windows.Forms.ColumnHeader]::new()
            # With CustomCreds_LVG ColumnHeader object...
            $CustomCreds_LVG | ForEach-Object {
                $_.Text = "Custom credentials"
                $_.TextAlign = 0
                $_.Width = 200
                $_.ImageIndex = 13
            }
            # Creation of SharedCreds_LVG ColumnHeader object (Directory column for Log directories Listview)
            $SharedCreds_LVG = [System.Windows.Forms.ColumnHeader]::new()
            # With SharedCreds_LVG ColumnHeader object...
            $SharedCreds_LVG | ForEach-Object {
                $_.Text = ""
                $_.TextAlign = 0
                $_.Width = 35
                $_.ImageIndex = 15
            }
            $_.Columns.AddRange(@($Computer_LVG, $CurrentUser_LVG, $CustomCreds_LVG, $SharedCreds_LVG))
            # Adding an IntemSelectionChanged event handler to Listview
            $_.Add_ItemSelectionChanged( {
                    # If some computer is selected in the list view...
                    if ($this.SelectedItems.Count -gt 0) {
                        # Enable the Remove button
                        $RemRC_Btn.Enabled = $true
                        # Firstly enable the custom credentials settings depending of usage of shared credentials.
                        $SetCreds_Btn.Enabled = !$script:UseSharedCreds
                        # If the Item from selection changed event is selected...
                        if ($_.Item.Selected) {
                            # Get corresponding computer name
                            $tmpsrvname = $_.Item.Subitems[0].Text
                            # Get corresponding Computer Object from collection
                            $tmpsrvobj = $script:ComputerList.where( { $tmpsrvname -contains $_.Name })
                            # If computer object already have a PSDrive mounted (then stored in its PSDrive property), then disable custom credentials setting button.
                            if (![string]::IsNullOrEmpty($tmpsrvobj.PSDrive)) { $SetCreds_Btn.Enabled = $false }
                        }
                    }
                    else {
                        # Else (not any item selected in the listview), disable both buttons
                        $RemRC_Btn.Enabled = $false
                        $SetCreds_Btn.Enabled = $false
                    }
                })
            # First filling of listview with already existing computers in Collection
            $script:ComputerList | Add-ComputerToLV
        }
        # Creation of AddRC_Btn Button Object
        $AddRC_Btn = [System.Windows.Forms.Button]::new()
        # With AddRC_Btn Control...
        $AddRC_Btn | ForEach-Object {
            # Control properties definition
            $_.Name = 'AddRC_Btn'
            $_.Text = "+"
            $_.AutoSize = $false
            $_.Size = '50,50'
            $_.Location = '460,109'
            $_.BackColor = "Control"
            $_.ForeColor = "LimeGreen"
            $_.Font = "Segoe IU,20,style=Bold"
            $_.TextAlign = "MiddleCenter"
            $_.UseVisualStyleBackColor = $true
            $_.TabStop = $true
            $_.TabIndex = 2
            $_.Enabled = $true
            $_.Visible = $true
            # Adding Click event handler to the control.
            $_.Add_Click( {
                    # Display the computer adding form and store returned ID in the list of recently added computers.
                    $script:RecentlyAdded += Show-AddRcForm
                } )
        }
        # Creation of RemRC_Btn Button Object
        $RemRC_Btn = [System.Windows.Forms.Button]::new()
        # With RemRC_Btn Control...
        $RemRC_Btn | ForEach-Object {
            # Control properties definition
            $_.Name = 'RemRC_Btn'
            $_.Text = "-"
            $_.AutoSize = $false
            $_.Size = '50,50'
            $_.Location = '460,169'
            $_.BackColor = "Control"
            $_.ForeColor = "Red"
            $_.Font = "Segoe IU,25,style=Bold"
            $_.TextAlign = "MiddleCenter"
            $_.UseVisualStyleBackColor = $true
            $_.TabStop = $true
            $_.TabIndex = 3
            $_.Enabled = $false
            $_.Visible = $true
            # Adding Click event handler to the control.
            $_.Add_Click( {
                    # convert list of selected computers to a multiline string.
                    $Computers_LVw.SelectedItems | ForEach-Object { [string]$remlist += '- ' + $_.Subitems[0].Text + "`n"; $remlist }
                    # Request user confirmation before deleting any computer from the list
                    $RemConfirm = Show-Question $("By clicking this button you will remove following computers from de list and all related data (log directories list, Log files list, etc.):`n`n{0}`nDo you want to continue?" -f $remlist) "Are you sure ?"
                    # If user confirm computer(s) deletion...
                    if ($RemConfirm -eq 6) {
                        # For each selected item in the listview...
                        $Computers_LVw.SelectedItems | ForEach-Object {
                            # Get Computer Object corresponding to list item
                            $tmpCompObj = $script:ComputerList | Where-Object -Property Name -EQ $_.SubItems[0].Text
                            # Removing the id of the computer from recently added computers
                            $script:RecentlyAdded = $RecentlyAdded.Where( { $_ -ne $tmpCompObj.ID })
                            # Get count of log directories in collection
                            $beforecount = $script:LogDirectoryList.Count
                            # Removing log directories corresponding to removed computer
                            $script:LogDirectoryList.Where( { $_.ServerID -eq $tmpCompObj.ID }) | ForEach-Object { $script:LogDirectoryList.Remove($_) }
                            # If there are less log directories in collection after computer deletion than before...
                            if ($beforecount -gt $script:LogDirectoryList.Count) {
                                # Clear all Log Directory listview items in main from
                                $LogDirectories_LVw.SelectedItems.Clear()
                                $LogDirectories_LVw.Items.Clear()
                                # Reload log Directories form collection to Listview in main form
                                $script:LogDirectoryList | Add-DirectoryToLV
                            }
                            if (![string]::IsNullOrEmpty($tmpCompObj.PSDrive)) {
                                Get-PSDrive -Name $tmpCompObj.PSDrive | Remove-PSDrive -Force
                            }
                            # Removing the computer object from the computer list
                            $script:ComputerList.Remove($tmpCompObj)
                            # Removing the computer item from the listview
                            $Computers_LVw.Items.Remove($_)
                        }
                    }
                })
        }
        # Creation of SetCreds_Btn Button Object
        $SetCreds_Btn = [System.Windows.Forms.Button]::new()
        # With SetCreds_Btn Control...
        $SetCreds_Btn | ForEach-Object {
            # Control properties definition
            $_.Name = 'SetCreds_Btn'
            $_.Text = ""
            $_.Size = '50,50'
            $_.Location = '460,260'
            $_.TextAlign = "MiddleCenter"
            $_.TextImageRelation = "Overlay"
            $_.ImageAlign = "MiddleCenter"
            $_.ImageList = $SmallImgList
            $_.ImageIndex = 20
            $_.UseVisualStyleBackColor = $true
            $_.TabStop = $true
            $_.TabIndex = 4
            $_.Enabled = $false
            $_.Visible = $true
            # Adding Click event handler to the control.
            $_.Add_Click( {
                    # convert list of selected computers to a multiline string.
                    $Computers_LVw.SelectedItems | ForEach-Object { [string]$remlist += '- ' + $_.Subitems[0].Text + "`n"; $remlist }
                    # If at least 1 of the selected computers already have custom credentials...
                    if ($Computers_LVw.SelectedItems.where( { $_.Subitems[2].Text -ne "(None)" })) {
                        # Request user confirmation before clearing custom credentials to selected computers.
                        $RemConfirm = Show-Question $("By clicking this button, because at least 1 of the selected computers already have custom credentials defined, you will clear custom credentials for following computers:`n`n {0}`nDo you want to continue?" -f $remlist) "Are you sure ?"
                        # If user confirm computer(s) custom credentials deletion...
                        if ($RemConfirm -eq 6) {
                            $Computers_LVw.SelectedItems | ForEach-Object {
                                $script:ComputerList | Where-Object -Property Name -EQ $_.Subitems[0].Text | ForEach-Object {
                                    $_.Credential = $null
                                    # Get computer ID
                                    $tmpID = $_.ID
                                    # If computer ID is not already in RecentlyAdded array, add the ID in the array.
                                    if (!$script:RecentlyAdded.Where( { $tmpID -contains $_ })) { $script:RecentlyAdded += $tmpID }
                                }
                                $_.Subitems[2].Text = "(None)"
                                $_.SubItems[2].Font = [string][GuiFont]::ToItalic($MyFonts.Normal)
                                $_.SubItems[2].ForeColor = "WindowText"
                            }
                        }
                    }
                    else {
                        # Request user confirmation before applying custom credentials to selected computers.
                        $RemConfirm = Show-Question $("By clicking this button you will apply custom credentials to connect on following computers:`n`n {0}`nDo you want to continue?" -f $remlist) "Are you sure ?"
                        # If user confirm to apply new custom credentials for selected computers...
                        if ($RemConfirm -eq 6) {
                            # Set temporary credentials variable to $null
                            $tmpCred = $null
                            # Get tested credentials.
                            $tmpCred = Get-TestedCreds
                            # If credentials were provided...
                            if ($tmpCred) {
                                # For each selected item in the listview...
                                $Computers_LVw.SelectedItems | ForEach-Object {
                                    $computerObj = $script:ComputerList | Where-Object -Property Name -EQ $_.Subitems[0].Text
                                    $computerObj.Credential = $tmpCred
                                    # Get computer ID
                                    $tmpID = $computerObj.ID
                                    # If computer ID is not already in RecentlyAdded array, add the ID in the array.
                                    if (!$script:RecentlyAdded.Where( { $tmpID -contains $_ })) { $script:RecentlyAdded += $tmpID }
                                    # If no custom credentials are set for the computer...
                                    if ($computerObj.Credential -eq "(None)") {
                                        # Add column of custom credentials containing "(None)" in italic
                                        $_.Subitems[2].Text = $computerObj.Credential
                                        $_.Subitems[2].Font = [string][GuiFont]::ToItalic($MyFonts.Normal)
                                        $_.Subitems[2].ForeColor = "WindowText"
                                    }
                                    # Else (a custom credentials is set)...
                                    else {
                                        # Add column of Authentication status for custom credential
                                        if ($computerObj.Authentications.withCredentials) {
                                            # "OK" and username in regular and green
                                            $siTxt = "OK ({0})" -f $computerObj.Credential
                                            $_.Subitems[2].Font = [string]$MyFonts.Normal
                                            $_.Subitems[2].ForeColor = "Green"
                                        }
                                        else {
                                            # "KO" and username in bold and red
                                            $siTxt = "KO ({0})" -f $computerObj.Credential
                                            $_.Subitems[2].Font = [string]$MyFonts.Strong
                                            $_.Subitems[2].ForeColor = "Red"
                                        }
                                        $_.Subitems[2].Text = $siTxt
                                    }
                                }
                            }
                        }
                    }
                })
        }
        # Creation of ComputerUCol_Lbl Label Object
        $ComputerUCol_Lbl = [System.Windows.Forms.Label]::new()
        # With ComputerUCol_Lbl Control...
        $ComputerUCol_Lbl | ForEach-Object {
            # Control properties definition
            $_.Name = 'ComputerUCol_Lbl'
            $_.Text = "      represents an unreachable remote computer."
            $_.Size = '420,20'
            $_.Location = '10,315'
            $_.Font = [string]$MyFonts.Normal
            $_.TextAlign = "MiddleLeft"
            $_.ImageList = $SmallImgList
            $_.ImageAlign = "MiddleLeft"
            $_.ImageIndex = 1
            $_.TabStop = $false
            $_.TabIndex = 0
            $_.Enabled = $true
            $_.Visible = $true
        }
        # Creation of ComputerCCol_Lbl Label Object
        $ComputerCCol_Lbl = [System.Windows.Forms.Label]::new()
        # With ComputerCCol_Lbl Control...
        $ComputerCCol_Lbl | ForEach-Object {
            # Control properties definition
            $_.Name = 'ComputerCCol_Lbl'
            $_.Text = "      represents a reachable remote computer."
            $_.Size = '420,20'
            $_.Location = '10,335'
            $_.Font = [string]$MyFonts.Normal
            $_.TextAlign = "MiddleLeft"
            $_.ImageList = $SmallImgList
            $_.ImageAlign = "MiddleLeft"
            $_.ImageIndex = 2
            $_.TabStop = $false
            $_.TabIndex = 0
            $_.Enabled = $true
            $_.Visible = $true
        }
        # Creation of CurUserCol_Lbl Label Object
        $CurUserCol_Lbl = [System.Windows.Forms.Label]::new()
        # With CurUserCol_Lbl Control...
        $CurUserCol_Lbl | ForEach-Object {
            # Control properties definition
            $_.Name = 'CurUserCol_Lbl'
            $_.Text = "      authentication status for current user ({0}\{1})." -f $env:USERDOMAIN, $env:USERNAME
            $_.Size = '420,20'
            $_.Location = '10,365'
            $_.Font = [string]$MyFonts.Normal
            $_.TextAlign = "MiddleLeft"
            $_.ImageList = $SmallImgList
            $_.ImageAlign = "MiddleLeft"
            $_.ImageIndex = 14
            $_.TabStop = $false
            $_.TabIndex = 0
            $_.Enabled = $true
            $_.Visible = $true
        }
        # Creation of CustCredsCol_Lbl Label Object
        $CustCredsCol_Lbl = [System.Windows.Forms.Label]::new()
        # With CustCredsCol_Lbl Control...
        $CustCredsCol_Lbl | ForEach-Object {
            # Control properties definition
            $_.Name = 'CustCredsCol_Lbl'
            $_.Text = "      custom provided credentials and corresponding authentication status."
            $_.Size = '420,20'
            $_.Location = '10,385'
            $_.Font = [string]$MyFonts.Normal
            $_.TextAlign = "MiddleLeft"
            $_.ImageList = $SmallImgList
            $_.ImageAlign = "MiddleLeft"
            $_.ImageIndex = 13
            $_.TabStop = $false
            $_.TabIndex = 0
            $_.Enabled = $true
            $_.Visible = $true
        }
        # Creation of SharedCredsCol_Lbl Label Object
        $SharedCredsCol_Lbl = [System.Windows.Forms.Label]::new()
        # With SharedCredsCol_Lbl Control...
        $SharedCredsCol_Lbl | ForEach-Object {
            # Control properties definition
            $_.Name = 'SharedCredsCol_Lbl'
            $_.Text = "      authentication status for defined ""Shared Credentials""."
            $_.Size = '420,20'
            $_.Location = '10,405'
            $_.Font = [string]$MyFonts.Normal
            $_.TextAlign = "MiddleLeft"
            $_.ImageList = $SmallImgList
            $_.ImageAlign = "MiddleLeft"
            $_.ImageIndex = 15
            $_.TabStop = $false
            $_.TabIndex = 0
            $_.Enabled = $true
            $_.Visible = $true
        }
        # Creation of Close_Btn Button Object
        $Close_Btn = [System.Windows.Forms.Button]::new()
        # With Close_Btn Control...
        $Close_Btn | ForEach-Object {
            # Control properties definition
            $_.Name = 'Close_Btn'
            $_.Text = "Close"
            $_.AutoSize = $false
            $_.Size = '75,50'
            $_.Location = '435,380'
            $_.Font = [string]$MyFonts.Normal
            $_.TextAlign = "MiddleCenter"
            $_.UseVisualStyleBackColor = $true
            $_.TabStop = $true
            $_.TabIndex = 5
            $_.Enabled = $true
            $_.Visible = $true
            # Adding Click event handler to the control.
            $_.Add_Click( { $($this.FindForm()).Close() })
        }
        # Adding Controls to Form Object
        $_.Controls.AddRange(@($ComputerMgmtExplain_Lbl, $Computers_LVw, $AddRC_Btn, $RemRC_Btn, $SetCreds_Btn, $ComputerUCol_Lbl, $ComputerCCol_Lbl, $CurUserCol_Lbl, $CustCredsCol_Lbl, $SharedCredsCol_Lbl, $Close_Btn))

        # Creation of ComputerMgmt_TTp ToolTip Object
        $ComputerMgmt_TTp = [System.Windows.Forms.ToolTip]::new()
        # With ComputerMgmt_TTp Object...
        $ComputerMgmt_TTp | ForEach-Object {
            $_.isBalloon = $true
            $_.AutoPopDelay = 5000
            $_.InitialDelay = 500
            $_.ReshowDelay = 100
            $_.ToolTipIcon = "Info"
            # Set help messages to display in tooltips for each control of the form.
            $_.SetToolTip($AddRC_Btn, "Allow you to add a remote computer to the list.")
            $_.SetToolTip($RemRC_Btn, "Allow you to remove selected computers from the list.")
            $_.SetToolTip($SetCreds_Btn, "Allow you to set/remove custom credentials for selected computers.`nIf one of selected computer already have custom credentials defined, it delete credentials for all selected computers.")

        }
        [void]$_.ShowDialog()
    }
    # Get Recently Added variable content and set it in retArr variable (for function returned value)
    $retArr = $script:RecentlyAdded
    # removing of script scoped RecentlyAdded variable
    Remove-Variable -Name RecentlyAdded -Scope script -Force -ErrorAction SilentlyContinue
    # Return an array of computer objects IDs of recently added computer.
    return $retArr
}
#endregion
#-------------------------------------------------------------------------------
#region: Show-AddRcForm function (Display a dialog box to provide name of a new computer to add in the list)
function Show-AddRcForm () {
    $ARCret = $null
    # Creation of AddRC_Frm Form Object
    $AddRC_Frm = [System.Windows.Forms.Form]::new()
    # With AddRC_Frm Form Object...
    $AddRC_Frm | ForEach-Object {
        # Form Object properties definition
        $_.Name = "AddRC_Frm"
        $_.Text = "Add a remote computer..."
        $_.ClientSize = '350,100'
        $_.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
        $_.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedToolWindow
        # Creation of HostName_Lbl Label Object
        $HostName_Lbl = [System.Windows.Forms.Label]::new()
        # With HostName_Lbl Control...
        $HostName_Lbl | ForEach-Object {
            # Control properties definition
            $_.Name = 'HostName_Lbl'
            $_.Text = "Computer's hostname:"
            $_.AutoSize = $false
            $_.Size = '140,20'
            $_.Location = '10,10'
            $_.Font = [string]$MyFonts.Normal
            $_.TextAlign = "MiddleLeft"
            $_.TabStop = $false
            $_.TabIndex = 0
            $_.Enabled = $true
            $_.Visible = $true
        }
        # Creation of HostName_Txt TextBox Object
        $HostName_Txt = [System.Windows.Forms.TextBox]::new()
        # With HostName_Txt Control...
        $HostName_Txt | ForEach-Object {
            # Control properties definition
            $_.Name = 'HostName_Txt'
            $_.Location = '160,10'
            $_.Size = '180,20'
            $_.Font = [string]$MyFonts.Normal
            $_.TextAlign = "Left"
            $_.TabStop = $true
            $_.TabIndex = 1
            $_.Visible = $true
            $_.Enabled = $true
            # Adding an onTextChanged event handler (it disable OK button until an hostname is provided)
            $_.Add_TextChanged( { if (!([string]::IsNullOrEmpty($this.Text))) { $Ok_Btn.Enabled = $true } else { $Ok_Btn.enabled = $false } })
            # Adding a KeyDown event handler (click on OK button if Enter Key is pressed while the control has focus)
            $_.Add_KeyDown( { if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) { $OK_Btn.PerformClick() } })
        }
        # Creation of SetCustomCreds_CBx CheckBox Object
        $SetCustomCreds_CBx = [System.Windows.Forms.CheckBox]::new()
        # With SetCustomCreds_CBx Control...
        $SetCustomCreds_CBx | ForEach-Object {
            # Control properties definition
            $_.Name = 'SetCustomCreds_CBx'
            $_.Text = "Use alternate credentials to connect this computer."
            $_.Size = '330,20'
            $_.Location = '10,40'
            $_.Font = [string]$MyFonts.Normal
            $_.AutoCheck = $true
            $_.CheckAlign = "TopLeft"
            $_.Checked = $false
            $_.UseVisualStyleBackColor = $true
            $_.TextAlign = "MiddleLeft"
            $_.TabStop = $true
            $_.TabIndex = 2
            $_.Enabled = !$script:UseSharedCreds
            $_.Visible = $true
            # Adding a KeyDown event handler (click on OK button if Enter Key is pressed while the control has focus)
            $_.Add_KeyDown( { if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) { $OK_Btn.PerformClick() } })
        }
        # Creation of OK_Btn Button Object
        $OK_Btn = [System.Windows.Forms.Button]::new()
        # With OK_Btn Control...
        $OK_Btn | ForEach-Object {
            # Control properties definition
            $_.Name = 'OK_Btn'
            $_.Text = "OK"
            $_.AutoSize = $false
            $_.Size = '75,25'
            $_.Location = '265,65'
            $_.Font = [string]$MyFonts.Normal
            $_.TextAlign = "MiddleCenter"
            $_.UseVisualStyleBackColor = $true
            $_.TabStop = $true
            $_.TabIndex = 3
            $_.Enabled = $false
            $_.Visible = $true
            # Adding a Click event handler to control ()
            $_.Add_Click( {
                    # If computer name is not a valid hostname string, display a warning message, then return to adding computer form.
                    if ($HostName_Txt.Text -notmatch '^(([a-zA-Z0-9]|[a-zA-Z0-9][a-zA-Z0-9\-]*[a-zA-Z0-9])\.)*([A-Za-z0-9]|[A-Za-z0-9][A-Za-z0-9\-]*[A-Za-z0-9])$') {
                        Show-Warning "This is not a valid hostname!" "Bad hostname syntax"
                        return
                    }
                    # If the computer name already exists in the list, display a warning message, then return to adding computer form.
                    if ($script:ComputerList.Where( { $_.Name -eq $HostName_Txt.Text.ToUpper() })) {
                        Show-Warning "This computer name already exists in the list!`nPlease retry with another name." "Computer already in the list"
                        return
                    }
                    # Set temporary credentials variable to $null
                    $tmpCred = $null
                    # If "Use alternate credentials..." is checked, then try to get tested credentials.
                    if ($SetCustomCreds_CBx.Checked) { $tmpCred = Get-TestedCreds }
                    # Get the next available ID for new computer
                    $newID = $($script:ComputerList | Measure-Object -Property ID -Maximum).Maximum + 1
                    # Create new Computer Object (creation process is different with credentials than without)
                    if ($tmpCred) { $CompObj = [Computer]::new($HostName_Txt.Text, $tmpCred, $newID) } else { $CompObj = [Computer]::new($HostName_Txt.Text, $newID) }
                    # If Computer Object is correctly created...
                    if ($CompObj) {
                        # Add the computer object to the collection
                        $script:ComputerList.Add($CompObj)
                        # Add computer item to ListView in Computer management window
                        Add-ComputerToLV $CompObj
                        # Set function return value to new computer ID
                        Set-Variable -Name ARCret -Value $newID -Scope 2
                        #$ARCret = $newID
                    }
                    $this.FindForm().Close()
                })
        }
        # Creation of Cancel_Btn Button Object
        $Cancel_Btn = [System.Windows.Forms.Button]::new()
        # With Cancel_Btn Control...
        $Cancel_Btn | ForEach-Object {
            # Control properties definition
            $_.Name = 'Cancel_Btn'
            $_.Text = "Cancel"
            $_.AutoSize = $false
            $_.Size = '75,25'
            $_.Location = '180,65'
            $_.Font = [string]$MyFonts.Normal
            $_.TextAlign = "MiddleCenter"
            $_.DialogResult = "Cancel"
            $_.UseVisualStyleBackColor = $true
            $_.TabStop = $true
            $_.TabIndex = 4
            $_.Enabled = $true
            $_.Visible = $true

        }
        # Don't forget to add this new control to its parent (Form or Container control).
        # Adding Controls to Form Object
        $_.Controls.AddRange(@($HostName_Lbl, $HostName_Txt, $SetCustomCreds_CBx, $OK_Btn, $Cancel_Btn))
        # Set form cancel button
        $_.CancelButton = $Cancel_Btn
        [void]$_.ShowDialog()
    }
    return $ARCret
}
#endregion
#-------------------------------------------------------------------------------
#region: Show-AddLogDirForm function (Display a dialog box to select alternate log directory to add in the analysis scope)
function Show-AddLogDirForm {
    # Creation of AddLogDirectory_Frm Form Object
    $AddLogDirectory_Frm = [System.Windows.Forms.Form]::new()
    # With AddLogDirectory_Frm Form Object...
    $AddLogDirectory_Frm | ForEach-Object {
        # Form Object properties definition
        $_.Name = "AddLogDirectory_Frm"
        $_.Text = "{0} - Add alternate log directory" -f $MyScript.Name
        $_.ClientSize = '350,385'
        $_.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
        $_.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
        #$_.ControlBox = $false
        $_.MaximizeBox = $false
        $_.MinimizeBox = $false
        # Get Form Icon from SmallImgList ImageList Object
        $_.Icon = [System.Drawing.Icon]::FromHandle([System.Drawing.Bitmap]::new($SmallImgList.Images[18]).GetHicon())
        # Creation of AddLogDir_Lbl Label Object
        $AddLogDir_Lbl = [System.Windows.Forms.Label]::new()
        # With AddLogDir_Lbl Control...
        $AddLogDir_Lbl | ForEach-Object {
            # Control properties definition
            $_.Name = 'AddLogDir_Lbl'
            $_.Text = "Select a log files directory:"
            $_.Size = '330,20'
            $_.Location = '10,10'
            $_.Font = [string]$MyFonts.Normal
            $_.TextAlign = "TopLeft"
            $_.TabStop = $false
            $_.TabIndex = 0
            $_.Enabled = $true
            $_.Visible = $true
        }
        # Creation of AddLogDir_TVw TreeView Object
        $AddLogDir_TVw = [System.Windows.Forms.TreeView]::new()
        # With AddLogDir_TVw Control...
        $AddLogDir_TVw | ForEach-Object {
            # Control properties definition
            $_.Name = 'AddLogDir_TVw'
            $_.Size = '330,300'
            $_.Location = '10,35'
            $_.Scrollable = $true
            $_.Font = [string]$MyFonts.Normal
            $_.BorderStyle = "Fixed3D"
            $_.BackColor = "Window"
            $_.LineColor = "Black"
            $_.Enabled = $true
            $_.Visible = $true
            $_.TabStop = $true
            $_.ImageList = $SmallImgList
            $_.Add_AfterExpand( {
                    if ($_.Node.Level -ge 1) {
                        $this.BeginUpdate()
                        # Show Application Loading form
                        Show-AppLoading
                        # Update loading message
                        $MsgDisplay_Lbl.Text = "Loading"
                        $MsgDisplay_Lbl.Refresh()
                        $updtCount = 0
                        $_.Node.Nodes | ForEach-Object {
                            if ($_.Nodes.Count -eq 0 -and $_.Text -ne "WinSxS") {
                                $ParentNode = $_
                                Get-ChildItem $ParentNode.Tag -Directory -ErrorAction SilentlyContinue | ForEach-Object {
                                    # Update loading message
                                    if ($updtCount -ge 10) {
                                        if ($MsgDisplay_Lbl.Text.Length -gt 10) { $MsgDisplay_Lbl.Text = "Loading" } else { $MsgDisplay_Lbl.Text += "." }
                                        $updtCount = 0
                                    }
                                    $updtCount++
                                    $MsgDisplay_Lbl.Refresh()
                                    #Get current directory object
                                    $Target = $_
                                    # Create TreeNode with Directory name
                                    $DirectoryNode = [System.Windows.Forms.TreeNode]::new($_.Name)
                                    # With new Directory TreeNode...
                                    $DirectoryNode | ForEach-Object {
                                        # Set Images Indexes
                                        $_.ImageIndex = 18
                                        $_.SelectedImageIndex = 18
                                        # Add Directory filesystem object to Tag property of TreeNode
                                        $_.Tag = Get-Item $Target.FullName
                                    }
                                    # Adding of Directory TreeNode to its corresponding Parent directory TreeNode
                                    [void]$ParentNode.Nodes.Add($DirectoryNode)
                                }
                            }
                        }
                        # Closing Application loading form.
                        $AppLoading_Frm.Close()
                        $this.EndUpdate()
                    }
                })
            #region : Populate the 2 first levels of the treeview
            # For each FileSystem PSDrive...
            Get-PSDrive -PSProvider FileSystem | ForEach-Object {
                # Get correspondinf computer object from collection
                $tmpCompObj = $script:ComputerList | Where-Object -Property PSDrive -EQ $_.Name
                # If a computer object was found...
                if ($tmpCompObj) {
                    # Set Computer string to computer name
                    $ComputerStr = $tmpCompObj.Name
                    # Set Drive string (ie. for \\MyServer\c$ mounted as ILADRV00 PSDrive -> "C drive (ILADrv00:)")
                    $DriveStr = "{0} drive ({1}:)" -f $_.Root.Substring($_.Root.IndexOf('\', 2) + 1).ToUpper(), $_.Name
                    # Set drive image index (remote disk icon)
                    $DrvImgId = 8
                    # Set Drive path (ie. "\\MyServer\c$\")
                    $DrvPath = "{0}:\" -f $_.Name
                    # Else (no computer object found ==> then the local computer - PSDrive name is only stored for remote computers Objects)...
                }
                else {
                    # If the PSDrive root string contains ":"...
                    if ($_.Root -like '*:*') {
                        # Get computer object of Local computer from collection
                        $tmpCompObj = $script:ComputerList.Where( { $_.ID -eq 0 })
                        # Set computer string
                        $ComputerStr = "Local Computer ({0})" -f $tmpCompObj.Name
                        # If not empty string, extract last directory from DisplayRoot property of PSDrive
                        $DrvDRoot = if (![string]::IsNullOrEmpty($_.DisplayRoot)) { $_.DisplayRoot.Substring($_.DisplayRoot.LastIndexOf('\') + 1) }
                        # If PSDrive description is not an empty string, set Drive name to PSDrive Description
                        if (![string]::IsNullOrEmpty($_.Description)) { $DrvName = $_.Description }
                        # Else (no description available), if extracted PSDrive's DisplayRoot is not an empty string, set Drive name to the extracted string
                        elseif (![string]::IsNullOrEmpty($DrvDRoot)) { $DrvName = $DrvDRoot }
                        # Else (no Description and no DisplayRoot available), set Drive name to "Unknown"
                        else { $DrvName = "Unknown" }
                        # Set Drive string to previously defined Drive name and drive letter between parentheses
                        $DriveStr = "{1} ({0}:)" -f $_.Name, $DrvName
                        # Set drive image index (local disk icon)
                        $DrvImgId = 7
                        # Set drive path to the PSDrive root path
                        $DrvPath = "{0}:\" -f $_.Name
                    }
                }
                # If Computer object was found
                if ($tmpCompObj) {
                    # Get TreeNode corresponding to computer with Computer string from the Treeview nodes
                    $ComputerNode = $($AddLogDir_TVw.Nodes.Where( { $_.Text -eq $ComputerStr }))
                    # If not found...
                    if (!$ComputerNode) {
                        # Create a TreeNode with Computer string
                        $ComputerNode = [System.Windows.Forms.TreeNode]::new($ComputerStr)
                        # With new computer TreeNode...
                        $ComputerNode | ForEach-Object {
                            # Set Images Indexes
                            $_.ImageIndex = 1
                            $_.SelectedImageIndex = 1
                            # Add Computer Object to Tag property of TreeNode
                            $_.Tag = $tmpCompObj
                        }
                        # Adding of Computer TreeNode to Treeview control
                        [void]$AddLogDir_TVw.Nodes.Add($ComputerNode)
                    }
                    # Create TreeNode for current PSDrive with Drive string
                    $DriveNode = [System.Windows.Forms.TreeNode]::new($DriveStr)
                    # With the new Drive TreeNode...
                    $DriveNode | ForEach-Object {
                        # Set Images Indexes with previously defined one
                        $_.ImageIndex = $DrvImgId
                        $_.SelectedImageIndex = $DrvImgId
                        # Get FileSystem item corresponding to drive path and Add it to Tag property of TreeNode
                        $_.Tag = $DrvPath
                    }
                    # Adding Drive TreeNode to its corresponding Computer TreeNode
                    [void]$ComputerNode.Nodes.Add($DriveNode)
                    # For each Directory in the current Drive...
                    Get-ChildItem $DriveNode.Tag -Directory -ErrorAction SilentlyContinue | ForEach-Object {
                        #Get current directory object
                        $Target = $_
                        # Create TreeNode with Directory name
                        $DirectoryNode = [System.Windows.Forms.TreeNode]::new($_.Name)
                        # With new Directory TreeNode...
                        $DirectoryNode | ForEach-Object {
                            # Set Images Indexes
                            $_.ImageIndex = 18
                            $_.SelectedImageIndex = 18
                            # Add Directory filesystem object to Tag property of TreeNode
                            $_.Tag = Get-Item $Target.FullName
                        }
                        # Adding of Directory TreeNode to its corresponding Drive TreeNode
                        [void]$DriveNode.Nodes.Add($DirectoryNode)
                    }
                }
            }
            #endregion
        }
        # Creation of OK_Btn Button Object
        $OK_Btn = [System.Windows.Forms.Button]::new()
        # With OK_Btn Control...
        $OK_Btn | ForEach-Object {
            # Control properties definition
            $_.Name = 'OK_Btn'
            $_.Text = "OK"
            $_.AutoSize = $false
            $_.Size = '75,30'
            $_.Location = '265,345'
            $_.Font = [string]$MyFonts.Normal
            $_.DialogResult = "OK"
            $_.UseVisualStyleBackColor = $true
            $_.TabStop = $true
            $_.TabIndex = 0
            $_.Enabled = $true
            $_.Visible = $true
            # Adding a Click event handler to OK button
            $_.Add_Click( {
                    # If the selected item level in the treeview is less than 2 (then it is a computer or a drive)...
                    if ($AddLogDir_TVw.SelectedNode.Level -lt 2) {
                        # Display a warning message then return to form
                        Show-Warning "The selected item is not a directory!`nOnly directories can be selected as logs directory." "Bad selection"
                        return
                    }
                    # Get Selected node
                    $tmpNode = $AddLogDir_TVw.SelectedNode
                    # Get Filesystem object of selected item from its Tag property
                    $SelectedDir = $AddLogDir_TVw.SelectedNode.Tag
                    # Get FullName (path) of directory corresponding to selected item.
                    $SelDilPath = $SelectedDir.FullName + '\'
                    # For each LogDirectory object in the collection, if the selected item directory path is shorter than the LogDirectory object's path, and then if item path is in the LogDirectory object path
                    # (in fact if the selected item directory is a parent directory of an existing LogDirectory in the collection)...
                    if ($script:LogDirectoryList.Where( { $_.Path.Length -ge $SelDirPath.Length }).where( { $_.Path.Substring(0, $SelDirPath.Length) -eq $SelDirPath })) {
                        # Display a warning message then return to form
                        Show-Warning "The selected directory is already in the log directories list or is a parent directory of one of them!`nPlease select another directory." "Bad selection"
                        return
                    }
                    # Loop to Root node of the selected node
                    For ($i = 0 ; $i -lt $AddLogDir_TVw.SelectedNode.Level ; $i++) {
                        $tmpNode = $tmpNode.Parent
                    }
                    # Show Application Loading form
                    Show-AppLoading
                    # Update loading message
                    $MsgDisplay_Lbl.Text = "Loading"
                    $MsgDisplay_Lbl.Refresh()
                    # Update loading message
                    $MsgDisplay_Lbl.Text += "."
                    $MsgDisplay_Lbl.Refresh()
                    # Closing Application loading form.
                    $AppLoading_Frm.Close()

                    # Set counter to 0
                    $localldc = 0
                    # Get Log directories in selected directory, then for each found log directory...
                    Get-LogDirectories -Path $SelDilPath | ForEach-Object {
                        # Add the found log directory to LogDirectory objects collection with computer ID
                        $script:LogDirectoryList.Add([LogDirectory]::New($_, $tmpNode.Tag.ID))
                        # Increase counter
                        $localldc++
                    }
                    # If no log directory found (counter equals 0), Display a warning message box, else...
                    if ($localldc -eq 0) { Show-Warning "The selected directory seems to not contain any IIS logs directory!" "No IIS logs found" } else {
                        # For each LogDirectory Object which does not yet have any ID (in fact not yet added in the directory list view) in the collection...
                        $script:LogDirectoryList | Where-Object -Property ID -EQ -1 | ForEach-Object {
                            # Update loading message
                            $MsgDisplay_Lbl.Text = "Please wait"
                            $MsgDisplay_Lbl.Refresh()
                            # Get log files list by calling GetFiles method of LogDirectory Object.
                            $_.GetFiles()
                            $_
                        }
                    }
                })
        }
        # Creation of Cancel_Btn Button Object
        $Cancel_Btn = [System.Windows.Forms.Button]::new()
        # With Cancel_Btn Control...
        $Cancel_Btn | ForEach-Object {
            # Control properties definition
            $_.Name = 'Cancel_Btn'
            $_.Text = "Cancel"
            $_.AutoSize = $false
            $_.Size = '75,30'
            $_.Location = '10,345'
            $_.Font = [string]$MyFonts.Normal
            $_.DialogResult = "Cancel"
            $_.UseVisualStyleBackColor = $true
            $_.TabStop = $true
            $_.TabIndex = 0
            $_.Enabled = $true
            $_.Visible = $true
        }
        # Don't forget to add this new control to its parent (Form or Container control).
        # Adding Controls to Form Object
        $_.Controls.AddRange(@($AddLogDir_Lbl, $AddLogDir_TVw, $OK_Btn, $Cancel_Btn))
        $_.ShowDialog()
    }
}
#endregion
#-------------------------------------------------------------------------------
#region: Show-AppLoading function (show a small form to let user wait by displaying short message)
function Show-AppLoading {
    # Creation of AppLoading_Frm Form Object
    $script:AppLoading_Frm = [System.Windows.Forms.Form]::new()
    # With AppLoading_Frm Form Object...
    $AppLoading_Frm | ForEach-Object {
        # Form Object properties definition
        $_.Name = "AppLoading_Frm"
        $_.Text = "Looking for log directories on local computer"
        $_.ClientSize = '250,90'
        $_.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
        $_.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
        # Creation of a Label Control
        $script:MsgDisplay_Lbl = [System.Windows.Forms.Label]::new()
        # With Label Control...
        $script:MsgDisplay_Lbl | ForEach-Object {
            # Control properties definition
            $_.Name = "MsgDisplay_Lbl"
            $_.Text = ""
            $_.Size = '230,70'
            $_.Location = '10,10'
            $_.Font = "Segoe UI,22,style=Bold,Italic"
            $_.TextAlign = "MiddleCenter"
        }
        # Adding Controls to Form Object
        $_.Controls.AddRange(@($MsgDisplay_Lbl))
        $_.Add_Shown( { $this.Activate() })
        [void]$_.Show()
    }
}
#endregion
#-------------------------------------------------------------------------------
#region: Show-Warning function (Dispays a MessageBox with only Ok button and Excalamtion icon)
function Show-Warning ([string]$message, [string]$title = "") {
    [System.Windows.MessageBox]::Show($message, $title, 0, 48) | Out-Null
}
#endregion
#-------------------------------------------------------------------------------
#region: Show-Question function (displays a question dialog with yes and no buttons and, by default Exclamation icon)
function Show-Question ([string]$message, [string]$title = "", [System.Windows.MessageBoxImage]$icon = [System.Windows.MessageBoxImage]::Exclamation) {
    return [System.Windows.MessageBox]::Show($message, $title, 4, $icon)
}
#endregion
#-------------------------------------------------------------------------------
#region: Update-StatusBarScope function (to refresh status bar with amount of files in scope of the script analysis)
function Update-StatusBarScope {
    # Set Total and scope amounts of files coounters to 0
    $TotalFilesCount = 0
    $script:ScopeFilesCount = 0
    # For each selected LogDirectory in the collection...
    $script:LogDirectoryList.where( { $_.Selected }) | ForEach-Object {
        # Increase the total counter with amount of all IIS log files in the LogDirectory object
        $TotalFilesCount += $_.LogFiles.Where( { $_.isIISLog }).Count
        # If LogDirectory object's filter is not set...
        if ($_.Filter.Oldest -eq "Not set") {
            # Increase the scope counter with amount of all IIS log files in the LogDirectory object
            $script:ScopeFilesCount += $_.LogFiles.Where( { $_.isIISLog }).Count
            # Else (a filter is set)...
        }
        else {
            # Get Filtering start and end dates from LogDirectory object
            [datetime]$startFilter = (Get-Date -Date $_.Filter.Oldest)
            [datetime]$endFilter = (Get-Date -Date $_.Filter.Latest)
            # Increase the scope counter with amount of only IIS log files which are between filtering dates
            $script:ScopeFilesCount += $_.LogFiles.Where( { $_.isIISLog -and $_.CreationDate -ge $startFilter -and $_.CreationDate -le $endFilter }).Count
        }
    }
    # If there are more than 1 IIS log files in scope...
    if ($script:ScopeFilesCount -gt 1) {
        # Set strings to complete status text with plurals
        $scopetxtsuffix = "s"
        $scopetxtverb = "are"
    }
    else {
        # else, set strings to complete status text with singles
        $scopetxtsuffix = ""
        $scopetxtverb = "is"
    }
    # If there are more than 1 IIS log files in total set string with plural, else set string with single
    if ($TotalFilesCount -gt 1) { $totaltxtsuffix = "s" } else { $totaltxtsuffix = "" }
    # Build text to display in status bar for script scope of analysis
    [string]$StatusBarScopeTxt = "{0} IIS log file{1} on a total of {2} found file{3} {4} in scope of the analysis." -f $ScopeFilesCount, $scopetxtsuffix, $TotalFilesCount, $totaltxtsuffix, $scopetxtverb
    # Set status bar scope text
    $MainStatusLabel1_TSL.Text = $StatusBarScopeTxt
}
#endregion
#-------------------------------------------------------------------------------
#region: Update-SharedCredsControls function (update and refresh Shared credentials control appearance)
function Update-SharedCredsControls {
    # If "Force use of common credentials..." checkbox is checked...
    if ($UseSharedCreds_CBx.Checked) {
        # Enable "Shared credentials" controls.
        $SCUsername_Lbl.Enabled = $true
        $SCUsername_Txt.Enabled = $true
        $SCPassword_Lbl.Enabled = $true
        $SCPassword_Txt.Enabled = $true
        $SharedCredsDetails_Lbl.Enabled = $true
        $SCSave_Btn.Enabled = $true
        $SCClear_Btn.Enabled = $true
    }
    # else...
    else {
        # Disable "Shared credentials" controls.
        $SCUsername_Lbl.Enabled = $false
        $SCUsername_Txt.Enabled = $false
        $SCPassword_Lbl.Enabled = $false
        $SCPassword_Txt.Enabled = $false
        $SharedCredsDetails_Lbl.Enabled = $false
        $SCSave_Btn.Enabled = $false
        $SCClear_Btn.Enabled = $false
    }
}
#endregion
#-------------------------------------------------------------------------------
#region: Update-FilterDates function (apply Filtering dates changes to listview item and to LogDirectory object)
function Update-FilterDates ($startDate = $null, $endDate = $null) {
    # Get currently selected directory item in listView
    $Item = $LogDirectories_LVw.SelectedItems[0]
    # If startDate or endDate provided...
    if ((!$startDate) -or (!$endDate)) {
        # Clear LogDirectory object's filtering dates and wipe 2 last columns in list view
        $Item.Tag.Filter.ClearRange()
        $Item.SubItems[4].Text = ""
        $Item.SubItems[5].Text = ""
    }
    else {
        # Set Filtering Dates to LogDirectory Object and to Listview Item
        $Item.Tag.Filter.SetRange($startDate, $endDate, $Item.Tag.LogsDates)
        $Item.SubItems[4].Text = [string](Get-Date -Date $startDate -Format d)
        $Item.SubItems[5].Text = [string](Get-Date -Date $endDate -Format d)
    }
}
#endregion
#-------------------------------------------------------------------------------
#region: Update-SelDirFilter function (udpate and refresh aspect and content of filtering dates controls in details panel)
function Update-SelDirFilter ($Item) {
    # Define limits of DateTimePicker objects
    $OldestDate_DTP.MinDate = (Get-Date -Date $Item.SubItems[2].Text)
    $OldestDate_DTP.MaxDate = (Get-Date -Date $Item.SubItems[3].Text)
    $LatestDate_DTP.MinDate = (Get-Date -Date $Item.SubItems[2].Text)
    $LatestDate_DTP.MaxDate = (Get-Date -Date $Item.SubItems[3].Text)
    # If no filter is yet defined for selected directory (if it contains --- or is empty instead of a date)...
    if (($Item.SubItems[4].Text -eq "---") -or ($Item.SubItems[4].Text -eq "")) {
        # Set Filtering DateTimePicker objects value to first and last files dates.
        $OldestDate_DTP.Value = (Get-Date -Date $Item.SubItems[2].Text)
        $LatestDate_DTP.Value = (Get-Date -Date $Item.SubItems[3].Text)
        # Enable DateTimePicker controls and Apply button
        $OldestDate_DTP.Enabled = $true
        $LatestDate_DTP.Enabled = $true
        $ApplyFilter_Btn.Visible = $true
        # Disable Clear button
        $ClearFilter_Btn.Visible = $False
    }
    else {
        # a Filter is already applied for this directory.
        # Set Filtering DateTimePicker objects value to defined filtering dates.
        $OldestDate_DTP.Value = (Get-Date -Date $Item.SubItems[4].Text)
        $LatestDate_DTP.Value = (Get-Date -Date $Item.SubItems[5].Text)
        # Disable DateTimePicker controls and Apply button
        $OldestDate_DTP.Enabled = $False
        $LatestDate_DTP.Enabled = $False
        $ApplyFilter_Btn.Visible = $False
        # Enable Clear button
        $ClearFilter_Btn.Visible = $true
    }
    # Set values of analysis scope dates to filtering DateTimePicker dates.
    $ScopeStartV_Lbl.Text = [string](Get-Date -Date $OldestDate_DTP.Value -Format d)
    $ScopeEndV_Lbl.Text = [string](Get-Date -Date $LatestDate_DTP.Value -Format d)
    # Get amount of log files in the scope and display it (log files which are IIS Log, and file date is between filtering oldest and latest dates)
    $LogsAmountInScopeV_Lbl.Text = [string]$Item.Tag.LogFiles.Where( { $_.isIISLog -EQ $true -and $_.CreationDate -ge $OldestDate_DTP.Value -and $_.CreationDate -le $LatestDate_DTP.Value }).Count
}
#endregion
#-------------------------------------------------------------------------------
#region: Refresh-SelDirDetails function (update and refresh content of directory details panel)
function Update-SelDirDetails ($sender, $Item) {
    # If an Item is selected in Log Directories Listview...
    if ($sender.SelectedItems.Count -gt 0) {
        # Displaying directory path is details "panel" from selected item.
        $SelectedLogDir_Lbl.Text = $Item.SubItems[0].Text
        # Make directory details fields and labels visible
        $SelectedLogDir_Lbl.Visible = $true
        $LogsAmountInDirL_Lbl.Visible = $true
        $LogsAmountInDirV_Lbl.Visible = $true
        # Display amount of log files found in selected directory from item.
        $LogsAmountInDirV_Lbl.Text = $Item.SubItems[1].Text
        # If there are found file(s)...
        if ([int]$Item.SubItems[1].Text -gt 0) {
            # Set first and last log files dates to corresponding displaying labels
            $SDFirstFileDateV_Lbl.Text = $Item.SubItems[2].Text
            $SDLastFileDateV_Lbl.Text = $Item.SubItems[3].Text
            # Make Dates, filtering and analysis scope details visible
            $SDFirstFileDateL_Lbl.Visible = $true
            $SDFirstFileDateV_Lbl.Visible = $true
            $SDLastFileDateL_Lbl.Visible = $true
            $SDLastFileDateV_Lbl.Visible = $true
            $LogsFiltering_Grp.Visible = $true
            $DirAnalysisScope_Grp.Visible = $true
        }
        # Refresh and update filtering fields and controls.
        Update-SelDirFilter $Item
        # Set tooltip text to path of the selected directory.
        $MainToolTip_TTp.SetToolTip($SelectedLogDir_Lbl, $Item.SubItems[0].Text)
    }
    else {
        # If no Item selected, wipe directory path field and hide every directory details fields and labels.
        $SelectedLogDir_Lbl.Text = ""
        $SelectedLogDir_Lbl.Visible = $False
        $LogsAmountInDirL_Lbl.Visible = $False
        $LogsAmountInDirV_Lbl.Visible = $False
        $SDFirstFileDateL_Lbl.Visible = $False
        $SDFirstFileDateV_Lbl.Visible = $False
        $SDLastFileDateL_Lbl.Visible = $False
        $SDLastFileDateV_Lbl.Visible = $False
        $LogsFiltering_Grp.Visible = $False
        $DirAnalysisScope_Grp.Visible = $False
        # Reset limits of DateTimePicker objects (-100 years for Minimums & +100 years for Maximums of 2 DateTimePicker objects)
        $OldestDate_DTP.MinDate = (Get-Date -Date $OldestDate_DTP.MinDate).AddYears(-100)
        $OldestDate_DTP.MaxDate = (Get-Date -Date $OldestDate_DTP.MaxDate).AddYears(100)
        $LatestDate_DTP.MinDate = (Get-Date -Date $LatestDate_DTP.MinDate).AddYears(-100)
        $LatestDate_DTP.MaxDate = (Get-Date -Date $LatestDate_DTP.MaxDate).AddYears(100)
    }
}
#endregion
#-------------------------------------------------------------------------------
#region: Add-ComputerToLV function (add computer item in computer management form's ListView)
function Add-ComputerToLV {
    param (
        # Specifies one or more LogDirectory objects.
        [Parameter(Mandatory = $true,
            Position = 0,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true,
            HelpMessage = "One or more Computer objects.")]
        [ValidateNotNullOrEmpty()]
        [Computer[]]
        $Computer
    )
    # Allow processing each input object from Pipeline
    process {
        foreach ($computerObj in $Computer) {
            # Only if Type property of Computer Object is Remote...
            #if ($computerObj.Type -EQ [ComputerType]::Local) {
            if ($computerObj.Type -EQ [ComputerType]::Remote) {
                # Create temporary listview item with computer name
                $tmpLVItem = [System.Windows.Forms.ListViewItem]::new($computerObj.Name)
                # Disable style inheritance from item ot subitems
                $tmpLVItem.UseItemStyleForSubItems = $false
                # Set item icon depending of reachability status of computer
                if ($computerObj.IsReachable) { $tmpLVItem.ImageIndex = 2 } else { $tmpLVItem.ImageIndex = 1 }
                # Add column of Authentication status for current user
                if ($computerObj.Authentications.LocalUser) {
                    # "OK" in regular and green
                    $tmpSubItem = $tmpLVItem.SubItems.Add("OK")
                    $tmpSubItem.Font = [string]$MyFonts.Normal
                    $tmpSubItem.ForeColor = "Green"
                }
                else {
                    # "KO" in bold and red
                    $tmpSubItem = $tmpLVItem.SubItems.Add("KO")
                    $tmpSubItem.Font = [string]$MyFonts.Strong
                    $tmpSubItem.ForeColor = "Red"
                }
                # If no custom credentials are set for the computer...
                if ($computerObj.Credential -eq "(None)") {
                    # Add column of custom credentials containing "(None)" in italic
                    $tmpSubItem = $tmpLVItem.SubItems.Add($computerObj.Credential)
                    $tmpSubItem.Font = [string][GuiFont]::ToItalic($MyFonts.Normal)
                    $tmpSubItem.ForeColor = "WindowText"
                }
                # Else (a custom credentials is set)...
                else {
                    $tmpSubItem = $tmpLVItem.SubItems.Add("")
                    # Add column of Authentication status for custom credential
                    if ($computerObj.Authentications.withCredentials) {
                        # "OK" and username in regular and green
                        $siTxt = "OK ({0})" -f $computerObj.Credential
                        $tmpSubItem.Font = [string]$MyFonts.Normal
                        $tmpSubItem.ForeColor = "Green"
                    }
                    else {
                        # "KO" and username in bold and red
                        $siTxt = "KO ({0})" -f $computerObj.Credential
                        $tmpSubItem.Font = [string]$MyFonts.Strong
                        $tmpSubItem.ForeColor = "Red"
                    }
                    $tmpSubItem.Text = $siTxt
                }
                # If Shared Credentials usage setting is activated...
                if ($script:UseSharedCreds) {
                    # If Shared credentials is not empty...
                    if ($script:SharedCredentials -ne [System.Management.Automation.PSCredential]::Empty) {
                        # Test authentication with shared credentials and store result
                        $siStatus = $computerObj.TestAuthentication($script:SharedCredentials)
                    }
                    else {
                        # else (so similar than using current user), set result to current user status
                        $siStatus = $computerObj.Authentications.LocalUser
                    }
                    # Add column of Authentication status for shared credentials
                    if ($siStatus) {
                        # "OK" in regular and green
                        $tmpSubItem = $tmpLVItem.SubItems.Add("OK")
                        $tmpSubItem.Font = [string]$MyFonts.Normal
                        $tmpSubItem.ForeColor = "Green"
                    }
                    else {
                        # "KO" in bold and red
                        $tmpSubItem = $tmpLVItem.SubItems.Add("KO")
                        $tmpSubItem.Font = [string]$MyFonts.Strong
                        $tmpSubItem.ForeColor = "Red"
                    }
                }
                else {
                    # If shred Credentials usage setting is not activated, then add column for shared credentials containing "n/a" in italic.
                    $tmpSubItem = $tmpLVItem.SubItems.Add("n/a")
                    $tmpSubItem.Font = [string][GuiFont]::ToItalic($MyFonts.Normal)
                    $tmpSubItem.ForeColor = "WindowText"
                }
                # Adding of computer listview item in the list.
                [void]$Computers_LVw.Items.Add($tmpLVItem)
            }
        }
    }
}
#--
#$Computers_LVw
#--
#endregion
#-------------------------------------------------------------------------------
#region: Add-DirectoryToLV function (add log directory item in main from's ListView)
function Add-DirectoryToLV {
    param (
        # Specifies one or more LogDirectory objects.
        [Parameter(Mandatory = $true,
            Position = 0,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true,
            HelpMessage = "One or more LogDirectory objects.")]
        [Alias("LogDirectory")]
        [ValidateNotNullOrEmpty()]
        [LogDirectory[]]
        $LogDir
    )
    # Allow processing each input object from Pipeline
    process {
        foreach ($ldObj in $LogDir) {
            # Get Computer Object corresponding to current LogDirectory object
            $computerObj = $ComputerList | Where-Object -Property ID -EQ $ldObj.ServerID
            # Get/Set ListView Grouping Name according to Computer corresponding to ServerID property of directory object
            $grpName = "on {0} server" -f $computerObj.Name
            # Get ListViewGroup Object from ListView groups correspoding to previously Grouping Name
            $tmpLVGroup = $LogDirectories_LVw.Groups | Where-Object -Property header -EQ $grpName
            # If no group found, create then add new group to ListView
            if (!$tmpLVGroup) {
                $tmpLVGroup = [System.Windows.Forms.ListViewGroup]::New($grpName)
                [void]$LogDirectories_LVw.Groups.Add($tmpLVGroup)
            }
            if (![string]::IsNullOrEmpty($computerObj.PSDrive)) {
                $root_path = (Get-PSDrive -Name $computerObj.PSDrive).Root
                $tmpLDFPath = $ldObj.Path.Replace($root_path, 'C:')
            }
            else { $tmpLDFPath = $ldObj.Path }
            # Creation of a ListViewItem with firendly path and in previously defined Group
            $tmpLVItem = [System.Windows.Forms.ListViewItem]::new($tmpLDFPath, $tmpLVGroup)
            $tmpLVItem.ImageIndex = 18
            # Get amount of IIS Log files found in the directory
            $tmpLogsCount = $($ldObj.LogFiles | Where-Object -Property isIISLog -EQ $true).Count
            # Add amount of IIS log files in 2nd column of the Item
            [void]$tmpLVItem.SubItems.Add($tmpLogsCount.ToString())
            # If there is no IIS Log File in the directory...
            if ($tmpLogsCount -eq 0) {
                # Change item text color to LightGray
                $tmpLVItem.ForeColor = "LightGray"
                # Uncheck the item
                $tmpLVItem.Checked = $false
                # Set "---" string in all dates columns.
                [void]$tmpLVItem.SubItems.Add("---")
                [void]$tmpLVItem.SubItems.Add("---")
                [void]$tmpLVItem.SubItems.Add("---")
                [void]$tmpLVItem.SubItems.Add("---")
            }
            # Else (if some IIS log files found)...
            else {
                # Check by default the directory item
                $tmpLVItem.Checked = $true
                # Get log files dates for the directory and add them in corresponding columns
                [void]$tmpLVItem.SubItems.Add([string](Get-Date -Date $ldObj.LogsDates.Oldest -Format d))
                [void]$tmpLVItem.SubItems.Add([string](Get-Date -Date $ldObj.LogsDates.Latest -Format d))
                if ($ldObj.Filter.Oldest -eq "Not set") {
                    # Set filtering scope date to empty columns
                    [void]$tmpLVItem.SubItems.Add("")
                    [void]$tmpLVItem.SubItems.Add("")
                }
                else {
                    # Set filtering scope dates columns
                    [void]$tmpLVItem.SubItems.Add([string](Get-Date -Date $ldObj.Filter.Oldest -Format d))
                    [void]$tmpLVItem.SubItems.Add([string](Get-Date -Date $ldObj.Filter.Latest -Format d))
                }
            }
            # Add directory Item in Listview, get its index in ListView and store it in ID property of Directory Object
            $ldObj.ID = $LogDirectories_LVw.Items.Add($tmpLVItem).Index
            # Set Selected property value of Directory Object to Checked status of item
            $ldObj.Selected = $TmpLVItem.Checked
            # Adding the LogDirectory Object to Item's Tag property
            $tmpLVItem.Tag = $ldObj
        }
    }
}
#endregion
#-------------------------------------------------------------------------------
