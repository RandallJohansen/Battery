

#
#  shows a notification when battery is charged, or needs charged, alexa enabled
#


                
        # Lithium-ion batteries don't suffer from the memory effect.
        # HP recomends 80% cuttoff for max battery life

[int]$HiThreshold = 80           # pct batt to trigger kill power

[int]$LoThreshold = 10           # pct batt to trigger start power


[int]$DelayBetweenLoops = 600  # seconds



[console]::Title = ( "Battery Alert" )


# [void][system.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')  # .net old school
#           -OR-
Add-Type -Assembly System.Windows.Forms

Add-Type -AssemblyName System.Drawing   # for icon




##############################################
#                                            #
#   Sleep w/events and watch for app exit
#                                            #
##############################################

Function Pause( [int]$Seconds , $NotifyIcon ) {

        
        $ExitTime = ( Get-Date ).AddSeconds( $Seconds )


        while ( ( ( Get-Date ) -lt $ExitTime ) -and $NotifyIcon.Visible ) { 
        
            sleep -Milliseconds 100
            
            [System.Windows.Forms.Application]::DoEvents() 

        }
}


##############################################
#                                            #
#              Alexa enabled
#                                            #
##############################################
#region begin 


$UseVoice = "GB_HAZEL"  #  hush  US_DAVID  GB_HAZEL  US_ZIRA 

$Criteria = "Gender = Female"  # just don't want a guy in my ear

$Voice = New-Object -com SAPI.SpVoice

$Personalities = [string[]]@( $Voice.GetVoices( $Criteria ) | % { $_.id } )

$Chosen = [Array]::FindIndex( $Personalities , [Predicate[string]]{ $args[0] -match $UseVoice } )

if( $Chosen -eq -1 ) {  # would prefer to default to none, after all who wants a complete stranger addressing them?

} else { $Voice.Voice = $Voice.GetVoices( $Criteria ).Item( $Chosen ) }

$Voice.Rate = 0   # -1 sounds drunk

#$Voice.Volume = 80



[void] $Voice.Speak( "Hi" ) # extra word, sometimes misses first sylable(s)
[void] $Voice.Speak( "Cheers Randall, it's me" )
[void] $Voice.Speak( "I'll be watching your battery, that'll be fun for me." )
[void] $Voice.Speak( "Enjoy being worry free, I know you'll like that." )



Function Alexa( [string]$Text2Speak ) {


[void] $Voice.Speak( "Terribly sorry, need a word with the staff." )  # extra word, sometimes misses first sylable(s)

sleep -m 500

[void] $Voice.Speak( 'alexa' )  

sleep -m 200

[void] $Voice.Speak( $Text2Speak )

sleep -s 3

[void] $Voice.Speak( "There, it's done.  That'll be nice for you." )  


} 


#endregion


##############################################
#                                            #
#              Generate icon
#                                            #
##############################################

Function CreateIcon( [int]$Pct , $NotifyIcon , $Charging ) {


        $Image = new-object System.Drawing.Bitmap 32,32                # image of type .bmp

        $font = new-object System.Drawing.Font Verdana,14              # font to use
        
        $Bold = [System.Drawing.Font]::new("Verdana", 10, [System.Drawing.FontStyle]::Bold)

        $graphics = [System.Drawing.Graphics]::FromImage($Image)       # use bitmap as canvas

        $graphics.SmoothingMode = "AntiAlias"

        
        if ( $Pct -gt 75 ) { $Text2Draw = ''


                # smiley face


                $Image.MakeTransparent(  )                                 # to blend away corners of icon on taskbar



                if ( $Charging ) { 
                
                        $brushBg = [System.Drawing.Brushes]::green                 # background
                        
                        $BlackPen = new-object System.Drawing.Pen "white", 3       # foreground, 3 pixels diameter pen tip

                } else {

                        $brushBg = [System.Drawing.Brushes]::yellow                # background yellow

                        $BlackPen = new-object System.Drawing.Pen "Black", 3       # foreground black, 3 pixels diameter pen tip
                }


                # solid yellow circle
                $graphics.FillEllipse($brushBg,0,0,$Image.Width,$Image.Height)

                # black outline for Circle
                $graphics.DrawEllipse( $BlackPen , 3 , 3 , 27 , 27 )        # Draws elipse (circle) at upper-left x/y, width, height
                
                # smile
                $graphics.DrawArc( $BlackPen, 10 , 14 , 12 , 10 , 0 , 180 ) # pen, upper-left x/y, width, height, startAngle, sweepAngle
                
                # left eye
                $graphics.DrawEllipse( $BlackPen , 11 , 11 , 3 , 3 )        # Draws elipse (circle) at upper-left x/y, width, height
                
                # right eye
                $graphics.DrawEllipse( $BlackPen , 19 , 11 , 3 , 3 )        # Draws elipse (circle) at upper-left x/y, width, height


        } else { $Text2Draw = ($Pct).ToString() 


                # solid background with ( one or ) two digit number 

                
                $BgColors = @( 'green' , 'green' , 'yellow' , 'red' )   # background color pallete

                $FgColors = @( 'white' , 'white' , 'black' , 'white' )  # forground color pallete
        
                $Ptr = [int]( 50 / $Pct )                               # set pointer to fore/background color pair

                if ( $Ptr -gt 3 ) { $Ptr = 3 }                          # correct out of bounds ( charge below critical )
        
                $BgColor = $BgColors[ $Ptr ]                            # set background color to match balloon icon criteria
                                                                         
                $FgColor = $FgColors[ $Ptr ]                            # set foreground color to contrast background color

                $brushBg = [System.Drawing.Brushes]::$BgColor                     # background color
                                                                                  
                $brushFg = [System.Drawing.Brushes]::$FgColor                     # foreground color
                                                                                  
                $format = [System.Drawing.StringFormat]::GenericDefault           # allocate a string format
                $format.Alignment = [System.Drawing.StringAlignment]::Center      # .. set string centered  left/right
                $format.LineAlignment = [System.Drawing.StringAlignment]::Center  # .. set string centered  top/bottom

                $Rectangle = [System.Drawing.RectangleF]::FromLTRB(0, 0, $Image.Width, $Image.Height)  # basically entire icon

                $graphics.FillRectangle($brushBg,$Rectangle)                       # Fill background

                $graphics.DrawString($Text2Draw,$font,$brushFg,$Rectangle,$format) # Draws text string in the specified rectangle

                if ( $Charging ) {  # Draws band across bottom to show charging                                                              
                        
                        $Pen = new-object System.Drawing.Pen $FgColor , 2         # pixels diameter pen tip

                        $VertPlace = $Image.Height - 4

                        $graphics.DrawLine( $Pen , 2 , $VertPlace , ( $Image.Width - 2 ) , $VertPlace )           
                        
                } # if

        } # else


        $graphics.Dispose() 

        $icon = [System.Drawing.Icon]::FromHandle($Image.GetHicon())
        
        $Image.Dispose()


  #      $filename = "$home\desktop\foo.ico"   # $home = C:\Users\admin
  #
  #
  #      # write to file
  #
  #      $fileStream = [IO.File]::Create("$filename") 
  #                                    
  #      $icon.Save($fileStream) 
  #      
  #      $fileStream.Close()      
           
      
        $NotifyIcon.Icon = $icon


        $icon.Dispose()
                      
}

    
##############################################
#                                            #
#            get battery data
#                                            #
##############################################

Function GetBatteryData( $NotifyIcon , $Charging ) {   # object type, so by default passed to func by reference

        $battery = Get-CimInstance -ClassName Win32_Battery 

        <#

        Caption                     : Internal Battery
        Description                 : Internal Battery
        InstallDate                 :
        Name                        : Primary
        Status                      : OK
        Availability                : 2
        ConfigManagerErrorCode      :
        ConfigManagerUserConfig     :
        CreationClassName           : Win32_Battery
        DeviceID                    : 01508 2020/12/22Hewlett-PackardPrimary
        ErrorCleared                :
        ErrorDescription            :
        LastErrorCode               :
        PNPDeviceID                 :
        PowerManagementCapabilities : {1}
        PowerManagementSupported    : False
        StatusInfo                  :
        SystemCreationClassName     : Win32_ComputerSystem
        SystemName                  : DESKTOP-RJ2F399
        BatteryStatus               : 2
        Chemistry                   : 2
        DesignCapacity              :
        DesignVoltage               : 13259
        EstimatedChargeRemaining    : 98
        EstimatedRunTime            : 71582788     # equivalent to 0x04444444, means battery being charged
        ExpectedLife                :
        FullChargeCapacity          :
        MaxRechargeTime             :
        SmartBatteryVersion         :
        TimeOnBattery               :
        TimeToFullCharge            :
        BatteryRechargeTime         :
        ExpectedBatteryLife         :
        PSComputerName              :

        #>
 
        $charge = $battery.EstimatedChargeRemaining
 
        $StatusCode = $battery.BatteryStatus
 
        <#

        batteryStatus
        -------------

        Other (1)    The battery is discharging.

        Unknown (2)  The system has access to AC so no battery is being discharged. However, the battery is not necessarily charging.

        Fully Charged (3)

        Low (4)

        Critical (5)

        Charging (6)

        Charging and High (7)

        Charging and Low (8)

        Charging and Critical (9)

        Undefined (10)

        Partially Charged (11)

        #>

        $BatteryStatus = @( 'unknown' , 
                            'discharging          ' ,     
                            'access to AC         ' ,
                            'Fully Charged        ' ,
                            'Low                  ' ,
                            'Critical             ' ,
                            'Charging             ' ,
                            'Charging and High    ' ,
                            'Charging and Low     ' ,
                            'Charging and Critical' ,
                            'Undefined            ' ,
                            'Partially Charged    ' )


        $Status = $BatteryStatus[ $StatusCode ] 

        #    if ( $battery.EstimatedRunTime -eq 71582788 ) {     # 71582788 means battery being charged
        #
        #            $Runtime = "infinite minutes remaining"
        #
        #    } else { 
        #    
        #        [int]$MinutesLeft = New-TimeSpan -minutes $( $battery.EstimatedRunTime / 1 )
        #
        #        [string]$Runtime  = "$MinutesLeft minutes remaining" }
        #
        #    $NotifyIcon.BalloonTipText = "$charge% $Runtime, $Status (code=$StatusCode)"

        $NotifyIcon.BalloonTipText = "$charge% $Status (code=$StatusCode)"
                      
        $NotifyIcon.Text = " $charge% remaining`n $Status `n (code=$StatusCode) "


        $BalloonTipIcons = @( 'None' , 'Info' , 'Warning' , 'Error' )
                        
        $Ptr = [int]( 50 / $charge )

        if ( $Ptr -gt 3 ) { $Ptr = 3 }

        $NotifyIcon.BalloonTipIcon = $BalloonTipIcons[ $Ptr ]
        

        CreateIcon $charge $NotifyIcon $Charging



        Write-Host $NotifyIcon.BalloonTipText


        $Hash = @{

            'charge'     = $charge

            'StatusCode' = $StatusCode

            'Status'     = $Status

        }

        return $Hash

}


##############################################
#                                            #
#               Notify Icon
#                                            #
##############################################
#region begin 

$NotifyIcon = New-Object System.Windows.Forms.NotifyIcon

$BatteryStats = GetBatteryData $NotifyIcon

#endregion


##############################################
#                                            #
#         Notify Icon's Balloon
#                                            #
##############################################
#region begin 

$NotifyIcon.BalloonTipTitle = "Battery remaining"

# the text and icon are set every time battery is checked in function

# actual show is done by left-click

#endregion


##############################################
#                                            #
#      Notify Icon's Left-Click Menu
#                                            #
##############################################
#region begin 

$notifyicon.add_Click( { 

    if ( $_.Button -eq [Windows.Forms.MouseButtons]::Left ) { 

        $NotifyIcon.ShowBalloonTip(0)  # milliseconds to show, 0 = use windows default
    } 
} ) 

#endregion


##############################################
#                                            # 
#  Notify Icon Balloon's Right-Click Menu
#                                            #
##############################################
#region begin 

$menuitem = New-Object System.Windows.Forms.MenuItem 

$menuitem.Text = "Exit"        # When Exit is clicked, close the PowerShell process 

$contextmenu = New-Object System.Windows.Forms.ContextMenu 

$notifyicon.ContextMenu = $contextmenu 

$notifyicon.contextMenu.MenuItems.AddRange($menuitem) 

$menuitem.add_Click( { 

    $notifyicon.Visible = $false   # Note: If “phantom” icon remains after closing, set the Visible property to False

    $notifyicon.dispose()   

 } ) 

 
#endregion
 
 
##############################################
#                                            #
#             Float a Baloon
#                                            #
##############################################

Function FloatBaloon( $NotifyIcon ) {


               Write-Host "Floating a baloon."
               

    #           $balloon = New-Object System.Windows.Forms.NotifyIcon -Property @{
    #
    #                    Icon = [System.Drawing.SystemIcons]::Error
    #
    #                    BalloonTipTitle = $NotifyIcon.BalloonTipTitle
    #
    #                    BalloonTipText = $NotifyIcon.BalloonTipText
    #
    #                    Visible = $True 
    #            }


    #            $balloon.ShowBalloonTip(0)  # parameter deprecated as of Vista, display times based on system accessibility settings


    # already have notifyicon, so no need to create a new one
    #
                $NotifyIcon.ShowBalloonTip(0)  # parameter deprecated as of Vista, display times based on system accessibility settings
                


                ##############################################
                #                                            #
                #       Balloon Tip Clicked Actions
                #                                            #
                ##############################################

    #            $null = Register-ObjectEvent $balloon BalloonTipClicked -SourceIdentifier event_BalloonTipClicked -Action {
                $null = Register-ObjectEvent $NotifyIcon BalloonTipClicked -SourceIdentifier event_BalloonTipClicked -Action {

                        
                        Write-Host  -ForeGround Yellow "event_BalloonTipClicked occured !"        

    # redundant
    #                    Unregister-Event -SourceIdentifier $event.SourceIdentifier -Force
    #                    Remove-Job $event.SourceIdentifier -Force

                        # unregister event and remove job object
                        Unregister-Event -SourceIdentifier event_BalloonTipClosed -Force
                        Remove-Job event_BalloonTipClosed -Force
        
                        # unregister other event and remove job object
                        Unregister-Event -SourceIdentifier event_BalloonTipClicked -Force
                        Remove-Job event_BalloonTipClicked -Force

    #                    $balloon.Visible = $false
    #
    #                    $balloon.Dispose()
                }


                ##############################################
                #                                            #
                #        Balloon Tip Closed Actions
                #                                            #
                ##############################################

    #            $null = Register-ObjectEvent $balloon BalloonTipClosed -SourceIdentifier event_BalloonTipClosed -Action {
                $null = Register-ObjectEvent $NotifyIcon BalloonTipClosed -SourceIdentifier event_BalloonTipClosed -Action {


                        Write-Host -ForeGround Yellow "event_BalloonTipClosed occured !"

    # redundant
    #                    Unregister-Event -SourceIdentifier $event.SourceIdentifier -Force
    #                    Remove-Job $event.SourceIdentifier -Force

                        # unregister event and remove job object
                        Unregister-Event -SourceIdentifier event_BalloonTipClicked -Force
                        Remove-Job event_BalloonTipClicked -Force

                        # unregister other event and remove job object
                        Unregister-Event -SourceIdentifier event_BalloonTipClosed -Force
                        Remove-Job event_BalloonTipClosed -Force

    #                    $balloon.Visible = $false
    #                    
    #                    $balloon.Dispose()

                }

                
                ##############################################
                #                                            #
                #        Wait for Balloon to be done
                #                                            #
                ##############################################

                Write-Host "Waiting..."

                while ( ( Get-EventSubscriber -SourceIdentifier "event_BalloonTipClosed" -ErrorAction SilentlyContinue ) -and $NotifyIcon.Visible ) { 
  
                        sleep -Milliseconds 10
 
                        [System.Windows.Forms.Application]::DoEvents() 
                }
                
}





                        ##############################################
                        #                                            #
#########################                 main {}                    #########################
                        #                                            #
                        ##############################################


$ChargingCodes = @( 2 , 5 , 6 , 7 , 8 , 9 )   # swag from $BatteryStatus values


$NotifyIcon.Visible = $true


while ( $NotifyIcon.Visible ) {


        Write-Host "$(Get-Date) Checking battery"

        $BatteryStats = GetBatteryData $NotifyIcon ( $ChargingCodes -contains $BatteryStats.StatusCode )


        ##############################################
        #                                            #
        #    turn off power if fully charged
        #                                            #
        ##############################################

        if ( ( $BatteryStats.charge -ge $HiThreshold ) -and ( $ChargingCodes -contains $BatteryStats.StatusCode ) ) { 
        
                Write-Host "Fully charged, cutting off power."

              #  Pause 600 $NotifyIcon       # give the battery an extra 10 minutes to be sure it's topped off

                FloatBaloon( $NotifyIcon )

                
                ##############################################
                #                                            #
                #        ask Alexa to turn off power
                #                                            #
                ##############################################

                Write-Host "Alexa: asking to kill power" -ForegroundColor Yellow

                Alexa "turn off laptop outlet"

                Pause 5 $NotifyIcon 

                $BatteryStats = GetBatteryData $NotifyIcon $False # update stats after unplugged by Alexa

        } # if

        
        ##############################################
        #                                            #
        #    notify if charge is below threshold,
        #    and not already charging, also
        #                                            #
        ##############################################
        
                
        if ( ( $BatteryStats.charge -le $LoThreshold ) -and ( $ChargingCodes -notcontains $BatteryStats.StatusCode ) ) {

                Write-Host "Charge is below threshold, and not already charging."

                FloatBaloon( $NotifyIcon )

                
                ##############################################
                #                                            #
                #        ask Alexa to turn on power
                #                                            #
                ##############################################

                Write-Host "Alexa: asking for power" -ForegroundColor Yellow

                Alexa "turn on laptop outlet"

                Pause 5 $NotifyIcon 

                $BatteryStats = GetBatteryData $NotifyIcon $True  # update stats after plugged in by Alexa

        } # if
 

        ##############################################
        #                                            #
        #            Wait to check again
        #                                            #
        ##############################################

        Write-Host "resting..."

        Pause $DelayBetweenLoops $NotifyIcon 


} # while


# clean-up
$obj_tt.Dispose()
