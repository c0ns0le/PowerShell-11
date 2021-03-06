################################################################################
#                                                                              #
# This script calls the imageGrabLite function of NVPR to rename the output    #
# file of a scheduled recording and moves it to the recorded tv folder.        #
#                                                                              #
# Created Date: 27 SEP 2013                                                    #
# Version: 3.0                                                                 #
#                                                                              #
################################################################################

################################################################################
# VARIABLES                                                                    #
################################################################################

# Include functions module.
$recording = "D:\Videos\Recorded_TV\"
$date = Get-Date -Format "yyyy-MM-dd"

################################################################################
# PROCESS FLOW                                                                 #
################################################################################

sl $recording
$show = @(Get-ChildItem *.ts)[0]

& .\TVShows.ps1 $recording $show "20" "20" "single" $null "Recording"