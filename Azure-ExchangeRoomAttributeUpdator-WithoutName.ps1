<# 
Adjust the attribues of Exchange Rooms in bulk using CSV input. It adds attribute values based on what is found in the columns of the
csv and assigns them in the proper place in Exchange. It also adds Zoom Room date to the Phone attribute for better Zoom Room tracking. 
For On Prem Active Directory Room Objects, it skips them and adds to a CSV that
is exported at the end. This also assigns rooms to room lists (for Room Finder) if defined in the appropriate column. 

This script does not affect the display name of the room, although this could be easily modified to include it. 

Ensure that the source csv has the following column names for it to function properly -

Name | Email | Location | City | Floor | FloorLabel | ZoomRoom | Phone | Capacity | RoomList

#>


# Import necessary modules
Import-Module ExchangeOnlineManagement

# Assign paths
$sourcePath = "C:\Path\To\SourceFile"
$sourceCSV = "RoomAttributes.csv"
$destinationPath = "C:\Path\To\Output"

# Reset arrays and counters
$startTime = 0
$elapsedTime = 0
$roomCounter = 0
$onPremCounter = 0
$onPremRooms = @()
$updatedRooms = @()

# Start a timer
$startTime = Get-Date

# Get CSV input and convert 'Capacity' column to integer
$roomsToAdjust = Import-Csv "$sourcePath\$sourceCSV" | ForEach-Object {
    # Convert Capacity column to integer
    $_.Capacity = [int]$_.Capacity
    $_.Floor = [int]$_.Floor
    
    # Return the updated object
    $_                              
}

# Counts the number of rooms to adjust
$roomTotal = $roomsToAdjust.Count - 1

# Begin adjustment & start transcript

Write-Host "Beginning Room Attribute Adjustment --`nThere are $roomTotal rooms to process`n" -ForegroundColor Yellow

# Connect-ExchangeOnline -ShowBanner:$false

foreach ($room in $roomsToAdjust) {
    $roomName = $room.Name
    $roomEmail = $room.Email
    $roomLocation = $room.Location
    $roomCity = $room.City
    $roomNumFloor = $room.Floor
    $roomAlphaFloor = $room.FloorLabel
    $roomZoomSwitch = $room.ZoomRoom
    $roomPhone = $room.Phone
    $roomCapacity = $room.Capacity
    $roomList = $room.RoomList

    # Handle null values for Capacity
    if (-not $roomCapacity) {
        $roomCapacity = 0
    }

    # Handle null values for room floor numbers
    if (-not $roomNumFloor) {
        $roomNumFloor = 1
    }

    # Retrieve the mailbox
    $mailbox = Get-Mailbox "$roomEmail" -ErrorAction SilentlyContinue

    # Create a new mailbox if one is not found - uncomment to run, left lines 69-74 commented just for safety
    if ($mailbox -eq $null) {
        Write-Host "Mailbox not found for $roomEmail! Creating a new mailbox" -ForegroundColor Yellow

        # Handle null values for Capacity
        if (-not $roomCapacity) {
            $roomCapacity = 0
        }

        # Handle null values for Capacity
        if (-not $roomNumFloor) {
            $roomCapacity = 1
        }

        New-Mailbox -Name $roomName -DisplayName $roomName -PrimarySmtpAddress $roomEmail -Office $roomLocation -Phone $roomZoomSwitch -ResourceCapacity $roomCapacity -Room

        # Sets attributes of existing mailbox
    }
    else {

        # Checks to see if the room is on-prem synced (AD)
        $onPrem = (Get-User $mailbox).IsDirSynced

        # Adds on-prem rooms to array of skipped rooms
        if ($onPrem -eq $true) {
            Write-Host "$roomName found - $mailbox" -ForegroundColor Cyan
            Write-Host "This mailbox is on prem only! Adding to list of rooms skipped`n" -ForegroundColor Red

            $onPremCounter ++
            $onPremRooms += [PSCustomObject]@{
                Name          = $room.Name
                Email         = $room.Email
                Location      = $room.Location
                City          = $room.City
                Floor         = $room.Floor
                FloorLabel    = $room.FloorLabel
                ZoomRoom      = $room.ZoomRoom
                Phone         = $room.Phone
                Capacity      = $room.Capacity
                RoomList      = $room.RoomList
                DirectorySync = "OnPrem"
            }
            # Display number of skipped rooms
            Write-Host "[$onPremCounter of $roomTotal] rooms are on prem so far.`n" -ForegroundColor Magenta

            # Process room changes in 3 steps    
        }
        else {
            Write-Host "$roomName found - $mailbox`nSetting attributes of $mailbox" -ForegroundColor Cyan
            Write-Host "[1 of 3]`nSetting location & capacity..."
            Set-Mailbox -Identity $mailbox -Office $roomLocation -ResourceCapacity $roomCapacity
            Write-Host "Done!" -ForegroundColor Green
    
            Write-Host "[2 of 3]`nSetting city, phone, and floor data..."
            Set-Place -Identity $roomEmail -City $roomCity -Capacity $roomCapacity -Phone $roomPhone -Floor $roomNumFloor -FloorLabel $roomAlphaFloor
            Write-Host "Done!" -ForegroundColor Green
    
            Write-Host "[3 of 3]`nAdding to the appropriate room list ($roomList)..."
            Add-DistributionGroupMember -Identity $roomList -Member $roomEmail -ErrorAction SilentlyContinue
            Write-Host "Done!`n" -ForegroundColor Green
            
            # Add processed rooms to an array
            $updatedRooms += [PSCustomObject]@{
                Name          = $room.Name
                Email         = $room.Email
                Location      = $room.Location
                City          = $room.City
                Floor         = $room.Floor
                FloorLabel    = $room.FloorLabel
                ZoomRoom      = $room.ZoomRoom
                Phone         = $room.Phone
                Capacity      = $room.Capacity
                RoomList      = $room.RoomList
                DirectorySync = "Cloud"
            }    
        }

        # Increase counter
        $roomCounter ++
        Write-Host "Processed [$roomCounter of $roomTotal] rooms.`n" -ForegroundColor Yellow

    }
}

# Export output CSVs
Write-Host "Complete! Exporting all completed rooms and [$onPremCounter] skipped, on prem rooms..." -ForegroundColor Cyan
$onPremRooms | Export-Csv "$destinationPath\OnPremRooms.csv" -NoTypeInformation
$updatedRooms | Export-Csv "$destinationPath\UpdatedRooms.csv" -NoTypeInformation

# Get elapsed time and display output
$elapsedTime = (Get-Date) - $startTime
Write-Host "All tasks complete! Time elapsed: $($elapsedTime.Minutes) mins" -ForegroundColor Green