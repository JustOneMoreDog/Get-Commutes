<#
.SYNOPSIS
Returns a best estimate for the travel time between two points at peak commute hours

.DESCRIPTION
Uses Google's Distance Matrix API to get a best estimate for how long it would take to get to City from House assuming you wanted to get there by 9am on the next Monday.
Then uses the same API to find out how long it would take to get from City to House assuming you left City at 5pm on a Friday.
The estimates are designed to be as close to worst case as possible.  

.PARAMETER Cities
Allows you to specify your destination cities.  If not used it will use a list of demo cities.

.PARAMETER ApiKey
Your ApiKey from Google.  Make sure to enable the Distance Matrix API for your apikey or else this will error out.

.PARAMETER House
Allows you to specify your house location. If not used it will use a demo house.

.PARAMETER v
Verbosity switch that will allow you to see a lot more information.  Use only if you are debugging or else you will get flooded.

.EXAMPLE
PS> Get-Commutes -Cities "Chantilly VA","Washington DC" -House "College Park MD" -ApiKey "abcdefghijklmnopqrstuvwxyz" -v
File.txt

.LINK
https://github.com/picnicsecurity/Get-Commutes

#>
function Get-Commutes {
    
    param(
        [Parameter(Mandatory=$false)][string[]]$Cities,
        [Parameter(Mandatory=$false)][string]$House,
        [Parameter(Mandatory=$true)][string]$ApiKey,
        [Parameter(Mandatory=$false)][switch]$v
    )

    # In case we want to see what is going with the api calls for debugging
    if ($v) { 
        $Global:VerbosePreference = 'Continue'
    } else {
        $Global:VerbosePreference = 'SilentlyContinue'
    }

    <# ############################# #>
     #     Setting Everything Up     #
    <# ############################# #> 

    # Setting our URL and Arraylist
    $base_url = 'https://maps.googleapis.com/maps/api/distancematrix/json' # Note this does not need to be json but it is setup to handle it
    $masterList = New-Object System.Collections.ArrayList

    # We can use a base list of cities as a demo or we can take in a list of cities and just check to make sure they are formatted correctly
    if(!$Cities){ 
        $Cities = @(
            "McLean+VA",
            "Alexandria+VA",
            "Arlington+VA",
            "Fort+Belvoir+VA",
            "Springfield+VA",
            "Annandale+VA",
            "Falls+Church+VA",
            "Arlington+VA",
            "Chantilly+VA",
            "Vienna+VA",
            "Reston+VA",
            "Herndon+VA",
            "Dulles+VA"
            "Dupont+Circle+Washington+DC",
            "College+Park+MD",
            "Bethesda+MD",
            "Silver+Spring+MD",
            "Adelphi+MD",
            "Rockville+MD",
            "Beltsville+MD",
            "Laurel+MD",
            "Annapolis+MD",
            "Fort+Meade+MD",
            "Bowie+MD",
            "Crofton+MD",
            "Baltimore+MD",
            "Ellicott+City+MD",
            "Woodstock+MD",
            "Owings+Mills+MD",
            "College+Park+MD"
        )
    } else {
    	# We need to ensure that the cities passed in have no spaces
        $tmp = @()
        $Cities | ForEach-Object {
            $tmp += [string]$($_.ToString().Replace(' ','+'))
        }
        $Cities = $tmp
    }
    
    if(!$House){
        $House = "College+Park+MD"
    } else {
    	# Same deal with the cities above
        $House = $House.Replace(' ','+')
    }
    
    <# ############################# #>
     #      Setting Up The Days      #
    <# ############################# #> 
    
    ## Note these need to be in GMT/UTC and we also want them to be on the right days 
    switch ($(Get-Date -UFormat %u)) {
        1 { [int]$toMonday = 7; [int]$toFriday = 4; } # Monday
        2 { [int]$toMonday = 6; [int]$toFriday = 3; } # Tuesday
        3 { [int]$toMonday = 5; [int]$toFriday = 2; } # Wednesday
        4 { [int]$toMonday = 4; [int]$toFriday = 1; } # Thursday
        5 { [int]$toMonday = 3; [int]$toFriday = 7; } # Friday
        6 { [int]$toMonday = 2; [int]$toFriday = 6; } # Saturday 
        7 { [int]$toMonday = 1; [int]$toFriday = 5; } # Sunday
    }
    
    # Our first api call will be to see what time we should depart at in the morning
    # Setting arrival time in the api does not allow us to use the pessimistic traffic model and thus is not getting us a good estimate of the worst case
    # So we just make a 2nd api call but set our depart time to be 9:00am - time to destination from 1rst api 
    $arrivalTime = Get-Date -Date $(Get-Date -Date 9:00am).AddDays($toMonday) -UFormat %s
    $departureTime = Get-Date -Date $(Get-Date -Date 5:00pm).AddDays($toFriday) -UFormat %s

    <# ############################# #>
     #       Static Parameters       #
    <# ############################# #>
    $apikey = 'key=$ApiKey'
    $units = "units=imperial"
    $mode = "mode=driving"
    $traffic_model = "traffic_model=pessimistic"
    $arrival = "arrival_time=$arrivalTime"
    $departure = "departure_time=$departureTime"

    <# ############################# #>
     #  Looping through the Cities   #
    <# ############################# #>

    $Cities | ForEach-Object {

        # Variable Parameters
        $destinations = "destinations=$($_)"
        $origins = "origins=$House"

        # temporary api call to destination 
        $params = @($origins,$destinations,$arrival,$units,$mode,$apikey) -join "&"
        Write-Verbose "Making API call TO: $($base_url)?$($params)" 
        $apicall = Invoke-RestMethod "$($base_url)?$($params)"

        # times to destination
        $durationTo = $apicall | Select-Object -ExpandProperty rows | Select-Object -ExpandProperty elements | Select-Object -ExpandProperty duration | Select-Object -ExpandProperty text 
        $trafficTo = $apicall | Select-Object -ExpandProperty rows | Select-Object -ExpandProperty elements | Select-Object -ExpandProperty duration_in_traffic -ErrorAction SilentlyContinue | Select-Object -ExpandProperty text 
        Write-Verbose "Duration to is $durationTo and traffic to is $trafficTo"
        
        # format the commute times so that they are in minutes
        [int]$time = 0
        ## time to work
        if($durationTo -like "*hour*"){
            [int]$time = [int]$([int]$($durationTo.Split(' ')[0]) * 60) 
            [int]$time += [int]$($durationTo.Split(' ')[2])
        } else {
            if($durationTo){
                [int]$time += [int]$([int]$($durationTo.Split(' ')[0]))
            }
        }
        if($trafficTo -like "*hour*"){
            [int]$time += [int]$([int]$($trafficTo.Split(' ')[0]) * 60) 
            [int]$time += [int]$($trafficTo.Split(' ')[2])
        } else {
            if($trafficTo) {
                [int]$time += [int]$([int]$($trafficTo.Split(' ')[0]))
            }
        }
        $minutesTo = "$time mins"
	# Now we set our departure time to be what our arrival time told us it should be so that we can access the traffic model api 
        $departureToTime = "departure_time=$(Get-Date -Date $($(Get-Date 9:00am).AddMinutes(0-[int]$($minutesTo.Split(' ')[0])).AddDays($toMonday)) -UFormat %s)"
        
        # final api call to destination 
        $params = @($origins,$destinations,$departureToTime,$units,$mode,$traffic_model,$apikey) -join "&"
        Write-Verbose "Making API call TO: $($base_url)?$($params)" 
        $apicall = Invoke-RestMethod "$($base_url)?$($params)"

        # times to destination
        $durationTo = $apicall | Select-Object -ExpandProperty rows | Select-Object -ExpandProperty elements | Select-Object -ExpandProperty duration | Select-Object -ExpandProperty text 
        $trafficTo = $apicall | Select-Object -ExpandProperty rows | Select-Object -ExpandProperty elements | Select-Object -ExpandProperty duration_in_traffic -ErrorAction SilentlyContinue | Select-Object -ExpandProperty text 
        Write-Verbose "Duration to is $durationTo and traffic to is $trafficTo"
        
        # format the commute times so that they are in minutes
        [int]$time = 0
        ## time to work
        if($durationTo -like "*hour*"){
            [int]$time = [int]$([int]$($durationTo.Split(' ')[0]) * 60) 
            [int]$time += [int]$($durationTo.Split(' ')[2])
        } else {
            if($durationTo){
                [int]$time += [int]$([int]$($durationTo.Split(' ')[0]))
            }
        }
        if($trafficTo -like "*hour*"){
            [int]$time += [int]$([int]$($trafficTo.Split(' ')[0]) * 60) 
            [int]$time += [int]$($trafficTo.Split(' ')[2])
        } else {
            if($trafficTo) {
                [int]$time += [int]$([int]$($trafficTo.Split(' ')[0]))
            }
        }
        $minutesTo = "$time mins"

        <# ############################# #>
	 #      Coming Home Times        # 
        <# ############################# #>

        # api call from destination
        $tmp = $origins
        $origins = "origins=$($destinations.Split("=")[1])"
        $destinations = "destinations=$($tmp.Split("=")[1])"
        Write-Verbose "We are going from $origins to $destinations"
        $params = @($origins,$destinations,$departure,$units,$mode,$traffic_model,$apikey) -join "&"
        Write-Verbose "Making API call FROM: $($base_url)?$($params)"
        $apicall = Invoke-RestMethod "$($base_url)?$($params)"

        # times from destination
        $durationFrom = $apicall | Select-Object -ExpandProperty rows | Select-Object -ExpandProperty elements | Select-Object -ExpandProperty duration | Select-Object -ExpandProperty text 
        $trafficFrom = $apicall | Select-Object -ExpandProperty rows | Select-Object -ExpandProperty elements | Select-Object -ExpandProperty duration_in_traffic -ErrorAction SilentlyContinue | Select-Object -ExpandProperty text 
        Write-Verbose "Duration from is $durationFrom and traffic from is $trafficFrom"

        # format the commute times so that they are in minutes
        [int]$time = 0
        ## time from work
        if($durationFrom -like "*hour*"){
            [int]$time = [int]$([int]$($durationFrom.Split(' ')[0]) * 60)
            [int]$time += [int]$($durationFrom.Split(' ')[2])
        } else {
            if($durationFrom){
                [int]$time += [int]$([int]$($durationFrom.Split(' ')[0]))
            }
        }    
        if($trafficFrom -like "*hour*"){
            [int]$time += [int]$([int]$($trafficFrom.Split(' ')[0]) * 60)
            [int]$time += [int]$($trafficFrom.Split(' ')[2])
        } else {
            if($trafficFrom){
                [int]$time += [int]$([int]$($trafficFrom.Split(' ')[0]))
            }
        }
        $minutesFrom = "$time mins"

        <# ############################# #>
	 #      Creating our object      #
        <# ############################# #>

        # Creating temporary object to be stored in arraylist
        $objectProperties = @{}
        $objectProperties.Add("City",$($($origins.Split("=")[1]).Replace('+',' ')))
        $objectProperties.Add("Arrive By 9am",$minutesTo)
        $objectProperties.Add("Depart House At",$($(Get-Date 9:00am).AddMinutes(0-[int]$($minutesTo.Split(' ')[0])).ToString("hh:mm")))
        $objectProperties.Add("Leave At 5pm",$minutesFrom)
        $objectProperties.Add("Get Home By",$($(Get-Date 5:00pm).AddMinutes([int]$($minutesFrom.Split(' ')[0])).ToString("hh:mm")))
        $object = New-Object -TypeName PSObject -Property $objectProperties
     	
	# Adding our object to our Arraylist
        $masterList.Add($object) | Out-Null

    }
	
    <# output #>
    $masterList | Format-Table -Property 'City','Arrive by 9am','Depart House At','Leave At 5pm','Get Home By' #| Sort-Object -Property City
					
}

