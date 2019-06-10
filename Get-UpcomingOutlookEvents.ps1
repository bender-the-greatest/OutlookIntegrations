# start time and end time
$now = Get-Date
$startTime = "$($now.ToShortDateString()) $($now.ToShortTimeString())"
$endTime = ( Get-Date ).AddDays( 5 ).ToShortDateString()

# open outlook
$outlook = New-Object -ComObject Outlook.Application

# get calendar items
$calendarItems = $outlook.getNamespace( 'MAPI' ).GetDefaultFolder( [Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderCalendar ).Items

# sort and show recurring items
$calendarItems.Sort( '[Start]' )
$calendarItems.IncludeRecurrences = $true

# get items for the day
$upcomingEvents = $calendarItems.Restrict( "[Start] >= ""$startTime"" and [Start] <= ""$endTime""" ) | Select-Object Start, End, Subject

# build output string
$builder = New-Object System.Text.StringBuilder
$lastDay = $now.Day - 1

foreach( $event in $upcomingEvents ){
    # new day gets a date separator line
    if( ( $currentDate = ([DateTime]( $event.Start ) ) ).Day -ne $lastDay ) {
        $lastDay = $currentDate.Day
        if( $builder.Length -gt 0 ) {
            [void]$builder.AppendLine()
        }
        [void]$builder.AppendLine( $currentDate.ToString( 'D' ) )
        [void]$builder.AppendLine( '-' * $currentDate.ToString( 'D' ).Length )
    }

    [void]$builder.AppendLine( "* $($currentDate.ToShortTimeString()): $($event.Subject)")
}

$builder.ToString()
