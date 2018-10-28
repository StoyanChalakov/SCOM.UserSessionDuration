param(
 [Parameter(mandatory=$true)][string]$DomainSuffix,
 [Parameter(mandatory=$true)][string]$Username
 )

#Preparing all the xml queries for the different scenarios

#Query Session Logon events (Connected) for a particular user (Suffix\SamAccountName)
$ConnectedUserSessionsQuery = @"
<QueryList>
  <Query Id="0" Path="Operations Manager">
    <Select Path="Operations Manager">*[System[(EventID=26328)]]
    and  *[EventData[Data='$DomainSuffix\$Username']]</Select>
  </Query>
</QueryList>
"@ 

#Query Session LogOff events (Disconnected) for a particular user (Suffix\SamAccountName)
$DisconnectedUserSessionsQuery = @"
<QueryList>
  <Query Id="0" Path="Operations Manager">
    <Select Path="Operations Manager">*[System[(EventID=26329)]]
    and  *[EventData[Data='$DomainSuffix\$Username']]</Select>
  </Query>
</QueryList>
"@ 

try
{
    #Get connected session events for the user
    $ConnectedUserSessionEvents = Get-WinEvent -FilterXml $ConnectedUserSessionsQuery -ErrorAction Stop
}
catch
{
    $ErrorMessage = $_.Exception.Message
    $ErrorMessageForDisplay = $ErrorMessage + "Either the DOMAIN\UserName combination is incorrect or the events for this user have been overwritten."
    Write-Host -ForegroundColor White -BackgroundColor Red $ErrorMessageForDisplay
}

try
{
    #Get disconnected session events for the user
    $DisconnectedUserSessionEvents = Get-WinEvent -FilterXml $DisconnectedUserSessionsQuery -ErrorAction Stop
}
catch
{
    $ErrorMessage = $_.Exception.Message
    $ErrorMessageForDisplay = $ErrorMessage + "Either the DOMAIN\UserName combination is incorrect or the events for this user have been overwritten."
    Write-Host -ForegroundColor White -BackgroundColor Red $ErrorMessageForDisplay
}

#Initialize an array for storing the disconnected session id. Used for comparisson
$DisconnectedSessionsArray = @()

foreach ($DisconnetcedEvent in $DisconnectedUserSessionevents)
{
    #Get the event message
    $DisconnectedMessage = $DisconnetcedEvent.Message

    #Parse the Session Id from the message
    $Dispos = $DisconnectedMessage.IndexOf(";")
    $DisconnectedSessionID = $DisconnectedMessage.Substring($Dispos +1)

    $DisconnectedSessionsArray += $DisconnectedSessionID        
}


#iterate through all the connections and check for an disconnect event
foreach ($ConnectedEvent in $ConnectedUserSessionevents)
{
    #Get the event message
    $ConnectedMessage = $ConnectedEvent.Message

    #Parse the Session Id from the message
    $pos = $ConnectedMessage.IndexOf(";")
    $ConnectedSessionID = $ConnectedMessage.Substring($pos +1)    
    
    #if there is a an disconnected event for the same session id, then the session duration can be claculated
    if ($DisconnectedSessionsArray.Contains($ConnectedSessionID))
    {            
        #Get the TimeCreated property of the respective Disconnected event
        $DisconnectedEventTimeCreated =  ($DisconnectedUserSessionevents | where {$_.Message -match "$ConnectedSessionID"}).TimeCreated    
        $ConnectedEventTimeCreated = ($ConnectedUserSessionevents | where {$_.Message -match "$ConnectedSessionID"}).TimeCreated

        #Claculate and formnat the sesion duration
        $SessionDuration = $DisconnectedEventTimeCreated - $ConnectedEventTimeCreated
        $SessionDurationFormatted = $SessionDuration.ToString("hh':'mm':'ss")
        Write-Host "--------------------------------------------------------"
        Write-Host -ForegroundColor Green "Completed session:"
        Write-Host "Session established on $ConnectedEventTimeCreated and disconnected on $DisconnectedEventTimeCreated"
        Write-Host "The duration of the session was $SessionDurationFormatted"    
    }    
 }   
 
if (($ConnectedUserSessionEvents) -and ($DisconnectedUserSessionEvents))
{   
    #Check for stanalone sessions
    #Side Indicator '<=' Opened Sessions with no session disconnected event
    #Side Indicator '=>' Disconnected Sessions with no session connection event
    $StandaloneSesion =  Compare-Object $ConnectedUserSessionEvents $DisconnectedUserSessionEvents
    $StandaloneConnected = ($StandaloneSesion |  Where-Object {$_.SideIndicator -eq '<='}).InputObject
    $StandaloneDisconnected = ($StandaloneSesion |  Where-Object {$_.SideIndicator -eq '=>'}).InputObject
}    

    #if there isn't a respective disconnected event for the same session id, then user session has not been treminated yet (Could be console session or a PowerShell session, you cannot differentiate between those two)
    if ($StandaloneConnected)
    {
    
       Write-Host -ForegroundColor Cyan "Currently Open user sessions"

      foreach ($StandloneConnSession in $StandaloneConnected)
      {
      
        #Get the event message
        $StandloneConnSessionMessage = $StandloneConnSession.Message

        #Parse the Session Id from the message
        $Pos = $StandloneConnSessionMessage.IndexOf(";")
        $StandloneConnSessionID = $StandloneConnSessionMessage.Substring($Pos +1)
        $StandloneConnSessionIDCreated = $StandloneConnSession.TimeCreated
        
        Write-Host "--------------------------------------------------------"
        Write-Host -ForegroundColor Yellow "Open Session:"
        Write-Host "The user '$DomainSuffix\$Username' has an open session with the ID '$StandloneConnSessionID' started on '$StandloneConnSessionIDCreated'"
      }
    }
   
    #if there is a disconnected event for which there is no connected event, this means that the connected event has been overriden in the logs and we are not able to calculate the session duration.
    if ($StandaloneDisconnected)
    {
       Write-Host -ForegroundColor Yellow "Disconnected user sessions with no duration information (session open event has been overwritten)"
       
       foreach ($StandloneDisConnSession in $StandaloneDisconnected)
       {      
        #Get the event message
        $StandloneDisConnSessionMessage = $StandloneDisConnSession.Message

        #Parse the Session Id from the message
        $DisPos = $StandloneDisConnSessionMessage.IndexOf(";")
        $StandloneDisConnSessionID = $StandloneDisConnSessionMessage.Substring($DisPos +1)
        $StandloneDisConnSessionIDCreated = $StandloneDisConnSession.TimeCreated
        
        Write-Host "The user '$DomainSuffix\$Username' has an open session with the ID '$StandloneDisConnSessionID' started on '$StandloneDisConnSessionIDCreated'"
        }
    }