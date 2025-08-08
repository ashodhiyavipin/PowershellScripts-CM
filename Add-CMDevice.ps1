# Define the collection name (replace 'YourCollectionName' with your actual collection name)
$CollectionName = "YourCollectionName"

# Read the hostnames from the text file
$Hostnames = Get-Content -Path "C:\Path\To\Your\Hostnames.txt"

# Get the collection object
$Collection = Get-CMDeviceCollection | Where-Object { $_.Name -eq $CollectionName }

# Check if collection exists
if ($Collection -eq $null) {
    Write-Host "Collection '$CollectionName' not found!" -ForegroundColor Red
    exit
}

# Loop through each hostname and add it to the collection
foreach ($Hostname in $Hostnames) {
    try {
        # Add the device to the collection
        Add-CMDeviceCollectionDirectMembershipRule -CollectionId $Collection.CollectionID -ResourceId (Get-CMDevice -Name $Hostname).ResourceId
        Write-Host "Added $Hostname to $CollectionName" -ForegroundColor Green
    } catch {
        Write-Host "Error adding $Hostname: $_" -ForegroundColor Yellow
    }
}