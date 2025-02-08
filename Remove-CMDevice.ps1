# Path to the text file containing machine names
$machineListPath = "C:\path\to\your\machines.txt" # Replace with the actual path to your .txt file

# Read the machine names from the text file
$machines = Get-Content -Path $machineListPath

# Loop through each machine name and remove it from SCCM
foreach ($machine in $machines) {
    # Call the Remove-CMDevice cmdlet for each machine
    Remove-CMDevice -DeviceName $machine
    Write-Host "Removed device: $machine"
}