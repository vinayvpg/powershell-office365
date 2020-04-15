$csvFile = "C:\Users\prabhvx\OneDrive - Murphy Oil\Desktop\mocconnect\reports\O365 Migration Prep\SW to VP email mapping.csv"

#Create a table based on the csv
$table = Import-CSV $csvFile -Delimiter ","

#Declaration of the mapping settings:
$mappingSettings = New-MappingSettings

#Cycle through each row of the CSV
foreach ($row in $table)
{
    #Add the current row source user and destination user to the mapping list
    Set-UserAndGroupMapping -MappingSettings $mappingSettings -Source $row.SourceValue -Destination $row.DestinationValue 
}

#The user and group mappings are exported to C:\FolderName\FileName.sgum
Export-UserAndGroupMapping -MappingSettings $mappingSettings -Path "C:\Users\prabhvx\OneDrive - Murphy Oil\Desktop\mocconnect\reports\O365 Migration Prep\SWToVPUserMappings.sgum"