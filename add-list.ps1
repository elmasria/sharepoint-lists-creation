# Import required librearies for Sharepoint Client
Add-Type -Path "libraries\Microsoft.SharePoint.Client.dll"
Add-Type -Path "libraries\Microsoft.SharePoint.Client.Runtime.dll"

$Logfile = "listCreators.log"

function LogWrite([string]$logstring) {


	Add-content $Logfile -value $logstring
}

function New-Context([String]$WebUrl) {
	$context = New-Object Microsoft.SharePoint.Client.ClientContext($WebUrl)
	$context.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
	$context
}

function CreateCustomList([Microsoft.SharePoint.Client.ClientContext]$Context,  
	[String]$listDescription, 
	[String]$listName, 
	[String]$TemplateType = "100") {    
	# Create required list
	$ListInfo = New-Object Microsoft.SharePoint.Client.ListCreationInformation
	$ListItemInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
	$ListInfo.Title = $listName
	$ListInfo.TemplateType = $TemplateType
	$List = $context.Web.Lists.Add($ListInfo)
	$List.Description = $listDescription
	$List.Update()
	$context.ExecuteQuery() 
	$List
}  

function AddField([Microsoft.SharePoint.Client.ClientContext]$targetContext, 
	[Microsoft.SharePoint.Client.List]$targetList,
	[String]$targetFieldXML,
	[String]$targetOption){
	#create required field
	$targetList.Fields.AddFieldAsXml($targetFieldXML, $true, $targetOption)
	$targetList.Update()
	$targetContext.ExecuteQuery()
}

$castToMethodGeneric = [Microsoft.SharePoint.Client.ClientContext].GetMethod("CastTo")
$castToMethodLookup = $castToMethodGeneric.MakeGenericMethod([Microsoft.SharePoint.Client.FieldLookup])

$configurations = Get-Content 'configurations.json' | Out-String | ConvertFrom-Json
LogWrite "Json has been Loadded successfully."


$siteCollectionUrl= $configurations.webURL
$context = New-Context -WebUrl $siteCollectionUrl
LogWrite "Site collection has been Loadded successfully. - $($siteCollectionUrl)"

foreach ($element in $configurations.Lists) {
	# Create list
	$createdList = CreateCustomList -Context $context -listDescription $element.Description -listName  $element.Title
	
	LogWrite "List was created successfully. List Name:  $($createdList.Title)"	

	# Add Fields
	foreach ($field in $element.Fields) {
		LogWrite  "Start Adding  $($field.DisplayName)  to:  $($createdList.Title)"

		$option = [Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldToDefaultView
		[string] $FieldXML = "<Field DisplayName=`"$($field.DisplayName)`" Type=`"$($field.Type)`" />"

		# Load the current list
		$currentList = $context.Web.Lists.GetByTitle($createdList.Title)
		$context.Load($currentList)
		$context.ExecuteQuery() 
		LogWrite  "created List was loadded successfully. List Name:  $($createdList.Title)"

		if ($field.LookUpInfo) {
			LogWrite  "$($field.DisplayName) is a $($field.Type) Field"

			#Load the list that we need to get the lookup from it
			$targetList = $context.Web.Lists.GetByTitle($field.LookUpInfo.TargetList)
			$context.Load($targetList)
			LogWrite "target List was loadded successfully. List Name: $($targetList.Title)"

			# Prepare the Lookup Field			
			$newLookupField = $currentList.Fields.AddFieldAsXml($FieldXML, $true, $option)
			$context.Load($newLookupField)

			$lookupField = $castToMethodLookup.Invoke($context, $newLookupField)

			$lookupField.Title = $field.DisplayName
			$lookupField.LookupList = $targetList.Id
			$lookupField.LookupField = $field.LookUpInfo.targetField
			$lookupField.Update()
			$context.ExecuteQuery()
			LogWrite  "$($field.DisplayName) was created successfully."

			}else{			
				LogWrite  "$($field.DisplayName) is a $($field.Type) Field"
				AddField -targetContext $context -targetList $currentList -targetFieldXML $FieldXML -targetOption $option
				LogWrite  "$($field.DisplayName) was created successfully."
			}
		}
	}

	$context.Dispose()
	LogWrite "Completed successfully."