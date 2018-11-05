# Import required librearies for Sharepoint Client
Add-Type -Path "libraries\Microsoft.SharePoint.Client.dll"
Add-Type -Path "libraries\Microsoft.SharePoint.Client.Runtime.dll"

$Logfile = "listCreators.log"

function LogWrite([string]$logstring, [string]$logstringColor = "Green") {
	Add-content $Logfile -value $logstring
	Write-Host $logstring -ForegroundColor $logstringColor
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

function HandleMixedModeWebApplication(){
	  param([Parameter(Mandatory=$true)][object]$clientContext)
	  Add-Type -TypeDefinition @"
	  using System;
	  using Microsoft.SharePoint.Client;

	  namespace Toth.SPOHelpers
	  {
	      public static class ClientContextHelper
	      {
	          public static void AddRequestHandler(ClientContext context)
	          {
	              context.ExecutingWebRequest += new EventHandler<WebRequestEventArgs>(RequestHandler);
	          }

	          private static void RequestHandler(object sender, WebRequestEventArgs e)
	          {
	              //Add the header that tells SharePoint to use Windows authentication.
	              e.WebRequestExecutor.RequestHeaders.Remove("X-FORMS_BASED_AUTH_ACCEPTED");
	              e.WebRequestExecutor.RequestHeaders.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
	          }
	      }
	  }
"@ -ReferencedAssemblies "libraries\Microsoft.SharePoint.Client.dll", "libraries\Microsoft.SharePoint.Client.Runtime.dll";
	  [Toth.SPOHelpers.ClientContextHelper]::AddRequestHandler($clientContext);
}

Try {

	$castToMethodGeneric = [Microsoft.SharePoint.Client.ClientContext].GetMethod("CastTo")
	$castToMethodLookup = $castToMethodGeneric.MakeGenericMethod([Microsoft.SharePoint.Client.FieldLookup])
	$castToMethodCalculated = $castToMethodGeneric.MakeGenericMethod([Microsoft.SharePoint.Client.FieldCalculated])
	$castToMethodUser = $castToMethodGeneric.MakeGenericMethod([Microsoft.SharePoint.Client.FieldUser])

	$configurations = Get-Content 'configurations.json' | Out-String | ConvertFrom-Json
	LogWrite "Json has been Loadded successfully."  -logstringColor  Green

	$siteCollectionUrl= $configurations.webURL
	$context = New-Context -WebUrl $siteCollectionUrl
	$isNotValidCredentials = $true

	if ($configurations.useStaticCredentials) {
		$context.Credentials = New-Object System.Net.NetworkCredential($configurations.username, $configurations.password, $configurations.domain)
		$isNotValidCredentials = $false
	}else{
		# Check the validity of credential and repeat the proccess
		# if they are not valid
		while($isNotValidCredentials -eq $true){
			try{
				LogWrite "Input your credentials:" -logstringColor  Green
				$credentials = Get-Credential
				$context.Credentials = $credentials.GetNetworkCredential()

				# Check if the credentials is valid
				$web = $context.Web
				$context.Load($web)
				$context.ExecuteQuery()

				# Valid Credential
				$isNotValidCredentials = $false

			}catch{
				$ErrorMessage = $_.Exception.Message
				LogWrite  "$($ErrorMessage)" -logstringColor  Red
				if ($ErrorMessage -like "*(401) Unauthorized*") {
					# not valid credentials
					LogWrite  "Wrong Credential" -logstringColor Red
				}else {
					# Error other Credential validity
				    $isNotValidCredentials = $false
				}
			}
		}
	}

	if ($configurations.MixedAuthenticationMode) {
		$context.AuthenticationMode = [Microsoft.SharePoint.Client.ClientAuthenticationMode]::Default
		HandleMixedModeWebApplication $context
		$isNotValidCredentials = $false
	}

	LogWrite "Connection to Site collection has been done successfully. - $($siteCollectionUrl)" -logstringColor Green

	foreach ($element in $configurations.Lists) {
		# Create list
		$createdList = ""
		$ListExists = $false
		Try
		{
			$createdList = $context.Web.Lists.GetByTitle($element.Title)
			$context.Load($createdList)
			$context.ExecuteQuery()
			$ListExists = $true
			LogWrite "List was Loadded successfully. List Name:  $($createdList.Title)" -logstringColor Green
		}
		Catch
		{
				$createdList = CreateCustomList -Context $context -listDescription $element.Description -listName  $element.Title
				LogWrite "List was created successfully. List Name:  $($createdList.Title)" -logstringColor Green
		}

		# Add Fields
		foreach ($field in $element.Fields) {

			if ($ListExists)
			{
				Try
				{
						$context.Load($createdList.Fields)
						$context.ExecuteQuery()

						$Column = $createdList.Fields.GetByInternalNameOrTitle($field.Name)
						$Column.DeleteObject()
	    			$context.ExecuteQuery()

						LogWrite  "Deleting  $($field.DisplayName)  to:  $($createdList.Title)" -logstringColor Green
				}
				Catch
				{
					LogWrite  "Error Deleting  $($field.DisplayName)  to:  $($createdList.Title)" -logstringColor Red
				}
			}

			LogWrite  "Start Adding  $($field.DisplayName)  to:  $($createdList.Title)" -logstringColor Green

			$option = [Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldInternalNameHint

			if($field.Mult){
				$mult  = "Mult=`"$($field.Mult)`""
			}
			if($field.UserSelectionMode){
				$userSelectionMode=  "UserSelectionMode=`"$($field.UserSelectionMode)`""
			}
			if($field.NumLines){
				$numLines  = "NumLines=`"$($field.NumLines)`""
			}
			if($field.RichText){
				$richText  = "RichText=`"$($field.RichText)`""
			}
			if($field.Sortable){
				$sortable  = "Sortable=`"$($field.Sortable)`""
			}
			if($field.Min){
				$min  = "Min=`"$($field.Min)`""
			}
			if($field.Format){
				$format  = "Format=`"$($field.Format)`""
			}

			[string] $FieldXML = "<Field $($format) $($min) $($mult) $($sortable) $($RichText) $($description) $($numLines) $($userSelectionMode)  DisplayName=`"$($field.DisplayName)`"  StaticName=`"$($field.StaticName)`" Name=`"$($field.Name)`"   Type=`"$($field.Type)`" />"

			# Load the current list
			$currentList = $context.Web.Lists.GetByTitle($createdList.Title)
			$context.Load($currentList)
			$context.ExecuteQuery()
			LogWrite  "created List was loadded successfully. List Name:  $($createdList.Title)" -logstringColor Blue
			LogWrite  "$($field.DisplayName) is a $($field.Type) Field" -logstringColor Blue

			if ($field.LookUpInfo) {
				Try {
					#Load the list that we need to get the lookup from it
					$targetList = $context.Web.Lists.GetByTitle($field.LookUpInfo.TargetList)
					$context.Load($targetList)
					LogWrite "target List was loadded successfully. List Name: $($targetList.Title)" -logstringColor Blue

					# Prepare the Lookup Field
					$newLookupField = $currentList.Fields.AddFieldAsXml($FieldXML, $true, $option)
					$context.Load($newLookupField)

					$lookupField = $castToMethodLookup.Invoke($context, $newLookupField)

					$lookupField.Title = $field.DisplayName
					$lookupField.LookupList = $targetList.Id
					$lookupField.LookupField = $field.LookUpInfo.targetField
					$lookupField.Update()
					$context.ExecuteQuery()
					LogWrite  "$($field.DisplayName) was created successfully." -logstringColor Green

				}Catch {
					$ErrorMessage = $_.Exception.Message
					$FailedItem = $_.Exception.ItemName
					LogWrite  "Error $($ErrorMessage)." -logstringColor Red
					Read-Host "Press Enter to exit"
				}

			}elseif($field.CalculatedInfo){
				Try {

					# Prepare the Calculated Field
					$newCalculatedField = $currentList.Fields.AddFieldAsXml($FieldXML, $true, $option)
					$context.Load($newCalculatedField)

					$CalculatedField = $castToMethodCalculated.Invoke($context, $newCalculatedField)


					$CalculatedField.Title = $field.DisplayName
					$CalculatedField.Formula = "=UPPER([Title])"
					$CalculatedField.OutputType = $field.CalculatedInfo.OutputType
					$CalculatedField.Update()
					$context.ExecuteQuery()
					LogWrite  "$($field.DisplayName) was created successfully."

				}Catch {
					$ErrorMessage = $_.Exception.Message
					$FailedItem = $_.Exception.ItemName
					LogWrite "$($_.Exception| Out-String)"
					LogWrite  "Error: $($ErrorMessage)." -logstringColor Red
					Read-Host "Press Enter to exit"
				}
			}else{
				Try{
					AddField -targetContext $context -targetList $currentList -targetFieldXML $FieldXML -targetOption $option
					LogWrite  "$($field.DisplayName) was created successfully." -logstringColor Green
				}Catch {
					$ErrorMessage = $_.Exception.Message
					$FailedItem = $_.Exception.ItemName
					LogWrite  "Error: $($ErrorMessage)."-logstringColor Red
				}
			}


			[Microsoft.SharePoint.Client.View]$view = $currentList.Views.GetByTitle("All Items")

			if ($field.NotInDefaultView) {
				if($view -eq $Null) {
					LogWrite "View doesn't exists!"  -logstringColor Blue
				}else{
					$viewF = $view.ViewFields

					$viewF.Remove($field.StaticName)
					$view.Update()
					$context.ExecuteQuery()
				}
			}

		}
	}

	LogWrite "Completed successfully." -logstringColor Green

	$context.Dispose()

}Catch {
	$ErrorMessage = $_.Exception.Message
	$FailedItem = $_.Exception.ItemName
	LogWrite  "Error: $($ErrorMessage)." -logstringColor red
}

Read-Host "Press Enter to exit"
