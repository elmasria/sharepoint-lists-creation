# SharePoint Lists Creation Client Side

This Application help SharePoint users to create SharePoint lists. The Idea is that the user can move the required lists from any site collection or sub-site easily, by changing the **_SwebURL_** in the configuration file **_configuration.json_**.

## Run Application

Each Project or set of lists that user need to create should have their own **list-creation** folder. So user can return any time and run the application to create the configured lists.

1. Clone / download the project
	* git clone https://github.com/elmasria/sharesoint-lists-creation.git
2. List and Fields should be configured in the **_configuration.json_** file.

### Example

##### basic list creation

1. Create **configurations.json** file
2. add below json configurations

```json

 {
	"webURL": "http://WEB_APPLICATION_URL/",
	"username":"USERNAME",
	"password":"PASSWORD",
	"domain":"DOMAIN_NAME",
	"useStaticCredentials": false,
	"Lists": [
		{
      "Title":"IRMSNotificationSystem",
      "TemplateType":"100",
      "Description":"list that send notifications (email and mobile) for investors",
      "Fields": [
        {
          "Type":"Text",
          "Description" : "FIELD DESCRIPTION",
          "DisplayName":"NAME FOR LIST COLUMN",
          "Required":"False",
          "StaticName":"FIELD_NAME",
          "Name":"FIELD_NAME",
					"NotInDefaultView": true
        },{
          "Type":"Note",
          "NumLines": "6",
          "RichText":"FALSE",
          "Sortable":"FALSE",
          "Description" :"FIELD DESCRIPTION",
          "DisplayName":"",
          "Required":"False",
          "StaticName":"",
          "Name":""
        },{
          "Type":"Boolean",
          "Description" : "",
          "DisplayName":"",
          "Required":"False",
          "StaticName":"",
          "Name":""
        }
      ]
    }
	]
}


```

##### user need to create two lists

1. List A with Title "A"
	* Field "fa" type ``` Single Line of text ```
2. List A with Title "A"
	* Field "fb" type ``` Lookup ``` lookup to *fa* in list *A*

```json

 {
	"webURL": "http://WEB_APPLICATION_URL/",
	"username":"USERNAME",
	"password":"PASSWORD",
	"domain":"DOMAIN_NAME",
	"Lists": [
		{
			"Title":"A",
			"TemplateType":"100",
			"Description":"a",
			"Fields": [
				{
					"Type":"Text",
					"DisplayName":"fa",
					"Required":"False",
					"StaticName":"fa",
					"Name":"fa"
				},{
					"Type":"User",  // Or you may choose "UserMulti" for Select more than one user
					"UserSelectionMode": "PeopleOnly", // this attribute can be set to PeopleOnly or to PeopleAndGroups
					"Mult":"FALSE",
					"Description" : "abcd",
					"DisplayName":"Secretary",
					"Required":"False",
					"StaticName":"Secretary",
					"Name":"Secretary"
				}
			]
		},
		{
			"Title":"B",
			"TemplateType":"100",
			"Description":"b",
			"Fields": [
				{
					"Type":"Lookup",
					"DisplayName":"fb",
					"Required":"False",
					"StaticName":"fb",
					"Name":"Claim",
					"LookUpInfo": {
						"TargetList":"A",
						"targetField":"fa"
					}
				}
			]
		}
	]
}


```

1. **webURL** : The target site collection user need to create the lists in.
2. **Lists** : Array that contains list of Lists that should be created.
	1. **Title** : List name
	2. **TemplateType** : Currently sets to 100 (Custom List).  [SPListTemplateType enumeration](https://msdn.microsoft.com/en-us/library/microsoft.sharepoint.splisttemplatetype.aspx)
	3. **Description** : List description
	4. **Fields** : List of the fields that need to be added to that list
		1. **Type** : Site column types. [Site column types and options](https://support.office.com/en-us/article/Site-column-types-and-options-0d8ddb7b-7dc7-414d-a283-ee9dca891df7?ui=en-US&rs=en-US&ad=US)
		2. **DisplayName** : Name of the filed that will appear in the list view.
		3. **LookUpInfo** : This Key is optional, it is used when you need to add a lookup. User need to configure the
			* **TargetList** : List that list will read from
			* **targetField** : Field that list will display.


> Important: List that contains Lookup field should be after the list that will read from.


## TODO

* Add more Attributes for the field
* Solve calculated value issue when creating a field in another site collection
