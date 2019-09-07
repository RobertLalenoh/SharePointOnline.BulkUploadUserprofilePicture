# Office 365 SharePoint Online | Bulk upload userprofile pictures

With this solution you can do an bulkupload from userprofile pictures in the SharePoint Online User profile database.

Office365 not automatically does an initial synchronization from user profile pictures in the User profile database  as it does with Delve and Outlook.  This manual and solution helps you to do an initial load yourself.

**1. Download  the profile pictures from Outlook Online. **

  Use [this powershell script ](https://drive.google.com/file/d/1mzGZdV_xXQrvJ7iX_WXns688R8nK0WZM/view?usp=sharing "this") to download all the  thumbnails from Outlook. 

1. Check if the paths in the script at $PictureStorageDir and $ExportLocation exists.
2. Make sure you execute this script with admin permissions otherwise you are only able to download your own thumbnail.

The script wil download all the thumbnails from Outlook an also wil create an CSV file with thumbnail location and user info.

The content of the CSV file wil look like this:

		UserPrincipalName,SourceUrl
		testuser@company.com,d:\PictureUploader\Pictures\john.doe@company.nl.jpg

**2. Configure the console app**

Compile the Github version or download an already compiled one [here](https://drive.google.com/open?id=1yjc8b4bZRvIPZclLf0kgsmNeVozZ0_Rq "here").

All the settings are set in the ConsoleApp.exe.config. 

		setting name="Tenant_admin_url" serializeAs="String">
                <value>https://tenantname-admin.sharepoint.com</value>
            </setting>
            <setting name="Tenant_admin_username" serializeAs="String">
                <value>admin@tenantname.onmicrosoft.com</value>
            </setting>
            <setting name="Tenant_admin_password" serializeAs="String">
                <value />
            </setting>
            <setting name="Profile_prefix" serializeAs="String">
                <value>i:0#.f|membership|</value>
            </setting>
            <setting name="Small_thumbwidth" serializeAs="String">
                <value>48</value>
            </setting>
            <setting name="Medium_thumbwidth" serializeAs="String">
                <value>72</value>
            </setting>
            <setting name="Large_thumbwidth" serializeAs="String">
                <value>200</value>
            </setting>
            <setting name="Import_csv_location" serializeAs="String">
                <value>userlist.csv</value>
            </setting>
            <setting name="Sleep_time_millisecond" serializeAs="String">
                <value>5000</value>
            </setting>
            <setting name="Mysite_url" serializeAs="String">
                <value>https://tenantname-my.sharepoint.com</value>
            </setting>
            <setting name="Relative_path_userprofile_picture_library" serializeAs="String">
                <value>/User%20Photos/Profielafbeeldingen</value>
            </setting>


1.  For every setting that uses tenantname change this to the name of your own tenant. 
2.  Change the import_csv_location to the location from the CSV file you created.
3.  Run the ConsoleApp.exe.

It can take a while befor all the user profile pictures are loaded into the User profile database. 

If you made an employee directory in the search sitecollection, you will noticed that the profile pictures are not shown immediately. This is because the search enginge has not crawled the user profle databse yet. Be patience, they will be showned after a couple of hours. 






  
















