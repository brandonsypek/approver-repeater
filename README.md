# approver-repeater
Custom Nintex plugin that allows the user to select an unlimited amount of Approvers, change the order of approvers, and remove rows and users.
The plugin usees Azure AD App Client ID and your Tenand ID to gain access to the users to work like an OOTB People Picker.
This custom plugin saves data in json and uses a multi line text field and its assosiated ID to save to SharePoint to get called back on in Nintex Workflow.
Use the Force Editable Mode control within the 'Approvers Repeater' and set where you want the control to be "editable" eg. [Form mode].[Is New mode]
