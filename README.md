#azure-resource-office-365
Enhances the [azure-active-directory](https://github.com/djluck/azure-active-directory) & [accounts-azure-active-directory](https://github.com/djluck/accounts-azure-active-directory) packages, allowing users to authorize themselves against the office 365 rest APIs.

### Configuration
You will need to assign your Azure AD application the necessary permissions. [This guide](https://msdn.microsoft.com/en-us/office/office365/howto/add-common-consent-manually) describes the process pretty well.

### Example
The following demonstrates how to register a Meteor method called "getEvents" that would call the office 365 rest api to fetch the current user's upcoming events.
You will need to assign the Calendar.Read permission to your Azure AD application for this to work.

    Meteor.methods({
        "getEvents": function () {
            var url = "https://outlook.office365.com/api/v1.0/me/events";
            var user = Meteor.users.findOne(this.userId);
            var accessToken = AzureAd.resources.getOrUpdateUserAccessToken(AzureAd.resources.office365.friendlyName, user);
            return AzureAd.http.callAuthenticated("GET", url, accessToken);
        }
    });

The method could then be called on the client side, like this:

    Meteor.call("getEvents", function(err, calendars){
        console.log(calendars);
    });
