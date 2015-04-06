AzureAd.resources.office365 = {};
AzureAd.resources.office365.friendlyName = "office365";
AzureAd.resources.office365.resourceUri = "https://outlook.office365.com/";

if (Meteor.isServer){
    Meteor.startup(function(){
        AzureAd.resources.registerResource(AzureAd.resources.office365.friendlyName, AzureAd.resources.office365.resourceUri);
    });
}
