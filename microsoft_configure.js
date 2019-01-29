Template.configureLoginServiceDialogForMicrosoft.helpers({
  siteUrl: function () {
    return Meteor.absoluteUrl();
  }
});

Template.configureLoginServiceDialogForMicrosoft.fields = function () {
  return [
    { property: 'clientID', label: 'Client Id (App Id)' },
    { property: 'tenantID', label: 'Tenant Id' },
    { property : 'secret', label: 'App Secret' }
  ];
};
