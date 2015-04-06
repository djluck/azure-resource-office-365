Package.describe({
    name: 'wiseguyeh:azure-resource-office-365',
    summary: "Allows users to authorize themselves against the office 365 rest APIs",
    version: "0.1.0",
    git: "https://github.com/djluck/azure-resource-office-365"
});

Package.onUse(function(api) {
  api.versionsFrom('1.1');
  api.use('wiseguyeh:azure-active-directory@0.3.0', ['server']);
  api.addFiles('office365Resource.js', 'server');
});
