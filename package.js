Package.describe({
  name: 'uibalabs:microsoft',
  version: '1.0.0',
  // Brief, one-line summary of the package.
  summary: 'An implementation of the Microsoft OAuth 2.0 login service.',
  // URL to the Git repository containing the source code for this package.
  git: 'https://github.com/uiba/uiba-microsoft.git',
  // By default, Meteor will default to using README.md for documentation.
  // To avoid submitting documentation, set this field to null.
  documentation: null
});

Package.onUse(function(api) {
  api.versionsFrom('1.2.1');
  api.use('ecmascript');

  api.use('accounts-ui', ['client', 'server']);
  api.use('oauth2', ['client', 'server']);
  api.use('oauth', ['client', 'server']);
  api.use('http', ['server']);
  api.use(['underscore', 'service-configuration'], ['client', 'server']);
  api.use(['random', 'templating'], 'client');
  Npm.depends({
    "atob": "2.1.2"
  });

  api.export('Microsoft');

  api.addFiles(
    ['microsoft_configure.html', 'microsoft_configure.js'],
    'client');

  api.addFiles('microsoft_server.js', 'server');
  api.addFiles('microsoft_client.js', 'client');
});

Package.onTest(function(api) {
  api.use('ecmascript');
  api.use('tinytest');
  api.use('uibalabs:microsoft@1.0.0');
  api.mainModule('microsoft-tests.js');
});
