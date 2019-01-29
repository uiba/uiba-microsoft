'use strict';

/**
 * Define the base object namespace. By convention we use the service name
 * in PascalCase (aka UpperCamelCase). Note that this is defined as a package global (boilerplate).
 */
Microsoft = {};

/**
 * Request Microsoft credentials for the user (boilerplate).
 * Called from accounts-microsoft.
 *
 * @param {Object}    options                             Optional
 * @param {Function}  credentialRequestCompleteCallback   Callback function to call on completion. Takes one argument, credentialToken on success, or Error on error.
 */
Microsoft.requestCredential = function(options, credentialRequestCompleteCallback) {
  /**
   * Support both (options, callback) and (callback).
   */
  if (!credentialRequestCompleteCallback && typeof options === 'function') {
    credentialRequestCompleteCallback = options;
    options = {};
  } else if (!options) {
    options = {};
  }

  /**
   * Make sure we have a config object for subsequent use (boilerplate)
   */
  const config = ServiceConfiguration.configurations.findOne({
    service: 'microsoft'
  });
  if (!config) {
    credentialRequestCompleteCallback && credentialRequestCompleteCallback(
      new ServiceConfiguration.ConfigError()
    );
    return;
  }

  /**
   * Boilerplate
   */

  const credentialToken = Random.secret();
  const loginStyle = OAuth._loginStyle('microsoft', config, options);

  // Create nonce GUID https://docs.microsoft.com/en-us/azure/active-directory/develop/v1-protocols-openid-connect-code
  const nonce = Random.secret();

  /**
   * Microsoft requires response_type and client_id
   * We use state to roundtrip a random token to help protect against CSRF (boilerplate)
   */
    var redirectUri = Meteor.absoluteUrl() + "_oauth/microsoft";
    console.log("redirectUri", redirectUri);
    const loginUrl = 'https://login.microsoftonline.com/' +
      config.tenantID +
      '/oauth2/v2.0/authorize' +
      // '?response_type=id_token' +
      '?response_type=code' +
      '&client_id=' + config.clientID +
      // '&scope=' + encodeURIComponent(config.graphScopes) +
      '&scope=' + encodeURIComponent(config.graphScopes.join(' ') + " openid profile") +
      '&redirect_uri=' + encodeURIComponent(redirectUri) +
      '&nonce=' + encodeURIComponent(nonce) +
      '&state=' + encodeURIComponent(OAuth._stateParam(loginStyle, credentialToken));

  /**
   * Client initiates OAuth login request (boilerplate)
  */
  OAuth.launchLogin({
    loginService: 'microsoft',
    loginStyle: loginStyle,
    loginUrl: loginUrl,
    credentialRequestCompleteCallback: credentialRequestCompleteCallback,
    credentialToken: credentialToken,
    popupOptions: {
      height: 600
    }
  });
};
