'use strict';

/**
 * Define the base object namespace. By convention we use the service name
 * in PascalCase (aka UpperCamelCase). Note that this is defined as a package global.
 */
Microsoft = {};

/*
  Boilerplate hook for use by underlying Meteor code
 */
Microsoft.retrieveCredential = (credentialToken, credentialSecret) => {
  return OAuth.retrieveCredential(credentialToken, credentialSecret);
};

/**
 * Define the fields we want.
 * Note that we *must* have an id. Also, this array is referenced in the
 * accounts-microsoft package, so we should probably keep this name and structure.
 */
Microsoft.whitelistedFields = ['id', 'givenName', 'surname', 'displayName',
  'mail', 'idToken', 'businessPhones', 'jobTitle', 'mobilePhone',
  'officeLocation', 'preferredLanguage', 'userPrincipalName'];

/**
 * Register this service with the underlying OAuth handler
 * (name, oauthVersion, urls, handleOauthRequest):
 *  name = 'microsoft'
 *  oauthVersion = 2
 *  urls = null for OAuth 2
 *  handleOauthRequest = function(query) returns {serviceData, options} where options is optional
 * serviceData will end up in the user's services.microsoft
 */
OAuth.registerService('microsoft', 2, null, function(query) {

  /**
   * Make sure we have a config object for subsequent use (boilerplate)
   */
  const config = ServiceConfiguration.configurations.findOne({
    service: 'microsoft'
  });
  if (!config) {
    throw new ServiceConfiguration.ConfigError();
  }

  /**
   * Get the token and username (Meteor handles the underlying authorization flow).
   * Note that the username comes from from this request in Microsoft.
   */
  // console.log("microsoft_server.js config", config);

  const response = getTokens(config, query);
  const accessToken = response.accessToken;
  const refreshToken = response.refreshToken;
  const username = response.username;
  // console.log("response", response);

  /**
   * If we got here, we can now request data from the account endpoints
   * to complete our serviceData request.
   * The identity object will contain the username plus *all* properties
   * retrieved from the account and settings methods.
  */
  // const identity = _.extend(
  //   {username},
  //   getAccount(config, accessToken)
  //   // getSettings(config, username, accessToken)
  // );
  // const identity = {};
  const identity = getAccount(config, accessToken);

  /**
   * Build our serviceData object. This needs to contain
   *  accessToken
   *  expiresAt, as a ms epochtime
   *  refreshToken, if there is one
   *  id - note that there *must* be an id property for Meteor to work with
   *  mail
   *  username
   */
  const serviceData = {
    accessToken,
    expiresAt: (+new Date) + (1000 * response.expiresIn)
  };
  if (response.refreshToken) {
    serviceData.refreshToken = response.refreshToken;
  }
  _.extend(serviceData, _.pick(identity, Microsoft.whitelistedFields));

  /**
   * Return the serviceData object along with an options object containing
   * the initial profile object with the username.
   */

  return {
    serviceData: serviceData,
    options: {
      profile: {
        name: response.username // comes from the token request
      }
    }
  };
});

/**
 * The following three utility functions are called in the above code to get
 *  the access_token, refresh_token and username (getTokens)
 *  account data (getAccount)
 *  settings data (getSettings)
 * repectively.
 */

/** getTokens exchanges a code for a token in line with Microsoft's documentation
 *
 *  returns an object containing:
 *   accessToken        {String}
 *   expiresIn          {Integer}   Lifetime of token in seconds
 *   refreshToken       {String}    If this is the first authorization request
 *   account_username   {String}    User name of the current user
 *   token_type         {String}    Set to 'Bearer'
 *
 * @param   {Object} config       The OAuth configuration object
 * @param   {Object} query        The OAuth query object
 * @return  {Object}              The response from the token request (see above)
 */
const getTokens = function(config, query) {

  const endpoint = `https://login.microsoftonline.com/${config.tenantID}/oauth2/token`

  /**
   * Attempt the exchange of code for token
   */
  const redirectUri = Meteor.absoluteUrl() + '_oauth/microsoft';
  let response;
  try {
    response = HTTP.post(
      endpoint, {
        params: {
          code: query.code,
          client_id: config.clientID,
          client_secret: OAuth.openSecret(config.secret),
          grant_type: 'authorization_code',
          redirect_uri: redirectUri,
          resource: 'https://graph.microsoft.com/'
        }
      });
  } catch (err) {
    throw _.extend(new Error(`Failed to complete OAuth handshake with Microsoft. ${err.message}`), {
      response: err.response
    });
  }

  if (response.data.error) {

    /**
     * The http response was a json object with an error attribute
     */
    throw new Error(`Failed to complete OAuth handshake with Microsoft. ${response.data.error}`);

  } else {

    /** The exchange worked. We have an object containing
     *   access_token
     *   refresh_token
     *   expires_in
     *   token_type
     *   account_username
     *
     * Return an appropriately constructed object
     */
    return {
      accessToken: response.data.access_token,
      refreshToken: response.data.refresh_token,
      expiresIn: response.data.expires_in,
      username: response.data.account_username
    };
  }
};

/**
 * getAccount gets the basic Microsoft account data
 *
 *  returns an object containing:
 *   id                {Integer}         The user's Microsoft id
 *   displayName       {String}          The account username
 *   givenName         {String}
 *   surname           {String}          A basic description the user has filled out
 *   created           {Integer}         The epoch time of account creation
 *   pro_expiration    {Integer/Boolean} False if not a pro user, their expiration date if they are.
 *   mail              {String}          The user's email address used on the account.
 *   businessPhones
 *   jobTitle          {String}
 *   mobilePhone       {String}
 *   officeLocation    {String}
 *   preferredLanguage {String}
 *   userPrincipalName {String}
 *   accessToken       {String}
 *   refreshToken      {String}
 *   expiresAt         {Integer}
 *
 * @param   {Object} config       The OAuth configuration object
 * @param   {String} accessToken  The OAuth access token
 * @return  {Object}              The response from the account request (see above)
 */
const getAccount = function(config, accessToken) {

  // const endpoint = `https://api.microsoft.com/3/account/${username}`;
  const endpoint = "https://graph.microsoft.com/v1.0/me" // Graph Endpoint
  let accountObject;

  /**
   * Note the strange .data.data - the HTTP.get returns the object in the response's data
   * property. Also, Microsoft returns the data we want in a data property of the response data
   * Hence (response).data.data
   */

  try {
   //  var options = {}
   //  var headers = {
   //    Authorization : 'Bearer ' + accessToken
   //  }
   //
   // options.headers = headers;
   //  accountObject = HTTP.call("GET", endpoint, options).data;
    accountObject = HTTP.get(
      endpoint, {
        headers: {
          Authorization: `Bearer ${accessToken}`
        }
      }
    ).data;
    return accountObject;

  } catch (err) {
    throw _.extend(new Error(`Failed to fetch account data from Microsoft. ${err.message}`), {
      response: err.response
    });
  }
};


/**
 * getSettings gets the basic Microsoft account/settings data
 *
 *  returns an object containing:
 *
 * @param   {Object} config       The OAuth configuration object
 * @param   {String} username     The Microsoft username
 * @param   {String} accessToken  The OAuth access token
 * @return  {Object}              The response from the account request (see above)
 */
const getSettings = function(config, username, accessToken) {

  const endpoint = `https://api.microsoft.com/3/account/${username}/settings`;
  let settingsObject;

  /**
   * Note the strange .data.data - the HTTP.get returns the object in the response's data
   * property. Also, Microsoft returns the data we want in a data property of the response data
   * Hence (response).data.data
   */
  try {
    settingsObject = HTTP.get(
      endpoint, {
        headers: {
          Authorization: `Bearer ${accessToken}`
        }
      }
    ).data.data;
    return settingsObject;

  } catch (err) {
    throw _.extend(new Error(`Failed to fetch settings data from Microsoft. ${err.message}`), {
      response: err.response
    });
  }
};
