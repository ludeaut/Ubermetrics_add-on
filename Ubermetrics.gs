/**
 * @OnlyCurrentDoc
 */

/**
 * Apps Script trigger. Runs when the add-on is installed.
 */
function onInstall(e) {
  /**
   * The document is already open, so after installation is complete
   * the ˙onOpen()` trigger must be called manually in order for the
   * add-on to execute.
   */
  onOpen(e);
}

/**
 * Apps Script trigger. Runs when an editable document is opened.
 */
function onOpen(e) {

  /**
   * Create the Google Sheets add-on menu item in the navigation bar, and have it
   * call `createSetupSheet_()` when clicked.
   */

  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem('Create a setup sheet', 'createSetupSheet_')
    .addToUi();
}

/**
 * Creates and configures a new Setup sheet if it doesn't exist
 * Documentation link only available for INTK's employees
 */
function createSetupSheet_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Setup UberMetrics');
  if (sheet == null) {
    SpreadsheetApp.getActiveSpreadsheet().insertSheet('Setup UberMetrics');
    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Setup UberMetrics');
    var values = [['Documentation'], ['UberAnalytics arguments'], ['Start date'], ['End date'], ['Metrics'], ['Dimensions'],['Sort'], ['Filters']
                  , ['Max results'], [''], ['UberAds arguments'], ['Manager Account ID'], ['Start date'], ['End date'], ['Fields'], ['Resource'],['Conditions'], ['Sort']
                  , ['Max results']];
    sheet.getRange(1,1, 19).setValues(values);
    sheet.getRange(1,1, 19).setFontWeight("bold");
    sheet.getRange(1,2).setValues([['https://app.asana.com/0/1131272848348008/1133839332226129/f']]);
    sheet.autoResizeColumns(1,2);
    sheet.setFrozenRows(1);
    sheet.getRange(2,1,1,2).merge();
    sheet.getRange(11,1,1,2).merge();
    sheet.getRange(2,1,1,2).getMergedRanges()[0].setHorizontalAlignment("center");
    sheet.getRange(11,1,1,2).getMergedRanges()[0].setHorizontalAlignment("center");
  }
}

/**
 * Fetch data from Google Analytics.
 * @param {string} viewId The Id of the starred view in the client Google Analytics
 * @param {string} startDate The start of the date range
 * @param {string} endDate The end of the date range
 * @param {string} metrics The metrics separated by commas
 * @param {string} dimensions The dimensions separated by commas
 * @param {string} sort The sorting order and the sorting direction example:"sessions_desc"
 * @param {string} filters The filters separated by commas (OR) or semicolons (AND), example:"country==France"
 * @param {integer} maxResults The maximum number of results to return
 * @return The value of the given metrics for the given arguments.
 * @customfunction
 */
function UberAnalytics(viewId, startDate, endDate, metrics, dimensions, sort, filters, maxResults){
  return UberAnalyticsV3Function_(viewId, startDate, endDate, metrics, dimensions, sort, filters, maxResults);
}

/**
 * Fetch data from Google Ads.
 * @param {string} managerAccountId The Id of the manager account in Google Ads, format: xxx-xxx-xxxx
 * @param {string} accountId The Id of the client account in Google Ads, format: xxx-xxx-xxxx
 * @param {string} startDate The start of the date range or a literal date range
 * @param {string} endDate The end of the date range or empty if a literal date range was previously given
 * @param {string} fields The fields (metrics or segments) separated by commas
 * @param {string} resource The resource where to fetch the data from
 * @param {string} conditions The conditions separated by commas
 * @param {string} sort The sorting order and the sorting direction example:"segments.date_desc"
 * @param {integer} maxResults The maximum number of results to return
 * @return The value of the given metrics for the given arguments.
 * @customfunction
 */
function UberAds(managerAccountId, accountId, startDate, endDate, fields, resource, conditions, sort, maxResults){
  return UberGoogleAds_(managerAccountId, accountId, startDate, endDate, fields, resource, conditions, sort, maxResults);
}

/**
 * Fetch data from Google Analytics using the Google Analytics API v3 by calling API function.
 * @param {string} viewId The Id of the starred view in the client Google Analytics
 * @param {string} startDate The start of the date range
 * @param {string} endDate The end of the date range
 * @param {string} metrics The metrics separated by commas
 * @param {string} dimensions The dimensions separated by commas
 * @param {string} sort The sorting order and the sorting direction example:"sessions_desc"
 * @param {string} filters The filters separated by commas (OR) or semicolons (AND), example:"country==France"
 * @param {integer} maxResults The maximum number of results to return
 * @return The value of the given metrics for the given arguments.
 * @customfunction
 */
function UberAnalyticsV3Function_(viewId, startDate, endDate, metrics, dimensions, sort, filters, maxResults) {
  /**
   * Handles dates given to add the right format to the API function.
   * First, gets the timezone of the script
   * Then, if the start date entered can be, it's converted to 'yyyy-MM-dd' format in the timezone found before.
   * If it can't, the start date entered is kept.
   * Then, does the same thing for the end date given.
   */
  var timeZone = Session.getScriptTimeZone();
  if(startDate != '' && new Date(startDate) != 'Invalid Date'){
    var startDateString = Utilities.formatDate(new Date(startDate), timeZone, 'yyyy-MM-dd');
  } else {
    var startDateString = startDate;
  }
  if (endDate != '' && new Date(endDate) != 'Invalid Date'){
    var endDateString = Utilities.formatDate(new Date(endDate), timeZone, 'yyyy-MM-dd');
  } else {
    var endDateString = endDate;
  }
  /**
   * Configures the two others required parameters: view ID and metric(s) to be readable for the API function.
   * For other parameters, first tests if a value was entered and if it's the case, add it to the options.
   */
  var ids = 'ga:' + viewId.toString().trim();
  var metrics = 'ga:' + metrics.replace(/,/g, ',ga:');
  var options = {};
  (dimensions != "" && dimensions != null) ? options['dimensions'] = 'ga:' + dimensions.replace(/,/g, ',ga:') : {};
  if (sort != "" && sort != null){
    sort = sort.split(',');
    for (var i = 0; i < sort.length; i++){
      var sortParams = sort[i].split('_');
      sort[i] = sortParams[1] == 'desc' ? '-' : '';
      sort[i] += 'ga:' + sortParams[0];
    }
    options['sort'] = sort.join(',');
  }
  (filters != "" && filters != null) ? options['filters'] = 'ga:' + filters.replace(/,/g, ',ga:').replace(/;/g, ';ga:') : {};
  (maxResults != "" && maxResults != null) ? options['max-results'] = maxResults.toString().replace(' ', '') : {};
  /**
   * Handles the access token refresh.
   * Tests if the script have an expires_at property and therefore an access_token property as they're created at the same time.
   * If not, that's the first time the add-on is used. If it has one but the property value is expired, the token is no longer valid
   * and needs to be refreshed.
   * In these both cases, an access token and an expiration date are created/refreshed and stored in the script properties.
   * In all other cases, nothing happens.
   */
  var expires_at = PropertiesService.getScriptProperties().getProperty('expires_at');
  if (expires_at == null || expires_at < new Date()){
    Logger.log('Access token has been refreshed');
    PropertiesService.getScriptProperties().setProperty('access_token', refreshAccessToken_());
    PropertiesService.getScriptProperties().setProperty('expires_at', new Date(new Date().getTime()+3600*1000));
  }
  /**
   * Parameter that handles the authentication process.
   *
   */
  options['access_token'] = PropertiesService.getScriptProperties().getProperty('access_token');
  /**
   * Parameter that handles the quota limits.
   *
   */
  options['quotaUser'] = parseFloat(Math.random());
  /**
   * Performs the request a first time
   * If an error occured, does different actions depending the error message
   */
  try {
    var json = Analytics.Data.Ga.get(ids, startDateString, endDateString, metrics, options);
    var data = JSON.parse(json);
  } catch (error) {
    var message = error.message.split(':').slice(1).join(':').trim();
    if(message == 'Invalid Credentials' || message == 'Login Required'){
      Logger.log('Token has been refreshed');
      PropertiesService.getScriptProperties().setProperty('access_token', refreshAccessToken_());
      PropertiesService.getScriptProperties().setProperty('expires_at', new Date(new Date().getTime()+3600*1000));
      options['access_token'] = PropertiesService.getScriptProperties().getProperty('access_token');
      options['quotaUser'] = parseFloat(Math.random());
      var json = Analytics.Data.Ga.get(ids, startDateString, endDateString, metrics, options);
      var data = JSON.parse(json);
    } else {
      return message;
    }
  }
  /**
   * Handles the two possible cases of the response received:
   * Values in an array,
   * No result for this request
   */
  if(data.rows){
    Logger.log(data.rows);
    return numericStringToNumber_(data.rows);
  }
  var noResultResponse = (dimensions != '' && dimensions != null) ? dimensions.split(',') : [];
  noResultResponse.length += metrics.split(',').length;
  noResultResponse.fill(0, noResultResponse.length - metrics.split(',').length);
  Logger.log([noResultResponse]);
  return [noResultResponse];
}



/**
 * Fetch data from Google Ads using the Google Ads API v2.
 * @param {string} managerAccountId The Id of the manager account in Google Ads, format: xxx-xxx-xxxx
 * @param {string} accountId The Id of the client account in Google Ads, format: xxx-xxx-xxxx
 * @param {string} startDate The start of the date range or a literal date range
 * @param {string} endDate The end of the date range or empty if a literal date range was previously given
 * @param {string} fields The metrics separated by commas
 * @param {string} resource The resource where to fetch the data from
 * @param {string} conditions The conditions separated by commas
 * @param {string} sort The sorting order and the sorting direction example:"segments.date_desc"
 * @param {integer} maxResults The maximum number of results to return
 * @return The value of the given metrics for the given arguments.
 * @customfunction
 */
function UberGoogleAds_(managerAccountId, accountId, startDate, endDate, fields, resource, conditions, sort, maxResults) {
  /**
   * Handles dates given to add the right format to the API function.
   * First, gets the timezone of the script
   * Then, if the start date entered can be, it's converted to 'yyyyMMdd' format in the timezone found before.
   * If it can't, the start date entered is kept.
   * Then, does the same thing for the end date given.
   */
  var timeZone = Session.getScriptTimeZone();
  if(startDate != '' && new Date(startDate) != 'Invalid Date'){
    var startDateString = Utilities.formatDate(new Date(startDate), timeZone, 'yyyy-MM-dd').split('-').join('');
  } else {
    var startDateString = startDate;
  }
  if (endDate != '' && new Date(endDate) != 'Invalid Date'){
    var endDateString = Utilities.formatDate(new Date(endDate), timeZone, 'yyyy-MM-dd').split('-').join('');
  } else {
    var endDateString = endDate;
  }
  /**
   * Builds the Google Ads query with the required parameters: field(s), date range and resource.
   * Tests if a value was entered for conditions before it adds them to the Google Ads query.
   * The last line deals with the date range either given as a literal or with two dates.
   * Documentation: https://developers.google.com/google-ads/api/docs/query/grammar
   */
  var googleAdsQuery = 'SELECT ' + fields.replace(/,/g, ', ') + ' FROM ' + resource;
  (conditions != '' && conditions != null) ? googleAdsQuery += ' WHERE ' + conditions.replace(/,/g, ' AND ') : {};
  googleAdsQuery += (endDateString == '' ?  ' DURING ' + startDateString :  ' BETWEEN ' + startDateString + ' AND ' + endDateString);
  (sort != '' && sort != null) ? googleAdsQuery += ' ORDER BY ' + sort.replace(/_desc/g, ' DESC').replace(/_asc/g, ' ASC') : {};
  (maxResults != '' && maxResults != null) ? googleAdsQuery += ' LIMIT ' + maxResults : {};
  Logger.log(googleAdsQuery);
  /**
   * Handles the access token refresh.
   * Tests if the script have an expires_at property and therefore an access_token property as they're created at the same time.
   * If not, that's the first time the add-on is used. If it has one but the property value is expired, the token is no longer valid
   * and needs to be refreshed.
   * In these both cases, an access token and an expiration date are created/refreshed and stored in the script properties.
   * In all other cases, nothing happens.
   */
  var expires_at = PropertiesService.getScriptProperties().getProperty('expires_at');
  if (expires_at == null || expires_at < new Date()){
    Logger.log('Access token has been refreshed');
    PropertiesService.getScriptProperties().setProperty('access_token', refreshAccessToken_());
    PropertiesService.getScriptProperties().setProperty('expires_at', new Date(new Date().getTime()+3600*1000));
  }
  /**
   * Builds the URL request body with the required parameters: query, our Google Ads manager account ID and developer token among others.
   */
  var googleAdsOptions = {
    contentType: "application/json",
    method: 'post',
    muteHttpExceptions: true,
    payload: JSON.stringify({"query": googleAdsQuery})
  };
  googleAdsOptions.headers = {
    "Authorization": "Bearer " + PropertiesService.getScriptProperties().getProperty('access_token'),
    "developer-token": 'oSk1hSegFeHFB_dADBt45w',
    "login-customer-id": "" + managerAccountId.replace(/-/g, ''),
    "User-Agent": "curl",
    "Accept": "application/json"
  };
  /**
   * Performs the request a first time
   * If an Authentication error is responded,
   * Refreshes the access token and the expiration date then performs the request a second time
   * The access token should rarely be refreshed here as its refresh is handled above.
   * If another error occured, return the error text.
   */
  var url = 'https://googleads.googleapis.com/v2/customers/' + accountId.replace(/-/g, '') + '/googleAds:search';
  var response = UrlFetchApp.fetch(url, googleAdsOptions);
  var json = response.getContentText();
  var data = JSON.parse(json);
  if (data.error != null){
    Logger.log(data.error.message);
    switch(data.error.code){
      /**
       * Code 401 : Invalid credentials error
       * Refreshes the access token and the expiration date then performs the request a second time.
       * The access token should rarely be refreshed here as its refresh is handled above.
       */
      case 401:
        PropertiesService.getScriptProperties().setProperty('access_token', refreshAccessToken_());
        PropertiesService.getScriptProperties().setProperty('expires_at', new Date(new Date().getTime()+3600*1000));
        googleAdsOptions.headers['Authorization'] = 'Bearer ' + PropertiesService.getScriptProperties().getProperty('access_token');
        response = UrlFetchApp.fetch(url, googleAdsOptions);
        json = response.getContentText();
        data = JSON.parse(json);
        break;
      default:
        return data.error.message;
    }
  }
  /**
   * Handles the response received.
   */
  Logger.log(data.results);
  if(data.results != null){
   fields = fields.split(',');
    var results = []
    for (var i = 0; i < data.results.length; i++){
      results.push([]);
      for (var j = 0; j < fields.length; j++){
        var metric = fields[j].split('.');
        results[i].push(data.results[i]);
        /**
         * Handles the cases where the name of the fields differs in the request and in the response.
         */
        if (metric[1].search('_') != -1){
          metric[1] = metric[1].split('_');
          for (var l = 1; l < metric[1].length; l++){
            metric[1][l] = metric[1][l].slice(0,1).toUpperCase() + metric[1][l].slice(1);
          }
          metric[1] = metric[1].join('');
        }
        for (var k = 0; k < metric.length; k++){
          results[i][j] = results[i][j][metric[k]];
        }
      }
    }
    /**
     * Include zero impressions excluded automatically with the segmentation.
     * The variable segment corresponds to the segment of date type and segmentIndex is its index in the fields.
     * The variable segments corresponds to all the segments given in the fields.
     * The variable order gives in which direction the dates will be checked depending on the specified sort order.
     */
    var segment = conditions.search(',') != -1 ? conditions[conditions.length-1] : conditions;
    var segmentIndex = fields.indexOf(segment);
    var segments = fields.filter(function(element){return element.search(/segments.*/g)!=-1});
    segments.forEach(function(element){
      segments[segments.indexOf(element)] = fields.indexOf(element);
    })
    var order = (sort != null && sort.search(/segments.(month|week|year|quarter|date|)_desc/g) != -1) ? -1 : 1;
    if (segmentIndex != -1){
      /**
      * Tests for all except the last date.
      */
      for (var i = 0; i < results.length-1; i++){
        /**
         * Tests for all except all the end date of the date range.
         */
        if (results[i][segmentIndex] != results[results.length - 1][segmentIndex]){
          var currentDateString = results[i][segmentIndex].split('-');
          var nextDateString = Utilities.formatDate(new Date(Number(currentDateString[0])+order*Number(segment=='segments.year'), Number(currentDateString[1])-1+order*Number(segment=='segments.month')+order*3*Number(segment=='segments.quarter'), Number(currentDateString[2])+order*Number(segment=='segments.date')+order*7*Number(segment=='segments.week')), timeZone, 'yyyy-MM-dd');
          var zeroImpressionsResults = [];
          zeroImpressionsResults.length = results[i].length;
          /**
           * Set the metric values to 0 and the segments value to the previous result one.
           */
          zeroImpressionsResults.fill(0);
          segments.forEach(function(element){
            zeroImpressionsResults[element] = results[i][element];
          })
          zeroImpressionsResults[segmentIndex] = nextDateString;
          nextDateString != results[i+1][segmentIndex] ? results.splice(i+1,0,zeroImpressionsResults) : {};
        /**
         * Tests for the end date of the date range if the next date is the start date of the date range.
         */
        } else if (results[i+1][segmentIndex] != results[0][segmentIndex]){
          var zeroImpressionsResults = [];
          zeroImpressionsResults.length = results[i+1].length;
          /**
           * Set the metric values to 0 and the segments value to the next result one.
           */
          zeroImpressionsResults.fill(0);
          segments.forEach(function(element){
            zeroImpressionsResults[element] = results[i+1][element];
          })
          zeroImpressionsResults[segmentIndex] = results[0][segmentIndex];
          results.splice(i+1,0,zeroImpressionsResults);
        }
      }
    }
    Logger.log(results);
    return numericStringToNumber_(results);
  }
  var noResultResponse = [data.fieldMask.split(','),data.fieldMask.split(',')];
  noResultResponse[1].fill(0);
  Logger.log(noResultResponse);
  return noResultResponse;
}

/**
 * Utility function that looks for string containing numeric values in the values given.
 * When one is found, changes it into a number.
 */
function numericStringToNumber_(values){
  for (var i = 0; i < values.length; i++){
    for(var j = 0; j < values[i].length; j++){
      if (!isNaN(values[i][j])){
        values[i][j] = Number(values[i][j], 10);
      }
    }
  }
  return values;
}

/**
 * This part of the script allows the stepping through of the Authorization Code Grant in
 * order to obtain a refresh token and then a access token.
 *
 * This script uses the out-of-band redirect URI, which is not part of the
 * OAuth2 standard, to allow not redirecting the user. If this does not work
 * with your API, try instead the OAuth playground:
 * https://developers.google.com/oauthplayground/
 *
 * Execute main function twice:
 * Execution 1: will result in a URL, which when placed in the browser will
 * issue a code.
 * Execution 2: place the code in "CODE" below and execute. If successful a
 * refresh token will be printed to the console.
 */
var CLIENT_ID = 'PASS YOUR CLIENT ID HERE';
var CLIENT_SECRET = 'PASS YOUR CLIENT SECRET HERE';

// Required scopes, e.g. ['https://www.googleapis.com/auth/drive']
var SCOPES = ['https://www.googleapis.com/auth/analytics','https://www.googleapis.com/auth/adwords'];

// Auth URL, e.g. https://accounts.google.com/o/oauth2/auth
var AUTH_URL = 'https://accounts.google.com/o/oauth2/auth';
// Token URL, e.g. https://accounts.google.com/o/oauth2/token
var TOKEN_URL = 'https://www.googleapis.com/oauth2/v4/token';

// After execution 1, enter the OAuth code inside the quotes below:
var CODE = '';

// After execution 2, enter the token inside the quotes below:
var REFRESH_TOKEN = '';

function main_() {
  if (CODE) {
    generateRefreshToken_();
  } else {
    generateAuthUrl_();
  }
}

/**
 * Creates the URL for pasting in the browser, which will generate the code
 * to be placed in the CODE variable.
 */
function generateAuthUrl_() {
  var payload = {
    scope: SCOPES.join(' '),
    // Specify that no redirection should take place
    // This is Google-specific and not part of the OAuth2 specification.
    redirect_uri: 'urn:ietf:wg:oauth:2.0:oob',
    response_type: 'code',
    access_type: 'offline',
    client_id: CLIENT_ID
  };
  var options = {payload: payload};
  var request = UrlFetchApp.getRequest(AUTH_URL, options);
  Logger.log(
      'Browse to the following URL: ' + AUTH_URL + '?' + request.payload
      + '\n Copy the obtained code at the end' );
}

/**
 * Generates a refresh token given the authorization code.
 */
function generateRefreshToken_() {
  var payload = {
    code: CODE,
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
    // Specify that no redirection should take place
    // This is Google-specific and not part of the OAuth2 specification.
    redirect_uri: 'urn:ietf:wg:oauth:2.0:oob',
    grant_type: 'authorization_code'
  };
  var options = {method: 'POST', payload: payload};
  var response = UrlFetchApp.fetch(TOKEN_URL, options);
  var data = JSON.parse(response.getContentText());
  Logger.log(data)
  if (data.refresh_token) {
    var msg = 'Success! Refresh token: ' + data.refresh_token +
      '\n\nThe following may also be a useful format for pasting into your ' +
      'script:\n\n' +
      'var CLIENT_ID = \'' + CLIENT_ID + '\';\n' +
      'var CLIENT_SECRET = \'' + CLIENT_SECRET + '\';\n' +
      'var REFRESH_TOKEN = \'' + data.refresh_token + '\';';
    Logger.log(msg);
  } else {
    Logger.log(
        'Error, failed to generate Refresh token: ' +
        response.getContentText());
  }
}

/**
 * Generates an access token given the refresh token.
 */
function refreshAccessToken_(){
  var payload = {
    refresh_token: REFRESH_TOKEN,
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
    // Specify that no redirection should take place
    // This is Google-specific and not part of the OAuth2 specification.
    redirect_uri: 'urn:ietf:wg:oauth:2.0:oob',
    grant_type: 'refresh_token'
  };
  var options = {method: 'POST', headers: {'content-type': 'application/x-www-form-urlencoded'}, payload: payload};
  var response = UrlFetchApp.fetch(TOKEN_URL, options);
  var data = JSON.parse(response.getContentText());
  Logger.log(data.access_token);
  return data.access_token;
}
