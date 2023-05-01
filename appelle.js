function sendCalls() {
    var sheet = SpreadsheetApp.getActiveSheet();
    var dataRange = sheet.getDataRange();
    var data = dataRange.getValues();
    
    var voice = getVoiceApi();
    
    for (var i = 1; i < data.length; i++) {
      var phoneNumber = data[i][0];
      var message = data[i][1];
      
      var call = voice.calls().create({
        'phoneNumber': phoneNumber,
        'text': message
      });
      
      Logger.log(call);
    }
  }
  
  function getVoiceApi() {
    var service = getGoogleService();
    return service.newAdvancedService('voice', 'v1');
  }
  
  function getGoogleService() {
    return OAuth2.createService('Google Voice')
        .setAuthorizationBaseUrl('https://accounts.google.com/o/oauth2/auth')
        .setTokenUrl('https://accounts.google.com/o/oauth2/token')
        .setClientId('YOUR_CLIENT_ID')
        .setClientSecret('YOUR_CLIENT_SECRET')
        .setCallbackFunction('authCallback')
        .setPropertyStore(PropertiesService.getUserProperties())
        .setScope('https://www.googleapis.com/auth/voice')
        .setParam('login_hint', Session.getActiveUser().getEmail())
        .setParam('access_type', 'offline');
  }
  
  function authCallback(request) {
    var googleService = getGoogleService();
    var isAuthorized = googleService.handleCallback(request);
    if (isAuthorized) {
      return HtmlService.createHtmlOutput('Success! You can now close this tab.');
    } else {
      return HtmlService.createHtmlOutput('Denied');
    }
  }
  function getMessages() {
    var voice = getVoiceApi();
    
    var messages = voice.voice.messages.list({
      'labelIds': ['SMS']
    });
    
    Logger.log(messages);
  }
  
  function getVoiceApi() {
    var service = getGoogleService();
    return service.newAdvancedService('voice', 'v1');
  }
  
  function getGoogleService() {
    return OAuth2.createService('Google Voice')
        .setAuthorizationBaseUrl('https://accounts.google.com/o/oauth2/auth')
        .setTokenUrl('https://accounts.google.com/o/oauth2/token')
        .setClientId('YOUR_CLIENT_ID')
        .setClientSecret('YOUR_CLIENT_SECRET')
        .setCallbackFunction('authCallback')
        .setPropertyStore(PropertiesService.getUserProperties())
        .setScope('https://www.googleapis.com/auth/voice')
        .setParam('login_hint', Session.getActiveUser().getEmail())
        .setParam('access_type', 'offline');
  }
  
  function authCallback(request) {
    var googleService = getGoogleService();
    var isAuthorized = googleService.handleCallback(request);
    if (isAuthorized) {
      return HtmlService.createHtmlOutput('Success! You can now close this tab.');
    } else {
      return HtmlService.createHtmlOutput('Denied');
    }
  }
  
