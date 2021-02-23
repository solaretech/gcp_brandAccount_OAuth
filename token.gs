// 随時コメント追記

const initialize = () => {
  const CLIENT_ID = Browser.inputBox("CLIENT_IDを入力してください。")
  const CLIENT_SECRET = Browser.inputBox("CLIENT_SECRETを入力してください。")
  PropertiesService.getScriptProperties().setProperty("CLIENT_ID", CLIENT_ID)
  PropertiesService.getScriptProperties().setProperty("CLIENT_SECRET", CLIENT_SECRET)

  // PropertyにCLIENT_ID, CLIENT_SECRETが入っているか確認
  const client_id = PropertiesService.getScriptProperties().getProperty("CLIENT_ID");
  const client_secret = PropertiesService.getScriptProperties().getProperty("CLIENT_SECRET");
  if(!!client_id && !!client_secret){
    Browser.msgBox("認証情報の登録が完了しました。")
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("C4").setValue("→完了");
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("C8").clearContent();
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("C12").clearContent();
  }else{
    Browser.msgBox("認証情報の登録に失敗しました。")
  }
  
}

const deployURL = "DEPLOY_URL"
const tokenUri = "https://accounts.google.com/o/oauth2/token"

const authorize = () => {
  const client_id = PropertiesService.getScriptProperties().getProperty("CLIENT_ID");
  const ui = SpreadsheetApp.getUi();
  let url = 'https://accounts.google.com/o/oauth2/auth?client_id=' + client_id + '&redirect_uri=' + deployURL + '&scope=https://www.googleapis.com/auth/youtube.readonly&response_type=code&access_type=offline';
  const dialog = HtmlService.createTemplateFromFile("auth_assignment");
  dialog.url = url;
  const html = dialog.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME).setWidth(520).setHeight(240);
  ui.showModalDialog(html, "認証情報の登録")
  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("C8").setValue("→完了");
  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("C12").clearContent();
}

const getNewAccessToken = () => {
  const clientID = PropertiesService.getScriptProperties().getProperty("CLIENT_ID")
  const clientSecret = PropertiesService.getScriptProperties().getProperty("CLIENT_SECRET")
  const code = Browser.inputBox("第一認証のURLにアクセスして表示されたコードを入力");
  if (code === "") {
    throw new Error("getNewAccessToken: code未入力が未定義です")
  }
  
  const payload = {
    'code': code,
    'client_id': clientID,
    'client_secret': clientSecret,
    'redirect_uri': deployURL,
    'grant_type': 'authorization_code'
  }
  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
  }
  const response = UrlFetchApp.fetch(tokenUri, options)
  
  if (response.getResponseCode() == 200) {
    const responseJson = JSON.parse(response.getContentText())
    setAccessToken(responseJson["access_token"], responseJson["expires_in"])
    if(responseJson["refresh_token"]){
      setRefreshToken(responseJson["refresh_token"])
    }
    Browser.msgBox("認証情報が保存されました。")
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("C12").setValue("→完了");
  }
  else {
    throw new Error("getNewAccessToken: response error")
  }
}

const setRefreshToken = (token) => {
  PropertiesService.getScriptProperties().setProperty("REFRESH_TOKEN",token);
}

const getRefreshToken = () => {
  return PropertiesService.getScriptProperties().getProperty("REFRESH_TOKEN");
}

const setAccessToken = (token, expiresIn) => {
  PropertiesService.getScriptProperties().setProperty("ACCESS_TOKEN",token);
  PropertiesService.getScriptProperties().setProperty("EXPIRES_IN",expiresIn);
  PropertiesService.getScriptProperties().setProperty("REFRESH_AT",new Date());
}

const getAccessToken = () => {
  const refreshLimit = new Date(PropertiesService.getScriptProperties().getProperty("REFRESH_AT"))
  const expiresIn = Number(PropertiesService.getScriptProperties().getProperty("EXPIRES_IN"))
  const now = new Date()

  refreshLimit.setSeconds(now.getSeconds() + expiresIn)

  if(refreshLimit < now) {
    refreshAccessToken()
  }
  
  return PropertiesService.getScriptProperties().getProperty("ACCESS_TOKEN")

}

const refreshAccessToken = () => {
  const clientID = PropertiesService.getScriptProperties().getProperty("CLIENT_ID")
  const clientSecret = PropertiesService.getScriptProperties().getProperty("CLIENT_SECRET")
  const payload = {
    'client_id': clientID,
    'client_secret': clientSecret,
    'refresh_token': getRefreshToken(),
    'grant_type': 'refresh_token'
  }
  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
  }
  const response = UrlFetchApp.fetch(tokenUri, options)
  
  if (response.getResponseCode() == 200) {
    const responseJson = JSON.parse(response.getContentText())
    setAccessToken(responseJson["access_token"], responseJson["expires_in"])
  }
  else {
    throw new Error("refreshAccessToken: response error")
  }
}
