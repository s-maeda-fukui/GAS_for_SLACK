/*
参考URL
http://tech.camph.net/slack-bot-with-gas/

初心者がGASでSlack Botをつくってみた
*/
function postSlackMessage() {
  //トークンの設定
  var token = PropertiesService.getScriptProperties().getProperty("SLACK_ACCESS_TOKEN");
  //SlackAppインスタンスの取得
  var slackApp = SlackApp.create(token);
  
  var options = {
    channelId: "#test_for_gasbot" ,//チャネル名
    userName:"bot", //投稿するbot名
    message:"Hello, World" //投稿メッセージ
  };
  
  slackApp.postMessage(options.channelId, options.message, {username: options.userName});n
}