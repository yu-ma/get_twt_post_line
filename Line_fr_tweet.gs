//スプレッドシートの取得
var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

// 最初にこの関数を実行し、ログに出力されたURLにアクセスしてOAuth認証する
function twitterAuthorizeUrl() {
  Twitter.oauth.showUrl();
}

// OAuth認証成功後のコールバック関数
function twitterAuthorizeCallback(request) {
  return Twitter.oauth.callback(request);
}

// OAuth認証のキャッシュをを削除する場合はこれを実行（実行後は再度認証が必要）
function twitterAuthorizeClear() {
  Twitter.oauth.clear();
}

var Twitter = {
  //各Keyはtwitter api 管理画面から確認
  projectKey: "*********************",

  consumerKey: "*********************",
  consumerSecret: "*********************",

  apiUrl: "https://api.twitter.com/1.1/",

  oauth: {
    name: "twitter",

    service: function(screen_name) {
      // 参照元：https://github.com/googlesamples/apps-script-oauth2

      return OAuth1.createService(this.name)
      // Set the endpoint URLs.
      .setAccessTokenUrl('https://api.twitter.com/oauth/access_token')
      .setRequestTokenUrl('https://api.twitter.com/oauth/request_token')
      .setAuthorizationUrl('https://api.twitter.com/oauth/authorize')

      // Set the consumer key and secret.
      .setConsumerKey(this.parent.consumerKey)
      .setConsumerSecret(this.parent.consumerSecret)

      // Set the project key of the script using this library.
      .setProjectKey(this.parent.projectKey)


      // Set the name of the callback function in the script referenced
      // above that should be invoked to complete the OAuth flow.
      .setCallbackFunction('twitterAuthorizeCallback')

      // Set the property store where authorized tokens should be persisted.
      .setPropertyStore(PropertiesService.getUserProperties());
    },

    showUrl: function() {
      var service = this.service();
      if (!service.hasAccess()) {
        Logger.log(service.authorize());
      } else {
        Logger.log("認証済みです");
      }
    },

    callback: function (request) {
      var service = this.service();
      var isAuthorized = service.handleCallback(request);
      if (isAuthorized) {
        return HtmlService.createHtmlOutput("認証に成功しました！このタブは閉じてかまいません。");
      } else {
        return HtmlService.createHtmlOutput("認証に失敗しました・・・");
      }
    },

    clear: function(){
      OAuth1.createService(this.name)
      .setPropertyStore(PropertiesService.getUserProperties())
      .reset();
    }
  },

  api: function(path, data) {
    var that = this, service = this.oauth.service();
    if (!service.hasAccess()) {
      Logger.log("先にOAuth認証してください");
      return false;
    }

    path = path.toLowerCase().replace(/^\//, '').replace(/\.json$/, '');

    var method = (
         /^statuses\/(destroy\/\d+|update|retweet\/\d+)/.test(path)
      || /^media\/upload/.test(path)
      || /^direct_messages\/(destroy|new)/.test(path)
      || /^friendships\/(create|destroy|update)/.test(path)
      || /^account\/(settings|update|remove)/.test(path)
      || /^blocks\/(create|destroy)/.test(path)
      || /^mutes\/users\/(create|destroy)/.test(path)
      || /^favorites\/(destroy|create)/.test(path)
      || /^lists\/[^\/]+\/(destroy|create|update)/.test(path)
      || /^saved_searches\/(create|destroy)/.test(path)
      || /^geo\/place/.test(path)
      || /^users\/report_spam/.test(path)
      ) ? "post" : "get";

    var url = this.apiUrl + path + ".json";
    var options = {
      method: method,
      muteHttpExceptions: true
    };

    if ("get" === method) {
      if (!this.isEmpty(data)) {
        url += '?' + Object.keys(data).map(function(key) {
            return that.encodeRfc3986(key) + '=' + that.encodeRfc3986(data[key]);
        }).join('&');
      }
    } else if ("post" == method) {
      if (!this.isEmpty(data)) {
        options.payload = Object.keys(data).map(function(key) {
          return that.encodeRfc3986(key) + '=' + that.encodeRfc3986(data[key]);
        }).join('&');

        if (data.media) {
          options.contentType = "multipart/form-data;charset=UTF-8";
        }
      }
    }

    try {
      var result = service.fetch(url, options);
      //Logger.log(result.getContentText())
      var json = JSON.parse(result.getContentText());
      if (json) {
        if (json.error) {
          throw new Error(json.error + " (" + json.request + ")");
        } else if (json.errors) {
          var err = [];
          for (var i = 0, l = json.errors.length; i < l; i++) {
            var error = json.errors[i];
            err.push(error.message + " (code: " + error.code + ")");
          }
          throw new Error(err.join("\n"));
        } else {
          return json;
        }
      }
    } catch(e) {
      this.error(e);
    }

    return false;
  },

  error: function(error) {
    var message = null;
    if ('object' === typeof error && error.message) {
      message = error.message + " ('" + error.fileName + '.gs:' + error.lineNumber +")";
    } else {
      message = error;
    }

    Logger.log(message);
  },

  isEmpty: function(obj) {
    if (obj == null) return true;
    if (obj.length > 0)    return false;
    if (obj.length === 0)  return true;
    for (var key in obj) {
        if (hasOwnProperty.call(obj, key)) return false;
    }
    return true;
  },

  encodeRfc3986: function(str) {
    return encodeURIComponent(str).replace(/[!'()]/g, function(char) {
      return escape(char);
    }).replace(/\*/g, "%2A");
  },

  init: function() {
    this.oauth.parent = this;
    return this;
  }
}.init();


/********************************************************************
以下はサポート関数
*/

// ツイート検索
Twitter.search = function(data) {
  if ("string" === typeof data) {
    data = {q: data};
  }

  return this.api("search/tweets", data);
};

// 自分のタイムライン取得
Twitter.tl = function(since_id) {
  var data = null;

  if ("number" === typeof since_id || /^\d+$/.test(''+since_id)) {
    data = {since_id: since_id};
  } else if("object" === typeof since_id) {
    data = since_id;
  }

  return this.api("statuses/home_timeline", data);
};

// ユーザーのタイムライン取得
Twitter.usertl = function(user, since_id) {
  var path = "statuses/user_timeline";
  var data = {};

  if (user) {
    if (/^\d+$/.test(user)) {
      data.user_id = user;
    } else {
      data.screen_name = user;
    }
  } else {
    var path = "statuses/home_timeline";
  }

  if (since_id) {
    data.since_id = since_id;
  }

  return this.api(path, data);
};

// フォロワー取得
Twitter.followers = function(user) {
  var path = "followers/ids";
  var data = {}
  if (user) {
    if (/^\d+$/.test(user)) {
      data.user_id = user;
    } else {
      data.screen_name = user;
    }
  } else {
    data.screen_name = "elena_bot161127";
  }

  return this.api(path, data);
}

// フォロー取得
Twitter.friends = function(user) {
  var path = "friends/ids";
  var data = {}
  if (user) {
    if (/^\d+$/.test(user)) {
      data.user_id = user;
    } else {
      data.screen_name = user;
    }
  } else {
    data.screen_name = "elena_bot161127";
  }

  return this.api(path, data);
}


// ツイートする
Twitter.tweet = function(data, reply) {
  var path = "statuses/update";
  if ("string" === typeof data) {
    data = {status: data};
  } else if(data.media) {
    path = "statuses/update_with_media ";
  }

  if (reply) {
    if("string" === typeof reply) {
      data.in_reply_to_status_id_str = reply;
    } else {
      data.in_reply_to_status_id = reply;
    }
  }

  return this.api(path, data);
};

// トレンド取得（日本）
Twitter.trends = function(woeid) {
  data = {id: woeid || 1118108};
  var res = this.api("trends/place", data);
  return (res && res[0] && res[0].trends && res[0].trends.length) ? res[0].trends : null;
};

// トレンドのワードのみ取得
Twitter.trendWords = function(woeid) {
  data = {id: woeid || 1118108};
  var res = this.api("trends/place", data);
  if (res && res[0] && res[0].trends && res[0].trends.length) {
    var trends = res[0].trends;
    var words = [];
    for(var i = 0, l = trends.length; i < l; i++) {
      words.push(trends[i].name);
    }
    return words;
  }
};


Twitter.replies = function(last_id){
  if(last_id) {
    return Twitter.api('statuses/mentions_timeline',{'since_id':last_id});
  } else {
    return Twitter.api('statuses/mentions_timeline');
  }
}

/*******************************************************************************************************************************/
//CHANNEL_ACCESS_TOKENを設定
//LINE developerで登録をした、CHANNEL_ACCESS_TOKEN
var CHANNEL_ACCESS_TOKEN = '****************';

//メッセージ送信先のid.自身のidはlineMessegingAPIのbot管理画面で確認できる
//var user_id = ‘*****’;
//ユーザーIDを格納する列情報を取得
var id_range = sheet.getRange(2,3,sheet.getLastRow(),1)

var Line = {
  //スプレッドシートをデータベースとして使用。userIdを二次元配列で取得
  get_userId: function() {
    var user_id = id_range.getValues();
    if (user_id[user_id.length-1][0] === ""){
      user_id.pop()
    }
    id_length = Math.ceil(user_id.length / 150);
    todata = []
    for (len = 0; len < id_length; len++){
      todata.push([])
    }
    i = 0
    for (i; i < id_length; i++) {
      for (j = 0; j < 150; j++) {
        if(150*i + j < user_id.length){
          todata[i].push(user_id[150*i + j][0])
        }
      }
    }
    return todata;
  },
  //userIdに紐付いたユーザーに対しtextを送る
  push: function(to,text) {
    //var url = "https://api.line.me/v2/bot/message/push";
    var url = "https://api.line.me/v2/bot/message/multicast";
    var headers = {
      "Content-Type" : "application/json; charset=UTF-8",
      'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN,
    };

    var postData = {
      "to" : to,
      "messages" : [
        {
          'type':'text',
          'text':text,
        }
      ]
    };

    var options = {
      "method" : "post",
      "headers" : headers,
      "payload" : JSON.stringify(postData)
    };

    return UrlFetchApp.fetch(url,options);
  },
  //ユーザーのdisplayNameを取得
  get_profile: function(userId) {
    var headers = {
      'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN
    };
    var options = {
      'headers': headers
    };
    url = 'https://api.line.me/v2/bot/profile/' + userId;
    var response = UrlFetchApp.fetch(url, options);
    var content = JSON.parse(response.getContentText());
    return content;
  }
}
//postリクエストを処理
function doPost(e) {
  var json = JSON.parse(e.postData.contents);

  var get_user_id = json.events[0].source.userId
  //友達追加したユーザーのuserIdを取得
  if (json.events[0].type === "follow") {
    append_tosheet(get_user_id);
  } else if (json.events[0].type === "unfollow") {
    delete_frsheet(get_user_id);
  }
}

/*******************************************************************************************************************************/
//スプレッドシートにユーザーデータを格納
function append_tosheet(uni_id) {
   value = id_range.getLastRow();
   content = Line.get_profile(uni_id)
   sheet.getRange(value,3).setValue(uni_id);
   sheet.getRange(value,6).setValue(content.displayName);
}
//スプレッドシートからユーザーデータを削除
function delete_frsheet(uni_id) {
  value = id_range.getValues()
  index = id_range.getLastRow()
  for (i = 2; i <= index; i++) {
    if(value[i-1][0] === uni_id) {
      sheet.deleteRow(i+1);
      break;
    }
  }
}
//特定ユーザーのツイートを取得する
//urlfetch:2 * 60 * 24
function line() {
  var value = sheet.getRange("A1").getValue();
  //ツイートを取得したいユーザーのIDを入れる
  var res = Twitter.usertl("********");
  var since_id;
  var log = []
  for (i = 0; i < res.length; i++) {
    if (res[i].text && res[i].id_str > value) {
      kaz_log[i] = res[i].text
    } else {
      break;
    }
    if(i===0) {
      since_id = res[i].id_str;
      sheet.getRange("A1").setValue(since_id);
    }
  }
  if (log.length > 0) {
    todata = Line.get_userId();
    for ( j = 0; j < todata.length; j++) {
      for (i = 0; i < log.length; i++) {
        Logger.log(log[(log.length -1) - i])
        Line.push(todata[j],log[(log.length - 1) - i])
      }
    }
  }
}
