// Initializes the MSAL configuration with client and tenant IDs.
const clientId = config.clientId; // Application (client) ID
const tenantId = config.tenantId; // Directory (tenant) ID
const tokenEndpoint = config.tokenEndpoint; // Endpoint to fetch Direct Line token

const msalConfig = {
  auth: { 
    clientId, 
    authority: `https://login.microsoftonline.com/${tenantId}`,
    redirectUri: window.location.origin + "/login.html"
  },
  cache: { cacheLocation: "localStorage", storeAuthStateInCookie: false }
};

const loginRequest = { scopes: [config.customScope] };
const msalInstance = new msal.PublicClientApplication(msalConfig);
let user = null;

// Redirect to login page if not authenticated
async function checkAuthentication() {
  await msalInstance.initialize();
  
  const accounts = msalInstance.getAllAccounts();
  
  if (accounts.length === 0) {
    // No user logged in, redirect to login page
    console.log("No user authenticated, redirecting to login page...");
    window.location.href = "/login.html";
    return false;
  }
  
  // User is authenticated
  user = accounts[0];
  msalInstance.setActiveAccount(user);
  console.log("User authenticated:", user);
  return true;
}

// Handles the sign-out process when the logout button is clicked.
async function onSignOutClick() {
  const logoutRequest = {
    account: user,
    postLogoutRedirectUri: window.location.origin + "/login.html"
  };
  await msalInstance.logoutRedirect(logoutRequest);
}

// Updates UI elements to reflect logged-in state
function updateUIForLoggedInUser() {
  document.getElementById("loginStatus").innerHTML = "Currently logged in as " + user.name + " on the website.";
  document.getElementById("login").style.display = "none";
  document.getElementById("logout").style.display = "inline";
}

// Extracts the OAuth card resource URI from the activity.
function getOAuthCardResourceUri(activity) {
  if (activity && activity.attachments && activity.attachments[0] &&
      activity.attachments[0].contentType === 'application/vnd.microsoft.card.oauth' &&
      activity.attachments[0].content.tokenExchangeResource) {
    return activity.attachments[0].content.tokenExchangeResource.uri;
  }
}

// Exchanges the token asynchronously using the resource URI.
async function exchangeTokenAsync(resourceUri) {
  let user = msalInstance.getAllAccounts();
  if (user.length <= 0) return null;
  const tokenRequest = { scopes: [resourceUri] };
  try { return (await msalInstance.acquireTokenSilent(tokenRequest)).accessToken; }
  catch (err) { console.log(err); return null; }
}

// Fetches JSON data from a given URL with optional options.
async function fetchJSON(url, options = {}) {
  const res = await fetch(url, { ...options, headers: { ...options.headers, accept: 'application/json' }});
  if (!res.ok) throw new Error(`Failed to fetch JSON due to ${res.status}`);
  return await res.json();
}

// Renders the chat widget and manages the chat interactions.
async function renderChatWidget() {
  var userID = user?.localAccountId ? user.localAccountId.substr(0,36) : (Math.random().toString() + Date.now().toString()).substr(0,64);
  const { token } = await fetchJSON(tokenEndpoint);
  const directLine = window.WebChat.createDirectLine({ token, domain: "https://europe.directline.botframework.com/v3/directline" });

  const store = WebChat.createStore({}, ({ dispatch }) => next => action => {
    if (action.type === "DIRECT_LINE/CONNECT_FULFILLED") {
      dispatch({
        type: "DIRECT_LINE/POST_ACTIVITY",
        meta: { method: "keyboard" },
        payload: { activity: { type: "event", name: "startConversation", channelData: { postBack: true } } }
      });
    }

    if (action.type === 'DIRECT_LINE/INCOMING_ACTIVITY') {
      const activity = action.payload.activity;
      let resourceUri;
      if (activity.from?.role === 'bot' && (resourceUri = getOAuthCardResourceUri(activity))) {
        exchangeTokenAsync(resourceUri).then((token) => {
          if (token) {
            directLine.postActivity({
              type: 'invoke',
              name: 'signin/tokenExchange',
              value: {
                id: activity.attachments[0].content.tokenExchangeResource.id,
                connectionName: activity.attachments[0].content.connectionName,
                token
              },
              from: { id: userID, name: user.name, role: "user" }
            }).subscribe(
              id => { if (id === 'retry') return next(action); },
              error => { return next(action); }
            );
            return;
          } else return next(action);
        });
      } else return next(action);
    } else return next(action);
  });

  // Configuration for the chat widget's appearance and behavior.
  const styleOptions = {
    "hideSigninButton": true,
    "accent":"#909da7",
    "autoScrollSnapOnPage":true,
    "autoScrollSnapOnPageOffset":0,
    "avatarBorderRadius":"7%",
    "avatarSize":31,
    "botAvatarBackgroundColor":"#ffffff00",
    "botAvatarImage":"assets/images/logo.png",
    "botAvatarInitials":"B",
    "bubbleAttachmentMaxWidth":480,
    "bubbleAttachmentMinWidth":250,
    "bubbleBackground":"#f0eded",
    "bubbleBorderColor":"#f5f5f5",
    "bubbleBorderRadius":41,
    "bubbleBorderStyle":"solid",
    "bubbleBorderWidth":1,
    "bubbleFromUserBackground":"#ebefff",
    "bubbleFromUserBorderColor":"#f5f5f5",
    "bubbleFromUserBorderRadius":41,
    "bubbleFromUserBorderStyle":"solid",
    "bubbleFromUserBorderWidth":1,
    "bubbleFromUserNubOffset":0,
    "bubbleFromUserNubSize":0,
    "bubbleFromUserTextColor":"#242424",
    "bubbleImageHeight":10,
    "bubbleImageMaxHeight":240,
    "bubbleImageMinHeight":240,
    "bubbleMessageMaxWidth":480,
    "bubbleMessageMinWidth":120,
    "bubbleMinHeight":50,
    "bubbleNubOffset":0,
    "bubbleTextColor":"#242424",
    "emojiSet":true,
    "fontSizeSmall":"70%",
    "hideUploadButton":true,
    "messageActivityWordBreak":"break-word",
    "monospaceFont":"Consolas",
    "paddingRegular":10,
    "paddingWide":10,
    "primaryFont":null,
    "sendBoxBorderTop":"solid 1px #808080",
    "sendBoxButtonColor":"#0078d4",
    "sendBoxButtonColorOnHover":"#006cbe",
    "sendBoxButtonShadeBorderRadius":40,
    "sendBoxButtonShadeColorOnHover":"",
    "sendBoxHeight":60,
    "sendBoxPlaceholderColor":"#171616",
    "sendBoxTextColor":"#2e2d2d",
    "showAvatarInGroup":"status",
    "spinnerAnimationHeight":16,
    "spinnerAnimationPadding":12,
    "spinnerAnimationWidth":16,
    "subtleColor":"#000000FF",
    "suggestedActionBackgroundColor":"#909da7",
    "suggestedActionBackgroundColorOnHover":"#0078D4",
    "suggestedActionBorderColor":"",
    "suggestedActionBorderRadius":10,
    "suggestedActionBorderWidth":1,
    "suggestedActionLayout":"flow",
    "suggestedActionTextColor":"#FFFFFFFF",
    "typingAnimationDuration":5000,
    "typingAnimationHeight":20,
    "typingAnimationWidth":64,
    "userAvatarBackgroundColor":"#FFFFF",
    "userAvatarImage": 'https://img.icons8.com/?size=100&id=98957&format=png&color=0078d4',
    "userAvatarInitials":"U"
  };

  // Renders the chat widget into the specified DOM element.
  window.WebChat.renderWebChat({ directLine, store, userID, styleOptions }, document.getElementById('webchat'));
}

// Initialize and check authentication
(async () => {
  const isAuthenticated = await checkAuthentication();
  
  if (isAuthenticated) {
    updateUIForLoggedInUser();
    await renderChatWidget();
  }
})();