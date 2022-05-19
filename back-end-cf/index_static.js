/*
    This is a modified version by devseed,
    support the url pattern without query, 
    such as https://xxx.yyy/path/to/file
*/

// user ondrive config
const ONEDRIVE_REFRESHTOKEN = "";
const clientId = "";
const clientSecret = "";

// user netdisk config
const EXPOSE_PATH = "/public";
const PASSWD_FILENAME = ".password";

// api uri 
const loginHost = "https://login.microsoftonline.com";
const apiHost = "https://graph.microsoft.com";
const redirectUri = "http://localhost/onedrive-login";

addEventListener('scheduled', event => {
  event.waitUntil(fetchAccessToken(event.scheduledTime));
});

addEventListener("fetch", (event) => {
  try {
    return event.respondWith(handleRequest(event.request));
  } catch (e) {
    return event.respondWith(new Response("Error thrown " + e.message));
  }
});

const OAUTH = {
  redirectUri: redirectUri,
  refreshToken: ONEDRIVE_REFRESHTOKEN,
  clientId: clientId,
  clientSecret: clientSecret,
  oauthUrl: loginHost + "/common/oauth2/v2.0/",
  apiUrl: apiHost + "/v1.0/me/drive/root",
  scope: apiHost + "/Files.ReadWrite.All offline_access",
};

async function handleRequest(request) {
  let querySplited, requestPath;
  let queryUrl = new URL(request.url);
  let queryString = decodeURIComponent(request.url.split("?")[1]);
  if (queryString) querySplited = queryString.split("=");

  // use for static path
  console.log(queryUrl.search)
  if(queryUrl.search=="" && queryUrl.pathname.length>1){
    // onedrive must pass non-encoded url
    const file = decodeURIComponent(queryUrl.pathname);
    const fileName = file.split("/").pop();
    if (fileName === PASSWD_FILENAME)
      return Response("Forbidden", 403);
    requestPath = file.replace("/" + fileName, "");
    console.log(requestPath, fileName)
    const url = await fetchFiles(requestPath, fileName);
    return Response.redirect(url, 302);
  }
}

async function gatherResponse(response) {
  const { headers } = response;
  const contentType = headers.get("content-type");
  if (contentType.includes("application/json")) {
    return await response.json();
  } else if (contentType.includes("application/text")) {
    return await response.text();
  } else if (contentType.includes("text/html")) {
    return await response.text();
  } else {
    return await response.text();
  }
}

async function cacheFetch(url, options) {
  return fetch(new Request(url, options), {
    cf: {
      cacheTtl: 3600,
      cacheEverything: true,
    },
  });
}

async function getContent(url) {
  const response = await cacheFetch(url);
  const result = await gatherResponse(response);
  return result;
}

async function getContentWithHeaders(url, headers) {
  const response = await cacheFetch(url, { headers: headers });
  const result = await gatherResponse(response);
  return result;
}

async function fetchFormData(url, data) {
  const formdata = new FormData();
  for (const key in data) {
    if (data.hasOwnProperty(key)) {
      formdata.append(key, data[key]);
    }
  }
  const requestOptions = {
    method: "POST",
    body: formdata,
  };
  const response = await cacheFetch(url, requestOptions);
  const result = await gatherResponse(response);
  return result;
}

async function fetchAccessToken() {
  let refreshToken = OAUTH["refreshToken"];
  if (typeof FODI_CACHE !== 'undefined') {
    const cache = JSON.parse(await FODI_CACHE.get('token_data'));
    if (cache?.refresh_token) {
      const passedMilis = Date.now() - cache.save_time;
      if (passedMilis / 1000 < cache.expires_in - 600) {
        return cache.access_token;
      }

      if (passedMilis < 6912000000) {
        refreshToken = cache.refresh_token;
      }
    }
  }

  const url = OAUTH["oauthUrl"] + "token";
  const data = {
    client_id: OAUTH["clientId"],
    client_secret: OAUTH["clientSecret"],
    grant_type: "refresh_token",
    requested_token_use: "on_behalf_of",
    refresh_token: refreshToken,
  };
  const result = await fetchFormData(url, data);

  if (typeof FODI_CACHE !== 'undefined' && result?.refresh_token) {
    result.save_time = Date.now();
    await FODI_CACHE.put('token_data', JSON.stringify(result));
  }
  return result.access_token;
}

async function fetchFiles(path, fileName, passwd) {
  if (path === "/") path = "";
  if (path || EXPOSE_PATH) path = ":" + EXPOSE_PATH + path;

  const accessToken = await fetchAccessToken();
  const uri =
    OAUTH.apiUrl +
    encodeURI(path) +
    "?expand=children(select=name,size,parentReference,lastModifiedDateTime,@microsoft.graph.downloadUrl)";
  const body = await getContentWithHeaders(uri, { // get all file urls in a dir
    Authorization: "Bearer " + accessToken,
  });
  if (fileName) {
    let thisFile = null;
    body.children.forEach((file) => { // check if files in cache
      if (file.name === decodeURIComponent(fileName)) {
        thisFile = file["@microsoft.graph.downloadUrl"]; // get download url
        return;
      }
    });
    return thisFile;
  } else {
    let files = [];
    let encrypted = false;
    for (let i = 0; i < body.children.length; i++) {
      const file = body.children[i];
      if (file.name === PASSWD_FILENAME) {
        const PASSWD = await getContent(file["@microsoft.graph.downloadUrl"]);
        if (PASSWD !== passwd) {
          encrypted = true;
          break;
        } else {
          continue;
        }
      }
      files.push({
        name: file.name,
        size: file.size,
        time: file.lastModifiedDateTime,
        url: file["@microsoft.graph.downloadUrl"],
      });
    }
    let parent = body.children.length
      ? body.children[0].parentReference.path
      : body.parentReference.path;
    parent = parent.split(":").pop().replace(EXPOSE_PATH, "") || "/";
    parent = decodeURIComponent(parent);
    if (encrypted) {
      return JSON.stringify({ parent: parent, files: [], encrypted: true });
    } else {
      return JSON.stringify({ parent: parent, files: files });
    }
  }
}