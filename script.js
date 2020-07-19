const fetch = require("node-fetch");
const fs = require("fs");

let file = "./sample.docx";
let onedrive_folder = "Geminid";
let onedrive_filename = "sample.docx";

// Replace the following variables as per the documentation
const ONEDRIVE_CONIFG = {
  clientId: "<client id>",
  clientSecret: "<client secret>",
  refreshToken: "<refresh token>",
};

const BASE_URL = "https://graph.microsoft.com/v1.0/me/drive";

async function uploadFile() {
  let url = `${BASE_URL}/root:/${onedrive_folder}/${onedrive_filename}:/content`;
  return sendApiRequest(url, {
    method: "PUT",
    body: fs.readFileSync(file),
  });
}

async function shareURL() {
  let url = `${BASE_URL}/items/root:/${onedrive_folder}/${onedrive_filename}:/createLink`;
  return sendApiRequest(url, {
    method: "POST",
    headers: {
      type: "edit",
      scope: "anonymous",
    },
  });
}

async function deleteURL(permId) {
  let url = `${BASE_URL}/items/root:/${onedrive_folder}/${onedrive_filename}:/permissions/${permId}`;
  return sendApiRequest(url, {
    method: "DELETE"
  });
}

let AUTH_TOKEN;
const LOGIN_URL = "https://login.microsoftonline.com/common/oauth2/v2.0/token";
async function getAuthToken() {
  if (!AUTH_TOKEN)
    AUTH_TOKEN = await fetch(LOGIN_URL, {
      method: "POST",

      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
      },

      body: new URLSearchParams({
        redirect_uri: "http://localhost/dashboard",
        client_id: ONEDRIVE_CONIFG.clientId,
        client_secret: ONEDRIVE_CONIFG.clientSecret,
        refresh_token: ONEDRIVE_CONIFG.refreshToken,
        grant_type: "refresh_token",
      }).toString(),
    }).then((e) => e.json());

  return AUTH_TOKEN;
}

async function sendApiRequest(url, options) {
  if (!options.headers) options.headers = {};

  if (typeof options.body == "object")
    options.body = JSON.stringify(options.body);

  let token = await getAuthToken();
  options.headers["Authorization"] = token.access_token;
  options.headers["Content-Type"] =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

  return fetch(url, options).then((e) => e.json());
}

async function main() {
  let uploadResponse = await uploadFile();
  console.log(uploadResponse);

  let shareResponse = await shareURL();
  console.log(shareResponse);

  let deleteResponse = await deleteURL(shareResponse.id);
  console.log(deleteResponse);
}
main();
