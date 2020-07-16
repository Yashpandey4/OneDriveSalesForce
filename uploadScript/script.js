const fetch = require("node-fetch");
const fs = require("fs");
const mime = require("mime");
const request = require("request");

var file = "./sample.docx";
var onedrive_folder = "Geminid";
var onedrive_filename = "sample.docx";

// Replace the following variables as per the documentation
var onedrive_client_id = "insert-here";
var onedrive_client_secret = "insert-here";
var onedrive_refresh_token = "insert-here";

function shareURL(token) {
    request.post(
        {
            url: `https://graph.microsoft.com/v1.0/me/drive/items/root:/${onedrive_folder}/${onedrive_filename}:/createLink`,
            headers: {
                Authorization: `bearer ${token}`,
                "Content-Type": mime.getType(file),
                "type": "edit",
                "scope": "anonymous"
            },
        },
        function(er, re, bo) {
            console.log(JSON.parse(bo).link.webUrl)
        }
    )
}

function uploadFile(token) {
  fs.readFile(file, function read(e, f) {
    request.put(
      {
        url: `https://graph.microsoft.com/v1.0/me/drive/root:/${onedrive_folder}/${onedrive_filename}:/content`,
        headers: {
          Authorization: `bearer ${token}`,
          "Content-Type": mime.getType(file),
        },
        body: f,
      },
      function (er, re, bo) {
        console.log(bo);
        shareURL(token);
      }
    );
  });
}

request.post(
  {
    url: "https://login.microsoftonline.com/common/oauth2/v2.0/token",
    form: {
      redirect_uri: "http://localhost/dashboard",
      client_id: onedrive_client_id,
      client_secret: onedrive_client_secret,
      refresh_token: onedrive_refresh_token,
      grant_type: "refresh_token",
    },
  },
  function (error, response, body) {
    uploadFile(JSON.parse(body).access_token);
  }
);
