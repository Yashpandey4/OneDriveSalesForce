const fetch = require("node-fetch");
const fs = require("fs");
const mime = require("mime");
const request = require("request");

var file = "./sample.docx";
var onedrive_folder = "Geminid";
var onedrive_filename = "sample.docx";

// Replace the following files as per the documentation
var onedrive_client_id = "insert-here";
var onedrive_client_secret = "insert-here";
var onedrive_refresh_token = "insert-here";

function uploadFile(error, response, body) {
  fs.readFile(file, function read(e, f) {
    request.put(
      {
        url: `https://graph.microsoft.com/v1.0/drive/root:/${onedrive_folder}/${onedrive_filename}:/content`,
        headers: {
          Authorization: `bearer ${JSON.parse(body).access_token}`,
          "Content-Type": mime.getType(file),
        },
        body: f,
      },
      function (er, re, bo) {
        confole.log(bo);
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
    uploadFile(response);
  }
);
