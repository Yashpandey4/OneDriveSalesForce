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

async function main() {
    try {
        var auth_response = await fetch("https://login.microsoftonline.com/common/oauth2/v2.0/token", {
            method: "POST",
            body: new URLSearchParams({
                redirect_uri: "http://localhost/dashboard",
                client_id: onedrive_client_id,
                client_secret: onedrive_client_secret,
                refresh_token: onedrive_refresh_token,
                grant_type: "refresh_token",
            }).toString(),
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded'
            }
        });
        var data = await response.json();
        var token = data.access_token;
        uploadFile(token);
        shareURL(token);
    }
    catch(err) {
        console.log(`Auth failed: ${err}`);
    }
    
}

async function uploadFile(token) {
    fs.readFile(file, function read(e, f) {
        try {
            var upload_response = await fetch(`https://graph.microsoft.com/v1.0/me/drive/root:/${onedrive_folder}/${onedrive_filename}:/content`, {
                method: "PUT",
                body: f,
                headers: {
                    Authorization: `bearer ${token}`,
                    "Content-Type": mime.getType(file),
                }
            })
            console.log(`File uploaded successfully! ${upload_response.json()}`);
        }
        catch(er) {
            console.log(`Upload failed: ${err}`);
        }
    });
  }

async function shareURL(token) {
    try {
        var share_response = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/root:/${onedrive_folder}/${onedrive_filename}:/createLink`, {
            method: 'POST',
            headers: {
                Authorization: `bearer ${token}`,
                "Content-Type": mime.getType(file),
                "type": "edit",
                "scope": "anonymous"
            }
        });
        var data = await share_response.json();
        console.log(`Shared URL of the File uploaded is ${data.link.webUrl}`);
    }
    catch(er) {
        console.log(`Share failed: ${err}`);
    }
}

