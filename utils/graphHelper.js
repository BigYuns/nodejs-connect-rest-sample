/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

const request = require('superagent');
/* create an empty dictionary for {key: contenturi, value: id_of_a_file}. 
This is a hacky way of finding an item by id. 
I chose to find an item by id because of the graph API.*/
var dict = {};

function getTheItemFields(accessToken, contenturi, dict, res, callback) {
  //populate the dict
  //getUserData: current version of popuating the dtaa
  //look for the id 
  var idOfTheItem = dict[contenturi]; //content uri should be unique
  request
   .get('https://graph.microsoft.com/v1.0/sites/microsoft.sharepoint.com,dc09680e-36f3-4162-ad41-5da49216ca9b,f8f26641-b58a-41ef-8ad8-f03e01aced55/lists/436260bc-bba7-4a12-91df-5a004155cdff/items/' + idOfTheItem)
   .set('Authorization', 'Bearer ' + accessToken)
   .end((err, res) => {
    var jsonFormattedItemFieldsInfo = JSON.parse(res.text);
    callback(err, jsonFormattedItemFieldsInfo.fields); 
   });
}

/**
 * Generates a GET request the user endpoint.
 * @param {string} accessToken The access token to send with the request.
 * @param {Function} callback
 */
function getFilesMetaData(accessToken, contentUri, callback) {
  request
   .get('https://graph.microsoft.com/v1.0/sites/microsoft.sharepoint.com,dc09680e-36f3-4162-ad41-5da49216ca9b,f8f26641-b58a-41ef-8ad8-f03e01aced55/lists/436260bc-bba7-4a12-91df-5a004155cdff/items?expand=allfields')
   .set('Authorization', 'Bearer ' + accessToken)
   .end((err, res) => {
    var jsonFormattedTextResponse = JSON.parse(res.text);
    var numberOfItems = jsonFormattedTextResponse.value.length
    for (i = 0; i < numberOfItems; i++) { 
      var key = jsonFormattedTextResponse.value[i].fields.ContentUri;
      var value = jsonFormattedTextResponse.value[i].id;
      dict[key]= value;
    }
    getTheItemFields(accessToken,contentUri,dict,res, callback);
  });
}

/**
 * Generates a GET request for the user's profile photo.
 * @param {string} accessToken The access token to send with the request.
 * @param {Function} callback
//  */
function getProfilePhoto(accessToken, callback) {
  // Get the profile photo of the current user (from the user's mailbox on Exchange Online).
  // This operation in version 1.0 supports only work or school mailboxes, not personal mailboxes.
  request
   .get('https://graph.microsoft.com/beta/me/photo/$value')
   .set('Authorization', 'Bearer ' + accessToken)
   .end((err, res) => {
     // Returns 200 OK and the photo in the body. If no photo exists, returns 404 Not Found.
     callback(err, res.body);
   });
}

/**
 * Generates a PUT request to upload a file.
 * @param {string} accessToken The access token to send with the request.
 * @param {Function} callback
//  */
function uploadFile(accessToken, file, callback) {
  // This operation only supports files up to 4MB in size.
  // To upload larger files, see `https://developer.microsoft.com/graph/docs/api-reference/v1.0/api/item_createUploadSession`.
  request
   .put('https://graph.microsoft.com/beta/me/drive/root/children/mypic.jpg/content')
   .send(file)
   .set('Authorization', 'Bearer ' + accessToken)
   .set('Content-Type', 'image/jpg')
   .end((err, res) => {
     // Returns 200 OK and the file metadata in the body.
     callback(err, res.body);
   });
}

/**
 * Generates a POST request to create a sharing link (if one doesn't already exist).
 * @param {string} accessToken The access token to send with the request.
 * @param {string} id The ID of the file to get or create a sharing link for.
 * @param {Function} callback
//  */
// See https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/api/item_createlink
function getSharingLink(accessToken, id, callback) {
  request
   .post('https://graph.microsoft.com/beta/me/drive/items/' + id + '/createLink')
   .send({ type: 'view' })
   .set('Authorization', 'Bearer ' + accessToken)
   .set('Content-Type', 'application/json')
   .end((err, res) => {
     // Returns 200 OK and the permission with the link in the body.
     callback(err, res.body.link);
   });
}

/**
 * Generates a POST request to the SendMail endpoint.
 * @param {string} accessToken The access token to send with the request.
 * @param {string} data The data which will be 'POST'ed.
 * @param {Function} callback
 * Per issue #53 for BadRequest when message uses utf-8 characters:
 * `.set('Content-Length': Buffer.byteLength(mailBody,'utf8'))`
 */
function postSendMail(accessToken, message, callback) {
  request
   .post('https://graph.microsoft.com/beta/me/sendMail')
   .send(message)
   .set('Authorization', 'Bearer ' + accessToken)
   .set('Content-Type', 'application/json')
   .set('Content-Length', message.length)
   .end((err, res) => {
     // Returns 202 if successful.
     // Note: If you receive a 500 - Internal Server Error
     // while using a Microsoft account (outlook.com, hotmail.com or live.com),
     // it's possible that your account has not been migrated to support this flow.
     // Check the inner error object for code 'ErrorInternalServerTransientError'.
     // You can try using a newly created Microsoft account or contact support.
     callback(err, res);
   });
}

exports.getFilesMetaData = getFilesMetaData;
exports.getProfilePhoto = getProfilePhoto;
exports.uploadFile = uploadFile;
exports.getSharingLink = getSharingLink;
exports.postSendMail = postSendMail;
