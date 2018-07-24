/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/**
* This sample shows how to:
*    - Get the current user's metadata
*    - Get the current user's profile photo
*    - Attach the photo as a file attachment to an email message
*    - Upload the photo to the user's root drive
*    - Get a sharing link for the file and add it to the message
*    - Send the email
*/
const express = require('express');
const router = express.Router();
const graphHelper = require('../utils/graphHelper.js');
const emailer = require('../utils/emailer.js');
const passport = require('passport');
var url = require('url');
// ////const fs = require('fs');
// ////const path = require('path');
var g_cachedContentUri = "";

// Get the home page.
router.get('/', (req, res) => {
  // save the contentUri query string to the global variable.
  var q = url.parse(req.url, true);
  var queryString = q.search;
  // trim the query string symbol.
  if(queryString[0]=='?')
  {
    g_cachedContentUri = queryString.substring(1);
  }

  // check if user is authenticated
  if (!req.isAuthenticated()) {
    res.render('login');
  } else {
    renderFieldsWrapper(req,res,g_cachedContentUri);
  }
});

// Authentication request.
router.get('/login',
  passport.authenticate('azuread-openidconnect', { failureRedirect: '/' }),
    (req, res) => {
      res.redirect('/');
    });

// Authentication callback.
// After we have an access token, get user data and load the fields page.
router.get('/token',
  passport.authenticate('azuread-openidconnect', { failureRedirect: '/' }),
    (req, res) => {
      renderFieldsWrapper(req,res,g_cachedContentUri);
    });

router.get('/fields', (req, res) => {
  renderFields(req,res);
});

function renderFieldsWrapper(req,res,g_cachedContentUri)
{
  graphHelper.getFilesMetaData(req.user.accessToken, g_cachedContentUri, (err, fields) => {
    if (!err) {
      renderFields(fields,res);
    } else {
      renderError(err, res);
    }
  });
}

// Load the page to show all fields(attributes) of a file.
function renderFields(fields,res) {
  res.render('fields',{
    fields : JSON.stringify(fields)
  });
}

// Load the sendMail page.
function renderSendMail(req, res) {
  res.render('sendMail', {
    display_name: req.user.profile.displayName,
    email_address: req.user.profile.emails[0].address
  });
}

// Do prep before building the email message.
// The message contains a file attachment and embeds a sharing link to the file in the message body.
function prepForEmailMessage(req, callback) {
  const accessToken = req.user.accessToken;
  const displayName = req.user.profile.displayName;
  const destinationEmailAddress = req.body.default_email;
  // Get the current user's profile photo.
  graphHelper.getProfilePhoto(accessToken, (errPhoto, profilePhoto) => {
    // //// TODO: MSA flow with local file (using fs and path?)
    if (!errPhoto) {
        // Upload profile photo as file to OneDrive.
        graphHelper.uploadFile(accessToken, profilePhoto, (errFile, file) => {
          // Get sharingLink for file.
          graphHelper.getSharingLink(accessToken, file.id, (errLink, link) => {
            const mailBody = emailer.generateMailBody(
              displayName,
              destinationEmailAddress,
              link.webUrl,
              profilePhoto
            );
            callback(null, mailBody);
          });
        });
      }
      else {
        var fs = require('fs');
        var readableStream = fs.createReadStream('public/img/test.jpg');
        var picFile;
        var chunk;
        readableStream.on('readable', function() {
          while ((chunk=readableStream.read()) != null) {
            picFile = chunk;
          }
      });
      
      readableStream.on('end', function() {

        graphHelper.uploadFile(accessToken, picFile, (errFile, file) => {
          // Get sharingLink for file.
          graphHelper.getSharingLink(accessToken, file.id, (errLink, link) => {
            const mailBody = emailer.generateMailBody(
              displayName,
              destinationEmailAddress,
              link.webUrl,
              picFile
            );
            callback(null, mailBody);
          });
        });
      });
      }
  });
}

// Send an email.
router.post('/sendMail', (req, res) => {
  const response = res;
  const templateData = {
    display_name: req.user.profile.displayName,
    email_address: req.user.profile.emails[0].address,
    actual_recipient: req.body.default_email
  };
  prepForEmailMessage(req, (errMailBody, mailBody) => {
    if (errMailBody) renderError(errMailBody);
    graphHelper.postSendMail(req.user.accessToken, JSON.stringify(mailBody), (errSendMail) => {
      if (!errSendMail) {
        response.render('sendMail', templateData);
      } else {
        if (hasAccessTokenExpired(errSendMail)) {
          errSendMail.message += ' Expired token. Please sign out and sign in again.';
        }
        renderError(errSendMail, response);
      }
    });
  });
});

router.get('/disconnect', (req, res) => {
  req.session.destroy(() => {
    req.logOut();
    res.clearCookie('graphNodeCookie');
    res.status(200);
    res.redirect('/');
  });
});

// helpers
function hasAccessTokenExpired(e) {
  let expired;
  if (!e.innerError) {
    expired = false;
  } else {
    expired = e.forbidden &&
      e.message === 'InvalidAuthenticationToken' &&
      e.response.error.message === 'Access token has expired.';
  }
  return expired;
}
/**
 * 
 * @param {*} e 
 * @param {*} res 
 */
function renderError(e, res) {
  e.innerError = (e.response) ? e.response.text : '';
  res.render('error', {
    error: e
  });
}

module.exports = router;
