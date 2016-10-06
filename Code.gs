/**
 * @OnlyCurrentDoc
 */

/*
    This file is part of MergeThisEmail.

    MergeThisEmail is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    Foobar is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with Foobar.  If not, see <http://www.gnu.org/licenses/>.

 */


var MPREFIX = "[merge]";
var MPREFIX_LENGTH = 7;

function onOpen() {
  SpreadsheetApp.getUi()
      .createAddonMenu()
      .addItem('Start', 'openDialog')
      .addToUi();
}

function onInstall() {
  onOpen();
}

function openDialog() {
  var html = HtmlService.createTemplateFromFile('Emails').evaluate();
  SpreadsheetApp.getUi()
      .showModalDialog(html, 'MergeThisEmail');
}

/* get data as an array of objects [{"header": "value", ...}, ...] 
 * startrow defaults to 2 and endrow defaults to the active sheet's last row
 */ 
function getData(startrow, endrow)
{
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastcol = sheet.getLastColumn();
  var startrow = startrow || 2;
  var lastrow = sheet.getLastRow();
  var endrow = Math.min(lastrow, endrow || lastrow);
  var headers = sheet.getRange(1, 1, 1, lastcol).getValues()[0];
  var rangearray = sheet.getRange(startrow, 1, endrow - startrow + 1, lastcol).getValues();
  var data = [];
  for (var i = 0; i < rangearray.length; i++)
  {
    rowarray = rangearray[i];
    row = {}
    for (var j = 0; j < rowarray.length; j++) {
      row[headers[j].toLowerCase()] = rowarray[j];
    }
    data.push(row);
  }
  return data;
}

function sendEmails(draftid)
{
  var draft = GmailApp.getMessageById(draftid);
  if (!draft.isDraft())
    throw "Message id does not correspond to a draft!";
  var rows = getData();
  var ii = getInlineImages(draft);
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var recipient = row["to"];
    var subject = Mustache.render(draft.getSubject().substring(MPREFIX_LENGTH).trim(), row);
    var plainbody = Mustache.render(draft.getPlainBody(), row);
    var body = Mustache.render(ii[0], row);
    var cc = draft.getCc();
    var bcc = draft.getBcc();
    var options = {"bcc": bcc, "cc": cc, "htmlBody": body, "inlineImages": ii[1]};
    var sender = getAlias(draft.getFrom());
    if (sender != "me")
      options["from"] = sender;
    GmailApp.sendEmail(recipient, subject, plainbody, options);
  }
}

function getInlineImages(draft)
{
  var body = draft.getBody();
  var imageid = [];
  var editted = body.replace(/<img src="\?[^"]+realattid=([\w_]+)(?:[\&\#][^"]+)?"([^>]*)>/g, function (match, realattid, options) { imageid.push(realattid); return '<img src="cid:' + realattid + '" ' + options+ '>';});
  var raw = draft.getRawContent();
  var blobs = {};
  for (var i = 0; i < imageid.length; i++) {
    id = imageid[i];
    var findStart = new RegExp(/Content-Type: (\w+\/\w+);(?:[^\r\n]+\r\n)*/.source +
                               "Content-ID: <"+id+">" + 
                               /\r\n(?:[^\n\r]+\r\n)*\r\n/.source, "g");
    var match = findStart.exec(raw);
    if (!match) 
      throw ("Error, inline image not found");
    var mimetype = match[1];
    var findEnd = /--/g;
    findEnd.lastIndex = findStart.lastIndex;
    findEnd.exec(raw);
    var start = findStart.lastIndex;
    var end = findEnd.lastIndex - 2;
    if (end == 0)
      end = raw.length;
    var attachbytes = Utilities.base64Decode(raw.substring(start, end));
    blobs[id] = Utilities.newBlob(attachbytes, mimetype, id);
  }
  return [editted, blobs];
}

function getAlias(sender) {
  var email = /[^<>]+@.*\.[^<>]+/;
  var stripped = email.exec(sender);
  var aliases = GmailApp.getAliases();
  for (var i = 0; i < aliases.length; i++) {
    if (aliases[i].lastIndexOf(stripped) != -1) {
      return aliases[i];
    }
  }
  return "me";
}


/*  
   --------------------------------------------------------------------------------
                 HTML Service Helper Functions (for async data)
   --------------------------------------------------------------------------------
 */

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

function getDrafts()
{
  var drafts = GmailApp.getDraftMessages();
  var draftarray = [];
  for (var i = 0; i < drafts.length; i++) {
    var d = drafts[i];
    if (d.getSubject().lastIndexOf(MPREFIX, 0) == 0)
      draftarray.push({"subject": d.getSubject(), "id": d.getId()});
  }
  return draftarray;
}

function getNumDataRows()
{
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  return sheet.getLastRow() - 1;
}

function getSampleEmail(draftid, index)
{
  var draft = GmailApp.getMessageById(draftid);
  if (!draft || !draft.isDraft())
    throw "Message id does not correspond to a draft!";
  var row = getData(2 + index, 2 + index)[0];
  var t = HtmlService.createTemplateFromFile('SampleEmail');
  t.from = getAlias(draft.getFrom());
  t.to = row["to"];
  t.subject = Mustache.render(draft.getSubject().substring(MPREFIX_LENGTH).trim(), row);
  t.body = Mustache.render(draft.getBody(), row);
  t.d = draft;
  return t.evaluate().getContent();
}
