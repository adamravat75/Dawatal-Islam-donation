function doGet(e) {
  var portal = e && e.parameter && e.parameter.portal ? e.parameter.portal : '';
  var scriptUrl = ScriptApp.getService().getUrl();
  
  if (portal === 'community') {
    var template = HtmlService.createTemplateFromFile('Community');
    template.adminUrl = scriptUrl + '?portal=admin';
    return template.evaluate()
        .setTitle('Dawatal Islam - Community Portal')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
  }
  
  if (portal === 'admin') {
    return HtmlService.createHtmlOutputFromFile('Admin')
        .setTitle('Dawatal Islam - Admin Portal')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  }
  
  var hostTemplate = HtmlService.createTemplateFromFile('Host');
  hostTemplate.scriptUrl = scriptUrl;
  return hostTemplate.evaluate()
      .setTitle('Dawatal Islam - Portal Hub')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

function findOrGetNotesColumn(pHeaders, peopleSheet) {
  var notesIdx = -1;
  for (var c = 0; c < pHeaders.length; c++) {
    var h = pHeaders[c] ? pHeaders[c].toString().trim().toLowerCase() : '';
    if (h === 'notes' || h === 'note' || h.includes('note')) {
      notesIdx = c;
      break;
    }
  }
  if (notesIdx === -1 && peopleSheet) {
    notesIdx = pHeaders.length;
    peopleSheet.getRange(1, notesIdx + 1).setValue('Notes');
  }
  return notesIdx;
}

function getDatabase() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      return { success: false, error: "Active spreadsheet not found.", members: [], donations: [], pendingUpdates: [] };
    }
    
    // 1. Load People Sheet
    var peopleSheet = ss.getSheetByName('People') || ss.getSheetByName('people') || ss.getSheets()[0];
    var members = [];
    if (peopleSheet) {
      var peopleData = peopleSheet.getDataRange().getValues();
      if (peopleData && peopleData.length > 0) {
        var pHeaders = peopleData[0];
        var pCol = { fullName: 1, houseNo: 2, address: 3, postcode: 4, phone: 5, email: 6, type: 7 };
        for (var c = 0; c < pHeaders.length; c++) {
          var h = pHeaders[c] ? pHeaders[c].toString().trim().toLowerCase() : '';
          if (h.includes('email')) pCol.email = c;
          else if (h.includes('full name') || h === 'name') pCol.fullName = c;
          else if (h.includes('house')) pCol.houseNo = c;
          else if (h.includes('address') || h.includes('street')) pCol.address = c;
          else if (h.includes('postcode')) pCol.postcode = c;
          else if (h.includes('contact number') || h.includes('phone') || h.includes('mobile')) pCol.phone = c;
          else if (h.includes('contact type') || h.includes('type')) pCol.type = c;
        }
        var notesColIdx = findOrGetNotesColumn(pHeaders, peopleSheet);
        for (var i = 1; i < peopleData.length; i++) {
          var row = peopleData[i];
          var fullName = row[pCol.fullName] ? row[pCol.fullName].toString().trim() : '';
          if (!fullName) continue;
          var parts = fullName.split(/\s+/);
          var member = {
            'MemberID': i,
            'RowIndex': i + 1,
            'Full Name': fullName,
            'First Name': parts[0] || '',
            'Surname': parts.slice(1).join(' ') || '',
            'House Number': row[pCol.houseNo] !== undefined ? row[pCol.houseNo].toString().trim() : '',
            'Address': row[pCol.address] !== undefined ? row[pCol.address].toString().trim() : '',
            'Postcode': row[pCol.postcode] !== undefined ? row[pCol.postcode].toString().trim() : '',
            'Contact Number': row[pCol.phone] !== undefined ? row[pCol.phone].toString().trim() : '',
            'Email Address': row[pCol.email] !== undefined ? row[pCol.email].toString().trim() : '',
            'Contact Type': row[pCol.type] !== undefined ? row[pCol.type].toString().trim() : '',
            'Notes': (notesColIdx !== -1 && row[notesColIdx] !== undefined) ? row[notesColIdx].toString().trim() : ''
          };
          for (var c = 0; c < pHeaders.length; c++) {
            var headerName = pHeaders[c] ? pHeaders[c].toString().trim() : '';
            if (headerName && headerName.toLowerCase().includes('membership')) {
              member[headerName] = row[c] !== undefined && row[c] !== '' ? row[c].toString().trim() : 'Unpaid';
            }
          }
          members.push(member);
        }
      }
    }

    // 2. Load Donations Sheet
    var donSheet = ss.getSheetByName('Donations') || ss.getSheetByName('Donations Log') || ss.getSheetByName('donations');
    var donations = [];
    if (donSheet) {
      var donData = donSheet.getDataRange().getValues();
      if (donData.length > 0) {
        var dHeaders = donData[0];
        var dCol = { date: -1, firstName: -1, surname: -1, type: -1, amount: -1, method: -1, giftAid: -1, houseNo: -1, postcode: -1, street: -1, bookNo: -1, invoiceNo: -1 };
        for (var c = 0; c < dHeaders.length; c++) {
          var h = dHeaders[c] ? dHeaders[c].toString().trim().toLowerCase() : '';
          if (h.includes('date')) dCol.date = c;
          else if (h.includes('first name') || h === 'name') dCol.firstName = c;
          else if (h.includes('surname') || h.includes('last name')) dCol.surname = c;
          else if (h.includes('donation type') || h === 'type') dCol.type = c;
          else if (h.includes('amount')) dCol.amount = c;
          else if (h.includes('payment') || h.includes('method')) dCol.method = c;
          else if (h.includes('gift aid')) dCol.giftAid = c;
          else if (h.includes('house')) dCol.houseNo = c;
          else if (h.includes('postcode')) dCol.postcode = c;
          else if (h.includes('street') || (h.includes('address') && !h.includes('email'))) dCol.street = c;
          else if (h.includes('book')) dCol.bookNo = c;
          else if (h.includes('invoice')) dCol.invoiceNo = c;
        }
        for (var i = 1; i < donData.length; i++) {
          var row = donData[i];
          var firstName = (dCol.firstName !== -1 && row[dCol.firstName] !== undefined) ? row[dCol.firstName].toString().trim() : '';
          var surname = (dCol.surname !== -1 && row[dCol.surname] !== undefined) ? row[dCol.surname].toString().trim() : '';
          var dateVal = (dCol.date !== -1 && row[dCol.date] !== undefined) ? row[dCol.date].toString().trim() : '';
          var rawAmt = (dCol.amount !== -1 && row[dCol.amount] !== undefined) ? row[dCol.amount].toString().replace('£', '').trim() : '0';
          if (!firstName && !dateVal && !rawAmt) continue;
          var amount = parseFloat(rawAmt);
          if (isNaN(amount)) amount = 0;
          donations.push({
            'RowIndex': i + 1,
            'Date Paid': dateVal,
            'First Name': firstName,
            'Surname': surname,
            'Donation Type': (dCol.type !== -1 && row[dCol.type] !== undefined) ? row[dCol.type].toString().trim() : '',
            'Amount': amount,
            'Payment Method': (dCol.method !== -1 && row[dCol.method] !== undefined) ? row[dCol.method].toString().trim() : 'Cash',
            'Gift Aided?': (dCol.giftAid !== -1 && row[dCol.giftAid] !== undefined) ? row[dCol.giftAid].toString().trim() : 'No',
            'House Number': (dCol.houseNo !== -1 && row[dCol.houseNo] !== undefined) ? row[dCol.houseNo].toString().trim() : '',
            'Postcode': (dCol.postcode !== -1 && row[dCol.postcode] !== undefined) ? row[dCol.postcode].toString().trim() : '',
            'Street': (dCol.street !== -1 && row[dCol.street] !== undefined) ? row[dCol.street].toString().trim() : '',
            'Book Number': (dCol.bookNo !== -1 && row[dCol.bookNo] !== undefined) ? row[dCol.bookNo].toString().trim() : '',
            'Invoice Number': (dCol.invoiceNo !== -1 && row[dCol.invoiceNo] !== undefined) ? row[dCol.invoiceNo].toString().trim() : ''
          });
        }
      }
    }

    // 3. Load Pending Updates Sheet Safely
    var pendingSheet = ss.getSheetByName('Pending_Updates') || ss.getSheetByName('Pending Updates');
    if (!pendingSheet) {
      pendingSheet = ss.insertSheet('Pending_Updates');
      pendingSheet.appendRow(['Action Type', 'Community Row', 'Area', 'Full Name', 'House Number', 'Street', 'Phone', 'Type', 'Status', 'Notes', 'Occupation', 'Email', 'Timestamp']);
    }
    
    var pendingUpdates = [];
    var pData = pendingSheet.getDataRange().getValues();
    if (pData && pData.length > 1) {
      for (var i = 1; i < pData.length; i++) {
        var r = pData[i];
        pendingUpdates.push({
          'RowIndex': i + 1,
          'ActionType': r[0] || 'Note Update',
          'CommunityRow': r[1] || '',
          'Area': r[2] || '',
          'Full Name': r[3] || '',
          'House Number': r[4] || '',
          'Street': r[5] || '',
          'Phone': r[6] || '',
          'Type': r[7] || '',
          'Status': r[8] || '',
          'Notes': r[9] || '',
          'Occupation': r[10] || '',
          'Email': r[11] || '',
          'Timestamp': r[12] || ''
        });
      }
    }

    return { success: true, members: members, donations: donations, pendingUpdates: pendingUpdates };
  } catch (err) {
    return { success: false, error: err.toString(), members: [], donations: [], pendingUpdates: [] };
  }
}

function addDonation(donation) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var donSheet = ss.getSheetByName('Donations') || ss.getSheetByName('Donations Log');
  if (!donSheet) {
    donSheet = ss.insertSheet('Donations');
    donSheet.appendRow(['Date Paid', 'First Name', 'Surname', 'Donation Type', 'Amount', 'Payment Method', 'Gift Aided?', 'House Number', 'Postcode', 'Street Name', 'Book Number', 'Invoice Number']);
  }
  var donData = donSheet.getDataRange().getValues();
  var dHeaders = donData[0];
  var colMap = {};
  for (var c = 0; c < dHeaders.length; c++) {
    var h = dHeaders[c] ? dHeaders[c].toString().trim().toLowerCase() : '';
    if (h.includes('date')) colMap['Date Paid'] = c + 1;
    else if (h.includes('first name') || h === 'name') colMap['First Name'] = c + 1;
    else if (h.includes('surname') || h.includes('last name')) colMap['Surname'] = c + 1;
    else if (h.includes('donation type') || h === 'type') colMap['Donation Type'] = c + 1;
    else if (h.includes('amount')) colMap['Amount'] = c + 1;
    else if (h.includes('payment') || h.includes('method')) colMap['Payment Method'] = c + 1;
    else if (h.includes('gift aid')) colMap['Gift Aided?'] = c + 1;
    else if (h.includes('house')) colMap['House Number'] = c + 1;
    else if (h.includes('postcode')) colMap['Postcode'] = c + 1;
    else if (h.includes('street') || (h.includes('address') && !h.includes('email'))) colMap['Street'] = c + 1;
    else if (h.includes('book')) colMap['Book Number'] = c + 1;
    else if (h.includes('invoice')) colMap['Invoice Number'] = c + 1;
  }
  
  var maxCol = dHeaders.length;
  var requiredFields = ['Date Paid', 'First Name', 'Surname', 'Donation Type', 'Amount', 'Payment Method', 'Gift Aided?', 'House Number', 'Postcode', 'Street', 'Book Number', 'Invoice Number'];
  requiredFields.forEach(function(field) {
    if (!colMap[field]) {
      maxCol++;
      donSheet.getRange(1, maxCol).setValue(field === 'Street' ? 'Street Name' : field);
      colMap[field] = maxCol;
    }
  });

  var newRow = new Array(donSheet.getLastColumn());
  var firstName = donation['First Name'] ? donation['First Name'].toString().trim() : '';
  var surname = donation['Surname'] ? donation['Surname'].toString().trim() : '';
  if (colMap['Date Paid']) newRow[colMap['Date Paid'] - 1] = donation['Date Paid'];
  if (colMap['First Name']) newRow[colMap['First Name'] - 1] = firstName;
  if (colMap['Surname']) newRow[colMap['Surname'] - 1] = surname;
  if (colMap['Donation Type']) newRow[colMap['Donation Type'] - 1] = donation['Donation Type'];
  if (colMap['Amount']) newRow[colMap['Amount'] - 1] = donation['Amount'];
  if (colMap['Payment Method']) newRow[colMap['Payment Method'] - 1] = donation['Payment Method'];
  if (colMap['Gift Aided?']) newRow[colMap['Gift Aided?'] - 1] = donation['Gift Aided?'];
  if (colMap['House Number']) newRow[colMap['House Number'] - 1] = donation['House Number'] || '';
  if (colMap['Postcode']) newRow[colMap['Postcode'] - 1] = donation['Postcode'] || '';
  if (colMap['Street']) newRow[colMap['Street'] - 1] = donation['Street'] || '';
  if (colMap['Book Number']) newRow[colMap['Book Number'] - 1] = donation['Book Number'] || '';
  if (colMap['Invoice Number']) newRow[colMap['Invoice Number'] - 1] = donation['Invoice Number'] || '';
  donSheet.appendRow(newRow);

  if (donation['Is Member Fee'] === 'Yes') {
    var datePaidStr = donation['Date Paid'] ? donation['Date Paid'].toString() : '';
    var yearMatch = datePaidStr.match(/\d{4}/);
    if (yearMatch) {
      var payYear = yearMatch[0];
      var peopleSheet = ss.getSheetByName('People') || ss.getSheets()[0];
      var peopleData = peopleSheet.getDataRange().getValues();
      var pHeaders = peopleData[0];
      var targetColIdx = -1;
      var targetHeaderName = 'Membership ' + payYear;
      for (var c = 0; c < pHeaders.length; c++) {
        if (pHeaders[c].toString().trim().toLowerCase() === targetHeaderName.toLowerCase()) {
          targetColIdx = c + 1;
          break;
        }
      }
      if (targetColIdx === -1) {
        peopleSheet.insertColumnAfter(pHeaders.length);
        targetColIdx = pHeaders.length + 1;
        peopleSheet.getRange(1, targetColIdx).setValue(targetHeaderName);
      }
      var fullTargetName = (firstName + ' ' + surname).toLowerCase();
      for (var r = 1; r < peopleData.length; r++) {
        var sheetName = peopleData[r][1] ? peopleData[r][1].toString().trim().toLowerCase() : '';
        if (sheetName === fullTargetName) {
          peopleSheet.getRange(r + 1, targetColIdx).setValue(donation['Donation Type']);
          break;
        }
      }
    }
  }
  return 'Success';
}

function deleteDonationByRow(rowIndex) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var donSheet = ss.getSheetByName('Donations') || ss.getSheetByName('Donations Log');
  if (!donSheet || rowIndex <= 1) return 'Error: Row not found';
  donSheet.deleteRow(rowIndex);
  return 'Success';
}

function addMember(member) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var peopleSheet = ss.getSheetByName('People') || ss.getSheets()[0];
  var peopleData = peopleSheet.getDataRange().getValues();
  var pHeaders = peopleData[0];
  var notesColIdx = findOrGetNotesColumn(pHeaders, peopleSheet) + 1;
  var nextId = peopleData.length;
  var fullName = (member['First Name'] + ' ' + (member['Surname'] || '')).trim();
  
  var rowValues = [
    nextId,
    fullName,
    member['House Number'] || '',
    member['Address'] || '',
    member['Postcode'] || '',
    member['Contact Number'] || '',
    member['Email Address'] || ''
  ];
  peopleSheet.appendRow(rowValues);
  return 'Success';
}

function updateMember(member) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var peopleSheet = ss.getSheetByName('People') || ss.getSheets()[0];
  var rowIndex = member['RowIndex'];
  if (!rowIndex || rowIndex <= 1) return 'Error: Invalid row';
  
  var fullName = (member['First Name'] + ' ' + (member['Surname'] || '')).trim();
  peopleSheet.getRange(rowIndex, 2).setValue(fullName);
  peopleSheet.getRange(rowIndex, 3).setValue(member['House Number'] || '');
  peopleSheet.getRange(rowIndex, 4).setValue(member['Address'] || '');
  peopleSheet.getRange(rowIndex, 5).setValue(member['Postcode'] || '');
  peopleSheet.getRange(rowIndex, 6).setValue(member['Contact Number'] || '');
  peopleSheet.getRange(rowIndex, 7).setValue(member['Email Address'] || '');
  return 'Success';
}

function deleteMemberByRow(rowIndex) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var peopleSheet = ss.getSheetByName('People') || ss.getSheets()[0];
  if (!peopleSheet || rowIndex <= 1) return 'Error';
  peopleSheet.deleteRow(rowIndex);
  return 'Success';
}

function submitCommunityUpdate(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pendingSheet = ss.getSheetByName('Pending_Updates') || ss.getSheetByName('Pending Updates');
  if (!pendingSheet) {
    pendingSheet = ss.insertSheet('Pending_Updates');
    pendingSheet.appendRow(['Action Type', 'Community Row', 'Area', 'Full Name', 'House Number', 'Street', 'Phone', 'Type', 'Status', 'Notes', 'Occupation', 'Email', 'Timestamp']);
  }

  pendingSheet.appendRow([
    payload.actionType || 'Note Update',
    payload.communityRow || '',
    payload.area || '',
    payload.fullName || '',
    payload.houseNum || '',
    payload.street || '',
    payload.phone || '',
    payload.type || '',
    payload.status || '',
    payload.notes || '',
    payload.occupation || '',
    payload.email || '',
    new Date()
  ]);
  return 'Success';
}

function processPendingUpdate(pendingRowIndex, decision, updatedData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pendingSheet = ss.getSheetByName('Pending_Updates') || ss.getSheetByName('Pending Updates');
  if (!pendingSheet) return 'Error';

  if (decision === 'approve') {
    var pData = pendingSheet.getRange(pendingRowIndex, 1, 1, 13).getValues()[0];
    var actionType = pData[0];
    var area = pData[2];
    var communityRow = pData[1];

    if (actionType === 'New Resident Sign-up') {
      var targetSheet = ss.getSheetByName(area) || ss.insertSheet(area);
      if (targetSheet.getLastRow() === 0) {
        targetSheet.appendRow(['First Name', 'Last Name', 'House Number', 'Street Name', 'Telephone Number', 'Email Address', 'Occupation', 'Type', 'Payment Status', 'Notes']);
      }
      targetSheet.appendRow([
        updatedData.fullName.split(' ')[0] || '',
        updatedData.fullName.split(' ').slice(1).join(' ') || '',
        updatedData.houseNum || '',
        updatedData.street || '',
        updatedData.phone || '',
        updatedData.email || '',
        updatedData.occupation || '',
        updatedData.type || 'Member',
        updatedData.status || 'Unpaid',
        updatedData.notes || ''
      ]);
    } else if (communityRow && area) {
      var targetSheet = ss.getSheetByName(area);
      if (targetSheet) {
        if (actionType === 'Delete') {
          targetSheet.deleteRow(parseInt(communityRow));
        } else {
          targetSheet.getRange(parseInt(communityRow), 1).setValue(updatedData.fullName.split(' ')[0] || '');
          targetSpeed.getRange(parseInt(communityRow), 2).setValue(updatedData.fullName.split(' ').slice(1).join(' ') || '');
          targetSheet.getRange(parseInt(communityRow), 3).setValue(updatedData.houseNum || '');
          targetSheet.getRange(parseInt(communityRow), 4).setValue(updatedData.street || '');
          targetSheet.getRange(parseInt(communityRow), 5).setValue(updatedData.phone || '');
          targetSheet.getRange(parseInt(communityRow), 6).setValue(updatedData.email || '');
          targetSheet.getRange(parseInt(communityRow), 7).setValue(updatedData.occupation || '');
          targetSheet.getRange(parseInt(communityRow), 8).setValue(updatedData.type || '');
          targetSheet.getRange(parseInt(communityRow), 9).setValue(updatedData.status || '');
          targetSheet.getRange(parseInt(communityRow), 10).setValue(updatedData.notes || '');
        }
      }
    }
  }

  pendingSheet.deleteRow(pendingRowIndex);
  return 'Success';
}