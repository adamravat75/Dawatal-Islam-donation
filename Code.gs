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
    
    checkAndResetForNewYear(ss);
    
    var masterDirectory = {};

    // Helper to generate a collision-free key using Name + House Number + Street Name (ignoring missing postcodes)
    function makeUniqueKey(fullName, houseNo, street) {
      var cleanName = (fullName || '').toLowerCase().trim();
      var cleanHouse = (houseNo || '').toLowerCase().trim();
      var cleanStreet = (street || '').toLowerCase().trim();
      
      return cleanName + '|' + cleanHouse + '|' + cleanStreet;
    }

    // 1. SAFE PEOPLE FETCH
    try {
      var peopleSheet = ss.getSheetByName('People') || ss.getSheetByName('people') || ss.getSheets()[0];
      if (peopleSheet && peopleSheet.getLastRow() > 1) {
        var peopleData = peopleSheet.getDataRange().getValues();
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
          var houseNo = row[pCol.houseNo] !== undefined && row[pCol.houseNo] !== null ? row[pCol.houseNo].toString().trim() : '';
          var address = row[pCol.address] !== undefined && row[pCol.address] !== null ? row[pCol.address].toString().trim() : '';
          var postcode = row[pCol.postcode] !== undefined && row[pCol.postcode] !== null ? row[pCol.postcode].toString().trim() : '';
          var key = makeUniqueKey(fullName, houseNo, address);
          
          masterDirectory[key] = {
            'MemberID': i,
            'PeopleRowIndex': i + 1,
            'CommunityRowIndex': null,
            'Full Name': fullName,
            'First Name': parts[0] || '',
            'Surname': parts.slice(1).join(' ') || '',
            'House Number': houseNo,
            'Address': address,
            'Postcode': postcode,
            'Contact Number': row[pCol.phone] !== undefined && row[pCol.phone] !== null ? row[pCol.phone].toString().trim() : '',
            'Email Address': row[pCol.email] !== undefined && row[pCol.email] !== null ? row[pCol.email].toString().trim() : '',
            'Contact Type': 'Member',
            'Notes': (notesColIdx !== -1 && row[notesColIdx] !== undefined && row[notesColIdx] !== null) ? row[notesColIdx].toString().trim() : ''
          };
          
          for (var c = 0; c < pHeaders.length; c++) {
            var headerName = pHeaders[c] ? pHeaders[c].toString().trim() : '';
            if (headerName && headerName.toLowerCase().includes('membership')) {
              var val = row[c];
              masterDirectory[key][headerName] = (val !== undefined && val !== null && val !== '') ? val.toString().trim() : 'Unpaid';
            }
          }
        }
      }
    } catch (errP) {
      console.error("People read error: " + errP.toString());
    }

    // 2. SAFE ALL COMMUNITY FETCH
    try {
      var commSheet = ss.getSheetByName('All Community') || ss.getSheetByName('All_Community');
      if (commSheet && commSheet.getLastRow() > 1) {
        var cData = commSheet.getDataRange().getValues();
        var cHeaders = cData[0];
        var cCol = { area: 0, fullName: 1, houseNo: 2, street: 3, postcode: 4, phone: 5, type: 6, status: 7, notes: 8, occupation: 10, email: 11 };
        
        for (var c = 0; c < cHeaders.length; c++) {
          var h = cHeaders[c] ? cHeaders[c].toString().trim().toLowerCase() : '';
          if (h.includes('area')) cCol.area = c;
          else if (h.includes('full name') || h === 'name' || h.includes('resident')) cCol.fullName = c;
          else if (h.includes('house')) cCol.houseNo = c;
          else if ((h.includes('street') || h.includes('address')) && !h.includes('email')) cCol.street = c;
          else if (h.includes('postcode')) cCol.postcode = c;
          else if (h.includes('phone') || h.includes('contact') || h.includes('mobile')) cCol.phone = c;
          else if (h.includes('type') || h.includes('member/non')) cCol.type = c;
          else if (h.includes('status') || h.includes('payment')) cCol.status = c;
          else if (h.includes('note')) cCol.notes = c;
          else if (h.includes('occupation')) cCol.occupation = c;
          else if (h.includes('email')) cCol.email = c;
        }

        for (var i = 1; i < cData.length; i++) {
          var row = cData[i];
          var fullName = row[cCol.fullName] !== undefined ? row[cCol.fullName].toString().trim() : '';
          if (!fullName) continue;
          
          var houseNo = row[cCol.houseNo] !== undefined && row[cCol.houseNo] !== null ? row[cCol.houseNo].toString().trim() : '';
          var street = row[cCol.street] !== undefined && row[cCol.street] !== null ? row[cCol.street].toString().trim() : '';
          var postcode = row[cCol.postcode] !== undefined && row[cCol.postcode] !== null ? row[cCol.postcode].toString().trim() : '';
          var key = makeUniqueKey(fullName, houseNo, street);
          var parts = fullName.split(/\s+/);
          var typeVal = row[cCol.type] !== undefined ? row[cCol.type].toString().trim() : 'Member';
          
          if (masterDirectory[key]) {
            masterDirectory[key]['CommunityRowIndex'] = i + 1;
            masterDirectory[key]['Area'] = row[cCol.area] !== undefined ? row[cCol.area].toString().trim() : '';
            masterDirectory[key]['Contact Type'] = typeVal;
            masterDirectory[key]['Occupation'] = row[cCol.occupation] !== undefined ? row[cCol.occupation].toString().trim() : '';
            if (!masterDirectory[key]['Postcode'] && postcode) masterDirectory[key]['Postcode'] = postcode;
            if (!masterDirectory[key]['Email Address'] && row[cCol.email]) masterDirectory[key]['Email Address'] = row[cCol.email].toString().trim();
          } else {
            masterDirectory[key] = {
              'MemberID': 'C-' + i,
              'PeopleRowIndex': null,
              'CommunityRowIndex': i + 1,
              'Full Name': fullName,
              'First Name': parts[0] || '',
              'Surname': parts.slice(1).join(' ') || '',
              'House Number': houseNo,
              'Address': street,
              'Postcode': postcode,
              'Contact Number': row[cCol.phone] !== undefined && row[cCol.phone] !== null ? row[cCol.phone].toString().trim() : '',
              'Contact Type': typeVal,
              'Notes': row[cCol.notes] !== undefined && row[cCol.notes] !== null ? row[cCol.notes].toString().trim() : '',
              'Area': row[cCol.area] !== undefined && row[cCol.area] !== null ? row[cCol.area].toString().trim() : '',
              'Occupation': row[cCol.occupation] !== undefined && row[cCol.occupation] !== null ? row[cCol.occupation].toString().trim() : '',
              'Email Address': row[cCol.email] !== undefined && row[cCol.email] !== null ? row[cCol.email].toString().trim() : ''
            };
          }
        }
      }
    } catch (errC) {
      console.error("Community read error: " + errC.toString());
    }

    var members = Object.keys(masterDirectory).map(function(k) { return masterDirectory[k]; });

    // 3. SAFE DONATIONS FETCH
    var donations = [];
    try {
      var donSheet = ss.getSheetByName('Donations') || ss.getSheetByName('Donations Log') || ss.getSheetByName('donations');
      if (donSheet && donSheet.getLastRow() > 1) {
        var donData = donSheet.getDataRange().getValues();
        var dHeaders = donData[0];
        var dCol = { date: 0, firstName: 1, surname: 2, type: 3, amount: 4, method: 5, giftAid: 6, houseNo: 7, postcode: 8, street: 9, bookNo: 10, invoiceNo: 11 };
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
          var dateVal = row[dCol.date] instanceof Date ? row[dCol.date].toISOString().split('T')[0] : (row[dCol.date] ? row[dCol.date].toString().trim() : '');
          var firstName = row[dCol.firstName] ? row[dCol.firstName].toString().trim() : '';
          var surname = row[dCol.surname] ? row[dCol.surname].toString().trim() : '';
          var typeVal = row[dCol.type] ? row[dCol.type].toString().trim() : '';
          var rawAmt = row[dCol.amount] !== undefined && row[dCol.amount] !== null ? row[dCol.amount].toString().replace('£', '').trim() : '0';
          var amount = parseFloat(rawAmt);
          if (isNaN(amount)) amount = 0;
          
          donations.push({
            'RowIndex': i + 1,
            'Date Paid': dateVal,
            'First Name': firstName,
            'Surname': surname,
            'Donation Type': typeVal,
            'Amount': amount,
            'Payment Method': row[dCol.method] !== undefined && row[dCol.method] !== null ? row[dCol.method].toString().trim() : 'Cash',
            'Gift Aided?': row[dCol.giftAid] !== undefined && row[dCol.giftAid] !== null ? row[dCol.giftAid].toString().trim() : 'No',
            'House Number': row[dCol.houseNo] !== undefined && row[dCol.houseNo] !== null ? row[dCol.houseNo].toString().trim() : '',
            'Postcode': row[dCol.postcode] !== undefined && row[dCol.postcode] !== null ? row[dCol.postcode].toString().trim() : '',
            'Street': row[dCol.street] !== undefined && row[dCol.street] !== null ? row[dCol.street].toString().trim() : '',
            'Book Number': row[dCol.bookNo] !== undefined && row[dCol.bookNo] !== null ? row[dCol.bookNo].toString().trim() : '',
            'Invoice Number': row[dCol.invoiceNo] !== undefined && row[dCol.invoiceNo] !== null ? row[dCol.invoiceNo].toString().trim() : ''
          });
        }
      }
    } catch (errD) {
      console.error("Donations read error: " + errD.toString());
    }

    // 4. SAFE PENDING UPDATES FETCH
    var pendingUpdates = [];
    try {
      var pendingSheet = ss.getSheetByName('Pending_Updates') || ss.getSheetByName('Pending Updates');
      if (pendingSheet && pendingSheet.getLastRow() > 1) {
        var pData = pendingSheet.getDataRange().getValues();
        for (var i = 1; i < pData.length; i++) {
          var r = pData[i];
          var ts = r[12] instanceof Date ? r[12].toISOString() : (r[12] ? r[12].toString() : '');
          pendingUpdates.push({
            'RowIndex': i + 1,
            'ActionType': r[0] ? r[0].toString() : 'Note Update',
            'CommunityRow': r[1] !== undefined ? r[1].toString() : '',
            'Area': r[2] !== undefined ? r[2].toString() : '',
            'Full Name': r[3] !== undefined ? r[3].toString() : '',
            'House Number': r[4] !== undefined ? r[4].toString() : '',
            'Street': r[5] !== undefined ? r[5].toString() : '',
            'Postcode': r[6] !== undefined ? r[6].toString() : '',
            'Phone': r[7] !== undefined ? r[7].toString() : '',
            'Type': r[8] !== undefined ? r[8].toString() : '',
            'Status': r[9] !== undefined ? r[9].toString() : '',
            'Notes': r[10] !== undefined ? r[10].toString() : '',
            'Occupation': r[11] !== undefined ? r[11].toString() : '',
            'Email': r[12] !== undefined ? r[12].toString() : '',
            'Timestamp': ts
          });
        }
      }
    } catch (errU) {
      console.error("Pending updates read error: " + errU.toString());
    }

    return { success: true, members: members, donations: donations, pendingUpdates: pendingUpdates };
  } catch (err) {
    return { success: false, error: err.toString(), members: [], donations: [], pendingUpdates: [] };
  }
}

function checkAndResetForNewYear(ss) {
  var props = PropertiesService.getDocumentProperties();
  var currentYear = new Date().getFullYear().toString();
  var lastCheckedYear = props.getProperty('LAST_ACTIVE_YEAR');

  if (lastCheckedYear && lastCheckedYear !== currentYear) {
    try {
      var peopleSheet = ss.getSheetByName('People') || ss.getSheetByName('people');
      if (peopleSheet && peopleSheet.getLastRow() > 1) {
        var pData = peopleSheet.getDataRange().getValues();
        var pHeaders = pData[0];
        for (var c = 0; c < pHeaders.length; c++) {
          var h = pHeaders[c] ? pHeaders[c].toString().toLowerCase() : '';
          if (h.includes('membership') || h.includes(lastCheckedYear)) {
            peopleSheet.getRange(1, c + 1).setValue('Membership ' + currentYear);
            for (var r = 1; r < pData.length; r++) {
              peopleSheet.getRange(r + 1, c + 1).setValue('Unpaid');
            }
          }
        }
      }

      var commSheet = ss.getSheetByName('All Community');
      if (commSheet && commSheet.getLastRow() > 1) {
        var cData = commSheet.getDataRange().getValues();
        var cHeaders = cData[0];
        var typeColIdx = 6;
        var statusColIdx = 7;
        for (var c = 0; c < cHeaders.length; c++) {
          var h = cHeaders[c] ? cHeaders[c].toString().toLowerCase() : '';
          if (h.includes('type')) typeColIdx = c;
          if (h.includes('status')) statusColIdx = c;
        }

        for (var r = 1; r < cData.length; r++) {
          var typeVal = cData[r][typeColIdx] ? cData[r][typeColIdx].toString().toLowerCase() : '';
          if (typeVal.includes('member') && !typeVal.includes('non')) {
            commSheet.getRange(r + 1, statusColIdx + 1).setValue('Unpaid');
          } else {
            commSheet.getRange(r + 1, statusColIdx + 1).setValue('N/A');
          }
        }
      }
    } catch (e) {
      console.error("Year reset error: " + e.toString());
    }
  }
  props.setProperty('LAST_ACTIVE_YEAR', currentYear);
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

  var dType = donation['Donation Type'] ? donation['Donation Type'].toString().trim() : '';
  var dTypeLower = dType.toLowerCase();
  var isMemberFeePayment = dTypeLower.includes('member fee') || dTypeLower.includes('student member fee');
  
  if (isMemberFeePayment) {
    var datePaidStr = donation['Date Paid'] ? donation['Date Paid'].toString() : '';
    var yearMatch = datePaidStr.match(/\d{4}/);
    var payYear = yearMatch ? yearMatch[0] : new Date().getFullYear().toString();
    
    if (firstName && surname) {
      var peopleSheet = ss.getSheetByName('People') || ss.getSheets()[0];
      if (peopleSheet && peopleSheet.getLastRow() > 1) {
        var pData = peopleSheet.getDataRange().getValues();
        var pHeaders = pData[0];
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
        for (var r = 1; r < pData.length; r++) {
          var sheetName = pData[r][1] ? pData[r][1].toString().trim().toLowerCase() : '';
          if (sheetName === fullTargetName) {
            peopleSheet.getRange(r + 1, targetColIdx).setValue(dType);
            break;
          }
        }
      }

      var commSheet = ss.getSheetByName('All Community');
      if (commSheet && commSheet.getLastRow() > 1) {
        var cData = commSheet.getDataRange().getValues();
        var cHeaders = cData[0];
        var statusColIdx = 7;
        var typeColIdx = 6;
        for(var c=0; c<cHeaders.length; c++) {
          if (cHeaders[c].toString().toLowerCase().includes('status')) statusColIdx = c;
          if (cHeaders[c].toString().toLowerCase().includes('type')) typeColIdx = c;
        }

        for (var i = 1; i < cData.length; i++) {
          var cName = cData[i][1] ? cData[i][1].toString().trim().toLowerCase() : '';
          var typeVal = cData[i][typeColIdx] ? cData[i][typeColIdx].toString().toLowerCase() : '';
          
          if (cName.includes(firstName.toLowerCase()) && cName.includes(surname.toLowerCase())) {
            if (typeVal.includes('member') && !typeVal.includes('non')) {
              commSheet.getRange(i + 1, statusColIdx + 1).setValue('Paid');
            }
            break;
          }
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
  
  var donData = donSheet.getDataRange().getValues();
  var targetDonation = donData[rowIndex - 1];
  
  if (targetDonation) {
    var dHeaders = donData[0];
    var dCol = { date: 0, firstName: 1, surname: 2, type: 3 };
    for (var c = 0; c < dHeaders.length; c++) {
      var h = dHeaders[c] ? dHeaders[c].toString().trim().toLowerCase() : '';
      if (h.includes('date')) dCol.date = c;
      else if (h.includes('first name') || h === 'name') dCol.firstName = c;
      else if (h.includes('surname') || h.includes('last name')) dCol.surname = c;
      else if (h.includes('donation type') || h === 'type') dCol.type = c;
    }

    var datePaidStr = targetDonation[dCol.date] instanceof Date ? targetDonation[dCol.date].toISOString().split('T')[0] : (targetDonation[dCol.date] ? targetDonation[dCol.date].toString() : '');
    var firstName = targetDonation[dCol.firstName] ? targetDonation[dCol.firstName].toString().trim() : '';
    var surname = targetDonation[dCol.surname] ? targetDonation[dCol.surname].toString().trim() : '';
    var dType = targetDonation[dCol.type] ? targetDonation[dCol.type].toString().trim() : '';
    
    var dTypeLower = dType.toLowerCase();
    var isMemberFeePayment = dTypeLower.includes('member fee') || dTypeLower.includes('student member fee');
    
    if (isMemberFeePayment && firstName && surname) {
      var yearMatch = datePaidStr.match(/\d{4}/);
      var payYear = yearMatch ? yearMatch[0] : new Date().getFullYear().toString();
      
      var peopleSheet = ss.getSheetByName('People') || ss.getSheets()[0];
      if (peopleSheet && peopleSheet.getLastRow() > 1) {
        var pData = peopleSheet.getDataRange().getValues();
        var pHeaders = pData[0];
        var targetColIdx = -1;
        var targetHeaderName = 'Membership ' + payYear;
        
        for (var c = 0; c < pHeaders.length; c++) {
          if (pHeaders[c].toString().trim().toLowerCase() === targetHeaderName.toLowerCase()) {
            targetColIdx = c + 1;
            break;
          }
        }
        
        if (targetColIdx !== -1) {
          var fullTargetName = (firstName + ' ' + surname).toLowerCase();
          for (var r = 1; r < pData.length; r++) {
            var sheetName = pData[r][1] ? pData[r][1].toString().trim().toLowerCase() : '';
            if (sheetName === fullTargetName) {
              peopleSheet.getRange(r + 1, targetColIdx).setValue('Unpaid');
              break;
            }
          }
        }
      }

      var commSheet = ss.getSheetByName('All Community');
      if (commSheet && commSheet.getLastRow() > 1) {
        var cData = commSheet.getDataRange().getValues();
        var cHeaders = cData[0];
        var statusColIdx = 7;
        var typeColIdx = 6;
        for(var c=0; c<cHeaders.length; c++) {
          if (cHeaders[c].toString().toLowerCase().includes('status')) statusColIdx = c;
          if (cHeaders[c].toString().toLowerCase().includes('type')) typeColIdx = c;
        }

        for (var i = 1; i < cData.length; i++) {
          var cName = cData[i][1] ? cData[i][1].toString().trim().toLowerCase() : '';
          var typeVal = cData[i][typeColIdx] ? cData[i][typeColIdx].toString().toLowerCase() : '';
          
          if (cName.includes(firstName.toLowerCase()) && cName.includes(surname.toLowerCase())) {
            if (typeVal.includes('member') && !typeVal.includes('non')) {
              commSheet.getRange(i + 1, statusColIdx + 1).setValue('Unpaid');
            }
            break;
          }
        }
      }
    }
  }

  donSheet.deleteRow(rowIndex);
  return 'Success';
}

function addMember(member) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var fullName = (member['First Name'] + ' ' + (member['Surname'] || '')).trim();
  var isMemberType = member['Type'] ? member['Type'].toString().toLowerCase().includes('member') && !member['Type'].toString().toLowerCase().includes('non') : true;
  var initialStatus = isMemberType ? 'Unpaid' : 'N/A';

  if (isMemberType) {
    var peopleSheet = ss.getSheetByName('People') || ss.getSheets()[0];
    var peopleData = peopleSheet.getDataRange().getValues();
    var pHeaders = peopleData[0];
    var notesColIdx = findOrGetNotesColumn(pHeaders, peopleSheet) + 1;
    var nextId = peopleData.length;
    
    var rowValues = [
      nextId,
      fullName,
      member['House Number'] || '',
      member['Address'] || '',
      member['Postcode'] || '',
      member['Contact Number'] || '',
      member['Email Address'] || ''
    ];
    
    while (rowValues.length < notesColIdx - 1) {
      rowValues.push('');
    }
    rowValues.push(member['Notes'] || '');
    peopleSheet.appendRow(rowValues);
  }

  var commSheet = ss.getSheetByName('All Community') || ss.insertSheet('All Community');
  commSheet.appendRow([
    member['Area'] || 'Unallocated',
    fullName,
    member['House Number'] || '',
    member['Address'] || '',
    member['Postcode'] || '',
    member['Contact Number'] || '',
    member['Type'] || 'Member',
    initialStatus,
    member['Notes'] || '',
    new Date(),
    member['Occupation'] || '',
    member['Email Address'] || ''
  ]);

  return 'Success';
}

function updateMember(member) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var fullName = (member['First Name'] + ' ' + (member['Surname'] || '')).trim();
  var isMember = (member['Type'] || '').toLowerCase().includes('member') && !(member['Type'] || '').toLowerCase().includes('non');
  
  // 1. Update All Community Sheet
  var commSheet = ss.getSheetByName('All Community');
  if (commSheet && commSheet.getLastRow() > 1) {
    var cData = commSheet.getDataRange().getValues();
    var foundCommRow = -1;
    
    if (member['CommunityRowIndex'] && member['CommunityRowIndex'] > 1) {
      foundCommRow = parseInt(member['CommunityRowIndex']);
    } else {
      for (var r = 1; r < cData.length; r++) {
        var cName = cData[r][1] ? cData[r][1].toString().trim().toLowerCase() : '';
        if (cName === fullName.toLowerCase()) {
          foundCommRow = r + 1;
          break;
        }
      }
    }

    if (foundCommRow !== -1) {
      commSheet.getRange(foundCommRow, 1).setValue(member['Area'] || 'Unallocated');
      commSheet.getRange(foundCommRow, 2).setValue(fullName);
      commSheet.getRange(foundCommRow, 3).setValue(member['House Number'] || '');
      commSheet.getRange(foundCommRow, 4).setValue(member['Address'] || '');
      commSheet.getRange(foundCommRow, 5).setValue(member['Postcode'] || '');
      commSheet.getRange(foundCommRow, 6).setValue(member['Contact Number'] || '');
      commSheet.getRange(foundCommRow, 7).setValue(member['Type'] || 'Member');
      commSheet.getRange(foundCommRow, 9).setValue(member['Notes'] || '');
      commSheet.getRange(foundCommRow, 11).setValue(member['Occupation'] || '');
      commSheet.getRange(foundCommRow, 12).setValue(member['Email Address'] || '');
      
      if (!isMember) {
        commSheet.getRange(foundCommRow, 8).setValue('N/A');
      }
    } else {
      commSheet.appendRow([
        member['Area'] || 'Unallocated',
        fullName,
        member['House Number'] || '',
        member['Address'] || '',
        member['Postcode'] || '',
        member['Contact Number'] || '',
        member['Type'] || 'Member',
        isMember ? 'Unpaid' : 'N/A',
        member['Notes'] || '',
        new Date(),
        member['Occupation'] || '',
        member['Email Address'] || ''
      ]);
    }
  }

  // 2. Update People Sheet if Member
  var peopleSheet = ss.getSheetByName('People') || ss.getSheetByName('people');
  if (peopleSheet && peopleSheet.getLastRow() > 1) {
    var pData = peopleSheet.getDataRange().getValues();
    var pHeaders = pData[0];
    var foundPeopleRow = -1;
    
    if (member['PeopleRowIndex'] && member['PeopleRowIndex'] > 1) {
      foundPeopleRow = parseInt(member['PeopleRowIndex']);
    } else {
      for (var r = 1; r < pData.length; r++) {
        var pName = pData[r][1] ? pData[r][1].toString().trim().toLowerCase() : '';
        if (pName === fullName.toLowerCase()) {
          foundPeopleRow = r + 1;
          break;
        }
      }
    }

    var notesColIdx = findOrGetNotesColumn(pHeaders, peopleSheet) + 1;

    if (isMember) {
      if (foundPeopleRow !== -1) {
        peopleSheet.getRange(foundPeopleRow, 2).setValue(fullName);
        peopleSheet.getRange(foundPeopleRow, 3).setValue(member['House Number'] || '');
        peopleSheet.getRange(foundPeopleRow, 4).setValue(member['Address'] || '');
        peopleSheet.getRange(foundPeopleRow, 5).setValue(member['Postcode'] || '');
        peopleSheet.getRange(foundPeopleRow, 6).setValue(member['Contact Number'] || '');
        peopleSheet.getRange(foundPeopleRow, 7).setValue(member['Email Address'] || '');
        if (notesColIdx) {
          peopleSheet.getRange(foundPeopleRow, notesColIdx).setValue(member['Notes'] || '');
        }
      } else {
        var nextId = peopleSheet.getLastRow();
        var newRowVals = [nextId, fullName, member['House Number'] || '', member['Address'] || '', member['Postcode'] || '', member['Contact Number'] || '', member['Email Address'] || ''];
        while(newRowVals.length < notesColIdx - 1) newRowVals.push('');
        newRowVals.push(member['Notes'] || '');
        peopleSheet.appendRow(newRowVals);
      }
    } else {
      if (foundPeopleRow !== -1) {
        peopleSheet.deleteRow(foundPeopleRow);
      }
    }
  }

  return 'Success';
}

function deleteMemberRecord(peopleRowIndex, communityRowIndex, fullName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if (peopleRowIndex && !isNaN(peopleRowIndex) && peopleRowIndex > 1) {
    try {
      var peopleSheet = ss.getSheetByName('People') || ss.getSheetByName('people');
      if (peopleSheet) {
        peopleSheet.deleteRow(parseInt(peopleRowIndex));
      }
    } catch (e1) {
      console.error("Error deleting from People sheet: " + e1.toString());
    }
  }

  if (communityRowIndex && !isNaN(communityRowIndex) && communityRowIndex > 1) {
    try {
      var commSheet = ss.getSheetByName('All Community');
      if (commSheet) {
        commSheet.deleteRow(parseInt(communityRowIndex));
      }
    } catch (e2) {
      console.error("Error deleting from All Community sheet: " + e2.toString());
    }
  } else if (fullName) {
    try {
      var commSheet = ss.getSheetByName('All Community');
      if (commSheet && commSheet.getLastRow() > 1) {
        var cData = commSheet.getDataRange().getValues();
        for (var r = 1; r < cData.length; r++) {
          var cName = cData[r][1] ? cData[r][1].toString().trim().toLowerCase() : '';
          if (cName === fullName.toLowerCase()) {
            commSheet.deleteRow(r + 1);
            break;
          }
        }
      }
    } catch (e3) {
      console.error("Error deleting fallback from All Community: " + e3.toString());
    }
  }

  return 'Success';
}

function getFilteredCommunityData(selectedAreas, statusFilter) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('All Community') || ss.getSheetByName('All_Community') || ss.getSheetByName('Community');
    if (!sheet) return { error: "'All Community' sheet tab not found." };
    
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return {};

    var headers = data[0];
    var col = { area: 0, fullName: 1, houseNo: 2, street: 3, postcode: 4, phone: 5, type: 6, status: 7, notes: 8, timestamp: 9, occupation: 10, email: 11 };

    for (var c = 0; c < headers.length; c++) {
      var h = headers[c] ? headers[c].toString().trim().toLowerCase() : '';
      if (h.includes('area')) col.area = c;
      else if (h.includes('full name') || h === 'name') col.fullName = c;
      else if (h.includes('house') || h.includes('door')) col.houseNo = c;
      else if (h.includes('email')) col.email = c;
      else if (h.includes('street') || (h.includes('address') && !h.includes('email'))) col.street = c;
      else if (h.includes('postcode')) col.postcode = c;
      else if (h.includes('phone') || h.includes('mobile') || h.includes('contact')) col.phone = c;
      else if (h.includes('type')) col.type = c;
      else if (h.includes('status')) col.status = c;
      else if (h.includes('note')) col.notes = c;
      else if (h.includes('occupation')) col.occupation = c;
    }

    var groupedData = {};

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var areaVal = row[col.area] !== undefined && row[col.area] !== '' ? row[col.area].toString().trim() : 'Unallocated / Blank Area';
      var includeArea = false;

      if (!selectedAreas || selectedAreas.length === 0 || selectedAreas.indexOf('All') !== -1 || selectedAreas.indexOf('all') !== -1) {
        includeArea = true;
      } else if (selectedAreas.indexOf(areaVal) !== -1) {
        includeArea = true;
      }

      if (!includeArea) continue;

      var fullName = row[col.fullName] ? row[col.fullName].toString().trim() : '';
      if (!fullName) continue;

      var parts = fullName.split(/\s+/);
      var memberType = row[col.type] !== undefined ? row[col.type].toString().trim() : 'Member';

      if (statusFilter && statusFilter !== 'All') {
        var isMem = memberType.toLowerCase().includes('member') && !memberType.toLowerCase().includes('non');
        if (statusFilter === 'Member' && !isMem) continue;
        if (statusFilter === 'Non-Member' && isMem) continue;
      }

      var street = row[col.street] !== undefined && row[col.street] !== '' ? row[col.street].toString().trim() : 'General Area';

      var person = {
        id: i + 1,
        firstName: parts[0] || '',
        lastName: parts.slice(1).join(' ') || '',
        fullName: fullName,
        houseNum: row[col.houseNo] !== undefined ? row[col.houseNo].toString().trim() : '',
        street: street,
        postcode: row[col.postcode] !== undefined ? row[col.postcode].toString().trim() : '',
        contact: row[col.phone] !== undefined ? row[col.phone].toString().trim() : '',
        email: (col.email !== undefined && row[col.email] !== undefined && row[col.email] !== '') ? row[col.email].toString().trim() : 'Not Provided',
        occupation: (col.occupation !== undefined && row[col.occupation] !== undefined && row[col.occupation] !== '') ? row[col.occupation].toString().trim() : 'Not Provided',
        status: memberType,
        paymentStatus: row[col.status] !== undefined && row[col.status] !== '' ? row[col.status].toString().trim() : 'Unpaid',
        notes: row[col.notes] !== undefined ? row[col.notes].toString().trim() : '',
        area: areaVal
      };

      if (!groupedData[street]) groupedData[street] = [];
      groupedData[street].push(person);
    }

    for (var st in groupedData) {
      groupedData[st].sort(function(a, b) {
        var numA = parseInt(a.houseNum.replace(/\D/g, '')) || 0;
        var numB = parseInt(b.houseNum.replace(/\D/g, '')) || 0;
        if (numA !== numB) return numA - numB;
        return a.houseNum.localeCompare(b.houseNum, undefined, {numeric: true});
      });
    }

    return groupedData;
  } catch (err) {
    return { error: err.toString() };
  }
}

function submitCommunityUpdate(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pendingSheet = ss.getSheetByName('Pending_Updates') || ss.getSheetByName('Pending Updates');
  if (!pendingSheet) {
    pendingSheet = ss.insertSheet('Pending_Updates');
    pendingSheet.appendRow(['Action Type', 'Community Row', 'Area', 'Full Name', 'House Number', 'Street', 'Postcode', 'Phone', 'Type', 'Status', 'Notes', 'Occupation', 'Email', 'Timestamp']);
  }

  pendingSheet.appendRow([
    payload.actionType || 'Note Update',
    payload.communityRow || '',
    payload.area || '',
    payload.fullName || '',
    payload.houseNum || '',
    payload.street || '',
    payload.postcode || '',
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
    var pData = pendingSheet.getRange(pendingRowIndex, 1, 1, 14).getValues()[0];
    var actionType = pData[0];
    var originalArea = pData[2];
    var communityRow = pData[1];

    var finalArea = updatedData.area || originalArea || 'Unallocated';
    var fullName = updatedData.fullName || '';
    var houseNum = updatedData.houseNum || '';
    var street = updatedData.street || '';
    var postcode = updatedData.postcode || '';
    var phone = updatedData.phone || '';
    var email = updatedData.email || '';
    var occupation = updatedData.occupation || '';
    var type = updatedData.type || 'Member';
    var status = updatedData.status || '';
    var notes = updatedData.notes || '';
    var currentDate = new Date();

    var isMember = type.toLowerCase().includes('member') && !type.toLowerCase().includes('non');
    if (status === 'N/A' || !isMember) status = 'N/A';

    if (actionType === 'New Resident Sign-up' || actionType === 'Add' || !communityRow) {
      var targetSheet = ss.getSheetByName('All Community') || ss.insertSheet('All Community');
      targetSheet.appendRow([
        finalArea,
        fullName,
        houseNum,
        street,
        postcode,
        phone,
        type,
        status,
        notes,
        currentDate,
        occupation,
        email
      ]);

      if (isMember) {
        var peopleSheet = ss.getSheetByName('People') || ss.getSheets()[0];
        var nextId = peopleSheet.getLastRow();
        peopleSheet.appendRow([
          nextId,
          fullName,
          houseNum,
          street,
          postcode,
          phone,
          email,
          notes
        ]);
      }
    } else if (communityRow) {
      var targetSheet = ss.getSheetByName('All Community');
      if (targetSheet) {
        if (actionType === 'Delete') {
          targetSheet.deleteRow(parseInt(communityRow));
        } else {
          var rIdx = parseInt(communityRow);
          targetSheet.getRange(rIdx, 1).setValue(finalArea);
          targetSheet.getRange(rIdx, 2).setValue(fullName);
          targetSheet.getRange(rIdx, 3).setValue(houseNum);
          targetSheet.getRange(rIdx, 4).setValue(street);
          targetSheet.getRange(rIdx, 5).setValue(postcode);
          targetSheet.getRange(rIdx, 6).setValue(phone);
          targetSheet.getRange(rIdx, 7).setValue(type);
          targetSheet.getRange(rIdx, 8).setValue(status);
          targetSheet.getRange(rIdx, 9).setValue(notes);
          targetSheet.getRange(rIdx, 10).setValue(currentDate);
          targetSheet.getRange(rIdx, 11).setValue(occupation);
          targetSheet.getRange(rIdx, 12).setValue(email);

          var peopleSheet = ss.getSheetByName('People') || ss.getSheets()[0];
          var peopleData = peopleSheet.getDataRange().getValues();
          var foundInPeopleRow = -1;
          for (var r = 1; r < peopleData.length; r++) {
            var pName = peopleData[r][1] ? peopleData[r][1].toString().trim().toLowerCase() : '';
            if (pName === fullName.toLowerCase()) {
              foundInPeopleRow = r + 1;
              break;
            }
          }

          if (isMember) {
            if (foundInPeopleRow !== -1) {
              peopleSheet.getRange(foundInPeopleRow, 3).setValue(houseNum);
              peopleSheet.getRange(foundInPeopleRow, 4).setValue(street);
              peopleSheet.getRange(foundInPeopleRow, 5).setValue(postcode);
              peopleSheet.getRange(foundInPeopleRow, 6).setValue(phone);
              peopleSheet.getRange(foundInPeopleRow, 7).setValue(email);
              peopleSheet.getRange(foundInPeopleRow, 8).setValue(notes);
            } else {
              var nextId = peopleSheet.getLastRow();
              peopleSheet.appendRow([nextId, fullName, houseNum, street, postcode, phone, email, notes]);
            }
          }
        }
      }
    }
  }

  pendingSheet.deleteRow(pendingRowIndex);
  return 'Success';
}