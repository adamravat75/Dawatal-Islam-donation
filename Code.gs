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
    
    // Map of all residents (combining People and All Community for unified search & editing)
    var masterDirectory = {};

    // 1. SAFE PEOPLE FETCH (Members with rolling fee statuses)
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
          var key = fullName.toLowerCase();
          
          masterDirectory[key] = {
            'MemberID': i,
            'PeopleRowIndex': i + 1,
            'CommunityRowIndex': null,
            'Full Name': fullName,
            'First Name': parts[0] || '',
            'Surname': parts.slice(1).join(' ') || '',
            'House Number': row[pCol.houseNo] !== undefined && row[pCol.houseNo] !== null ? row[pCol.houseNo].toString().trim() : '',
            'Address': row[pCol.address] !== undefined && row[pCol.address] !== null ? row[pCol.address].toString().trim() : '',
            'Postcode': row[pCol.postcode] !== undefined && row[pCol.postcode] !== null ? row[pCol.postcode].toString().trim() : '',
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

    // 2. SAFE ALL COMMUNITY FETCH (Includes both Members and Non-Members with exact shifted columns)
    try {
      var commSheet = ss.getSheetByName('All Community') || ss.getSheetByName('All_Community');
      if (commSheet && commSheet.getLastRow() > 1) {
        var cData = commSheet.getDataRange().getValues();
        // Columns: A=Area, B=Full Name, C=House No, D=Street, E=Postcode, F=Phone, G=Type, H=Status, I=Notes, J=Date, K=Occupation, L=Email
        for (var i = 1; i < cData.length; i++) {
          var row = cData[i];
          var fullName = row[1] ? row[1].toString().trim() : '';
          if (!fullName) continue;
          
          var key = fullName.toLowerCase();
          var parts = fullName.split(/\s+/);
          var typeVal = row[6] ? row[6].toString().trim() : 'Member';
          
          if (masterDirectory[key]) {
            // Already in People sheet, enrich with community row index & extra fields
            masterDirectory[key]['CommunityRowIndex'] = i + 1;
            masterDirectory[key]['Area'] = row[0] ? row[0].toString().trim() : '';
            masterDirectory[key]['Contact Type'] = typeVal;
            masterDirectory[key]['Occupation'] = row[10] ? row[10].toString().trim() : '';
            if (!masterDirectory[key]['Postcode'] && row[4]) masterDirectory[key]['Postcode'] = row[4].toString().trim();
            if (!masterDirectory[key]['Email Address'] && row[11]) masterDirectory[key]['Email Address'] = row[11].toString().trim();
          } else {
            // Non-member or resident only in All Community
            masterDirectory[key] = {
              'MemberID': 'C-' + i,
              'PeopleRowIndex': null,
              'CommunityRowIndex': i + 1,
              'Full Name': fullName,
              'First Name': parts[0] || '',
              'Surname': parts.slice(1).join(' ') || '',
              'House Number': row[2] !== undefined && row[2] !== null ? row[2].toString().trim() : '',
              'Address': row[3] !== undefined && row[3] !== null ? row[3].toString().trim() : '',
              'Postcode': row[4] !== undefined && row[4] !== null ? row[4].toString().trim() : '',
              'Contact Number': row[5] !== undefined && row[5] !== null ? row[5].toString().trim() : '',
              'Contact Type': typeVal,
              'Notes': row[8] !== undefined && row[8] !== null ? row[8].toString().trim() : '',
              'Area': row[0] !== undefined && row[0] !== null ? row[0].toString().trim() : '',
              'Occupation': row[10] !== undefined && row[10] !== null ? row[10].toString().trim() : '',
              'Email Address': row[11] !== undefined && row[11] !== null ? row[11].toString().trim() : ''
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
            'Phone': r[6] !== undefined ? r[6].toString() : '',
            'Type': r[7] !== undefined ? r[7].toString() : '',
            'Status': r[8] !== undefined ? r[8].toString() : '',
            'Notes': r[9] !== undefined ? r[9].toString() : '',
            'Occupation': r[10] !== undefined ? r[10].toString() : '',
            'Email': r[11] !== undefined ? r[11].toString() : '',
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
        for (var r = 1; r < cData.length; r++) {
          var typeVal = cData[r][6] ? cData[r][6].toString().toLowerCase() : '';
          if (typeVal.includes('member') && !typeVal.includes('non')) {
            commSheet.getRange(r + 1, 8).setValue('Unpaid');
          } else {
            commSheet.getRange(r + 1, 8).setValue('N/A');
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
        for (var i = 1; i < cData.length; i++) {
          var cName = cData[i][1] ? cData[i][1].toString().trim().toLowerCase() : '';
          var typeVal = cData[i][6] ? cData[i][6].toString().toLowerCase() : '';
          
          if (cName.includes(firstName.toLowerCase()) && cName.includes(surname.toLowerCase())) {
            if (typeVal.includes('member') && !typeVal.includes('non')) {
              commSheet.getRange(i + 1, 8).setValue('Paid');
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
        for (var i = 1; i < cData.length; i++) {
          var cName = cData[i][1] ? cData[i][1].toString().trim().toLowerCase() : '';
          var typeVal = cData[i][6] ? cData[i][6].toString().toLowerCase() : '';
          
          if (cName.includes(firstName.toLowerCase()) && cName.includes(surname.toLowerCase())) {
            if (typeVal.includes('member') && !typeVal.includes('non')) {
              commSheet.getRange(i + 1, 8).setValue('Unpaid');
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
      // If not in All Community yet, append it
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

  // 2. Update People Sheet if Member (or add if converted to member)
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
      // If converted from member to non-member, optionally remove from People tab or leave
      if (foundPeopleRow !== -1) {
        peopleSheet.deleteRow(foundPeopleRow);
      }
    }
  }

  return 'Success';
}

function deleteMemberByRow(rowIndex) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var peopleSheet = ss.getSheetByName('People') || ss.getSheets()[0];
  if (!peopleSheet || rowIndex <= 1) return 'Error';
  peopleSheet.deleteRow(rowIndex);
  return 'Success';
}