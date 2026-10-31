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

function getDatabase() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      return { success: false, error: "Active spreadsheet not found.", members: [], donations: [], pendingUpdates: [] };
    }
    
    // Check and perform Year-End reset check if needed
    checkAndResetForNewYear(ss);
    
    // 1. SAFE PEOPLE FETCH
    var members = [];
    try {
      var peopleSheet = ss.getSheetByName('People') || ss.getSheetByName('people') || ss.getSheets()[0];
      if (peopleSheet && peopleSheet.getLastRow() > 1) {
        var peopleData = peopleSheet.getDataRange().getValues();
        var pHeaders = peopleData[0];
        for (var i = 1; i < peopleData.length; i++) {
          var row = peopleData[i];
          var fullName = row[1] ? row[1].toString().trim() : (row[0] ? row[0].toString().trim() : '');
          if (!fullName) continue;
          var parts = fullName.split(/\s+/);
          var member = {
            'MemberID': i,
            'RowIndex': i + 1,
            'Full Name': fullName,
            'First Name': parts[0] || '',
            'Surname': parts.slice(1).join(' ') || '',
            'House Number': row[2] !== undefined && row[2] !== null ? row[2].toString().trim() : '',
            'Address': row[3] !== undefined && row[3] !== null ? row[3].toString().trim() : '',
            'Postcode': row[4] !== undefined && row[4] !== null ? row[4].toString().trim() : '',
            'Contact Number': row[5] !== undefined && row[5] !== null ? row[5].toString().trim() : '',
            'Email Address': row[6] !== undefined && row[6] !== null ? row[6].toString().trim() : '',
            'Notes': row[7] !== undefined && row[7] !== null ? row[7].toString().trim() : ''
          };
          for (var c = 0; c < pHeaders.length; c++) {
            var headerName = pHeaders[c] ? pHeaders[c].toString().trim() : '';
            if (headerName && headerName.toLowerCase().includes('membership')) {
              var val = row[c];
              member[headerName] = (val !== undefined && val !== null && val !== '') ? val.toString().trim() : 'Unpaid';
            }
          }
          members.push(member);
        }
      }
    } catch (errP) {
      console.error("People read error: " + errP.toString());
    }

    // 2. SAFE DONATIONS FETCH
    var donations = [];
    try {
      var donSheet = ss.getSheetByName('Donations') || ss.getSheetByName('Donations Log');
      if (donSheet && donSheet.getLastRow() > 1) {
        var donData = donSheet.getDataRange().getValues();
        for (var i = 1; i < donData.length; i++) {
          var row = donData[i];
          var dateVal = row[0] instanceof Date ? row[0].toISOString().split('T')[0] : (row[0] ? row[0].toString().trim() : '');
          var firstName = row[1] ? row[1].toString().trim() : '';
          var surname = row[2] ? row[2].toString().trim() : '';
          var typeVal = row[3] ? row[3].toString().trim() : '';
          var rawAmt = row[4] !== undefined && row[4] !== null ? row[4].toString().replace('£', '').trim() : '0';
          var amount = parseFloat(rawAmt);
          if (isNaN(amount)) amount = 0;
          
          donations.push({
            'RowIndex': i + 1,
            'Date Paid': dateVal,
            'First Name': firstName,
            'Surname': surname,
            'Donation Type': typeVal,
            'Amount': amount,
            'Payment Method': row[5] !== undefined && row[5] !== null ? row[5].toString().trim() : 'Cash',
            'Gift Aided?': row[6] !== undefined && row[6] !== null ? row[6].toString().trim() : 'No',
            'House Number': row[7] !== undefined && row[7] !== null ? row[7].toString().trim() : '',
            'Postcode': row[8] !== undefined && row[8] !== null ? row[8].toString().trim() : '',
            'Street': row[9] !== undefined && row[9] !== null ? row[9].toString().trim() : '',
            'Book Number': row[10] !== undefined && row[10] !== null ? row[10].toString().trim() : '',
            'Invoice Number': row[11] !== undefined && row[11] !== null ? row[11].toString().trim() : ''
          });
        }
      }
    } catch (errD) {
      console.error("Donations read error: " + errD.toString());
    }

    // 3. SAFE PENDING UPDATES FETCH
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

// Helper to check year-end reset: resets Members to Unpaid, keeps Non-Members as N/A
function checkAndResetForNewYear(ss) {
  var props = PropertiesService.getDocumentProperties();
  var currentYear = new Date().getFullYear().toString();
  var lastCheckedYear = props.getProperty('LAST_ACTIVE_YEAR');

  if (lastCheckedYear && lastCheckedYear !== currentYear) {
    try {
      // 1. Reset People sheet membership statuses for Members
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

      // 2. Reset All Community sheet: Members become 'Unpaid', Non-Members remain 'N/A'
      var commSheet = ss.getSheetByName('All Community');
      if (commSheet && commSheet.getLastRow() > 1) {
        var cData = commSheet.getDataRange().getValues();
        for (var r = 1; r < cData.length; r++) {
          var typeVal = cData[r][5] ? cData[r][5].toString().toLowerCase() : ''; // Col F: Type
          if (typeVal.includes('member') && !typeVal.includes('non')) {
            commSheet.getRange(r + 1, 7).setValue('Unpaid'); // Col G: Status -> Unpaid for Members
          } else {
            commSheet.getRange(r + 1, 7).setValue('N/A'); // Col G: Status -> N/A for Non-Members
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
  
  donSheet.appendRow([
    donation['Date Paid'],
    donation['First Name'],
    donation['Surname'],
    donation['Donation Type'],
    donation['Amount'],
    donation['Payment Method'],
    donation['Gift Aided?'],
    donation['House Number'] || '',
    donation['Postcode'] || '',
    donation['Street'] || '',
    donation['Book Number'] || '',
    donation['Invoice Number'] || ''
  ]);
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
  
  // 1. Add to People sheet
  var peopleSheet = ss.getSheetByName('People') || ss.getSheets()[0];
  var nextId = peopleSheet.getLastRow();
  var fullName = (member['First Name'] + ' ' + (member['Surname'] || '')).trim();
  
  peopleSheet.appendRow([
    nextId,
    fullName,
    member['House Number'] || '',
    member['Address'] || '',
    member['Postcode'] || '',
    member['Contact Number'] || '',
    member['Email Address'] || '',
    member['Notes'] || ''
  ]);

  // 2. Automatically sync and add to All Community sheet with exact columns A through K
  var isMemberType = member['Type'] ? member['Type'].toString().toLowerCase().includes('member') && !member['Type'].toString().toLowerCase().includes('non') : true;
  var initialStatus = isMemberType ? 'Unpaid' : 'N/A';

  var commSheet = ss.getSheetByName('All Community') || ss.insertSheet('All Community');
  commSheet.appendRow([
    member['Area'] || 'Unallocated',                // Col A: Area
    fullName,                                       // Col B: Full Name
    member['House Number'] || '',                   // Col C: House No
    member['Address'] || '',                        // Col D: Street
    member['Contact Number'] || '',                 // Col E: Phone
    member['Type'] || 'Member',                     // Col F: Type (Member / Non-Member)
    initialStatus,                                  // Col G: Status (Unpaid or N/A)
    member['Notes'] || '',                          // Col H: Notes
    new Date(),                                     // Col I: Timestamp / Date
    member['Occupation'] || '',                     // Col J: Occupation
    member['Email Address'] || ''                   // Col K: Email Address
  ]);

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
  if (member['Notes'] !== undefined) {
    peopleSheet.getRange(rowIndex, 8).setValue(member['Notes']);
  }

  // Check if membership fee status was updated (e.g., Paid) and sync to All Community
  var currentYear = new Date().getFullYear();
  var pHeaders = peopleSheet.getRange(1, 1, 1, peopleSheet.getLastColumn()).getValues()[0];
  var statusToSync = '';
  for (var c = 0; c < pHeaders.length; c++) {
    var h = pHeaders[c] ? pHeaders[c].toString().toLowerCase() : '';
    if (h.includes('membership') || h.includes(currentYear)) {
      statusToSync = peopleSheet.getRange(rowIndex, c + 1).getValue().toString();
    }
  }

  if (statusToSync) {
    var commSheet = ss.getSheetByName('All Community');
    if (commSheet && commSheet.getLastRow() > 1) {
      var cData = commSheet.getDataRange().getValues();
      for (var r = 1; r < cData.length; r++) {
        var cName = cData[r][1] ? cData[r][1].toString().trim().toLowerCase() : '';
        if (cName === fullName.toLowerCase()) {
          var typeVal = cData[r][5] ? cData[r][5].toString().toLowerCase() : '';
          // Only update status if they are a member; keep N/A for non-members
          if (typeVal.includes('member') && !typeVal.includes('non')) {
            commSheet.getRange(r + 1, 7).setValue(statusToSync); // Col G: Status
          } else {
            commSheet.getRange(r + 1, 7).setValue('N/A');
          }
          break;
        }
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

function getFilteredCommunityData(selectedAreas, statusFilter) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('All Community') || ss.getSheetByName('All_Community') || ss.getSheetByName('Community');
    if (!sheet) sheet = ss.getSheets()[0];
    
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return {};

    var headers = data[0];
    var col = { area: 0, fullName: 1, houseNo: 2, street: 3, contact: 4, type: 5, status: 6, notes: 7, date: 8, occupation: 9, email: 10 };

    for (var c = 0; c < headers.length; c++) {
      var h = headers[c] ? headers[c].toString().trim().toLowerCase() : '';
      if (h.includes('full name') || h === 'name' || h.includes('resident')) col.fullName = c;
      else if (h.includes('house')) col.houseNo = c;
      else if ((h.includes('street') || h.includes('address')) && !h.includes('email')) col.street = c;
      else if (h.includes('phone') || h.includes('contact') || h.includes('mobile')) col.contact = c;
      else if (h.includes('email')) col.email = c;
      else if (h.includes('occupation')) col.occupation = c;
      else if (h.includes('type') || h.includes('member/non')) col.type = c;
      else if (h.includes('status') || h.includes('payment')) col.status = c;
      else if (h.includes('note')) col.notes = c;
      else if (h.includes('date') || h.includes('timestamp')) col.date = c;
      else if (h.includes('area')) col.area = c;
    }

    var groupedData = {};

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var fullName = row[col.fullName] !== undefined ? row[col.fullName].toString().trim() : '';
      if (!fullName) continue;

      var rowArea = row[col.area] !== undefined && row[col.area] !== '' ? row[col.area].toString().trim() : 'Unallocated';
      if (!rowArea.toLowerCase().includes('area') && rowArea !== 'Unallocated') {
        rowArea = 'Area ' + rowArea;
      }

      if (selectedAreas && !selectedAreas.includes('All')) {
        var matchFound = selectedAreas.some(function(sel) {
          return sel.toLowerCase() === rowArea.toLowerCase() || 
                 rowArea.toLowerCase().includes(sel.toLowerCase()) || 
                 sel.toLowerCase().includes(rowArea.toLowerCase());
        });
        if (!matchFound) continue;
      }

      var statusType = row[col.type] !== undefined ? row[col.type].toString().trim() : 'Member';
      if (statusFilter && statusFilter !== 'All') {
        if (statusFilter === 'Member' && !statusType.toLowerCase().includes('member')) continue;
        if (statusFilter === 'Non-Member' && !statusType.toLowerCase().includes('non')) continue;
      }

      var streetName = row[col.street] !== undefined && row[col.street] !== '' ? row[col.street].toString().trim() : (rowArea + ' (General)');

      if (!groupedData[streetName]) groupedData[streetName] = [];

      groupedData[streetName].push({
        id: i + 1,
        area: rowArea,
        fullName: fullName,
        houseNum: row[col.houseNo] !== undefined ? row[col.houseNo].toString().trim() : '',
        street: streetName,
        contact: row[col.contact] !== undefined ? row[col.contact].toString().trim() : '',
        email: row[col.email] !== undefined ? row[col.email].toString().trim() : '',
        occupation: row[col.occupation] !== undefined ? row[col.occupation].toString().trim() : '',
        status: statusType,
        paymentStatus: row[col.status] !== undefined && row[col.status] !== '' ? row[col.status].toString().trim() : 'Unpaid',
        notes: row[col.notes] !== undefined ? row[col.notes].toString().trim() : ''
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
    var originalArea = pData[2];
    var communityRow = pData[1];

    var finalArea = updatedData.area || originalArea || 'Unallocated';
    var fullName = updatedData.fullName || '';
    var houseNum = updatedData.houseNum || '';
    var street = updatedData.street || '';
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
        finalArea,    // Col A (1)
        fullName,     // Col B (2)
        houseNum,     // Col C (3)
        street,       // Col D (4)
        phone,        // Col E (5)
        type,         // Col F (6)
        status,       // Col G (7)
        notes,        // Col H (8)
        currentDate,  // Col I (9)
        occupation,   // Col J (10)
        email         // Col K (11)
      ]);

      if (isMember) {
        var peopleSheet = ss.getSheetByName('People') || ss.getSheets()[0];
        var nextId = peopleSheet.getLastRow();
        peopleSheet.appendRow([
          nextId,
          fullName,
          houseNum,
          street,
          '', 
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
          targetSheet.getRange(rIdx, 5).setValue(phone);
          targetSheet.getRange(rIdx, 6).setValue(type);
          targetSheet.getRange(rIdx, 7).setValue(status);
          targetSheet.getRange(rIdx, 8).setValue(notes);
          targetSheet.getRange(rIdx, 9).setValue(currentDate);
          targetSheet.getRange(rIdx, 10).setValue(occupation);
          targetSheet.getRange(rIdx, 11).setValue(email);

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
              peopleSheet.getRange(foundInPeopleRow, 6).setValue(phone);
              peopleSheet.getRange(foundInPeopleRow, 7).setValue(email);
              peopleSheet.getRange(foundInPeopleRow, 8).setValue(notes);
            } else {
              var nextId = peopleSheet.getLastRow();
              peopleSheet.appendRow([nextId, fullName, houseNum, street, '', phone, email, notes]);
            }
          }
        }
      }
    }
  }

  pendingSheet.deleteRow(pendingRowIndex);
  return 'Success';
}