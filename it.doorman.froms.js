function onChange(event) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("data");

  const index = sheet.getLastRow();
  console.log("index", index);

  const results = sheet.getRange(index, 1, 1, 3).getValues();
  console.log("result", results);

  for (let inx = 0; inx < results.length; inx++) {
    setChipNumber(results[inx][1], results[inx][2]);
  }

  sheet.deleteRow(index);
}

function setChipNumber(email, chipnumber) {
  const userdata = AdminDirectory.Users.get(email, {
    "projection": "full",
  });

  if (!userdata.hasOwnProperty("customSchemas")) {
    userdata.customSchemas = {};
  }

  userdata.customSchemas["fuks"] = {
    "KIT_Card_Chipnummer": chipnumber
  };

  AdminDirectory.Users.update(userdata, email);

  console.log("update", {
    "email": email,
    "chipnumber": chipnumber,
    "customSchemas": userdata.customSchemas,
  });
}
