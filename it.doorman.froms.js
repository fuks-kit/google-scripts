function onSubmit(event) {
  const responseId = event.response.getId();
  Logger.log("responseId=%s", responseId);

  const active = FormApp.getActiveForm();
  const response = active.getResponse(responseId);
  const itemResponses = response.getItemResponses();

  const email = response.getRespondentEmail();

  for (var inx = 0; inx < itemResponses.length; inx++) {
    const itemResponse = itemResponses[inx];
    const item = itemResponse.getItem();
    const title = item.getTitle();

    if (title == "KIT Chipnummer") {
      const chipnumber = itemResponse.getResponse();

      Logger.log("email='%s' chipnumber='%s'", email, chipnumber);
      setChipNumber(email, chipnumber);

      break;
    }
  }
}

/**
 * @param {string} email
 * @param {string} chipnumber
 */
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

  Logger.log("update", {
      "email": email,
      "chipnumber": chipnumber,
      "customSchemas": userdata.customSchemas,
  });
}
