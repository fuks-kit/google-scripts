function onSubmit(event) {
  var responseId = event.response.getId();
  Logger.log("responseId=%s", responseId);

  var active = FormApp.getActiveForm();
  var response = active.getResponse(responseId);
  var itemResponses = response.getItemResponses();

  for (var inx = 0; inx < itemResponses.length; inx++) {
    var itemResponse = itemResponses[inx];
    Logger.log('Response to the question "%s" was "%s"',
        itemResponse.getItem().getTitle(),
        itemResponse.getResponse());
  }
}
