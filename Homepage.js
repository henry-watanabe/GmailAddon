function onDefaultHomePageOpen(e) {
  console.log('onDefaultHomePageOpen:');
  console.log(e);
  var builder = CardService.newCardBuilder();
  builder.setHeader(CardService.newCardHeader().setTitle('Log Your Expense'));
  return builder.build();
}

function getSettingsCard(e) {
  console.log('getSettingsCard:');
  console.log(e);
  var builder = CardService.newCardBuilder();
  builder.setHeader(CardService.newCardHeader().setTitle('Log Your Expense2'));
  return builder.build();
}
/**

 * Create a collection of cards to control the add-on settings and
 * present other information. These cards are displayed in a list when
 * the user selects the associated "Open settings" universal action.
 *
 * @param {Object} e an event object
 * @return {UniversalActionResponse}
 */
function createSettingsResponse(e) {
  return CardService.newUniversalActionResponseBuilder()
      .displayAddOnCards(
          [createSettingCard(), createAboutCard()])
      .build();
}

/**
 * Create and return a built settings card.
 * @return {Card}
 */
function createSettingCard() {
  return CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle('Settings'))
      .addSection(CardService.newCardSection()
          .addWidget(CardService.newSelectionInput()
              .setFieldName("checkBox")
              .setType(CardService.SelectionInputType.CHECK_BOX)
              .addItem("Ask before deleting contact", "contact", false)
              .addItem("Ask before deleting cache", "cache", false)
              .addItem("Preserve contact ID after deletion", "contactId", false))
           .addWidget(CardService.newTextInput()
             .setFieldName("text")
             .setTitle("search key"))
          // ... continue adding widgets or other sections here ...
      ).build();   // Don't forget to build the card!
}

/**
 * Create and return a built 'About' informational card.
 * @return {Card}
 */
function createAboutCard() {
  return CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle('About'))
      .addSection(CardService.newCardSection()
          .addWidget(CardService.newTextParagraph()
              .setText('This add-on manages contact information. For more '
                  + 'details see the <a href="https://www.example.com/help">'
                  + 'help page</a>.'))
      // ... add other information widgets or sections here ...
      ).build();  // Don't forget to build the card!
}
