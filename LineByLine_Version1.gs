/*Huge thank you to Jeff Everhart!
This project is mostly built off of their translation apps script.
Their original code and a helpful tutorial of theirs can be found at the following link:
https://jeffreyeverhart.com/2018/10/05/translate-text-google-doc-using-google-apps-script/
All of their commens will be marked with JE: */
/* Comments with three astericks (***) designate functionality that I haven't added yet but may be nice if you'd like to contribute ot the project*/
/*I've hardcoded in the translation from Spanish to English, not elegant but it works for me lol
To change the original and translate language, edit the lines bellow the comments marked with @@@*/

//this function will be automatically called when you open the Doc
function onOpen() {

//Add a menu item to start the translation when clicked
var ui = DocumentApp.getUi();

ui.createMenu('Translation Tools')
.addItem('Translate This', 'translateSelection')
//additional menu items could be added here, for example if you wanted to allow for the user to modify the languages being translated from and to
.addToUi();
}

function translateSelection() {
  var ui = DocumentApp.getUi();
  //This function depends on the user first highlighting at least a bit of text and then it will translate the entirety of the text that is in the same paragraph
  //***Would be nice to automatically select all the text in a document without the user highlighting text
  var selection = DocumentApp.getActiveDocument().getSelection();
  var rangeElements = selection.getRangeElements();
  //JE: This part is a little bit complicated. Since the selection in a Google Doc returns a range, not text,
  //JE: we have to do some massaging to get the text. Since rangeElements is an array data type, we'll map it
  //JE: to a new array that returns the text. Then, once we have an array with text, we'll join that into a sentence
  //JE: using array methods
  var text = rangeElements.map( function(element) {
    //JE: Here we have to walk down the classes to get to usable text:
    //JE: We go from rangeElement to Element to Text and then get the text as a string
    return element.getElement().asText().getText();
  })

  //initialize an array to store each sentence
  var sentence_array = [];
  /* following explanation came from https://stackoverflow.com/questions/18914629/split-string-into-sentences-in-javascript
   ([.?!]) = capture either . or ? or !
   \s* = capture 0 or more whitespace characters following the previous token ([.?!]).
   (?=[A-Z]) = The previous tokens only match if the next character is a capital letter.
   Find punction marks (./?/!) and capture them
   Punctuation marks can optionally include spaces after them
   After a punctuation mark, we expect a capital letter
   Replace the captured punctuation marks by appending a pip |
   Split the pipes to create an array of sentences*/
  sentence_array = text[0].replace(/([.?!])\s*(?=[A-Z])/g, "$1|").split("|");

  //JE: Here we'll get a user property called target_language
  //JE: PropertiesService is basically a key:value store like localStorage
  var userDefinedLanguage = PropertiesService.getUserProperties().getProperty('target_language')

  //JE: Next we do some checking and if userDefinedLanguage is undefined, we'll set a default
  //@@@ This configuration translates from Spanish to English but you can use any two language codes
  //Google's two letter language codes can be found here: https://cloud.google.com/translate/docs/languages
  var targetLanguage = userDefinedLanguage ? userDefinedLanguage : 'en';
  var translation = LanguageApp.translate(text, 'es', targetLanguage);

  //use a while loop for the translation, YESSS!!! VICTORYYY
  //size 10 font,should fit 70 characters per line

  /*Explanation of the loops below
  - As long we have sentences in the document, keep translating
  - Is the current sentence shorter than 70 characters? Nice, add the sentence and its translation to the document
  - If the current sentence is longer than 70 characters then we have to chop it off after the last full word
    - Start at character position 70 in the current sentence, is it a space? No ... 67? Yes. Okay, add the chopped sentence to the document
    - Delete that substring that we've already added to the document from the current sentence string
    - Go to the translatied sentence, is char 70 a space? No ... 68? Yes. Okay add the chopped translated sentence to the document
    - Delete that substring that we've already added to the document from the current translated sentence string
    - Restart for the next line  */
  var i = 0;
  while(i < sentence_array.length) {
    //@@@ Change the line below with the language code of the original language you want a translation of
    var translation = LanguageApp.translate(sentence_array[i], 'es', targetLanguage);
    var sentence = sentence_array[i]
    if (sentence.length > 70 ){
      while (sentence.length > 0) {
        var j = 0;
        var k = 0;
        if (sentence.length < 70){
          DocumentApp.getActiveDocument().getBody().appendParagraph(sentence.slice(0,70)).editAsText().setItalic(false).setBold(true);
          DocumentApp.getActiveDocument().getBody().appendParagraph(translation.slice(0,70)+'\n').editAsText().setItalic(true).setBold(false);
          sentence = "";
        }
        else {
          //find the space
          while (sentence.charAt(70 - j) != " ") {
            j++;
          }
          DocumentApp.getActiveDocument().getBody().appendParagraph(sentence.slice(0,70 - j)).editAsText().setItalic(false).setBold(true);
          while (translation.charAt(70 - k) != " "){
            k++;
          }
          DocumentApp.getActiveDocument().getBody().appendParagraph(translation.slice(0,70 - k)+'\n').editAsText().setItalic(true).setBold(false);
          sentence = sentence.substring(70-j);
          translation = translation.substring(70-k);
        }
      }
    }
    else {
      DocumentApp.getActiveDocument().getBody().appendParagraph(sentence_array[i]).editAsText().setItalic(false).setBold(true);
      DocumentApp.getActiveDocument().getBody().appendParagraph(translation+'\n').editAsText().setItalic(true).setBold(false);
    }
    i++;
  }
}
