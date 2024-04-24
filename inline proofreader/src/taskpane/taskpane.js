Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Office is ready
  }
});

// Function to check grammar using a local server
const checkGrammar = async (text) => {
  try {
    console.log(text);
    // Make a GET request to the grammar checking server
    const response = await axios.get(`http://localhost:8081/checkGrammar`, {
      params: {
        text: text,
      },
    });

    console.log(response.data);
    return response.data;
  } catch (error) {
    console.error("Error:", error);
    return null;
  }
};
// Function to paraphrase text using a local server
const paraphraseText = async (text) => {
  try {
    console.log('Paraphrasing request:', text);
    // Make a POST request to the paraphrasing server
    const response = await axios.post("http://localhost:8081/paraphrase", { text: text });

    console.log('Paraphrasing response:', response.data.paraphrasedTexts);
    return response.data.paraphrasedTexts;
  } catch (error) {
    console.error("Error in paraphrasing:", error);
    console.log(error.response?.data);
    return null;
  }
};

// Function to handle paraphrasing request and display
const paraphraseAndDisplayText = async (isAdvanced) => {
  const suggestionsDiv = document.getElementById("suggestions");
  suggestionsDiv.innerHTML = "<p>Paraphrasing text...</p>";

  Word.run(async (context) => {
    const range = context.document.getSelection();
    range.load("text");
    await context.sync();

    // Check if any text is selected
    if (range.text.trim().length === 0) {
      suggestionsDiv.innerHTML = "<p>Please select text to paraphrase.</p>";
      return;
    }

    // Call the paraphraseText function and display the results
    const paraphrasedTexts = await paraphraseText(range.text);
    if (paraphrasedTexts) {
      displayParaphrasedText(paraphrasedTexts, isAdvanced);
    } else {
      suggestionsDiv.innerHTML = "<p>Error in paraphrasing or no paraphrase available.</p>";
    }
  }).catch((error) => {
    console.error("Error:", error);
    suggestionsDiv.innerHTML = "<p>Error in processing the request.</p>";
  });
};

// Function to display paraphrased text in the suggestions div
const displayParaphrasedText = (paraphrasedTexts, isAdvanced) => {
  const suggestionsDiv = document.getElementById("suggestions");
  suggestionsDiv.innerHTML = "<p><strong>Paraphrased Text:</strong></p>";

  const textsToShow = isAdvanced ? paraphrasedTexts : [paraphrasedTexts[0]];
  textsToShow.forEach((text, index) => {
    const paraphrasedElement = document.createElement("p");
    paraphrasedElement.innerText = text;
    suggestionsDiv.appendChild(paraphrasedElement);
    if (index < textsToShow.length - 1) {
      suggestionsDiv.appendChild(document.createElement("br"));
    }
  });
};

// Event listener for the basic paraphrase button
document.getElementById("basicParaphrase").addEventListener("click", () => {
  paraphraseAndDisplayText(false); // false for basic paraphrasing
});

// Event listener for the advanced paraphrase button
document.getElementById("advancedParaphrase").addEventListener("click", () => {
  paraphraseAndDisplayText(true); // true for advanced paraphrasing
});

// Function to display grammar suggestions in the Word document
const displaySuggestions = async (suggestions) => {
  const suggestionsDiv = document.getElementById("suggestions");
  suggestionsDiv.innerHTML = "";

  if (suggestions.length === 0) {
    suggestionsDiv.innerHTML = "<p>No suggestions found.</p>";
    return;
  }

  await Word.run(async (context) => {
    const docBody = context.document.body;
    docBody.load('text');
    await context.sync();

    for (let suggestion of suggestions) {
      // Extract the text to be highlighted from the document based on the suggestion's offset and length
      const textToHighlight = docBody.text.substr(suggestion.offset, suggestion.length);

      // Search for the text in the document
      const searchResults = docBody.search(textToHighlight, { matchCase: false, matchWholeWord: false });
      context.load(searchResults, 'items');
      await context.sync();

      if (searchResults.items.length > 0) {
        // Highlight the first occurrence of the text
        const itemToHighlight = searchResults.items[0];
        itemToHighlight.font.highlightColor = 'yellow'; // Highlight color
        itemToHighlight.font.color = 'red'; // Text color
      }

      // Display the suggestion in the suggestions div
      const suggestionElement = document.createElement("div");
      suggestionElement.className = "suggestion";
      suggestionElement.innerHTML = `
        <p><strong>${suggestion.message}</strong></p>
        <p>Replace with: ${suggestion.replacements[0].value}</p>
      `;
      suggestionsDiv.appendChild(suggestionElement);
    }

    await context.sync();
  }).catch(error => console.error("Error:", error));
};

// Event listener for the grammar check button
document.getElementById("check").addEventListener("click", async () => {
  const suggestionsDiv = document.getElementById("suggestions");
  suggestionsDiv.innerHTML = "<p>Checking grammar...</p>";

  await Word.run(async (context) => {
    const range = context.document.getSelection();
    range.load("text");
    await context.sync();

    // Call the checkGrammar function and display the results
    const result = await checkGrammar(range.text);
    if (result) {
      if (result.matches && result.matches.length > 0) {
        await displaySuggestions(result.matches);
      } else {
        suggestionsDiv.innerHTML = "<p>No grammar suggestions found.</p>";
      }
    } else {
      suggestionsDiv.innerHTML = "<p>Error checking grammar. Please try again.</p>";
    }
  });
});

// Event listener for the paraphrase button
document.getElementById("paraphrase").addEventListener("click", async () => {
  const suggestionsDiv = document.getElementById("suggestions");
  suggestionsDiv.innerHTML = "<p>Paraphrasing text...</p>";

  Word.run(async (context) => {
    // Get the selected text in the Word document
    const range = context.document.getSelection();
    range.load("text");
    await context.sync();

    // Check if text is selected
    if (range.text.trim().length === 0) {
      suggestionsDiv.innerHTML = "<p>Please select text to paraphrase.</p>";
      return;
    }

    // Call the paraphraseText function
    const paraphrasedResponse = await paraphraseText(range.text);
    if (paraphrasedResponse) {
      // Display the paraphrased text
      suggestionsDiv.innerHTML = "<p><strong>Paraphrased Text:</strong></p>";
      const paraphrasedElement = document.createElement("div");
      paraphrasedElement.className = "paraphrased-text";
      paraphrasedElement.innerHTML = `<p>${paraphrasedResponse}</p>`;
      suggestionsDiv.appendChild(paraphrasedElement);
    } else {
      suggestionsDiv.innerHTML = "<p>No paraphrase available.</p>";
    }
  }).catch((error) => {
    console.error("Error:", error);
    suggestionsDiv.innerHTML = "<p>Error in processing the request.</p>";
  });
});
