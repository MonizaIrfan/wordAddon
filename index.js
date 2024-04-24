const express = require("express");
const axios = require("axios");
const cors = require("cors");

// Create an Express application
const app = express();

// Set up CORS options
const corsOptions = {
  origin: "*",
  methods: "GET,HEAD,PUT,PATCH,POST,DELETE",
};

// Apply CORS middleware and parse JSON in the request body
app.use(cors(corsOptions));
app.use(express.json());

// Endpoint to check grammar using a local grammar checking server
app.get("/checkGrammar", async (req, res) => {
  try {
    console.log('Received checkGrammar request');
    console.log(req.query.text);
    
    // Make a GET request to the local grammar checking server
    const response = await axios.get("http://127.0.0.1:8081/v2/check", {
      params: {
        language: "en-US",
        text: req.query.text,
      },
    });

    console.log("Grammar corrections", response.data);
    // Respond with the grammar checking results
    res.json(response.data);
  } catch (error) {
    console.error("Error:", error);
    // Handle errors and respond with an error status
    res.status(500).json({ error: "Internal Server Error" });
  }
});

// Endpoint to paraphrase text using the AI21 API
app.post("/paraphrase", async (req, res) => {
  try {
    // API key for authentication with the AI21 API
    const apiKey = "AdPQMCrA24GQfJwd3mqFWB0QKHcJyueB"; 
    const textToParaphrase = req.body.text;

    // Make a POST request to the AI21 paraphrasing API
    const response = await axios.post(
      "https://api.ai21.com/studio/v1/paraphrase",
      {
        text: textToParaphrase,
        startIndex: req.body.startIndex || 0,
        style: "general",
      },
      {
        headers: {
          "Content-Type": "application/json",
          "Authorization": `Bearer ${apiKey}`,
        },
      }
    );

    console.log("API Response:", response.data);

    // Process and respond with paraphrased text
    if (response.data && response.data.suggestions && response.data.suggestions.length > 0) {
      const paraphrasedTexts = response.data.suggestions.map(item => item.text);
      console.log("Paraphrased texts:", paraphrasedTexts);
      res.json({ paraphrasedTexts });
    } else {
      // Handle unexpected API response format
      console.error("Unexpected API response format:", response.data);
      res.status(500).json({
        error: "Internal Server Error in Paraphrasing",
        details: "Unexpected API response format",
      });
    }
  } catch (error) {
    console.error("Error in paraphrasing:", error);

    // Log the detailed error response
    if (error.response) {
      console.log("Error response from server:", error.response.data);
    }

    // Handle errors and respond with an error status
    res.status(error.response?.status || 500).json({
      error: "Internal Server Error in Paraphrasing",
      details: error.response?.data || "Unknown error",
    });
  }
});

// Set up the server to listen on port 8081
const port = 8081;
app.listen(port, () => {
  console.log(`Server is running on port ${port}`);
});

// Function to check grammar using a local grammar checking server (standalone)
const checkGrammar = async (text) => {
  try {
    const response = await axios.get("http://127.0.0.1:8081/v2/check", {
      params: {
        language: "en-US",
        text: text,
      },
    });

    return response.data;
  } catch (error) {
    console.error("Error:", error);
  }
};

// Read user input for standalone grammar checking
const readline = require("readline").createInterface({
  input: process.stdin,
  output: process.stdout,
});

readline.question("Enter the text to CheckGrammar: ", async (text) => {
  const result = await checkGrammar(text);
  if (result && result.matches) {
    console.log("Suggestions:", result.matches);
  } else {
    console.log("No suggestions available.");
  }
  readline.close();
});
