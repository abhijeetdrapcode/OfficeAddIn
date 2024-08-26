let addinClipboard = "";
let currentNumberIndex = 0; // Track the current number index in the line

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("selectNumberingButton").onclick = selectNumbering;
    document.getElementById("selectParagraphButton").onclick = selectParagraph;
  }
});

// Function to select and copy numbering from selected text
async function selectNumbering() {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const paragraphs = selection.paragraphs;
    paragraphs.load("items");

    await context.sync();

    let found = false;

    for (let i = 0; i < paragraphs.items.length; i++) {
      const paragraph = paragraphs.items[i];
      paragraph.load("text");

      await context.sync();

      const numberings = extractAllNumberings(paragraph.text);

      if (numberings.length > 0) {
        if (currentNumberIndex >= numberings.length) {
          currentNumberIndex = 0; // Reset the index if it exceeds the number of numbers
        }

        const numbering = numberings[currentNumberIndex];
        const startIndex = paragraph.text.indexOf(numbering, paragraph.text.indexOf(numbering) + currentNumberIndex);
        const endIndex = startIndex + numbering.length;

        // Get the range for the numbering
        const range = paragraph.getRange();
        const numberingRange = range.expandTo(startIndex, endIndex);

        // Highlight the selected numbering
        numberingRange.font.highlightColor = "#FFFF00"; // Yellow highlight color

        numberingRange.select();

        addinClipboard = numbering;
        console.log(`Numbering selected and copied to add-in's clipboard: ${numbering}`);

        // Update the <h1> tag with the selected number
        const selectedNumberElement = document.getElementById("selectedNumber");
        selectedNumberElement.textContent = ` ${addinClipboard}`;

        // Copy to clipboard using fallback method
        fallbackCopyToClipboard(addinClipboard);

        found = true;
        currentNumberIndex++; // Move to the next number for the next click
        break;
      }
    }

    if (!found) {
      console.log("No numbering found in the current paragraph.");
      // Clear the <h1> tag if no numbering is found
      const selectedNumberElement = document.getElementById("selectedNumber");
      selectedNumberElement.textContent = " None";
    }

    await context.sync();
  }).catch(errorHandler);
}

// Function to select and copy entire paragraph
async function selectParagraph() {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const paragraphs = selection.paragraphs;
    paragraphs.load("items");

    await context.sync();

    if (paragraphs.items.length > 0) {
      const paragraph = paragraphs.items[0];
      paragraph.select();

      paragraph.load("text");
      await context.sync();

      addinClipboard = paragraph.text;
      console.log("Paragraph selected and copied to add-in's clipboard.");

      // Update the <h1> tag with the truncated paragraph text
      const selectedParagraphElement = document.getElementById("selectedNumber");
      selectedParagraphElement.textContent = `${truncateText(addinClipboard, 10)}`;

      // Copy to clipboard using fallback method
      fallbackCopyToClipboard(addinClipboard);
    } else {
      console.log("No paragraph found at the current cursor position.");
    }

    await context.sync();
  }).catch(errorHandler);
}

// Function to extract all numberings from text
function extractAllNumberings(text) {
  const matches = text.match(/\d+(\.\d+)*|\d+/g); // Match all numbering patterns like "1.", "1.2", "1.2.3", etc.
  return matches || [];
}

// Function to truncate text to a specific number of words
function truncateText(text, wordLimit) {
  const words = text.split(" ");
  if (words.length > wordLimit) {
    return words.slice(0, wordLimit).join(" ") + "...";
  }
  return text;
}

// Error handling function
function errorHandler(error) {
  console.error("Error: " + error);
  if (error instanceof OfficeExtension.Error) {
    console.error("Debug info: " + JSON.stringify(error.debugInfo));
  }
}

// Fallback function to copy text to clipboard
function fallbackCopyToClipboard(text) {
  // Create a temporary textarea element
  const textArea = document.createElement("textarea");
  textArea.value = text;

  // Make the textarea invisible to the user
  textArea.style.position = "fixed";
  textArea.style.opacity = 0;

  // Append the textarea to the body
  document.body.appendChild(textArea);

  // Select the text inside the textarea
  textArea.focus();
  textArea.select();

  try {
    // Execute the copy command
    const successful = document.execCommand("copy");
    const msg = successful ? "successful" : "unsuccessful";
    console.log(`Fallback: Copying text command was ${msg}`);
    alert("Text copied to clipboard using fallback method!");
  } catch (err) {
    console.error("Fallback: Could not copy text: ", err);
  }

  // Remove the textarea from the document
  document.body.removeChild(textArea);
}
