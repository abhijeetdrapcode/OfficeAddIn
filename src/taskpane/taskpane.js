let addinClipboard = "";
let currentNumberIndex = 0;

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("selectNumberingButton").onclick = selectNumbering;
    document.getElementById("selectParagraphButton").onclick = selectParagraph;
  }
});

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
          currentNumberIndex = 0;
        }

        const numbering = numberings[currentNumberIndex];
        const startIndex = paragraph.text.indexOf(numbering, paragraph.text.indexOf(numbering) + currentNumberIndex);
        const endIndex = startIndex + numbering.length;

        const range = paragraph.getRange();
        const numberingRange = range.expandTo(startIndex, endIndex);

        // numberingRange.font.highlightColor = "#FFFFFF";

        numberingRange.select();

        addinClipboard = numbering;
        console.log(`Numbering selected and copied to add-in's clipboard: ${numbering}`);

        const selectedNumberElement = document.getElementById("selectedNumber");
        selectedNumberElement.textContent = ` ${addinClipboard}`;

        fallbackCopyToClipboard(addinClipboard);

        found = true;
        currentNumberIndex++;
        break;
      }
    }

    if (!found) {
      console.log("No numbering found in the current paragraph.");
      const selectedNumberElement = document.getElementById("selectedNumber");
      selectedNumberElement.textContent = " None";
    }

    await context.sync();
  }).catch(errorHandler);
}

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

      const selectedParagraphElement = document.getElementById("selectedNumber");
      selectedParagraphElement.textContent = `${truncateText(addinClipboard, 10)}`;

      fallbackCopyToClipboard(addinClipboard);
    } else {
      console.log("No paragraph found at the current cursor position.");
    }

    await context.sync();
  }).catch(errorHandler);
}

function extractAllNumberings(text) {
  const matches = text.match(/\d+(\.\d+)*|\d+/g);
  return matches || [];
}

function truncateText(text, wordLimit) {
  const words = text.split(" ");
  if (words.length > wordLimit) {
    return words.slice(0, wordLimit).join(" ") + "...";
  }
  return text;
}

function errorHandler(error) {
  console.error("Error: " + error);
  if (error instanceof OfficeExtension.Error) {
    console.error("Debug info: " + JSON.stringify(error.debugInfo));
  }
}

function fallbackCopyToClipboard(text) {
  const textArea = document.createElement("textarea");
  textArea.value = text;

  textArea.style.position = "fixed";
  textArea.style.opacity = 0;

  document.body.appendChild(textArea);

  textArea.focus();
  textArea.select();

  try {
    const successful = document.execCommand("copy");
    const msg = successful ? "successful" : "unsuccessful";
    console.log(`Fallback: Copying text command was ${msg}`);
  } catch (err) {
    console.error("Fallback: Could not copy text: ", err);
  }

  document.body.removeChild(textArea);
}
