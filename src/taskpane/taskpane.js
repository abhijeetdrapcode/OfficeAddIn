let extractedData = [];
let currentIndex = 0;

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("saveNumbersBtn").onclick = saveHierarchicalNumberedParagraphs;
  }
});

async function saveHierarchicalNumberedParagraphs() {
  // Remove the button after being clicked
  document.getElementById("saveNumbersBtn").style.display = "none";

  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const paragraphs = selection.paragraphs;
    paragraphs.load(["text"]);

    await context.sync();

    extractedData = [];
    currentIndex = 0;
    let currentParentNumbering = null;
    let currentSubNumbering = null;

    paragraphs.items.forEach((paragraph) => {
      const paragraphText = paragraph.text.trim();

      const parentNumberingMatch = paragraphText.match(/^(\d+(\.\d+)*)(.*)$/);
      const subNumberingMatch = paragraphText.match(/^([a-zA-Z])\)\s*(.*)$/);
      const subSubNumberingMatch = paragraphText.match(/^\((\w+)\)\s*(.*)$/);

      if (parentNumberingMatch) {
        currentParentNumbering = parentNumberingMatch[1];
        currentSubNumbering = null;
        const paragraphContent = parentNumberingMatch[3].trim();
        extractedData.push({ numbering: currentParentNumbering, paragraphContent });
      } else if (subNumberingMatch && currentParentNumbering) {
        const subNumbering = subNumberingMatch[1];
        currentSubNumbering = `${currentParentNumbering}.${subNumbering}`;
        const paragraphContent = subNumberingMatch[2];
        extractedData.push({ numbering: currentSubNumbering, paragraphContent });
      } else if (subSubNumberingMatch && currentSubNumbering) {
        const subSubNumbering = subSubNumberingMatch[1];
        const paragraphContent = subSubNumberingMatch[2];
        const combinedNumbering = `${currentSubNumbering}.${subSubNumbering}`;
        extractedData.push({ numbering: combinedNumbering, paragraphContent });
      } else {
        extractedData.push({ numbering: null, paragraphContent: paragraphText });
      }
    });

    populateDropdown();
    displayCurrentParagraph();
    showCloseButton();
  });
}

function populateDropdown() {
  const container = document.querySelector(".container");
  const dropdown = document.createElement("select");
  dropdown.id = "paragraphDropdown";
  dropdown.onchange = () => {
    currentIndex = dropdown.selectedIndex;
    displayCurrentParagraph();
  };

  extractedData.forEach((item, index) => {
    const option = document.createElement("option");
    option.value = index;
    option.text = item.numbering || "No Numbering";
    dropdown.appendChild(option);
  });

  container.appendChild(dropdown);
}

function displayCurrentParagraph() {
  const container = document.querySelector(".container");
  const existingContent = document.querySelector(".content");
  if (existingContent) {
    existingContent.remove();
  }

  const contentContainer = document.createElement("div");
  contentContainer.className = "content";

  if (currentIndex < extractedData.length) {
    const item = extractedData[currentIndex];

    const contentWrapper = document.createElement("div");
    contentWrapper.className = "content-wrapper";

    const numberingElement = document.createElement("div");
    numberingElement.className = "numbering-text";
    numberingElement.textContent = item.numbering || "";

    const textElement = document.createElement("div");
    textElement.className = "paragraph-text";
    const truncatedText =
      item.paragraphContent.length > 50 ? item.paragraphContent.substring(0, 50) + "..." : item.paragraphContent;
    textElement.textContent = truncatedText;

    const buttonContainer = document.createElement("div");
    buttonContainer.className = "button-container";

    const copyNumberButton = document.createElement("button");
    copyNumberButton.textContent = "Copy Number";
    copyNumberButton.className = "copy-buttons";
    copyNumberButton.onclick = () => copyToClipboard(item.numbering || "");

    const copyContentButton = document.createElement("button");
    copyContentButton.textContent = "Copy Content";
    copyContentButton.className = "copy-buttons";
    copyContentButton.onclick = () => copyToClipboard(item.paragraphContent);

    buttonContainer.appendChild(copyNumberButton);
    buttonContainer.appendChild(copyContentButton);

    contentWrapper.appendChild(numberingElement);
    contentWrapper.appendChild(textElement);

    contentContainer.appendChild(contentWrapper);
    contentContainer.appendChild(buttonContainer);

    container.appendChild(contentContainer);
  }
}

function showCloseButton() {
  const container = document.querySelector(".container");
  const closeButton = document.createElement("button");
  closeButton.textContent = "Close";
  closeButton.onclick = resetView;
  container.appendChild(closeButton);
}

function copyToClipboard(text) {
  const tempTextarea = document.createElement("textarea");
  tempTextarea.value = text;
  document.body.appendChild(tempTextarea);
  tempTextarea.select();
  document.execCommand("copy");
  document.body.removeChild(tempTextarea);
}

function resetView() {
  const container = document.querySelector(".container");
  container.innerHTML = '<button id="saveNumbersBtn">Save Numbered Paragraph</button>';
  document.getElementById("saveNumbersBtn").onclick = saveHierarchicalNumberedParagraphs;
}
