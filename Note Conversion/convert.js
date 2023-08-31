const fs = require("fs");
const path = require("path");
const mammoth = require("mammoth");
const cheerio = require("cheerio");
const { JSDOM } = require("jsdom");
const puppeteer = require("puppeteer");

function getRandomColor() {
  const letters = "0123456789ABCDEF";
  let color = "#";
  for (let i = 0; i < 6; i++) {
    color += letters[Math.floor(Math.random() * 16)];
  }
  return color;
}

function readLocalDocxToHtml(filePath) {
  return mammoth
    .convertToHtml({ path: filePath })
    .then((result) => {
      const html = result.value;
      console.log("Extracted HTML:", html);
      return html;
    })
    .catch((err) => {
      console.log("Error reading .docx file:", err);
      throw err;
    });
}

function groupContentByConcepts(htmlContent) {
  const concepts = [];

  const $ = cheerio.load(htmlContent);

  // grab outer OL tag
  const ol = $("ol")[0];
  ol.children.forEach((child) => {
    // handle these concepts
    let parts = $(child).prop("innerHTML").split("<ol>");
    // Each child is a list item
    // Grab the text of the first list item only as that is the concept title
    // Loop through and turn all of the nested list items into p tags
    // Add the concept to the concepts array
    const concept = {
      title: parts[0].replace("\t", ""),
      items: [],
    };
    $(child)
      .find("li")
      .each((index, listItem) => {
        // Check if listItem.parent is in concept.items
        // If it is replace the text from the parent with the text from the child with empte string
        // If it is not add the text from the child to the concept.items array
        const parent = $(listItem).parents("li")[0];
        const parentText = $(parent).text();
        const childText = $(listItem).text();
        if (concept.items.includes(parentText)) {
          concept.items.push(childText);
          concept.items = concept.items.filter((item) => item !== parentText);
        } else {
          concept.items.push(childText);
        }
      });
    concepts.push(concept);
  });

  return concepts;
}

const filePath = "C:/Users/wizar/Downloads/Intro To Senior Design.docx";

(async () => {
  const htmlContent = await readLocalDocxToHtml(filePath);
  console.log("Vanilla HTML:", htmlContent);

  const concepts = groupContentByConcepts(htmlContent);

  const browser = await puppeteer.launch();
  const page = await browser.newPage();

  const dom = new JSDOM(htmlContent);
  const document = dom.window.document;

  const $ = cheerio.load(htmlContent);

  $("head").append(`
    <link href="https://fonts.googleapis.com/css2?family=Oswald:wght@500&display=swap" rel="stylesheet">
    <link href="https://fonts.googleapis.com/css2?family=Poppins:wght@400;500&display=swap" rel="stylesheet">
  `);

  const style = document.createElement("style");
  style.innerHTML = `
  h1 {
    color: #fff;
    font-weight: 500;
    font-family: 'Oswald', sans-serif;
    text-align: center;
  }
  body {
    background-color: #333;
  }
  .container {
    display: flex;
    flex-direction: column;
    align-items: center;
  }
  .concept {
    border: 1px solid #333;
    padding: 20px;
    margin-bottom: 20px;
    background-color: #f0f0f0;
    text-align: center;
    border-radius: 5px;
    max-width: 800px;
  }
  .concept h2 {
    margin: 0;
    padding: 10px;
    background-color: #333;
    color: #fff;
    font-weight: 500;
    font-family: 'Oswald', sans-serif;
  }
  .items-container {
    display: flex;
    flex-wrap: wrap;
    justify-content: center;
  }
  .item {
    margin: 5px;
    border: 1px solid #ccc;
    padding: 10px;
    background-color: #fff;
    color: #333;
    font-weight: 400;
    font-family: 'Poppins', sans-serif;
  }
  `;
  document.head.appendChild(style);

  document.body.innerHTML = "";
  let queryVariable = cheerio.load(htmlContent);
  const title = document.createElement("h1");
  title.innerHTML = queryVariable("p").prop("innerHTML");
  document.body.appendChild(title);

  const container = document.createElement("div");
  container.className = "container";
  document.body.appendChild(container);

  concepts.forEach((concept) => {
    const conceptDiv = document.createElement("div");
    conceptDiv.className = "concept";
    conceptDiv.style.backgroundColor = getRandomColor();

    const title = document.createElement("h2");
    title.innerHTML = cheerio.load(concept.title).text();
    conceptDiv.appendChild(title);

    const itemsContainer = document.createElement("div");
    itemsContainer.className = "items-container";

    concept.items.forEach((item) => {
      const itemDiv = document.createElement("div");
      itemDiv.className = "item";
      itemDiv.innerHTML = cheerio.load(item).text();
      itemsContainer.appendChild(itemDiv);
    });

    conceptDiv.appendChild(itemsContainer);
    container.appendChild(conceptDiv);
  });

  await page.setContent(document.documentElement.outerHTML);

  const pdfOptions = {
    path: filePath.replace(".docx", ".pdf"), // Output PDF file path
    format: "A4", // Paper format
    printBackground: true, // Make background transparent
  };
  await page.pdf(pdfOptions);

  await browser.close();
})().catch((err) => {
  console.error(err);
});
