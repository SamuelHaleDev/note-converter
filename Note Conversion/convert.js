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
  const $ = cheerio.load(htmlContent);

  // Add the .container class to the outermost ol
  $("ol").first().addClass("container");

  // Add the .concept class to the first li elements directly under the first ol
  $("ol.container > li").each(function () {
    // Generate a random background color using the getRandomColor function
    const conceptBackgroundColor = getRandomColor();

    $(this).addClass("concept");

    // Set the background color for the concept div
    $(this).css("background-color", conceptBackgroundColor);

    // Add h2 for the concept title
    const conceptTitle = $(this).contents().first().text();
    // remove the concept title from the li element
    $(this).contents().first().remove();
    $(this).prepend(`<h2>${conceptTitle}</h2>`);

    // Wrap the nested ol in a div with the .items-container class
    $(this).find("> ol").wrap("<div class='items-container'></div>");

    // Add the .item class to the li elements under the nested ol
    $(this).find(".items-container ol li").addClass("item");
  });

  return $.html();
}

const filePath = "/Users/samuelhale/Downloads/Untitled document.docx";

(async () => {
  const htmlContent = await readLocalDocxToHtml(filePath);
  console.log("Vanilla HTML:", htmlContent);

  const styledHtmlContent = groupContentByConcepts(htmlContent);
  console.log("Styled HTML:", styledHtmlContent);

  const browser = await puppeteer.launch();
  const page = await browser.newPage();

  const dom = new JSDOM(styledHtmlContent);
  const document = dom.window.document;

  const $ = cheerio.load(styledHtmlContent);

  $("head").append(`
    <link href="https://fonts.googleapis.com/css2?family=Oswald:wght@500&display=swap" rel="stylesheet">
    <link href="https://fonts.googleapis.com/css2?family=Poppins:wght@400;500&display=swap" rel="stylesheet">
  `);

  const style = document.createElement("style");
  // Add Oswald to p tag and center it make it white and large as it is the header 1 of the page
  style.innerHTML = `
  p {
    font-family: 'Oswald', sans-serif;
    text-align: center;
    color: #fff;
    font-size: 30px; 
  }
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
  li {
    list-style: none;
  }
  ol {
    padding: 0;
  }
  `;
  document.head.appendChild(style);

  let queryVariable = cheerio.load(styledHtmlContent);
  const title = document.createElement("h1");
  title.innerHTML = queryVariable("p").prop("innerHTML");
  // place the title at the top of the body
  // remove document .body.firstChild
  document.body.firstChild.remove();
  document.body.insertBefore(title, document.body.firstChild);

  const container = document.createElement("div");
  container.className = "container";
  document.body.appendChild(container);

  console.log("Styled HTML:", document.documentElement.outerHTML)

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