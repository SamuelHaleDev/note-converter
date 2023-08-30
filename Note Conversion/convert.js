const fs = require("fs");
const path = require("path");
const mammoth = require("mammoth");
const cheerio = require("cheerio");
const { JSDOM } = require("jsdom");

function readLocalDocxToHtml(filePath) {
  return mammoth.convertToHtml({ path: filePath })
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

  $("li").each((index, listItem) => {
    const $listItem = $(listItem);
    const isConcept = $listItem.parent().is("ol");
    
    if (isConcept) {
      const conceptContent = $listItem.html();
      const conceptItems = [];

      $listItem.find("li").each((itemIndex, subListItem) => {
        conceptItems.push($(subListItem).html());
      });

      concepts.push({ content: conceptContent, items: conceptItems });
    }
  });

  return concepts;
}

const filePath = "/Users/samuelhale/Downloads/Intro To NLP.docx";

readLocalDocxToHtml(filePath)
  .then((htmlContent) => {
    console.log("Vanilla HTML:", htmlContent);

    const concepts = groupContentByConcepts(htmlContent);

    const dom = new JSDOM(htmlContent);
    const document = dom.window.document;

    const style = document.createElement("style");
    style.innerHTML = `
      .container {
        display: flex;
        flex-direction: column;
        align-items: flex-start;
      }
      .concept {
        border: 1px solid #333;
        padding: 10px;
        margin-bottom: 20px;
      }
      .concept h2 {
        background-color: #333;
        color: #fff;
        padding: 10px;
        margin-bottom: 10px;
      }
    `;
    document.head.appendChild(style);

    const container = document.createElement("div");
    container.className = "container";
    document.body.innerHTML = "";
    document.body.appendChild(container);

    concepts.forEach((concept) => {
      const conceptDiv = document.createElement("div");
      conceptDiv.className = "concept";

      const title = document.createElement("h2");
      title.innerHTML = cheerio.load(concept.content).text();
      title.style.backgroundColor = "#333";
      title.style.color = "#fff";
      title.style.padding = "10px";
      title.style.marginBottom = "10px";
      conceptDiv.appendChild(title);

      concept.items.forEach((item) => {
        const itemParagraph = document.createElement("p");
        itemParagraph.innerHTML = cheerio.load(item).text();
        conceptDiv.appendChild(itemParagraph);
      });

      container.appendChild(conceptDiv);
    });

    const modifiedHtml = document.documentElement.outerHTML;

    const fileNameWithoutExtension = path.parse(filePath).name;
    const outputFilePath = fileNameWithoutExtension + ".html";
    fs.writeFile(outputFilePath, modifiedHtml, (err) => {
      if (err) {
        console.error("Error writing HTML file:", err);
      } else {
        console.log("Modified HTML written to:", outputFilePath);
      }
    });
  })
  .catch((err) => {
    console.error(err);
  });
