import AdmZip from "adm-zip";
import fs from "fs";
import xml2js from "xml2js";
import xpath from "xpath";
import { DOMParser } from "xmldom";

const inputFileName = "input.pptx";
const outputFileName = "output.pptx";

// Read the existing PowerPoint file
const pptxBuffer = fs.readFileSync(inputFileName);
const zip = new AdmZip(pptxBuffer);

// Extract XML content from slides
const slides = zip.getEntries().filter((entry) => entry.name.startsWith("p"));

slides.forEach((slideEntry) => {
  const xmlContent = zip.readAsText(slideEntry);

  //   console.log("xmlContent", xmlContent);

  //   selectAllTextElements(xmlContent).then((textElements) => {
  //     console.log("textElements", textElements);
  //   });

  // Parse XML content
  xml2js.parseString(xmlContent, { explicitArray: false }, (err, result) => {
    if (err) {
      console.error("Error parsing XML:", err);
      return;
    }

    const textElements = selectElements(result, "//p:sp//p:txBody//a:t");

    console.log("textElements", textElements);

    // console.log("result", result);

    // Modify the XML structure to remove elements containing the specified text
    // For simplicity, let's assume the text to remove is 'TextToRemove'
    // removeTextElements(result);

    // Convert the modified XML structure back to a string
    const modifiedXmlContent = new xml2js.Builder().buildObject(result);

    // Update the slide in the zip file
    zip.updateFile(slideEntry.name, Buffer.from(modifiedXmlContent, "utf-8"));
  });
});

// Save the modified PowerPoint file
zip.writeZip(outputFileName);
console.log("Modified PowerPoint file saved successfully.");

// Function to remove text elements from the XML structure
function removeTextElements(xmlObject) {
  //   console.log("xmlObject", xmlObject);
  if (xmlObject && typeof xmlObject === "object") {
    for (const key in xmlObject) {
      //   console.log("found text", key);
      if (key === "_") {
        // Check and modify text content
        if (xmlObject[key].includes("Lesson 1")) {
          xmlObject[key] = xmlObject[key].replace(/TextToRemove/g, "");
        }
      } else if (Array.isArray(xmlObject[key])) {
        // Recursively process arrays of objects
        xmlObject[key].forEach((element) => {
          //   console.log("removing element", element);
          removeTextElements(element);
        });
      } else if (typeof xmlObject[key] === "object") {
        // Recursively process nested objects
        removeTextElements(xmlObject[key]);
      }
    }
  } else {
    console.log("no object");
  }
}

// // Utility function to select all text elements in a PowerPoint XML
// async function selectAllTextElements(xmlContent) {
//     return new Promise((resolve, reject) => {
//       xml2js.parseString(xmlContent, { explicitArray: false }, (err, result) => {
//         if (err) {
//           reject(err);
//         } else {
//           const textElements = selectElements(result, '//p:sp//p:txBody//a:t');
//           resolve(textElements);
//         }
//       });
//     });
//   }

//   // Utility function to select elements using XPath expressions
function selectElements(xmlObject, xpathExpression) {
  const builder = new xml2js.Builder();
  const xmlDoc = new DOMParser().parseFromString(
    builder.buildObject(xmlObject),
    "text/xml"
  );
  return xpath.select(xpathExpression, xmlDoc);
}

//   // Example usage
//   const xmlContent = /* your XML content here */;

//   selectAllTextElements(xmlContent)
//     .then(textElements => {
//       console.log('All text elements:', textElements);
//     })
//     .catch(error => {
//       console.error('Error:', error);
//     });
