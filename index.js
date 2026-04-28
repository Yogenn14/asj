const express = require("express");
const fs = require("fs");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");

const app = express();
app.use(express.json());

app.get("/", (req, res) => {
  res.send("Server running");
});

app.post("/generate", (req, res) => {
  try {
    const data = req.body;

    const defaultAddress = `1146, MUKIM 16, KAMPUNG TELUK,
SUNGAI DUA, 13800 BUTTERWORTH,
PENANG, MALAYSIA.`;

    let companyAddress = "";

    if (
      data.company === "IMPIAN MAWAR SDN BHD" ||
      data.companyAdd === "default"
    ) {
      companyAddress = defaultAddress;
    } else {
      companyAddress = data.address || "";
    }

    const content = fs.readFileSync("template.docx", "binary");
    const zip = new PizZip(content);

    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
    });

    doc.render({
      date: data.date,
      doc_type: data.doc_type,
      doc_number: data.doc_number,
      company: data.company,
      companyAddress,
      attentionTitle: data.attentionTitle,
      attention: data.attention,
      project: data.project,
      items: data.items,
      total: data.total,
      additionalText: data.additionalText
    });

    //doc.render();

    const buffer = doc.getZip().generate({ type: "nodebuffer" });
    fs.writeFileSync("output.docx", buffer);

    res.download("output.docx");

  } catch (error) {
    console.error(error);
    res.status(500).send("Error generating document");
  }
});

const PORT = process.env.PORT || 3000;

app.listen(PORT, "0.0.0.0", () => {
  console.log("Server running on port " + PORT);
});
