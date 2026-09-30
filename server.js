const express = require("express");
const multer = require("multer");
const fs = require("fs");
const path = require("path");
const { execFile } = require("child_process");

const app = express();

const upload = multer({
  dest: "/tmp/uploads"
});

app.get("/", (req, res) => {
  res.json({
    success: true,
    service: "iLovePDF4 Free Conversion Engine",
    status: "online"
  });
});

app.post("/convert", upload.single("file"), (req, res) => {
  if (!req.file) {
    return res.status(400).json({
      success: false,
      error: "No file uploaded."
    });
  }

  const tool = String(req.body.tool || "");
  const input = req.file.path;
  const originalName = req.file.originalname;

  let outputFormat = null;

  if (tool === "word-to-pdf") outputFormat = "pdf";
  if (tool === "powerpoint-to-pdf") outputFormat = "pdf";
  if (tool === "excel-to-pdf") outputFormat = "pdf";

  if (!outputFormat) {
    fs.unlink(input, () => {});

    return res.status(400).json({
      success: false,
      error: "Unsupported conversion tool."
    });
  }

  const outputDir = "/tmp/output";

  fs.mkdirSync(outputDir, {
    recursive: true
  });

  execFile(
    "libreoffice",
    [
      "--headless",
      "--convert-to",
      outputFormat,
      "--outdir",
      outputDir,
      input
    ],
    (error) => {

      if (error) {
        fs.unlink(input, () => {});

        return res.status(500).json({
          success: false,
          error: "LibreOffice conversion failed."
        });
      }

      const files = fs.readdirSync(outputDir);

      if (!files.length) {
        fs.unlink(input, () => {});

        return res.status(500).json({
          success: false,
          error: "Conversion output was not created."
        });
      }

      const outputFile = path.join(
        outputDir,
        files[0]
      );

      res.download(
        outputFile,
        `${path.parse(originalName).name}.${outputFormat}`,
        () => {
          fs.unlink(input, () => {});
          fs.unlink(outputFile, () => {});
        }
      );
    }
  );
});

const PORT = process.env.PORT || 10000;

app.listen(PORT, "0.0.0.0", () => {
  console.log(
    `Conversion engine running on port ${PORT}`
  );
});
