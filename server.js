async function convertScannedPdfToDocx(inputPath, outputPath) {
  const os = require("os");

  const tempDir = fs.mkdtempSync(
    path.join(os.tmpdir(), "pdf-render-")
  );

  try {
    const prefix = path.join(tempDir, "page");

    console.log("=================================");
    console.log("PDF -> WORD + OCR");
    console.log("Input:", inputPath);
    console.log("=================================");

    // --------------------------------------------------
    // STEP 1: Render PDF pages
    // --------------------------------------------------
    await new Promise((resolve, reject) => {
      execFile(
        "pdftoppm",
        [
          "-png",
          "-r",
          "96",
          inputPath,
          prefix
        ],
        {
          timeout: 180000,
          maxBuffer: 50 * 1024 * 1024
        },
        (error, stdout, stderr) => {
          if (error) {
            reject(
              new Error(
                stderr?.trim() ||
                error.message ||
                "PDF rendering failed."
              )
            );
            return;
          }

          resolve();
        }
      );
    });

    const files = fs
      .readdirSync(tempDir)
      .filter(file => /^page-\d+\.png$/i.test(file))
      .sort((a, b) => {
        const na = parseInt(a.match(/(\d+)/)[1], 10);
        const nb = parseInt(b.match(/(\d+)/)[1], 10);
        return na - nb;
      });

    if (!files.length) {
      throw new Error("No PDF pages were rendered.");
    }

    console.log("Rendered pages:", files.length);

    // --------------------------------------------------
    // PNG size reader
    // --------------------------------------------------
    function getPngSize(filePath) {
      const buffer = fs.readFileSync(filePath);

      if (
        buffer.length < 24 ||
        buffer.readUInt32BE(0) !== 0x89504e47
      ) {
        throw new Error("Invalid PNG output.");
      }

      return {
        width: buffer.readUInt32BE(16),
        height: buffer.readUInt32BE(20)
      };
    }

    // --------------------------------------------------
    // OCR helper
    // --------------------------------------------------
    async function runOCR(imagePath) {
      return new Promise((resolve) => {
        execFile(
          "tesseract",
          [
            imagePath,
            "stdout",
            "-l",
            "ben+eng",
            "--psm",
            "6"
          ],
          {
            timeout: 120000,
            maxBuffer: 20 * 1024 * 1024
          },
          (error, stdout, stderr) => {
            if (error) {
              console.error(
                "OCR failed:",
                stderr || error.message
              );

              resolve("");
              return;
            }

            resolve(
              (stdout || "")
                .replace(/\r/g, "")
                .trim()
            );
          }
        );
      });
    }

    // --------------------------------------------------
    // Create DOCX sections
    // --------------------------------------------------
    const sections = [];

    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      const imagePath = path.join(tempDir, file);

      console.log(
        `Processing page ${i + 1}/${files.length}`
      );

      const imageBuffer = fs.readFileSync(imagePath);

      const {
        width,
        height
      } = getPngSize(imagePath);

      // OCR
      console.log(
        `Running OCR on page ${i + 1}...`
      );

      const ocrText = await runOCR(imagePath);

      console.log(
        `OCR page ${i + 1}:`,
        ocrText
          ? `${ocrText.length} characters`
          : "No text detected"
      );

      // Keep original page appearance.
      const pageWidthTwips = Math.round(width * 15);
      const pageHeightTwips = Math.round(height * 15);

      const children = [];

      // ------------------------------------------------
      // Original rendered page image
      // ------------------------------------------------
      children.push(
        new Paragraph({
          spacing: {
            before: 0,
            after: 0,
            line: 0
          },
          indent: {
            left: 0,
            right: 0,
            firstLine: 0
          },
          children: [
            new ImageRun({
              type: "png",
              data: imageBuffer,
              transformation: {
                width: width,
                height: height
              }
            })
          ]
        })
      );

      // ------------------------------------------------
      // OCR text
      // ------------------------------------------------
      if (ocrText) {
        children.push(
          new Paragraph({
            spacing: {
              before: 100,
              after: 100
            },
            children: [
              new TextRun({
                text: ocrText,
                size: 20
              })
            ]
          })
        );
      }

      sections.push({
        properties: {
          page: {
            size: {
              width: pageWidthTwips,
              height: pageHeightTwips
            },
            margin: {
              top: 0,
              bottom: 0,
              left: 0,
              right: 0,
              header: 0,
              footer: 0,
              gutter: 0
            }
          }
        },
        children: children
      });
    }

    // --------------------------------------------------
    // Create DOCX
    // --------------------------------------------------
    const doc = new Document({
      sections: sections
    });

    const buffer = await Packer.toBuffer(doc);

    if (!buffer || buffer.length < 1000) {
      throw new Error(
        "DOCX output is invalid or empty."
      );
    }

    fs.writeFileSync(outputPath, buffer);

    console.log(
      "PDF -> WORD + OCR completed:",
      outputPath
    );

    console.log(
      "DOCX size:",
      buffer.length,
      "bytes"
    );

  } finally {
    cleanupDirectory(tempDir);
  }
}
