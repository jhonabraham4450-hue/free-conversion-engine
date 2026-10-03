async function convertScannedPdfToDocx(inputPath, outputPath) {
  const os = require("os");

  const tempDir = fs.mkdtempSync(
    path.join(os.tmpdir(), "pdf-render-")
  );

  try {
    const prefix = path.join(tempDir, "page");

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
          timeout: 120000,
          maxBuffer: 20 * 1024 * 1024
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

    const sections = [];

    for (const file of files) {
      const imagePath = path.join(tempDir, file);
      const imageBuffer = fs.readFileSync(imagePath);

      const {
        width,
        height
      } = getPngSize(imagePath);

      const pageWidthTwips = Math.round(width * 15);
      const pageHeightTwips = Math.round(height * 15);

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

        children: [
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
        ]
      });
    }

    const doc = new Document({
      sections
    });

    const buffer = await Packer.toBuffer(doc);

    fs.writeFileSync(outputPath, buffer);

  } finally {
    cleanupDirectory(tempDir);
  }
}
