async function convertScannedPdfToDocx(inputPath, outputPath) {
  const pdfBytes = new Uint8Array(fs.readFileSync(inputPath));

  const pdf = await pdfjsLib.getDocument({
    data: pdfBytes,
    disableWorker: true
  }).promise;

  const sections = [];

  for (
    let pageNumber = 1;
    pageNumber <= pdf.numPages;
    pageNumber++
  ) {
    const page = await pdf.getPage(pageNumber);

    const scale = 96 / 72;

    const viewport = page.getViewport({
      scale
    });

    const width = Math.ceil(viewport.width);
    const height = Math.ceil(viewport.height);

    const canvas = createCanvas(
      width,
      height
    );

    const context =
      canvas.getContext("2d");

    await page.render({
      canvasContext: context,
      viewport
    }).promise;

    const imageBuffer =
      canvas.toBuffer("image/png");

    const pageWidthTwips =
      Math.round(width * 15);

    const pageHeightTwips =
      Math.round(height * 15);

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

  const buffer =
    await Packer.toBuffer(doc);

  fs.writeFileSync(
    outputPath,
    buffer
  );
}
