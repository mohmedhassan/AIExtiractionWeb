const fs = require('fs');
const path = require('path');
const { createWorker } = require('tesseract.js');
const poppler = require('pdf-poppler');
const PDFDocument = require('pdfkit');
const { loadImage } = require('canvas');

const pdfPath = "sample.pdf";
const tempFolder = "pages";
const outputPdf = "searchable.pdf";
const arabicFont = "Tajawal-Regular.ttf"; // ضع خط عربي TTF هنا

if (!fs.existsSync(tempFolder)) fs.mkdirSync(tempFolder);

async function pdfToSearchablePDF() {
  try {
    console.log("🔹 Converting PDF to images...");
    await poppler.convert(pdfPath, {
      format: 'png',
      out_dir: tempFolder,
      out_prefix: 'page',
      page: null
    });

    const files = fs.readdirSync(tempFolder)
                    .filter(f => f.endsWith('.png'))
                    .sort();

    console.log(`🔹 Found ${files.length} pages`);

    // إنشاء PDF جديد
    const doc = new PDFDocument({ autoFirstPage: false });
    doc.pipe(fs.createWriteStream(outputPdf));

    // إعداد Tesseract worker
    const worker = await createWorker({
      logger: m => console.log(m)
    });
    await worker.load();
    await worker.loadLanguage('ara');
    await worker.initialize('ara');

    for (const file of files) {
      const imagePath = path.join(tempFolder, file);
      console.log(`🔹 Running OCR on ${file}...`);

      const img = await loadImage(imagePath);

      // إضافة صفحة جديدة
      doc.addPage({ size: [img.width, img.height] });

      // رسم الصورة على الصفحة
      doc.image(imagePath, 0, 0, { width: img.width, height: img.height });

      // OCR بالعربي
      const result = await worker.recognize(imagePath);
      let words = result.data.words;

        // إضافة نص مخفي لكل كلمة
        for (const w of words) {
          const x = w.bbox.x0;
          const y = img.height - w.bbox.y1;

          doc.font(arabicFont)
             .fontSize(10)
             .fillColor('black', 100) // شفاف
             .text(w.text, x, y);
        }
      
    }

    await worker.terminate();

    doc.end();
    console.log("✅ Searchable PDF created:", outputPdf);

    // تنظيف ملفات الصور المؤقتة
    for (const file of files) fs.unlinkSync(path.join(tempFolder, file));

  } catch (err) {
    console.error("❌ Error during OCR:", err);
  }
}

pdfToSearchablePDF();
