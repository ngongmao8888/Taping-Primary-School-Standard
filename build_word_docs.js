const fs = require('fs');
const path = require('path');

// Read data.js content
const dataFilePath = path.join(__dirname, 'js', 'data.js');
let dataContent = fs.readFileSync(dataFilePath, 'utf8');

// Strip out non-JS or wrap into evaluation
// In data.js: const STANDARDS_DATA = [ ... ];
let sandbox = {};
const fn = new Function('sandbox', dataContent + '; sandbox.STANDARDS_DATA = STANDARDS_DATA;');
fn(sandbox);

const STANDARDS_DATA = sandbox.STANDARDS_DATA;
console.log('Loaded standards:', STANDARDS_DATA.length);

const baseDir = __dirname;

// Sanitize folder/file names
function sanitizeFilename(name) {
  return name.replace(/[\\/:*?"<>|]/g, '_').trim();
}

let totalGenerated = 0;

STANDARDS_DATA.forEach(std => {
  // Folder for this standard
  const folderName = sanitizeFilename(`${std.title}`);
  const stdFolderPath = path.join(baseDir, folderName);
  
  if (!fs.existsSync(stdFolderPath)) {
    fs.mkdirSync(stdFolderPath);
  }

  std.keyIndicators.forEach(keyInd => {
    keyInd.subIndicators.forEach(sub => {
      const fileName = sanitizeFilename(`ឯកសារ_${sub.code}_${sub.templateTitle}.doc`);
      const filePath = path.join(stdFolderPath, fileName);

      // Build Table Headers
      let theadThs = sub.headers.map(h => `<th style="border: 1pt solid #000; padding: 6pt 8pt; background-color: #f2f2f2; font-weight: bold; text-align: center;">${h}</th>`).join('\n');

      // Build Table Rows
      let tbodyRows = sub.sampleRows.map(row => {
        let tds = row.map((c, i) => {
          let isCenter = (i === 0 || c.includes('%') || c.includes('/'));
          return `<td style="border: 1pt solid #000; padding: 6pt 8pt; text-align: ${isCenter ? 'center' : 'left'};">${c}</td>`;
        }).join('\n');
        return `<tr>${tds}</tr>`;
      }).join('\n');

      // Add extra blank editable rows for Word users
      for (let r = 1; r <= 3; r++) {
        let blankTds = sub.headers.map((h, i) => {
          return `<td style="border: 1pt solid #000; padding: 8pt 8pt; text-align: ${i === 0 ? 'center' : 'left'}; color: #999;">${i === 0 ? '...' : ''}</td>`;
        }).join('\n');
        tbodyRows += `<tr>${blankTds}</tr>\n`;
      }

      // Complete Word HTML Template with UTF-8 BOM
      const docHtml = `<!DOCTYPE html>
<html xmlns:o="urn:schemas-microsoft-com:office:office"
      xmlns:w="urn:schemas-microsoft-com:office:word"
      xmlns="http://www.w3.org/TR/REC-html40">
<head>
  <meta charset="utf-8">
  <title>${sub.templateTitle}</title>
  <!--[if gte mso 9]>
  <xml>
    <w:WordDocument>
      <w:View>Print</w:View>
      <w:Zoom>100</w:Zoom>
      <w:DoNotOptimizeForBrowser/>
    </w:WordDocument>
  </xml>
  <![endif]-->
  <style>
    @page {
      size: A4 portrait;
      margin: 1.5cm 1.5cm 1.5cm 1.5cm;
      mso-header-margin: 1cm;
      mso-footer-margin: 1cm;
    }
    body {
      font-family: 'Khmer OS Siemreap', 'Kantumruy Pro', 'Khmer OS', 'Arial', sans-serif;
      font-size: 11pt;
      color: #000000;
      line-height: 1.5;
    }
    .indicator-badge {
      float: right;
      border: 2pt solid #000000;
      padding: 4pt 12pt;
      text-align: center;
      margin-bottom: 10pt;
      background-color: #ffffff;
    }
    .indicator-label {
      font-size: 8pt;
      color: #555555;
      font-weight: bold;
    }
    .indicator-value {
      font-size: 16pt;
      font-weight: bold;
      color: #000000;
      letter-spacing: 1px;
    }
    .header-table {
      width: 100%;
      border: none;
      margin-bottom: 10pt;
    }
    .header-table td {
      border: none;
      vertical-align: top;
    }
    .moul {
      font-family: 'Khmer OS Muol Light', 'Khmer OS Muol', 'Moul', serif;
      font-weight: bold;
    }
    .title-area {
      text-align: center;
      margin-top: 15pt;
      margin-bottom: 15pt;
    }
    .doc-title {
      font-family: 'Khmer OS Muol Light', 'Khmer OS Muol', 'Moul', serif;
      font-size: 14pt;
      font-weight: bold;
      margin-bottom: 4pt;
    }
    .meta-box {
      width: 100%;
      border: 1pt solid #cccccc;
      background-color: #fcfcfc;
      padding: 8pt 10pt;
      margin-bottom: 12pt;
      font-size: 10pt;
    }
    .meta-item {
      margin-bottom: 4pt;
    }
    table.data-table {
      width: 100%;
      border-collapse: collapse;
      margin-top: 10pt;
      margin-bottom: 15pt;
    }
    .signatures-table {
      width: 100%;
      border: none;
      margin-top: 25pt;
    }
    .signatures-table td {
      border: none;
      text-align: center;
      vertical-align: top;
      width: 33.33%;
    }
  </style>
</head>
<body>

  <!-- លេខសម្គាល់សូចនាកររងនៅផ្នែកខាងស្តាំខាងលើ -->
  <div class="indicator-badge">
    <div class="indicator-label">លេខសម្គាល់</div>
    <div class="indicator-value">${sub.code}</div>
  </div>

  <!-- ក្បាលសំបុត្ររដ្ឋបាលផ្លូវការ -->
  <table class="header-table">
    <tr>
      <td style="width: 45%; font-size: 10pt; line-height: 1.4;">
        <div>ក្រសួងអប់រំ យុវជន និងកីឡា</div>
        <div>មន្ទីរអប់រំ យុវជន និងកីឡា ខេត្តកែប</div>
        <div>ការិយាល័យអប់រំ យុវជន និងកីឡា នៃរដ្ឋបាលស្រុកដំណាក់ចង្អើរ</div>
        <div style="font-weight: bold; font-size: 11pt;">សាលាបឋមសិក្សាតាពីង</div>
        <div style="color: #666; font-size: 9pt;">លេខកូដសាលា៖ 22020204012</div>
      </td>
      <td style="width: 55%; text-align: center;">
        <div class="moul" style="font-size: 11pt;">ព្រះរាជាណាចក្រកម្ពុជា</div>
        <div class="moul" style="font-size: 10pt;">ជាតិ សាសនា ព្រះមហាក្សត្រ</div>
        <div style="font-size: 10pt; letter-spacing: 2px;">*** 3 ***</div>
      </td>
    </tr>
  </table>

  <!-- ចំណងជើងឯកសារ -->
  <div class="title-area">
    <div class="doc-title">${sub.templateTitle}</div>
    <div style="font-size: 10pt; color: #444;">ឯកសារអនុវត្តស្តង់ដាសាលាបឋមសិក្សាគំរូ នៃក្រសួងអប់រំ យុវជន និងកីឡា</div>
    <div style="font-size: 9.5pt; color: #222; margin-top: 3pt;">
      <strong>${std.title}</strong> | <strong>${keyInd.title}</strong>
    </div>
  </div>

  <!-- ប្រអប់ព័ត៌មានលម្អិត -->
  <div class="meta-box">
    <div class="meta-item"><strong>• សូចនាកររង៖</strong> ${sub.code} - ${sub.name}</div>
    <div class="meta-item"><strong>• ឯកសារតម្រូវ (កូឡោនទី២ នៃសៀវភៅស្តង់ដា)៖</strong> ${sub.col2Docs}</div>
    <div class="meta-item"><strong>• ឆ្នាំសិក្សា៖</strong> ២០២៣ - ២០២៤</div>
  </div>

  <!-- តារាងទិន្នន័យឯកសារផ្លូវការ (អាចកែប្រែក្នុង Word បានពេញលេញ) -->
  <table class="data-table">
    <thead>
      <tr>
        ${theadThs}
      </tr>
    </thead>
    <tbody>
      ${tbodyRows}
    </tbody>
  </table>

  <!-- សេចក្តីណែនាំក្នុងការបំពេញ -->
  <div style="font-size: 9pt; color: #555555; background-color: #f9f9f9; border-left: 3pt solid #0284c7; padding: 6pt 8pt; margin-bottom: 20pt;">
    <strong>* កំណត់សម្គាល់៖</strong> លោកគ្រូ-អ្នកគ្រូ និងគណៈគ្រប់គ្រងសាលា អាចចុចកែសម្រួល ឬបន្ថែមទិន្នន័យតាមជួរដេកនីមួយៗក្នុងតារាងខាងលើនេះដោយផ្ទាល់ក្នុងកម្មវិធី Microsoft Word ឱ្យសមស្របតាមស្ថានភាពជាក់ស្តែង រួចបោះពុម្ព និងចុះហត្ថលេខារក្សាទុកជាភស្តុតាង។
  </div>

  <!-- ហត្ថលេខា និងការអនុម័តផ្លូវការ -->
  <table class="signatures-table">
    <tr>
      <td>
        <div style="font-style: italic; font-size: 10pt;">បានឃើញ និងអនុម័ត</div>
        <div style="font-weight: bold; margin-top: 3pt;">ប្រធាន គ.គ.ស</div>
        <div style="margin-top: 45pt;">................................................</div>
      </td>
      <td>
        <div style="font-style: italic; font-size: 10pt;">បានឃើញ និងឯកភាព</div>
        <div style="font-weight: bold; margin-top: 3pt;">នាយកសាលា</div>
        <div style="margin-top: 45pt;">................................................</div>
      </td>
      <td>
        <div style="font-style: italic; font-size: 10pt;">ថ្ងៃទី ១៥ ខែ វិច្ឆិកា ឆ្នាំ ២០២៣</div>
        <div style="font-weight: bold; margin-top: 3pt;">គ្រូបង្រៀន / មន្ត្រីទទួលបន្ទុក</div>
        <div style="margin-top: 45pt;">................................................</div>
      </td>
    </tr>
  </table>

</body>
</html>`;

      // Prepend UTF-8 BOM so Microsoft Word recognizes Khmer font immediately without mojibake
      const bom = Buffer.from([0xEF, 0xBB, 0xBF]);
      const fileBuffer = Buffer.concat([bom, Buffer.from(docHtml, 'utf8')]);

      fs.writeFileSync(filePath, fileBuffer);
      totalGenerated++;
    });
  });
});

console.log(`Successfully generated ${totalGenerated} Word (.doc) files across all 5 standards!`);
