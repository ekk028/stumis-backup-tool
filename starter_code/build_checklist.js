const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat, BorderStyle, WidthType,
  ShadingType, HeadingLevel, PageNumber
} = require('docx');

const border = { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" };
const borders = { top: border, bottom: border, left: border, right: border };

function cell(text, opts = {}) {
  const { width, bold = false, fill, align = AlignmentType.LEFT, size = 22 } = opts;
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: fill ? { fill, type: ShadingType.CLEAR } : undefined,
    margins: { top: 100, bottom: 100, left: 140, right: 140 },
    children: [new Paragraph({
      alignment: align,
      children: [new TextRun({ text, bold, size, font: "TH Sarabun New" })]
    })]
  });
}

function p(text, opts = {}) {
  const { bold = false, size = 28, heading, align = AlignmentType.LEFT, after = 80, bullet = false } = opts;
  const config = {
    alignment: align,
    spacing: { after, before: 0 },
    children: [new TextRun({ text, bold, size, font: "TH Sarabun New" })]
  };
  if (heading) config.heading = heading;
  if (bullet) config.numbering = { reference: "bullets", level: 0 };
  return new Paragraph(config);
}

const HEADER_FILL = "1F4E79";
const SECTION_FILL = "D9E2F3";
const WARN_FILL = "FFF2CC";

// 7-section table from manual
const items = [
  ["1", "เช็กชื่อนักเรียนมาเช้า", "Report เช็กชื่อ → เลือกเดือน → กด Print", "PDF"],
  ["2", "พฤติกรรมไม่เหมาะสม (คะแนนติดลบ)", "Report พฤติกรรม → เลือกห้องเรียน", "Excel"],
  ["3", "พฤติกรรมเหมาะสม / ความดี (คะแนนบวก)", "Report พฤติกรรม → เลือกห้องเรียน", "Excel"],
  ["4", "การเยี่ยมบ้านนักเรียน", "เมนูการเยี่ยมบ้าน → เลือกครั้งที่ → ปุ่ม รายงาน", "PDF"],
  ["5", "ผลประเมิน EQ", "เมนูประเมิน SDQ → EQ → Report EQ", "Excel"],
  ["6", "ผลประเมิน SDQ", "เมนูประเมิน SDQ → Report SDQ", "Excel"],
  ["7", "เช็กชื่อในคาบเรียน", "Report เช็กชื่อรายวิชา", "Excel"],
];

const tableRows = [
  new TableRow({
    tableHeader: true,
    children: [
      cell("ลำดับ",   { width:  900, bold: true, fill: HEADER_FILL, align: AlignmentType.CENTER, size: 24 }),
      cell("หมวดข้อมูล", { width: 3200, bold: true, fill: HEADER_FILL, size: 24 }),
      cell("ขั้นตอน",   { width: 3960, bold: true, fill: HEADER_FILL, size: 24 }),
      cell("ไฟล์",     { width: 1300, bold: true, fill: HEADER_FILL, align: AlignmentType.CENTER, size: 24 }),
    ].map(c => {
      // Override text color for header
      c.root[c.root.length-1].root.forEach(par => {
        if (par.root) par.root.forEach(r => {
          if (r.root && r.root[0] && r.root[0].root) {
            // shift color via raw insertion below if needed
          }
        });
      });
      return c;
    })
  })
];

// Build header row properly with white text
function headerCell(text, width, align = AlignmentType.LEFT) {
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: { fill: HEADER_FILL, type: ShadingType.CLEAR },
    margins: { top: 100, bottom: 100, left: 140, right: 140 },
    children: [new Paragraph({
      alignment: align,
      children: [new TextRun({ text, bold: true, size: 26, color: "FFFFFF", font: "TH Sarabun New" })]
    })]
  });
}

const tableRows2 = [
  new TableRow({
    tableHeader: true,
    children: [
      headerCell("ลำดับ", 900, AlignmentType.CENTER),
      headerCell("หมวดข้อมูล", 3200),
      headerCell("ขั้นตอน", 3960),
      headerCell("ไฟล์", 1300, AlignmentType.CENTER),
    ]
  })
];

items.forEach((row, i) => {
  const fill = i % 2 === 0 ? "FFFFFF" : "F4F6FB";
  tableRows2.push(new TableRow({
    children: [
      cell(row[0], { width:  900, align: AlignmentType.CENTER, fill, size: 26, bold: true }),
      cell(row[1], { width: 3200, fill, size: 26 }),
      cell(row[2], { width: 3960, fill, size: 24 }),
      cell(row[3], { width: 1300, align: AlignmentType.CENTER, fill, size: 24, bold: true }),
    ]
  }));
});

// Gap-analysis items
const gaps = [
  "เช็กชื่อตอนเย็น (คู่มือสอนเฉพาะตอนเช้า)",
  "การบันทึกตารางสอน",
  "รายวิชาที่บันทึกไว้ในปีการศึกษาที่ผ่านมา",
  "คะแนนรายวิชา (ปพ.5) — สำคัญที่สุด เป็นเอกสารราชการ",
  "การเลือกชุมนุม / กิจกรรมแบบเลือก",
  "การประเมินภาวะซึมเศร้า",
  "ข้อมูลการสแกนบัตรนักเรียนตอนเช้า-เย็น (ถ้าใช้ระบบสแกนบัตร)",
];

// Checklist before notifying admin
const finalSteps = [
  "สำรองข้อมูลครบทั้ง 7 หมวดตามตารางด้านบน",
  "สอบถามแอดมินเรื่องวิธี export ข้อมูล 7 รายการในตารางที่ 2",
  "เตรียมไฟล์ Excel รายชื่อนักเรียนใหม่จาก DMC/SGS (รหัสนักเรียน, ชื่อ-สกุล, ชั้น, เลขที่, รหัสประชาชน) เรียงต่อกันใน Sheet เดียว",
  "เก็บไฟล์สำรองทั้งหมดในโฟลเดอร์เดียว ตั้งชื่อชัดเจน เช่น \"backup_ปีการศึกษา_2568\"",
  "แจ้งวันที่พร้อมให้แอดมินดำเนินการ โทร. 0859043016",
];

const doc = new Document({
  styles: {
    default: { document: { run: { font: "TH Sarabun New", size: 28 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 40, bold: true, font: "TH Sarabun New", color: "1F4E79" },
        paragraph: { spacing: { before: 280, after: 160 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "TH Sarabun New", color: "2E5496" },
        paragraph: { spacing: { before: 220, after: 120 }, outlineLevel: 1 } },
    ]
  },
  numbering: {
    config: [
      { reference: "bullets",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } },
                   run: { font: "TH Sarabun New" } } }] },
      { reference: "checklist",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "☐", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } },
                   run: { font: "TH Sarabun New" } } }] },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }
      }
    },
    footers: {
      default: new Footer({ children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({ text: "Checklist สำรองข้อมูล STU-MIS  |  หน้า ", size: 20, font: "TH Sarabun New", color: "808080" }),
          new TextRun({ children: [PageNumber.CURRENT], size: 20, font: "TH Sarabun New", color: "808080" }),
        ]
      })] })
    },
    children: [
      // Title
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 80 },
        children: [new TextRun({ text: "Checklist สำรองข้อมูลก่อนเปิดปีการศึกษาใหม่", bold: true, size: 44, font: "TH Sarabun New", color: "1F4E79" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 240 },
        children: [new TextRun({ text: "ระบบ STU-MIS  |  สรุปจากคู่มือของแอดมิน (โทร. 0859043016)", size: 24, italics: true, font: "TH Sarabun New", color: "595959" })]
      }),

      // Section 1
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: "1. สรุปการสำรองข้อมูลจากคู่มือ (ทำตามนี้ก่อน)", bold: true, size: 36, font: "TH Sarabun New", color: "1F4E79" })] }),
      p("คู่มือฉบับนี้ครอบคลุมการดาวน์โหลด 7 หมวดข้อมูลหลัก ทำตามตารางทีละหมวดให้ครบก่อนแจ้งแอดมิน", { size: 26, after: 160 }),

      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [900, 3200, 3960, 1300],
        rows: tableRows2
      }),

      p("", { after: 160 }),

      // Tip box
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [9360],
        rows: [new TableRow({ children: [new TableCell({
          borders: {
            top: { style: BorderStyle.SINGLE, size: 4, color: "F4B400" },
            bottom: { style: BorderStyle.SINGLE, size: 4, color: "F4B400" },
            left: { style: BorderStyle.SINGLE, size: 24, color: "F4B400" },
            right: { style: BorderStyle.SINGLE, size: 4, color: "F4B400" },
          },
          shading: { fill: WARN_FILL, type: ShadingType.CLEAR },
          margins: { top: 160, bottom: 160, left: 200, right: 200 },
          width: { size: 9360, type: WidthType.DXA },
          children: [
            new Paragraph({ children: [new TextRun({ text: "เคล็ดลับการ Save PDF (หมวด 1 และ 4)", bold: true, size: 26, font: "TH Sarabun New", color: "7F6000" })], spacing: { after: 80 } }),
            new Paragraph({ children: [new TextRun({ text: "ที่หน้าต่าง Print ของเบราว์เซอร์ ให้ตั้งค่าดังนี้", size: 24, font: "TH Sarabun New" })], spacing: { after: 60 } }),
            new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "ปลายทาง (Destination) : Save as PDF", size: 24, font: "TH Sarabun New" })] }),
            new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "หน้ากระดาษ (Pages) : ทั้งหมด (All)", size: 24, font: "TH Sarabun New" })] }),
            new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "เปิดตัวเลือก Background graphics เพื่อให้ตารางและสีออกครบ", size: 24, font: "TH Sarabun New" })] }),
          ]
        })] })]
      }),

      p("", { after: 200 }),

      // Section 2 - Gap analysis
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: "2. ข้อมูลที่จะถูกลบ แต่ไม่มีในคู่มือ — ต้องสอบถามแอดมิน", bold: true, size: 36, font: "TH Sarabun New", color: "C00000" })] }),
      p("รายการต่อไปนี้แอดมินบอกว่าจะเคลียร์ แต่ในคู่มือฉบับนี้ไม่ได้สอนวิธี backup ควรโทรถามแอดมินก่อน (โทร. 0859043016) ว่าจะ export อย่างไร โดยเฉพาะ ปพ.5 ซึ่งเป็นเอกสารราชการ", { size: 26, after: 140 }),

      ...gaps.map(g => new Paragraph({
        numbering: { reference: "checklist", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: g, size: 26, font: "TH Sarabun New" })]
      })),

      p("", { after: 200 }),

      // Section 3 - Final checklist
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: "3. ขั้นตอนสุดท้ายก่อนแจ้งแอดมินดำเนินการ", bold: true, size: 36, font: "TH Sarabun New", color: "1F4E79" })] }),

      ...finalSteps.map((s, i) => new Paragraph({
        numbering: { reference: "checklist", level: 0 },
        spacing: { after: 80 },
        children: [new TextRun({ text: s, size: 26, font: "TH Sarabun New" })]
      })),

      p("", { after: 200 }),

      // Section 4 - Data not deleted
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: "4. ข้อมูลที่จะยังคงอยู่ในระบบ (ไม่ต้องสำรอง)", bold: true, size: 36, font: "TH Sarabun New", color: "548235" })] }),
      p("ข้อมูลด้านล่างนี้แอดมินยืนยันว่าจะไม่ถูกลบ ไม่ต้องเสียเวลาสำรอง", { size: 26, after: 120 }),

      ...[
        "ข้อมูลนักเรียน รูปประจำตัวนักเรียน",
        "ข้อมูลครู รูปประจำตัวครู (ใช้ Username/Password เดิม)",
        "ข้อมูลการเยี่ยมบ้านนักเรียน รูปภาพการเยี่ยมบ้าน",
        "ข้อมูลผู้บริหาร / แอดมิน (ใช้ Username/Password เดิม)",
      ].map(s => new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: s, size: 26, font: "TH Sarabun New" })]
      })),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("/sessions/compassionate-hopeful-cori/mnt/outputs/checklist_backup_STUMIS.docx", buffer);
  console.log("OK");
});
