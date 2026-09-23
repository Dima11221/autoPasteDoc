import PizZip from "pizzip";

interface ParaHit {
  start: number;
  end: number;
  text: string;
  xml: string
};

const paraText = (paraXml: string): string => {
  const parts: string[] = [];
  const re = /<w:t[^>]*>([^<]*)<\/w:t>/g;
  let match: RegExpExecArray | null;
  while ((match = re.exec(paraXml)) !== null) {
    parts.push(match[1]);
  }
  return parts.join('').replace(/\u00a0/g, ' ').replace(/\s+/g, " ").trim();
};

const listParagraphs = (documentXml: string): ParaHit[] => {
  const result: ParaHit[] = [];
  const re = /<w:p[\s>][\s\S]*?<\/w:p>/g;
  let match: RegExpExecArray | null;
  while ((match = re.exec(documentXml)) !== null) {
    result.push({
      start: match.index,
      end: match.index + match[0].length,
      text: paraText(match[0]),
      xml: match[0]
    });
  }
  return result;
};

const isTocPara = (paraXml: string): boolean => {
  return (
    paraXml.includes("<w:hyperlink") || /w:pStyle\s+[^>]*w:val="21"/.test(paraXml)
  );
};

const isAppendix2Photos = (text: string): boolean => {
  return (
    /^Приложение №\s*2\b/i.test(text) && /Фотоматериалы/i.test(text)
  );
};

const isAppendix3 = (text: string): boolean => {
  return (/^Приложение №\s*3\b/i.test(text));
};

const DEFAULT_PPR = `<w:pPr>
  <w:spacing w:after="120" w:line="276" w:lineRule="auto"/>
  <w:ind w:right="140"/>
  <w:jc w:val="both"/>
  <w:rPr>
    <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
    <w:color w:val="auto"/>
    <w:sz w:val="24"/>
    <w:szCs w:val="24"/>
  </w:rPr>
</w:pPr>`;

const DEFAULT_RPR = `<w:rPr>
  <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
  <w:color w:val="auto"/>
  <w:sz w:val="24"/>
  <w:szCs w:val="24"/>
</w:rPr>`;

const PAGE_BREAK_PARAGRAPH = `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`;

const extractPpr = (paraXml: string): string | null => {
  const m = paraXml.match(/<w:pPr>[\s\S]*?<\/w:pPr>/);
  return m ? m[0] : null;
};

const tagParagraph = (tag: string, pPr: string): string => {
  return (
    `<w:p>${pPr}<w:r>${DEFAULT_RPR}` +
    `<w:t xml:space="preserve">${tag}</w:t>` +
    `</w:r></w:p>`
  );
};

const withJc = (pPr: string, jc: "center" | "left" | "both"): string => {
  if (/<w:jc\b/.test(pPr)) {
    return pPr.replace(/<w:jc\b[^>]*\/>/, `<w:jc w:val="${jc}"/>`);
  }
  // если jc нет — вставить перед закрытием pPr
  return pPr.replace("</w:pPr>", `<w:jc w:val="${jc}"/></w:pPr>`);
};

const buildPhotosBlock = (imagePpr: string, captionPpr: string): string => {
  const bothImage = withJc(imagePpr, "both");
  const bothCaption = withJc(captionPpr, "both");
  const leftImage = withJc(imagePpr, "left");

  return (
    tagParagraph("{#photos}", bothImage) +
    // портрет 10.5 → слева
    tagParagraph("{#isPortrait}", leftImage) +
    tagParagraph("{%data}", leftImage) +
    tagParagraph("{caption}", bothCaption) +
    tagParagraph("{/isPortrait}", leftImage) +
    // альбом / квадрат → центр
    tagParagraph("{^isPortrait}", bothImage) +
    tagParagraph("{%data}", bothImage) +
    tagParagraph("{caption}", bothCaption) +
    tagParagraph("{/isPortrait}", bothImage) +
    tagParagraph("{/photos}", bothImage) +
    PAGE_BREAK_PARAGRAPH
  );
};


export const prepareTemplateZip = (docxBuffer: Buffer): PizZip => {
  const zip = new PizZip(docxBuffer);
  const file = zip.file("word/document.xml");
  if (!file) {
    throw new Error("В docx нет word/document.xml");
  }

  const documentXml = file.asText();
  const paragraphs = listParagraphs(documentXml);

  let startIdx = -1;
  for (let i = 0; i < paragraphs.length; i++) {
    const p = paragraphs[i];
    if (!isAppendix2Photos(p.text)) continue;
    if (isTocPara(p.xml)) continue;
    startIdx = i;
  }

  if (startIdx < 0) {
    throw new Error("Не найден раздел «Приложение № 2 к Акту обследования. Фотоматериалы»");
  }

  let endIdx = -1;
  for (let i = startIdx + 1; i < paragraphs.length; i++) {
    if (isAppendix3(paragraphs[i].text)) {
      endIdx = i;
      break;
    }
  }

  if (endIdx < 0) {
    throw new Error("Не найден раздел «Приложение № 3 ...» после блока фотоматериалов",);
  };

  const between = paragraphs.slice(startIdx + 1, endIdx);

  const imageSample =
    between.find((p) => p.xml.includes("<w:drawing")) ??
    between.find((p) => p.text.includes("{%data}")) ??
    between.find((p) => p.text.length > 0);

  const captionSample =
    between.find((p) => /^Фото\s+\d/i.test(p.text)) ??
    between.find((p) => p.text.includes("{caption}")) ??
    imageSample;

  const imagePpr = (imageSample && extractPpr(imageSample.xml)) || DEFAULT_PPR;
  const captionPpr =
    (captionSample && extractPpr(captionSample.xml)) || DEFAULT_PPR;

  const photosBlock = buildPhotosBlock(imagePpr, captionPpr);

  const cutFrom = paragraphs[startIdx].end;
  const cutTo = paragraphs[endIdx].start;

  const newXml =
    documentXml.slice(0, cutFrom) + photosBlock + documentXml.slice(cutTo);

  zip.file("word/document.xml", newXml);

  return zip;
};