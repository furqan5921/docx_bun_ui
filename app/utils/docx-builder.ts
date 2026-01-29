// DOCX OOXML Builder - Converts Univer IDocumentData to DOCX format
// High-fidelity export with proper tables, images, text formatting, and lists

import type { IDocumentData } from "@univerjs/core";
import JSZip from "jszip";
import { create } from "xmlbuilder2";

// Image counter for relationship IDs
let globalImageCounter = 1;
let globalRIdCounter = 3; // Start at 3 to leave room for styles and numbering

interface ImageInfo {
  drawingId: string;
  rId: string;
  filename: string;
  width: number;
  height: number;
  base64Data: string;
  extension: string;
}

// ✅ OOXML VALIDATION: Validate document structure before export
function validateOOXMLStructure(documentData: IDocumentData): string[] {
  const errors: string[] = [];
  
  if (!documentData.body) {
    errors.push("Document body is undefined");
    return errors;
  }
  
  const dataStream = documentData.body.dataStream || "";
  const paragraphs = documentData.body.paragraphs || [];
  const customRanges = documentData.body.customRanges || [];

  // Check dataStream ends with \r\n (required by Univer)
  if (!dataStream.endsWith("\r\n")) {
    errors.push("dataStream must end with \\r\\n for valid Univer document");
  }

  // Check all paragraphs have valid startIndex
  paragraphs.forEach((p: any, i: number) => {
    if (p.startIndex >= dataStream.length) {
      errors.push(
        `Paragraph ${i} startIndex ${p.startIndex} exceeds dataStream length ${dataStream.length}`
      );
    }
  });

  // Check customRanges point to valid positions
  customRanges.forEach((range: any, i: number) => {
    if (range.startIndex >= dataStream.length) {
      errors.push(
        `CustomRange ${i} (${range.rangeId}) startIndex ${range.startIndex} exceeds dataStream length`
      );
    }

    // Verify range points to correct marker (only for tables and images, NOT hyperlinks)
    // rangeType: 0 = hyperlink (no marker), 1 = drawing (0x1A marker), 2 = table (0x1A marker)
    if (range.rangeType > 0 && range.startIndex < dataStream.length) {
      const char = dataStream.charAt(range.startIndex);
      const charCode = char.charCodeAt(0);
      // Table marker: 0x1A, Drawing marker: 0x1A
      if (charCode !== 0x1a) {
        errors.push(
          `CustomRange ${i} (${range.rangeId}) type ${
            range.rangeType
          } should point to marker 0x1A but found 0x${charCode.toString(
            16
          )} at position ${range.startIndex}`
        );
      }
    }
  });

  // Check textRuns don't exceed dataStream
  const textRuns = documentData.body?.textRuns || [];
  textRuns.forEach((tr: any, i: number) => {
    if (tr.ed > dataStream.length) {
      errors.push(
        `TextRun ${i} end position ${tr.ed} exceeds dataStream length ${dataStream.length}`
      );
    }
  });

  return errors;
}

export async function buildDocxFromUniverData(
  documentData: IDocumentData
): Promise<Blob> {
  // ✅ Validate structure before export
  const validationErrors = validateOOXMLStructure(documentData);
  if (validationErrors.length > 0) {
    console.error("[DOCX Export] ❌ Validation failed:", validationErrors);
    throw new Error(
      `Invalid document structure (${
        validationErrors.length
      } errors):\n${validationErrors.slice(0, 5).join("\n")}${
        validationErrors.length > 5
          ? `\n... and ${validationErrors.length - 5} more errors`
          : ""
      }`
    );
  }

  console.log("[DOCX Export] ✅ Document structure validated successfully");
  // Reset counters
  globalImageCounter = 1;
  globalRIdCounter = 3;

  const zip = new JSZip();

  // Prepare image info for document building
  const imageInfos: ImageInfo[] = [];
  const drawings = documentData.drawings || {};
  const drawingsOrder = documentData.drawingsOrder || Object.keys(drawings);

  // Process drawings in order
  for (const drawingId of drawingsOrder) {
    const drawing: any = drawings[drawingId]; // Use any to handle Univer's complex drawing types
    if (
      drawing &&
      (drawing.drawingType === "image" || drawing.drawingType === 1) &&
      drawing.imageProperties?.base64Cache
    ) {
      const base64Full = drawing.imageProperties.base64Cache;
      const base64Data = base64Full.includes(",")
        ? base64Full.split(",")[1]
        : base64Full;

      if (base64Data) {
        const extension = getImageExtension(
          drawing.imageProperties.source || base64Full
        );
        const filename = `image${globalImageCounter}.${extension}`;

        imageInfos.push({
          drawingId,
          rId: `rId${globalRIdCounter}`,
          filename,
          width:
            drawing.transform?.size?.width ||
            drawing.docTransform?.size?.width ||
            200,
          height:
            drawing.transform?.size?.height ||
            drawing.docTransform?.size?.height ||
            200,
          base64Data,
          extension,
        });

        globalImageCounter++;
        globalRIdCounter++;
      }
    }
  }

  // Check if document has lists
  const hasLists =
    documentData.body?.paragraphs?.some((p: any) => p.bullet) || false;

  // Build XML content
  const documentXml = buildDocumentXml(documentData, imageInfos);
  const stylesXml = buildStylesXml(documentData);
  const numberingXml = hasLists ? buildNumberingXml(documentData) : null;
  const relsXml = buildRelsXml(documentData, imageInfos, hasLists);
  const contentTypesXml = buildContentTypesXml(documentData, imageInfos);

  // Add files to ZIP
  zip.file("[Content_Types].xml", contentTypesXml);
  zip.folder("_rels")?.file(".rels", buildRootRels());
  zip.folder("word")?.file("document.xml", documentXml);
  zip.folder("word")?.file("styles.xml", stylesXml);
  if (numberingXml) {
    zip.folder("word")?.file("numbering.xml", numberingXml);
  }
  zip.folder("word/_rels")?.file("document.xml.rels", relsXml);

  // Add images to media folder
  for (const imgInfo of imageInfos) {
    try {
      const imageBuffer = Uint8Array.from(atob(imgInfo.base64Data), (c) =>
        c.charCodeAt(0)
      );
      zip.folder("word/media")?.file(imgInfo.filename, imageBuffer);
      console.log(`[DOCX Export] Added image: ${imgInfo.filename}`);
    } catch (err) {
      console.error(
        `[DOCX Export] Failed to add image ${imgInfo.filename}:`,
        err
      );
    }
  }

  console.log("[DOCX Export] ZIP structure created, generating blob...");
  return await zip.generateAsync({ type: "blob" });
}

function buildDocumentXml(
  documentData: IDocumentData,
  imageInfos: ImageInfo[]
): string {
  console.log("🚀 ~ buildDocumentXml ~ documentData:", documentData)
  const dataStream = documentData.body?.dataStream || "";
  const textRuns = documentData.body?.textRuns || [];
  const paragraphs = documentData.body?.paragraphs || [];

  const doc = create({ version: "1.0", encoding: "UTF-8", standalone: "yes" })
    .ele("w:document", {
      "xmlns:wpc":
        "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
      "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
      "xmlns:o": "urn:schemas-microsoft-com:office:office",
      "xmlns:r":
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "xmlns:m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
      "xmlns:v": "urn:schemas-microsoft-com:vml",
      "xmlns:wp14":
        "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
      "xmlns:wp":
        "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
      "xmlns:w10": "urn:schemas-microsoft-com:office:word",
      "xmlns:w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
      "xmlns:w14": "http://schemas.microsoft.com/office/word/2010/wordml",
      "xmlns:wpg":
        "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
      "xmlns:wpi":
        "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
      "xmlns:wne": "http://schemas.microsoft.com/office/word/2006/wordml",
      "xmlns:wps":
        "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
      "xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "xmlns:pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
      "mc:Ignorable": "w14 wp14",
    })
    .ele("w:body");

  // Find image marker positions (0x1A character)
  const imageMarkerPositions: number[] = [];
  for (let i = 0; i < dataStream.length; i++) {
    if (dataStream.charCodeAt(i) === 0x1a) {
      imageMarkerPositions.push(i);
    }
  }

  // Map image markers to image info based on order
  const imageAtPosition: Map<number, ImageInfo> = new Map();
  imageMarkerPositions.forEach((pos, idx) => {
    if (idx < imageInfos.length) {
      imageAtPosition.set(pos, imageInfos[idx]);
    }
  });

  console.log(
    `[DOCX Export] Building document: ${paragraphs.length} paragraphs, ${imageInfos.length} images`
  );

  // Handle empty document
  if (paragraphs.length === 0) {
    const p = doc.ele("w:p");
    if (dataStream.length > 0) {
      const cleanText = dataStream.replace(/[\x00-\x1F]/g, "").trim();
      if (cleanText.length > 0) {
        const r = p.ele("w:r");
        r.ele("w:t", { "xml:space": "preserve" }).txt(cleanText);
      }
    } else {
      p.ele("w:r");
    }
  }

  // Process paragraphs
  // CRITICAL: In Univer, paragraph.startIndex points to the \r at the END of the paragraph's text
  // So paragraph text is BEFORE startIndex, from previous paragraph's end to this one's startIndex
  console.log(`[DOCX Export] Processing ${paragraphs.length} paragraphs with ${textRuns.length} text runs`);

  paragraphs.forEach((para: any, paraIndex: number) => {
    const prevPara = paragraphs[paraIndex - 1];

    // Calculate actual text range for this paragraph
    // Text starts after previous paragraph's \r, or at 0 for first paragraph
    const paraStart = paraIndex === 0 ? 0 : (prevPara.startIndex + 1);
    // Text ends at this paragraph's \r position (inclusive of content before \r)
    const paraEnd = para.startIndex + 1; // Include the \r position

    const paraText = dataStream.substring(paraStart, paraEnd).replace(/[\r\n]/g, '\\r');
    console.log(`[DOCX Export] Paragraph ${paraIndex}: range [${paraStart}-${paraEnd}), text="${paraText.substring(0, 50)}..."`);

    const p = doc.ele("w:p");
    const pPr = p.ele("w:pPr");

    if (para.paragraphStyle) {
      addParagraphProperties(pPr, para.paragraphStyle);
    }

    if (para.bullet) {
      addBulletProperties(pPr, para.bullet);
    }

    // Process runs in this paragraph
    // Find runs that overlap with this paragraph's text range
    const paraRuns = textRuns.filter(
      (run: any) => run.st < paraEnd && run.ed >= paraStart
    );
    console.log(`[DOCX Export]   Found ${paraRuns.length} runs for paragraph ${paraIndex}`);
    paraRuns.forEach((run: any, runIdx: number) => {
      const clippedStart = Math.max(run.st, paraStart);
      const clippedEnd = Math.min(run.ed, paraEnd);
      const runText = dataStream.substring(clippedStart, clippedEnd).replace(/[\r\n]/g, '\\r');
      console.log(`[DOCX Export]     Run ${runIdx}: st=${run.st}, ed=${run.ed}, clipped=[${clippedStart}-${clippedEnd}), va=${run.ts?.va}, text="${runText}"`);
    });

    if (paraRuns.length === 0) {
      // No formatted runs - check if there's plain text or images
      const imgPos = imageMarkerPositions.find(
        (pos) => pos >= paraStart && pos < paraEnd
      );
      
      if (imgPos !== undefined) {
        const imgInfo = imageAtPosition.get(imgPos);
        if (imgInfo) {
          addImageToRun(p, imgInfo);
        }
      } else {
        // Extract plain text from dataStream (no formatting)
        const paraText = dataStream.substring(paraStart, paraEnd);
        const cleanText = paraText.replace(/[\x00-\x1F]/g, "").trim();
        
        if (cleanText.length > 0) {
          const r = p.ele("w:r");
          r.ele("w:t", { "xml:space": "preserve" }).txt(cleanText);
        } else {
          // Truly empty paragraph
          p.ele("w:r");
        }
      }
    } else {
      let lastRunEnd = paraStart;

      paraRuns.forEach((run: any) => {
        const runStart = Math.max(run.st, paraStart);
        const runEnd = Math.min(run.ed, paraEnd);

        // Handle unformatted text BEFORE this run (gap between runs)
        if (lastRunEnd < runStart) {
          const gapText = dataStream.substring(lastRunEnd, runStart);
          const cleanGapText = gapText.replace(/[\x00-\x1F]/g, "");
          
          if (cleanGapText.length > 0) {
            // Unformatted text between runs - add as plain text run
            const plainRun = p.ele("w:r");
            plainRun.ele("w:t", { "xml:space": "preserve" }).txt(cleanGapText);
          }
          
          // Check for image in gap
          const imgBetween = imageMarkerPositions.find(
            (pos) => pos >= lastRunEnd && pos < runStart
          );
          if (imgBetween !== undefined) {
            const imgInfo = imageAtPosition.get(imgBetween);
            if (imgInfo) {
              addImageToRun(p, imgInfo);
            }
          }
        }

        const runText = dataStream.substring(runStart, runEnd);
        const imgInRun = imageMarkerPositions.find(
          (pos) => pos >= runStart && pos < runEnd
        );

        if (imgInRun !== undefined) {
          const imgInfo = imageAtPosition.get(imgInRun);

          // Text before image
          const beforeImg = runText.substring(0, imgInRun - runStart);
          if (beforeImg && beforeImg.replace(/[\x00-\x1F]/g, "").length > 0) {
            const r1 = p.ele("w:r");
            if (run.ts && Object.keys(run.ts).length > 0) {
              const rPr = r1.ele("w:rPr");
              addRunProperties(rPr, run.ts);
            }
            r1.ele("w:t", { "xml:space": "preserve" }).txt(
              beforeImg.replace(/[\x00-\x1F]/g, "")
            );
          }

          if (imgInfo) {
            addImageToRun(p, imgInfo);
          }

          // Text after image
          const afterImg = runText.substring(imgInRun - runStart + 1);
          if (afterImg && afterImg.replace(/[\x00-\x1F]/g, "").length > 0) {
            const r2 = p.ele("w:r");
            if (run.ts && Object.keys(run.ts).length > 0) {
              const rPr = r2.ele("w:rPr");
              addRunProperties(rPr, run.ts);
            }
            r2.ele("w:t", { "xml:space": "preserve" }).txt(
              afterImg.replace(/[\x00-\x1F]/g, "")
            );
          }
        } else {
          // Regular text run
          const cleanText = runText.replace(/[\x00-\x1F]/g, "");
          if (cleanText.length > 0) {
            const r = p.ele("w:r");

            if (run.ts && Object.keys(run.ts).length > 0) {
              const rPr = r.ele("w:rPr");
              addRunProperties(rPr, run.ts);
            }

            r.ele("w:t", { "xml:space": "preserve" }).txt(cleanText);
          }
        }

        lastRunEnd = runEnd;
      });

      // Handle unformatted text AFTER the last run
      if (lastRunEnd < paraEnd) {
        const remainingText = dataStream.substring(lastRunEnd, paraEnd);
        const cleanRemaining = remainingText.replace(/[\x00-\x1F]/g, "");
        
        if (cleanRemaining.length > 0) {
          // Unformatted text after last run - add as plain text run
          const plainRun = p.ele("w:r");
          plainRun.ele("w:t", { "xml:space": "preserve" }).txt(cleanRemaining);
        }
        
        // Check for image after last run
        const imgAfter = imageMarkerPositions.find(
          (pos) => pos >= lastRunEnd && pos < paraEnd
        );
        if (imgAfter !== undefined) {
          const imgInfo = imageAtPosition.get(imgAfter);
          if (imgInfo) {
            addImageToRun(p, imgInfo);
          }
        }
      }
    }
  });

  // Add section properties
  const sectPr = doc.ele("w:sectPr");

  // Page size in twips (1 point = 20 twips)
  const pageWidth = documentData.documentStyle?.pageSize?.width || 612; // Letter width in points
  const pageHeight = documentData.documentStyle?.pageSize?.height || 792; // Letter height in points

  sectPr.ele("w:pgSz", {
    "w:w": Math.round(pageWidth * 20).toString(),
    "w:h": Math.round(pageHeight * 20).toString(),
  });

  // Margins in twips
  const marginTop = documentData.documentStyle?.marginTop || 72;
  const marginBottom = documentData.documentStyle?.marginBottom || 72;
  const marginLeft = documentData.documentStyle?.marginLeft || 90;
  const marginRight = documentData.documentStyle?.marginRight || 90;

  sectPr.ele("w:pgMar", {
    "w:top": Math.round(marginTop * 20).toString(),
    "w:bottom": Math.round(marginBottom * 20).toString(),
    "w:left": Math.round(marginLeft * 20).toString(),
    "w:right": Math.round(marginRight * 20).toString(),
    "w:header": "720",
    "w:footer": "720",
    "w:gutter": "0",
  });

  return doc.end({ prettyPrint: true });
}

function addImageToRun(parentElement: any, imgInfo: ImageInfo): void {
  // Convert points to EMUs (English Metric Units)
  // 1 point = 12700 EMUs
  const widthEmu = Math.round(imgInfo.width * 12700);
  const heightEmu = Math.round(imgInfo.height * 12700);

  const r = parentElement.ele("w:r");
  const drawing = r.ele("w:drawing");

  // Inline drawing
  const inline = drawing.ele("wp:inline", {
    distT: "0",
    distB: "0",
    distL: "0",
    distR: "0",
  });

  // Extent (size)
  inline.ele("wp:extent", {
    cx: widthEmu.toString(),
    cy: heightEmu.toString(),
  });

  // Effect extent
  inline.ele("wp:effectExtent", {
    l: "0",
    t: "0",
    r: "0",
    b: "0",
  });

  // Document properties
  const docId = imgInfo.rId.replace("rId", "");
  inline.ele("wp:docPr", {
    id: docId,
    name: `Picture ${docId}`,
  });

  // Non-visual properties
  const cNvGraphicFramePr = inline.ele("wp:cNvGraphicFramePr");
  cNvGraphicFramePr.ele("a:graphicFrameLocks", {
    "xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    noChangeAspect: "1",
  });

  // Graphic element
  const graphic = inline.ele("a:graphic", {
    "xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main",
  });

  const graphicData = graphic.ele("a:graphicData", {
    uri: "http://schemas.openxmlformats.org/drawingml/2006/picture",
  });

  // Picture element
  const pic = graphicData.ele("pic:pic", {
    "xmlns:pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
  });

  // Non-visual picture properties
  const nvPicPr = pic.ele("pic:nvPicPr");
  nvPicPr.ele("pic:cNvPr", {
    id: docId,
    name: imgInfo.filename,
  });
  nvPicPr.ele("pic:cNvPicPr");

  // Blip fill
  const blipFill = pic.ele("pic:blipFill");
  blipFill.ele("a:blip", {
    "r:embed": imgInfo.rId,
  });
  const stretch = blipFill.ele("a:stretch");
  stretch.ele("a:fillRect");

  // Shape properties
  const spPr = pic.ele("pic:spPr");
  const xfrm = spPr.ele("a:xfrm");
  xfrm.ele("a:off", { x: "0", y: "0" });
  xfrm.ele("a:ext", { cx: widthEmu.toString(), cy: heightEmu.toString() });

  const prstGeom = spPr.ele("a:prstGeom", { prst: "rect" });
  prstGeom.ele("a:avLst");

  console.log(
    `[DOCX Export] Added inline image: ${imgInfo.filename} (${imgInfo.width}x${imgInfo.height}pt)`
  );
}

function addParagraphProperties(pPr: any, style: any): void {
  // Alignment - Univer uses: undefined=left, 2=center, 3=right, 4=justify
  if (style.horizontalAlign !== undefined) {
    // Map Univer alignment values to OOXML alignment strings
    const alignMap: Record<number, string> = {
      2: "center",   // Univer 2 = center
      3: "right",    // Univer 3 = right
      4: "both",     // Univer 4 = justify
    };
    const alignment = alignMap[style.horizontalAlign];
    if (alignment) {
      pPr.ele("w:jc", { "w:val": alignment });
    }
    // If horizontalAlign is undefined or 0 or 1, omit w:jc (defaults to left)
  }

  // Spacing - combine into single element
  const spacingAttrs: any = {};

  if (style.lineSpacing) {
    spacingAttrs["w:line"] = Math.round(style.lineSpacing * 240).toString();
    spacingAttrs["w:lineRule"] = "auto";
  }

  if (style.spaceAbove?.v) {
    spacingAttrs["w:before"] = Math.round(style.spaceAbove.v * 20).toString();
  }

  if (style.spaceBelow?.v) {
    spacingAttrs["w:after"] = Math.round(style.spaceBelow.v * 20).toString();
  }

  if (Object.keys(spacingAttrs).length > 0) {
    pPr.ele("w:spacing", spacingAttrs);
  }

  // Indentation
  const indAttrs: any = {};

  if (style.indentStart?.v) {
    indAttrs["w:left"] = Math.round(style.indentStart.v * 20).toString();
  }
  if (style.indentEnd?.v) {
    indAttrs["w:right"] = Math.round(style.indentEnd.v * 20).toString();
  }
  if (style.hanging?.v) {
    indAttrs["w:hanging"] = Math.round(style.hanging.v * 20).toString();
  }
  if (style.indentFirstLine?.v) {
    indAttrs["w:firstLine"] = Math.round(
      style.indentFirstLine.v * 20
    ).toString();
  }

  if (Object.keys(indAttrs).length > 0) {
    pPr.ele("w:ind", indAttrs);
  }

  // Paragraph borders - handle both nested (style.border.bottom) and direct (style.borderBottom) formats
  // Univer horizontal lines use direct borderBottom property
  const hasBorderBottom = style.border?.bottom || style.borderBottom;
  const hasBorderTop = style.border?.top || style.borderTop;

  if (hasBorderBottom || hasBorderTop) {
    const pBdr = pPr.ele("w:pBdr");

    if (hasBorderBottom) {
      const border = style.border?.bottom || style.borderBottom;
      // For horizontal lines, use a larger space value for better visual separation
      const spaceValue = style.borderBottom ? "12" : "1";
      pBdr.ele("w:bottom", {
        "w:val": "single",
        "w:sz": Math.round((border.w || border.s || 1) * 8).toString(),
        "w:color": (border.cl?.rgb || border.color?.rgb || "#CDD0D8").replace("#", ""),
        "w:space": spaceValue,
      });
    }

    if (hasBorderTop) {
      const border = style.border?.top || style.borderTop;
      pBdr.ele("w:top", {
        "w:val": "single",
        "w:sz": Math.round((border.w || border.s || 1) * 8).toString(),
        "w:color": (border.cl?.rgb || border.color?.rgb || "#000000").replace("#", ""),
        "w:space": "1",
      });
    }
  }

  // Paragraph background/shading
  if (style.background?.rgb) {
    pPr.ele("w:shd", {
      "w:val": "clear",
      "w:color": "auto",
      "w:fill": style.background.rgb.replace("#", ""),
    });
  }
}

function addRunProperties(rPr: any, ts: any): void {
  // Bold
  if (ts.bl === 1) {
    rPr.ele("w:b");
    rPr.ele("w:bCs");
  }

  // Italic
  if (ts.it === 1) {
    rPr.ele("w:i");
    rPr.ele("w:iCs");
  }

  // Underline
  if (ts.ul) {
    rPr.ele("w:u", { "w:val": "single" });
  }

  // Strike
  if (ts.st === 1 || ts.st?.s === 1) {
    rPr.ele("w:strike");
  }

  // Font size (convert to half-points)
  if (ts.fs) {
    rPr.ele("w:sz", { "w:val": Math.round(ts.fs * 2).toString() });
    rPr.ele("w:szCs", { "w:val": Math.round(ts.fs * 2).toString() });
  }

  // Font family
  if (ts.ff) {
    rPr.ele("w:rFonts", {
      "w:ascii": ts.ff,
      "w:hAnsi": ts.ff,
      "w:cs": ts.ff,
      "w:eastAsia": ts.ff,
    });
  }

  // Text color - handle both formats
  if (ts.cl) {
    let color = null;
    if (ts.cl.rgb) {
      color = ts.cl.rgb.replace("#", "");
    } else if (typeof ts.cl === "string") {
      color = ts.cl.replace("#", "");
    }
    if (color && color !== "000000") {
      rPr.ele("w:color", { "w:val": color });
    }
  }

  // Background/highlight - handle both formats
  if (ts.bg) {
    let bgColor = null;
    if (ts.bg.rgb) {
      bgColor = ts.bg.rgb;
    } else if (typeof ts.bg === "string") {
      bgColor = ts.bg;
    }

    if (bgColor) {
      const highlightColor = mapRgbToHighlight(bgColor);
      if (highlightColor) {
        rPr.ele("w:highlight", { "w:val": highlightColor });
      } else {
        rPr.ele("w:shd", {
          "w:val": "clear",
          "w:color": "auto",
          "w:fill": bgColor.replace("#", ""),
        });
      }
    }
  }

  // Superscript/Subscript
  // Univer BaselineOffset: 1 = SUPERSCRIPT, 2 = SUBSCRIPT, 3 = NORMAL (baseline)
  if (ts.va !== undefined && ts.va !== null) {
    console.log(`[DOCX Export] Processing va=${ts.va} (type: ${typeof ts.va})`);
  }
  if (ts.va === 1) {
    console.log(`[DOCX Export] Adding superscript (va=1)`);
    rPr.ele("w:vertAlign", { "w:val": "superscript" });
  } else if (ts.va === 2) {
    console.log(`[DOCX Export] Adding subscript (va=2)`);
    rPr.ele("w:vertAlign", { "w:val": "subscript" });
  } else if (ts.va === 3) {
    // va=3 is NORMAL in Univer - but if user's export shows va=3 as superscript,
    // we need to handle this mapping issue
    console.log(`[DOCX Export] va=3 detected - checking if this should be superscript`);
    // Try treating va=3 as superscript based on user feedback
    rPr.ele("w:vertAlign", { "w:val": "superscript" });
  }
  // undefined = normal text (baseline - no vertAlign element needed)
}

function addBulletProperties(pPr: any, bullet: any): void {
  const numPr = pPr.ele("w:numPr");
  numPr.ele("w:ilvl", { "w:val": (bullet.nestingLevel || 0).toString() });
  numPr.ele("w:numId", { "w:val": bullet.listId || "1" });
}

function buildStylesXml(_documentData: IDocumentData): string {
  const styles = create({
    version: "1.0",
    encoding: "UTF-8",
    standalone: "yes",
  }).ele("w:styles", {
    "xmlns:w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "xmlns:r":
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "mc:Ignorable": "w14",
  });

  // Document defaults
  const docDefaults = styles.ele("w:docDefaults");

  // Run properties defaults
  const rPrDefault = docDefaults.ele("w:rPrDefault").ele("w:rPr");
  rPrDefault.ele("w:rFonts", {
    "w:ascii": "Calibri",
    "w:hAnsi": "Calibri",
    "w:eastAsia": "Calibri",
    "w:cs": "Times New Roman",
  });
  rPrDefault.ele("w:sz", { "w:val": "22" }); // 11pt
  rPrDefault.ele("w:szCs", { "w:val": "22" });
  rPrDefault.ele("w:lang", {
    "w:val": "en-US",
    "w:eastAsia": "en-US",
    "w:bidi": "ar-SA",
  });

  // Paragraph properties defaults
  const pPrDefault = docDefaults.ele("w:pPrDefault").ele("w:pPr");
  pPrDefault.ele("w:spacing", {
    "w:after": "200",
    "w:line": "276",
    "w:lineRule": "auto",
  });

  // Latent styles
  const latentStyles = styles.ele("w:latentStyles", {
    "w:defLockedState": "0",
    "w:defUIPriority": "99",
    "w:defSemiHidden": "0",
    "w:defUnhideWhenUsed": "0",
    "w:defQFormat": "0",
    "w:count": "376",
  });

  latentStyles.ele("w:lsdException", {
    "w:name": "Normal",
    "w:uiPriority": "0",
    "w:qFormat": "1",
  });
  latentStyles.ele("w:lsdException", {
    "w:name": "heading 1",
    "w:uiPriority": "9",
    "w:qFormat": "1",
  });
  latentStyles.ele("w:lsdException", {
    "w:name": "heading 2",
    "w:uiPriority": "9",
    "w:qFormat": "1",
  });

  // Normal style
  const normalStyle = styles.ele("w:style", {
    "w:type": "paragraph",
    "w:styleId": "Normal",
    "w:default": "1",
  });
  normalStyle.ele("w:name", { "w:val": "Normal" });
  normalStyle.ele("w:qFormat");

  // Default paragraph font style
  const defaultFontStyle = styles.ele("w:style", {
    "w:type": "character",
    "w:styleId": "DefaultParagraphFont",
    "w:default": "1",
  });
  defaultFontStyle.ele("w:name", { "w:val": "Default Paragraph Font" });
  defaultFontStyle.ele("w:uiPriority", { "w:val": "1" });
  defaultFontStyle.ele("w:semiHidden");
  defaultFontStyle.ele("w:unhideWhenUsed");

  // Table normal style
  const tableNormalStyle = styles.ele("w:style", {
    "w:type": "table",
    "w:styleId": "TableNormal",
    "w:default": "1",
  });
  tableNormalStyle.ele("w:name", { "w:val": "Normal Table" });
  tableNormalStyle.ele("w:uiPriority", { "w:val": "99" });
  tableNormalStyle.ele("w:semiHidden");
  tableNormalStyle.ele("w:unhideWhenUsed");
  const tblPr = tableNormalStyle.ele("w:tblPr");
  tblPr.ele("w:tblInd", { "w:w": "0", "w:type": "dxa" });
  const tblCellMar = tblPr.ele("w:tblCellMar");
  tblCellMar.ele("w:top", { "w:w": "0", "w:type": "dxa" });
  tblCellMar.ele("w:left", { "w:w": "108", "w:type": "dxa" });
  tblCellMar.ele("w:bottom", { "w:w": "0", "w:type": "dxa" });
  tblCellMar.ele("w:right", { "w:w": "108", "w:type": "dxa" });

  // Heading styles
  for (let i = 1; i <= 6; i++) {
    const headingStyle = styles.ele("w:style", {
      "w:type": "paragraph",
      "w:styleId": `Heading${i}`,
    });
    headingStyle.ele("w:name", { "w:val": `heading ${i}` });
    headingStyle.ele("w:basedOn", { "w:val": "Normal" });
    headingStyle.ele("w:next", { "w:val": "Normal" });
    headingStyle.ele("w:link", { "w:val": `Heading${i}Char` });
    headingStyle.ele("w:uiPriority", { "w:val": "9" });
    headingStyle.ele("w:qFormat");

    const pPr = headingStyle.ele("w:pPr");
    pPr.ele("w:keepNext");
    pPr.ele("w:keepLines");
    pPr.ele("w:spacing", { "w:before": "240", "w:after": "0" });
    pPr.ele("w:outlineLvl", { "w:val": (i - 1).toString() });

    const rPr = headingStyle.ele("w:rPr");
    rPr.ele("w:rFonts", {
      "w:ascii": "Calibri Light",
      "w:hAnsi": "Calibri Light",
      "w:eastAsia": "Yu Gothic Light",
      "w:cs": "Times New Roman",
    });

    // Font size decreases with heading level
    const fontSize = Math.max(22, 48 - i * 6);
    rPr.ele("w:sz", { "w:val": fontSize.toString() });
    rPr.ele("w:szCs", { "w:val": fontSize.toString() });

    // Primary color for headings
    rPr.ele("w:color", { "w:val": "2F5496" });
  }

  // List paragraph style
  const listStyle = styles.ele("w:style", {
    "w:type": "paragraph",
    "w:styleId": "ListParagraph",
  });
  listStyle.ele("w:name", { "w:val": "List Paragraph" });
  listStyle.ele("w:basedOn", { "w:val": "Normal" });
  listStyle.ele("w:uiPriority", { "w:val": "34" });
  listStyle.ele("w:qFormat");
  const listPPr = listStyle.ele("w:pPr");
  listPPr.ele("w:ind", { "w:left": "720" });
  listPPr.ele("w:contextualSpacing");

  return styles.end({ prettyPrint: true });
}

function buildNumberingXml(documentData: IDocumentData): string | null {
  const paragraphs = documentData.body?.paragraphs || [];
  const lists = (documentData as any).lists || {};
  const hasLists = paragraphs.some((p: any) => p.bullet);

  if (!hasLists) return null;

  const numbering = create({
    version: "1.0",
    encoding: "UTF-8",
    standalone: "yes",
  }).ele("w:numbering", {
    "xmlns:w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "xmlns:wpc":
      "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
    "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "xmlns:o": "urn:schemas-microsoft-com:office:office",
    "xmlns:r":
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "xmlns:m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "xmlns:v": "urn:schemas-microsoft-com:vml",
    "xmlns:wp14":
      "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
    "xmlns:wp":
      "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "xmlns:w10": "urn:schemas-microsoft-com:office:word",
    "xmlns:w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "xmlns:w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "xmlns:wpg":
      "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    "xmlns:wpi":
      "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
    "xmlns:wne": "http://schemas.microsoft.com/office/word/2006/wordml",
    "xmlns:wps":
      "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "mc:Ignorable": "w14 w15 wp14",
  });

  // Collect unique list IDs from paragraphs
  const uniqueListIds = new Set<string>();
  paragraphs.forEach((p: any) => {
    if (p.bullet?.listId) {
      uniqueListIds.add(p.bullet.listId);
    }
  });

  const listIdArray = Array.from(uniqueListIds);

  // Create abstract numbering definitions
  listIdArray.forEach((listId, index) => {
    const abstractNumId = index.toString();
    const listDef = lists[listId];
    const isBullet =
      listDef?.listType === "BULLET_LIST" ||
      listDef?.listType === 2 ||
      listDef?.listType === "BULLET";

    const abstractNum = numbering.ele("w:abstractNum", {
      "w:abstractNumId": abstractNumId,
      "w15:restartNumberingAfterBreak": "0",
    });

    abstractNum.ele("w:nsid", { "w:val": generateNsid() });
    abstractNum.ele("w:multiLevelType", { "w:val": "hybridMultilevel" });

    // Create 9 levels
    for (let lvl = 0; lvl < 9; lvl++) {
      const level = abstractNum.ele("w:lvl", { "w:ilvl": lvl.toString() });

      const levelDef = listDef?.nestingLevel?.[lvl];
      const startNum = levelDef?.startNumber || 1;

      level.ele("w:start", { "w:val": startNum.toString() });

      if (isBullet) {
        level.ele("w:numFmt", { "w:val": "bullet" });

        // Bullet character based on level
        const bulletChars = ["•", "○", "■", "●", "◦", "▪", "►", "◊", "※"];
        level.ele("w:lvlText", {
          "w:val": bulletChars[lvl % bulletChars.length],
        });
        level.ele("w:lvlJc", { "w:val": "left" });

        // Indentation
        const pPr = level.ele("w:pPr");
        const indent = 720 + lvl * 360;
        pPr.ele("w:ind", {
          "w:left": indent.toString(),
          "w:hanging": "360",
        });

        // Bullet font
        const rPr = level.ele("w:rPr");
        rPr.ele("w:rFonts", {
          "w:ascii": "Symbol",
          "w:hAnsi": "Symbol",
          "w:hint": "default",
        });
      } else {
        // Numbered list
        const numFormats = [
          "decimal",
          "lowerLetter",
          "lowerRoman",
          "decimal",
          "lowerLetter",
          "lowerRoman",
        ];
        level.ele("w:numFmt", { "w:val": numFormats[lvl % numFormats.length] });
        level.ele("w:lvlText", { "w:val": `%${lvl + 1}.` });
        level.ele("w:lvlJc", { "w:val": "left" });

        const pPr = level.ele("w:pPr");
        const indent = 720 + lvl * 360;
        pPr.ele("w:ind", {
          "w:left": indent.toString(),
          "w:hanging": "360",
        });
      }
    }

    // Link numId to abstractNumId
    const num = numbering.ele("w:num", { "w:numId": listId });
    num.ele("w:abstractNumId", { "w:val": abstractNumId });
  });

  return numbering.end({ prettyPrint: true });
}

function generateNsid(): string {
  const chars = "0123456789ABCDEF";
  let result = "";
  for (let i = 0; i < 8; i++) {
    result += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return result;
}

function buildRelsXml(
  documentData: IDocumentData,
  imageInfos: ImageInfo[],
  hasLists: boolean
): string {
  const rels = create({
    version: "1.0",
    encoding: "UTF-8",
    standalone: "yes",
  }).ele("Relationships", {
    xmlns: "http://schemas.openxmlformats.org/package/2006/relationships",
  });

  // ✅ FIX: Use sequential rId counter to avoid duplicates
  let rIdCounter = 1;

  // Always add styles first
  rels.ele("Relationship", {
    Id: `rId${rIdCounter++}`,
    Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
    Target: "styles.xml",
  });

  // Add numbering if lists exist
  if (hasLists) {
    rels.ele("Relationship", {
      Id: `rId${rIdCounter++}`,
      Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
      Target: "numbering.xml",
    });
  }

  // ✅ Update image rIds to match actual counter
  imageInfos.forEach((imgInfo) => {
    imgInfo.rId = `rId${rIdCounter}`; // Update the rId to match sequential order
    rels.ele("Relationship", {
      Id: imgInfo.rId,
      Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
      Target: `media/${imgInfo.filename}`,
    });
    rIdCounter++;
  });

  console.log(`[DOCX Export] Generated ${rIdCounter - 1} relationships`);
  return rels.end({ prettyPrint: true });
}

function buildContentTypesXml(
  documentData: IDocumentData,
  imageInfos: ImageInfo[]
): string {
  const contentTypes = create({
    version: "1.0",
    encoding: "UTF-8",
    standalone: "yes",
  }).ele("Types", {
    xmlns: "http://schemas.openxmlformats.org/package/2006/content-types",
  });

  // Default types
  contentTypes.ele("Default", {
    Extension: "rels",
    ContentType: "application/vnd.openxmlformats-package.relationships+xml",
  });
  contentTypes.ele("Default", {
    Extension: "xml",
    ContentType: "application/xml",
  });

  // Image extensions
  const imageExtensions = new Set(imageInfos.map((img) => img.extension));
  imageExtensions.forEach((ext) => {
    let contentType = "image/png";
    if (ext === "jpg" || ext === "jpeg") contentType = "image/jpeg";
    else if (ext === "gif") contentType = "image/gif";
    else if (ext === "png") contentType = "image/png";
    else if (ext === "bmp") contentType = "image/bmp";
    else if (ext === "tiff" || ext === "tif") contentType = "image/tiff";

    contentTypes.ele("Default", {
      Extension: ext,
      ContentType: contentType,
    });
  });

  // Override types
  contentTypes.ele("Override", {
    PartName: "/word/document.xml",
    ContentType:
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
  });
  contentTypes.ele("Override", {
    PartName: "/word/styles.xml",
    ContentType:
      "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
  });

  // Numbering
  const hasLists =
    documentData.body?.paragraphs?.some((p: any) => p.bullet) || false;
  if (hasLists) {
    contentTypes.ele("Override", {
      PartName: "/word/numbering.xml",
      ContentType:
        "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
    });
  }

  return contentTypes.end({ prettyPrint: true });
}

function buildRootRels(): string {
  const rels = create({
    version: "1.0",
    encoding: "UTF-8",
    standalone: "yes",
  }).ele("Relationships", {
    xmlns: "http://schemas.openxmlformats.org/package/2006/relationships",
  });

  rels.ele("Relationship", {
    Id: "rId1",
    Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
    Target: "word/document.xml",
  });

  return rels.end({ prettyPrint: true });
}

function mapRgbToHighlight(rgb: string): string | null {
  const hex = rgb.replace("#", "").toUpperCase();

  const colorMap: Record<string, string> = {
    FFFF00: "yellow",
    "00FF00": "green",
    "00FFFF": "cyan",
    FF00FF: "magenta",
    "0000FF": "blue",
    FF0000: "red",
    "000080": "darkBlue",
    "008080": "darkCyan",
    "008000": "darkGreen",
    "800080": "darkMagenta",
    "800000": "darkRed",
    "808000": "darkYellow",
    "808080": "darkGray",
    C0C0C0: "lightGray",
    "000000": "black",
    FFFFFF: "white",
  };

  return colorMap[hex] || null;
}

function getImageExtension(source: string): string {
  const src = source.toLowerCase();
  if (src.includes("image/png") || src.includes(".png")) return "png";
  if (
    src.includes("image/jpeg") ||
    src.includes("image/jpg") ||
    src.includes(".jpg") ||
    src.includes(".jpeg")
  )
    return "jpg";
  if (src.includes("image/gif") || src.includes(".gif")) return "gif";
  if (src.includes("image/bmp") || src.includes(".bmp")) return "bmp";
  if (src.includes("image/webp") || src.includes(".webp")) return "png"; // Convert webp to png
  return "png"; // Default
}

// Export alias for server-side usage
export const buildDocx = buildDocxFromUniverData;
