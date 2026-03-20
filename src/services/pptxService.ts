import JSZip from "jszip";
import pptxgen from "pptxgenjs";
import { HymnSlot, LiturgyData } from "./geminiService";

export interface PPTXFile {
  name: string;
  data: ArrayBuffer;
}

// Helper to parse XML in browser
function parseXml(xmlStr: string): Document {
  const parser = new DOMParser();
  return parser.parseFromString(xmlStr, "application/xml");
}

export async function generateSampleMaster(): Promise<Blob> {
  const pres = new pptxgen();

  // Title Slide
  const slide1 = pres.addSlide();
  slide1.addText("{{Liturgy Sheet Title}}", { x: 1, y: 1, w: "80%", h: 1, fontSize: 44, align: "center", fontFace: "Georgia" });
  slide1.addText("Sunday Liturgy", { x: 1, y: 2, w: "80%", h: 0.5, fontSize: 24, align: "center", italic: true });

  // Entrance Hymn
  const slide2 = pres.addSlide();
  slide2.addText("Entrance Hymn", { x: 0.5, y: 0.5, fontSize: 18, bold: true });
  slide2.addText("{{Entrance Hymn}}", { x: 1, y: 2, w: "80%", h: 1, fontSize: 32, align: "center", color: "363636" });

  // First Reading
  const slide3 = pres.addSlide();
  slide3.addText("First Reading", { x: 0.5, y: 0.5, fontSize: 18, bold: true });
  slide3.addText("{{First Reading Verse}}", { x: 0.5, y: 1, w: "90%", fontSize: 14, italic: true });
  slide3.addText("{{FIRST_READING_TEXT}}", { x: 0.5, y: 1.5, w: "90%", h: 3, fontSize: 18, align: "left", valign: "top" });

  // Psalm
  const slide4 = pres.addSlide();
  slide4.addText("Responsorial Psalm", { x: 0.5, y: 0.5, fontSize: 18, bold: true });
  slide4.addText("Response:", { x: 0.5, y: 1.5, fontSize: 14, bold: true });
  slide4.addText("{{Response}}", { x: 0.5, y: 2, w: "90%", h: 1, fontSize: 24, align: "center", italic: true });

  // Psalm Verses (Placeholders for splitting)
  const slide4v = pres.addSlide();
  slide4v.addText("Psalm Verse", { x: 0.5, y: 0.5, fontSize: 18, bold: true });
  slide4v.addText("{{Responsorial Verse 1}}", { x: 0.5, y: 1.5, w: "90%", h: 3, fontSize: 20, align: "center" });

  // Gospel
  const slide5 = pres.addSlide();
  slide5.addText("Holy Gospel", { x: 0.5, y: 0.5, fontSize: 18, bold: true });
  slide5.addText("{{Gospel Verse Reference}}", { x: 0.5, y: 1, w: "90%", fontSize: 14, italic: true });
  slide5.addText("{{GOSPEL_TEXT}}", { x: 0.5, y: 1.5, w: "90%", h: 3, fontSize: 18, align: "left", valign: "top" });

  // Prayer
  const slide6p = pres.addSlide();
  slide6p.addText("{{Prayer Title}}", { x: 0.5, y: 0.5, fontSize: 18, bold: true });
  slide6p.addText("{{Prayer}}", { x: 0.5, y: 1.5, w: "90%", h: 3, fontSize: 20, align: "center" });

  // Gloria
  const slide6 = pres.addSlide();
  slide6.addText("Gloria", { x: 0.5, y: 0.5, fontSize: 18, bold: true });
  slide6.addText("{{Gloria}}", { x: 1, y: 2, w: "80%", h: 1, fontSize: 32, align: "center" });

  // Communion
  const slide7 = pres.addSlide();
  slide7.addText("Communion", { x: 0.5, y: 0.5, fontSize: 18, bold: true });
  slide7.addText("{{Communion Hymn}}", { x: 1, y: 2, w: "80%", h: 1, fontSize: 32, align: "center" });

  const buffer = await pres.write({ outputType: "arraybuffer" }) as ArrayBuffer;
  return new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" });
}

// Helper to calculate dynamic font size based on text length
function getDynamicFontSize(text: string, baseSize: number = 24): number {
  const length = text.length;
  if (length < 100) return baseSize;
  if (length < 200) return Math.max(baseSize * 0.85, 18);
  if (length < 400) return Math.max(baseSize * 0.7, 14);
  if (length < 600) return Math.max(baseSize * 0.6, 12);
  return Math.max(baseSize * 0.5, 10);
}

// Helper to split text into chunks for slides
function splitTextForSlides(text: string, maxChars: number = 500): string[] {
  if (text.length <= maxChars) return [text];
  
  const chunks: string[] = [];
  let current = text;
  
  while (current.length > 0) {
    if (current.length <= maxChars) {
      chunks.push(current);
      break;
    }
    
    // Try to find a good break point (period or newline)
    let breakPoint = current.lastIndexOf(". ", maxChars);
    if (breakPoint === -1) breakPoint = current.lastIndexOf("\n", maxChars);
    if (breakPoint === -1) breakPoint = current.lastIndexOf("; ", maxChars);
    if (breakPoint === -1) breakPoint = current.lastIndexOf(", ", maxChars);
    if (breakPoint === -1) breakPoint = current.lastIndexOf(" ", maxChars);
    if (breakPoint === -1) breakPoint = maxChars;
    
    chunks.push(current.substring(0, breakPoint + 1).trim());
    current = current.substring(breakPoint + 1).trim();
  }
  
  return chunks;
}

export async function mergePPTX(
  masterBuffer: ArrayBuffer,
  hymnFiles: Map<string, ArrayBuffer>, // slotName -> hymnData
  slotMappings: HymnSlot[],
  liturgyData?: LiturgyData,
  onLog?: (message: string, type?: "info" | "success" | "error") => void
): Promise<Blob> {
  const log = (msg: string, type: "info" | "success" | "error" = "info") => {
    console.log(msg);
    if (onLog) onLog(msg, type);
  };

  log("Starting presentation assembly...", "info");
  const masterZip = await JSZip.loadAsync(masterBuffer);
  log("Master template loaded.", "success");

  // 1. Pre-extract slide XMLs and Rels from matched hymns
  const hymnSlides = new Map<string, { xml: string; rels: string | null; media: Map<string, Uint8Array> }[]>();
  for (const mapping of slotMappings) {
    const buffer = hymnFiles.get(mapping.slot);
    
    if (!buffer) {
      if (mapping.fullLyrics) {
        log(`No file found for ${mapping.hymnName}. Will use lyrics from list.`, "info");
        continue;
      }
      log(`No file or lyrics found for hymn slot: ${mapping.slot}`, "info");
      continue;
    }

    log(`Extracting slides for ${mapping.slot}: ${mapping.hymnName}...`, "info");
    try {
      const hymnZip = await JSZip.loadAsync(buffer);
      const presXmlStr = await hymnZip.file("ppt/presentation.xml")?.async("string");
      if (!presXmlStr) continue;

      const presDoc = parseXml(presXmlStr);
      const sldIds = Array.from(presDoc.getElementsByTagName("p:sldId"));
      
      const presRelsStr = await hymnZip.file("ppt/_rels/presentation.xml.rels")?.async("string");
      const relsDoc = parseXml(presRelsStr || "");
      const relMap = new Map<string, string>();
      Array.from(relsDoc.getElementsByTagName("Relationship")).forEach(rel => {
        relMap.set(rel.getAttribute("Id") || "", rel.getAttribute("Target") || "");
      });

      let slides: { xml: string; rels: string | null; media: Map<string, Uint8Array> }[] = [];
      for (const sldId of sldIds) {
        const relId = sldId.getAttribute("r:id");
        const target = relMap.get(relId || "");
        if (target) {
          const sldPath = `ppt/${target.replace(/^\//, "")}`;
          const sldXml = await hymnZip.file(sldPath)?.async("string");
          const sldRelPath = `ppt/slides/_rels/${target.split("/").pop()}.rels`;
          const sldRelXml = await hymnZip.file(sldRelPath)?.async("string");
          
          const mediaMap = new Map<string, Uint8Array>();
          if (sldRelXml) {
            const sldRelsDoc = parseXml(sldRelXml);
            const sldRels = Array.from(sldRelsDoc.getElementsByTagName("Relationship"));
            for (const rel of sldRels) {
              const relTarget = rel.getAttribute("Target") || "";
              if (relTarget.includes("../media/")) {
                const mediaPath = `ppt/${relTarget.replace("..", "ppt").replace("ppt/ppt/", "ppt/")}`;
                const mediaData = await hymnZip.file(mediaPath)?.async("uint8array");
                if (mediaData) {
                  mediaMap.set(relTarget, mediaData);
                }
              }
            }
          }

          if (sldXml) {
            slides.push({ xml: sldXml, rels: sldRelXml || null, media: mediaMap });
          }
        }
      }
      // Use hymnName as key to handle multiple hymns in same slot
      hymnSlides.set(mapping.hymnName, slides);
      log(`Successfully extracted ${slides.length} slides for ${mapping.hymnName}.`, "success");
    } catch (err) {
      log(`Failed to extract slides for ${mapping.hymnName}`, "error");
      console.error(`Failed to extract slides for ${mapping.hymnName}:`, err);
    }
  }
  
  // 2. Parse master presentation.xml and rels
  log("Parsing master slide structure...", "info");
  const presXmlStr = await masterZip.file("ppt/presentation.xml")?.async("string");
  if (!presXmlStr) throw new Error("Invalid Master PPTX");
  const presDoc = parseXml(presXmlStr);
  const sldIdLst = presDoc.getElementsByTagName("p:sldIdLst")[0];
  const originalSldIds = Array.from(sldIdLst.getElementsByTagName("p:sldId"));

  const presRelsStr = await masterZip.file("ppt/_rels/presentation.xml.rels")?.async("string");
  const relsDoc = parseXml(presRelsStr || "");
  const relationships = relsDoc.getElementsByTagName("Relationships")[0];
  const originalRels = Array.from(relsDoc.getElementsByTagName("Relationship"));
  const relMap = new Map<string, Element>();
  originalRels.forEach(rel => relMap.set(rel.getAttribute("Id") || "", rel));

  // Find a valid slide layout relationship in the master to use as fallback
  let masterLayoutRel: { target: string; type: string } | null = null;
  const firstSlideRelPath = `ppt/slides/_rels/slide1.xml.rels`;
  const firstSlideRelStr = await masterZip.file(firstSlideRelPath)?.async("string");
  if (firstSlideRelStr) {
    const firstSlideRelDoc = parseXml(firstSlideRelStr);
    const rels = Array.from(firstSlideRelDoc.getElementsByTagName("Relationship"));
    const layoutRel = rels.find(r => r.getAttribute("Type")?.includes("slideLayout"));
    if (layoutRel) {
      masterLayoutRel = {
        target: layoutRel.getAttribute("Target") || "",
        type: layoutRel.getAttribute("Type") || ""
      };
    }
  }

  // 3. Parse [Content_Types].xml
  const contentTypesStr = await masterZip.file("[Content_Types].xml")?.async("string");
  if (!contentTypesStr) throw new Error("Missing [Content_Types].xml");
  const contentTypesDoc = parseXml(contentTypesStr);
  const typesNode = contentTypesDoc.getElementsByTagName("Types")[0];

  // Ensure common media types are in [Content_Types].xml
  const commonDefaults = [
    { Extension: "png", ContentType: "image/png" },
    { Extension: "jpg", ContentType: "image/jpeg" },
    { Extension: "jpeg", ContentType: "image/jpeg" },
    { Extension: "gif", ContentType: "image/gif" },
    { Extension: "rels", ContentType: "application/vnd.openxmlformats-package.relationships+xml" },
    { Extension: "xml", ContentType: "application/xml" }
  ];

  const existingDefaults = Array.from(typesNode.getElementsByTagName("Default"));
  const existingExts = new Set(existingDefaults.map(d => d.getAttribute("Extension")?.toLowerCase()));

  for (const def of commonDefaults) {
    if (!existingExts.has(def.Extension)) {
      const newDef = contentTypesDoc.createElement("Default");
      newDef.setAttribute("Extension", def.Extension);
      newDef.setAttribute("ContentType", def.ContentType);
      typesNode.appendChild(newDef);
    }
  }

  // Clear existing slide list to rebuild it
  while (sldIdLst.firstChild) sldIdLst.removeChild(sldIdLst.firstChild);

  let nextRIdNum = 1000;
  let nextSldIdNum = 1000;
  let nextSlideFileNum = 1000;

  const escapeRegExp = (str: string) => str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const serializer = new XMLSerializer();

  // 4. Rebuild presentation
  log("Rebuilding presentation slides...", "info");
  for (const [slideIdx, oldSldId] of originalSldIds.entries()) {
    const relId = oldSldId.getAttribute("r:id") || "";
    const relNode = relMap.get(relId);
    if (!relNode) continue;

    const target = relNode.getAttribute("Target") || "";
    const slidePath = `ppt/${target.replace(/^\//, "")}`;
    const slideXmlStr = await masterZip.file(slidePath)?.async("string");
    if (!slideXmlStr) continue;

    const slideRelPath = `ppt/slides/_rels/${target.split("/").pop()}.rels`;
    const slideRelStr = await masterZip.file(slideRelPath)?.async("string");

    // Check if this slide is a hymn placeholder
    let matchedSlot: string | null = null;
    for (const mapping of slotMappings) {
      if (slideXmlStr.includes(`{{${mapping.slot}}}`)) {
        matchedSlot = mapping.slot;
        break;
      }
    }

    if (matchedSlot) {
      const hymnsForSlot = slotMappings.filter(m => m.slot === matchedSlot);
      
      if (hymnsForSlot.length > 0) {
        log(`Processing ${hymnsForSlot.length} hymns for ${matchedSlot}...`, "info");
        
        for (const mapping of hymnsForSlot) {
          let slides = hymnSlides.get(mapping.hymnName);
          
          // If no slides found in archive, generate from lyrics
          if (!slides && mapping.fullLyrics) {
            log(`Generating slides for ${mapping.hymnName} from lyrics...`, "info");
            slides = await generateSlidesFromLyrics(mapping.hymnName, mapping.fullLyrics);
          }

          if (slides && slides.length > 0) {
            log(`Adding ${slides.length} slides for ${mapping.hymnName}...`, "info");
            for (const [idx, sldData] of slides.entries()) {
              const newRelId = `rId${nextRIdNum++}`;
              const newSldId = `${nextSldIdNum++}`;
              const slideFileNum = nextSlideFileNum++;
              const newFileName = `slides/slide${slideFileNum}.xml`;

              let processedSldXml = sldData.xml;
              let processedRelXml = sldData.rels;

              if (processedRelXml) {
                const relDoc = parseXml(processedRelXml);
                const rels = Array.from(relDoc.getElementsByTagName("Relationship"));
                for (const rel of rels) {
                  const type = rel.getAttribute("Type") || "";
                  const target = rel.getAttribute("Target") || "";

                  if (type.includes("slideLayout") && masterLayoutRel) {
                    rel.setAttribute("Target", masterLayoutRel.target);
                  } else if (target.includes("../media/")) {
                    const mediaData = sldData.media.get(target);
                    if (mediaData) {
                      const mediaName = target.split("/").pop();
                      const newMediaPath = `media/hymn_${mapping.hymnName.replace(/\s+/g, "_")}_s${idx}_${mediaName}`;
                      masterZip.file(`ppt/${newMediaPath}`, mediaData);
                      rel.setAttribute("Target", `../${newMediaPath}`);
                    }
                  }
                }
                processedRelXml = serializer.serializeToString(relDoc);
              }
              
              masterZip.file(`ppt/${newFileName}`, processedSldXml);
              if (processedRelXml) {
                masterZip.file(`ppt/slides/_rels/slide${slideFileNum}.xml.rels`, processedRelXml);
              } else if (masterLayoutRel) {
                const basicRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${masterLayoutRel.type}" Target="${masterLayoutRel.target}"/>
</Relationships>`;
                masterZip.file(`ppt/slides/_rels/slide${slideFileNum}.xml.rels`, basicRels);
              }

              const newOverride = contentTypesDoc.createElement("Override");
              newOverride.setAttribute("PartName", `/ppt/${newFileName}`);
              newOverride.setAttribute("ContentType", "application/vnd.openxmlformats-officedocument.presentationml.slide+xml");
              typesNode.appendChild(newOverride);

              const newRel = relsDoc.createElement("Relationship");
              newRel.setAttribute("Id", newRelId);
              newRel.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide");
              newRel.setAttribute("Target", newFileName);
              relationships.appendChild(newRel);

              const newSldIdNode = presDoc.createElement("p:sldId");
              newSldIdNode.setAttribute("id", newSldId);
              newSldIdNode.setAttribute("r:id", newRelId);
              sldIdLst.appendChild(newSldIdNode);
            }
          } else {
            log(`No slides or lyrics found for ${mapping.hymnName}, skipping.`, "info");
          }
        }
      } else {
        log(`No hymns mapped to ${matchedSlot}, using placeholder slide.`, "info");
        // Fallback: Use master slide if no hymn slides found
        const newRelId = `rId${nextRIdNum++}`;
        const newSldId = `${nextSldIdNum++}`;
        const slideFileNum = nextSlideFileNum++;
        const newFileName = `slides/slide${slideFileNum}.xml`;
        masterZip.file(`ppt/${newFileName}`, slideXmlStr.replace(new RegExp(escapeRegExp(`{{${matchedSlot}}}`), "g"), matchedSlot));
        
        const newOverride = contentTypesDoc.createElement("Override");
        newOverride.setAttribute("PartName", `/ppt/${newFileName}`);
        newOverride.setAttribute("ContentType", "application/vnd.openxmlformats-officedocument.presentationml.slide+xml");
        typesNode.appendChild(newOverride);

        const newRel = relsDoc.createElement("Relationship");
        newRel.setAttribute("Id", newRelId);
        newRel.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide");
        newRel.setAttribute("Target", newFileName);
        relationships.appendChild(newRel);

        const newSldIdNode = presDoc.createElement("p:sldId");
        newSldIdNode.setAttribute("id", newSldId);
        newSldIdNode.setAttribute("r:id", newRelId);
        sldIdLst.appendChild(newSldIdNode);
      }
    } else {
      // Regular slide (Liturgy or other)
      // Check for long text that needs splitting
      const placeholdersToSplit = [
        "{{FIRST_READING_TEXT}}", 
        "{{SECOND_READING_TEXT}}", 
        "{{GOSPEL_TEXT}}",
        "{{Responsorial Verse 1}}",
        "{{Responsorial Verse 2}}",
        "{{Responsorial Verse 3}}",
        "{{Responsorial Verse 4}}",
        "{{Responsorial Verse 5}}",
        "{{Prayer}}"
      ];
      let needsSplitting = false;
      let splitKey = "";
      let splitChunks: string[] = [];

      if (liturgyData) {
        for (const key of placeholdersToSplit) {
          if (slideXmlStr.includes(key)) {
            let text = "";
            if (key === "{{FIRST_READING_TEXT}}") text = liturgyData.firstReading.text;
            else if (key === "{{SECOND_READING_TEXT}}") text = liturgyData.secondReading.text;
            else if (key === "{{GOSPEL_TEXT}}") text = liturgyData.gospel.text;
            else if (key === "{{Prayer}}") text = liturgyData.prayerText || "";
            else if (key.startsWith("{{Responsorial Verse")) {
              const index = parseInt(key.match(/\d+/)![0]) - 1;
              text = liturgyData.psalm.verses[index] || "";
            }

            if (text && text.length > 600) {
              needsSplitting = true;
              splitKey = key;
              splitChunks = splitTextForSlides(text);
              break;
            }
          }
        }
      }

      if (needsSplitting) {
        log(`Splitting long text for ${splitKey} into ${splitChunks.length} slides...`, "info");
        for (const chunk of splitChunks) {
          let processedXml = slideXmlStr;
          const replacements = getLiturgyReplacements(liturgyData);
          replacements[splitKey] = chunk;
          
          for (const [key, val] of Object.entries(replacements)) {
            processedXml = processedXml.replace(new RegExp(escapeRegExp(key), "g"), val || "");
          }

          const newRelId = `rId${nextRIdNum++}`;
          const newSldId = `${nextSldIdNum++}`;
          const slideFileNum = nextSlideFileNum++;
          const newFileName = `slides/slide${slideFileNum}.xml`;
          masterZip.file(`ppt/${newFileName}`, processedXml);
          if (slideRelStr) masterZip.file(`ppt/slides/_rels/slide${slideFileNum}.xml.rels`, slideRelStr);

          const newOverride = contentTypesDoc.createElement("Override");
          newOverride.setAttribute("PartName", `/ppt/${newFileName}`);
          newOverride.setAttribute("ContentType", "application/vnd.openxmlformats-officedocument.presentationml.slide+xml");
          typesNode.appendChild(newOverride);

          const newRel = relsDoc.createElement("Relationship");
          newRel.setAttribute("Id", newRelId);
          newRel.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide");
          newRel.setAttribute("Target", newFileName);
          relationships.appendChild(newRel);

          const newSldIdNode = presDoc.createElement("p:sldId");
          newSldIdNode.setAttribute("id", newSldId);
          newSldIdNode.setAttribute("r:id", newRelId);
          sldIdLst.appendChild(newSldIdNode);
        }
      } else {
        let processedXml = slideXmlStr;
        if (liturgyData) {
          const replacements = getLiturgyReplacements(liturgyData);
          let replacedCount = 0;
          for (const [key, val] of Object.entries(replacements)) {
            if (processedXml.includes(key)) {
              const fontSize = getDynamicFontSize(val || "", 18);
              // Try to adjust font size in XML if it's a reading or prayer
              if (key.includes("_TEXT") || key === "{{Prayer}}") {
                processedXml = processedXml.replace(/sz="\d+"/g, `sz="${Math.round(fontSize * 100)}"`);
              }
              processedXml = processedXml.replace(new RegExp(escapeRegExp(key), "g"), val || "");
              replacedCount++;
            }
          }
          if (replacedCount > 0) {
            log(`Replaced ${replacedCount} placeholders on slide ${slideIdx + 1}.`, "info");
          }
        }

        const newRelId = `rId${nextRIdNum++}`;
        const newSldId = `${nextSldIdNum++}`;
        const slideFileNum = nextSlideFileNum++;
        const newFileName = `slides/slide${slideFileNum}.xml`;
        masterZip.file(`ppt/${newFileName}`, processedXml);
        if (slideRelStr) masterZip.file(`ppt/slides/_rels/slide${slideFileNum}.xml.rels`, slideRelStr);

        const newOverride = contentTypesDoc.createElement("Override");
        newOverride.setAttribute("PartName", `/ppt/${newFileName}`);
        newOverride.setAttribute("ContentType", "application/vnd.openxmlformats-officedocument.presentationml.slide+xml");
        typesNode.appendChild(newOverride);

        const newRel = relsDoc.createElement("Relationship");
        newRel.setAttribute("Id", newRelId);
        newRel.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide");
        newRel.setAttribute("Target", newFileName);
        relationships.appendChild(newRel);

        const newSldIdNode = presDoc.createElement("p:sldId");
        newSldIdNode.setAttribute("id", newSldId);
        newSldIdNode.setAttribute("r:id", newRelId);
        sldIdLst.appendChild(newSldIdNode);
      }
    }
  }

  // Update presentation.xml, rels, and [Content_Types].xml
  masterZip.file("ppt/presentation.xml", serializer.serializeToString(presDoc));
  masterZip.file("ppt/_rels/presentation.xml.rels", serializer.serializeToString(relsDoc));
  masterZip.file("[Content_Types].xml", serializer.serializeToString(contentTypesDoc));

  log("Finalizing PPTX file...", "info");
  const finalBlob = await masterZip.generateAsync({ type: "blob" });
  log("Presentation assembled successfully!", "success");
  return finalBlob;
}

function getLiturgyReplacements(liturgyData: any): Record<string, string> {
  // Clean title: remove "- Year A", "- Year B", etc.
  const cleanTitle = liturgyData.title.replace(/\s*-\s*Year\s*[A-C]/gi, "").trim();

  const replacements: Record<string, string> = {
    "{{Liturgy Sheet Title}}": cleanTitle,
    
    "{{First Reading Title}}": liturgyData.firstReading.title,
    "{{FIRST_READING_REF}}": liturgyData.firstReading.reference,
    "{{FIRST_READING_TEXT}}": liturgyData.firstReading.text,
    "{{First Reading Verse}}": liturgyData.firstReading.reference,
    
    "{{Second Reading Title}}": liturgyData.secondReading.title,
    "{{SECOND_READING_REF}}": liturgyData.secondReading.reference,
    "{{SECOND_READING_TEXT}}": liturgyData.secondReading.text,
    "{{Second Reading Verse}}": liturgyData.secondReading.reference,
    
    "{{Gospel Title}}": liturgyData.gospel.title,
    "{{GOSPEL_REF}}": liturgyData.gospel.reference,
    "{{GOSPEL_TEXT}}": liturgyData.gospel.text,
    "{{Gospel Verse Reference}}": liturgyData.gospel.reference,
    
    "{{PSALM_RESPONSE}}": liturgyData.psalm.response,
    "{{Response}}": liturgyData.psalm.response,
    
    "{{Gospel Antiphon From Liturgy Sheet}}": liturgyData.gospelAntiphon,
    "{{Prayer of the Faithful Response}}": liturgyData.prayerOfTheFaithfulResponse,
    
    "{{Prayer Title}}": liturgyData.prayerTitle || "Prayer of the Faithful",
    "{{Prayer}}": liturgyData.prayerText || liturgyData.prayerOfTheFaithfulResponse,
  };
  if (liturgyData.psalm.verses) {
    for (let i = 0; i < 5; i++) {
      replacements[`{{Responsorial Verse ${i + 1}}}`] = liturgyData.psalm.verses[i] || "";
    }
  }
  return replacements;
}

/**
 * Splits text into chunks suitable for slides and generates slide XML.
 * Ensures vertical centering.
 */
async function generateSlidesFromLyrics(title: string, lyrics: string): Promise<{ xml: string; rels: string | null; media: Map<string, Uint8Array> }[]> {
  const lines = lyrics.split("\n").map(l => l.trim()).filter(l => l.length > 0);
  const chunks: string[][] = [];
  let currentChunk: string[] = [];
  
  for (const line of lines) {
    // Split lyrics if a single line is too long, or if we have enough lines
    if (currentChunk.length >= 4 || line.length > 100) {
      if (currentChunk.length > 0) chunks.push(currentChunk);
      currentChunk = [];
    }
    
    if (line.length > 100) {
      // Split very long line into two
      const mid = line.lastIndexOf(" ", 50);
      if (mid !== -1) {
        chunks.push([line.substring(0, mid).trim()]);
        chunks.push([line.substring(mid).trim()]);
      } else {
        chunks.push([line]);
      }
    } else {
      currentChunk.push(line);
    }
  }
  if (currentChunk.length > 0) chunks.push(currentChunk);

  return chunks.map(chunk => {
    const combinedText = chunk.join(" ");
    const fontSize = getDynamicFontSize(combinedText, 32);
    const textLines = chunk.map(line => 
      `<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="${Math.round(fontSize * 100)}" b="1"/><a:t>${line.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")}</a:t></a:r></a:p>`
    ).join("");

    // Minimal slide XML with vertical centering (anchor="ctr")
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Title"/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="457200" y="274320"/><a:ext cx="8229600" cy="1143000"/></a:xfrm></p:spPr>
        <p:txBody>
          <a:bodyPr anchor="ctr" wrap="square"/><a:lstStyle/>
          <a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="2400" b="1"/><a:t>${title.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")}</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="3" name="Content"/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="457200" y="1600200"/><a:ext cx="8229600" cy="4521200"/></a:xfrm></p:spPr>
        <p:txBody>
          <a:bodyPr anchor="ctr" wrap="square"/><a:lstStyle/>
          ${textLines}
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`;
    return { xml, rels: null, media: new Map() };
  });
}

// Helper to extract text from a slide XML
export function getSlideText(xmlStr: string): string {
  const matches = xmlStr.match(/<a:t>([^<]+)<\/a:t>/g);
  return matches ? matches.map(m => m.replace(/<\/?a:t>/g, "")).join(" ") : "";
}

export async function extractPptxText(buffer: ArrayBuffer): Promise<string[]> {
  const zip = await JSZip.loadAsync(buffer);
  const presXmlStr = await zip.file("ppt/presentation.xml")?.async("string");
  if (!presXmlStr) return [];

  const presDoc = parseXml(presXmlStr);
  const sldIds = Array.from(presDoc.getElementsByTagName("p:sldId"));
  
  const presRelsStr = await zip.file("ppt/_rels/presentation.xml.rels")?.async("string");
  const relsDoc = parseXml(presRelsStr || "");
  const relMap = new Map<string, string>();
  Array.from(relsDoc.getElementsByTagName("Relationship")).forEach(rel => {
    relMap.set(rel.getAttribute("Id") || "", rel.getAttribute("Target") || "");
  });

  let slides: string[] = [];
  for (const sldId of sldIds) {
    const relId = sldId.getAttribute("r:id");
    const target = relMap.get(relId || "");
    if (target) {
      const sldPath = `ppt/${target.replace(/^\//, "")}`;
      const sldXml = await zip.file(sldPath)?.async("string");
      if (sldXml) {
        const text = getSlideText(sldXml);
        if (text.trim()) slides.push(text.trim());
      }
    }
  }
  return slides;
}
