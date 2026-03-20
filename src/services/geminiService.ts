import { GoogleGenAI, Type } from "@google/genai";

const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY || "" });

export interface HymnSlot {
  slot: string;
  hymnName: string;
  lyricsSnippet?: string;
  fullLyrics?: string;
  isAlternative?: boolean;
}

export interface LiturgyData {
  title: string;
  firstReading: { title: string; reference: string; text: string };
  secondReading: { title: string; reference: string; text: string };
  gospel: { title: string; reference: string; text: string };
  psalm: { response: string; verses: string[] };
  gospelAntiphon: string;
  prayerOfTheFaithfulResponse: string;
  prayerTitle?: string;
  prayerText?: string;
}

export async function parseHymnList(text: string): Promise<HymnSlot[]> {
  if (!process.env.GEMINI_API_KEY) {
    throw new Error("GEMINI_API_KEY is not configured.");
  }

  const prompt = `
    Parse the following Church Mass hymn list and identify which hymn is assigned to each slot.
    Use these EXACT slot names for the 'slot' property:
    - Entrance Hymn
    - LHM / Lord Have Mercy Hymn
    - Gloria Hymn
    - Gospel Acclamation hymn
    - Offertory hymn
    - Sanctus / Holy Hymn
    - Proclmation Hymn
    - Lamb Of God / LOG Hymn
    - Communion / Communion 1 Hymn
    - Communion / Communion 2 Hymn
    - Recessional Hymn
    
    CRITICAL RULES:
    1. If there is an "Alt." (Alternative) hymn listed, include it as a SEPARATE entry in the array with the same 'slot' name and set 'isAlternative' to true.
    2. Extract the 'fullLyrics' for EVERY hymn found in the text. This is crucial for when a PPTX file is missing.
    3. Return a JSON array of objects with 'slot', 'hymnName', 'fullLyrics', and 'isAlternative' properties.
    4. Include a 'lyricsSnippet' (first few words) as well.
    5. IMPORTANT: For the 'Entrance' slot, if both a primary and an 'Alt.' hymn are present, ensure the primary one (e.g., 'Rise up and Praise Him') is listed first and NOT marked as alternative.

    Hymn List:
    ${text}
  `;

  const response = await ai.models.generateContent({
    model: "gemini-3-flash-preview",
    contents: prompt,
    config: {
      responseMimeType: "application/json",
      responseSchema: {
        type: Type.ARRAY,
        items: {
          type: Type.OBJECT,
          properties: {
            slot: { type: Type.STRING },
            hymnName: { type: Type.STRING },
            lyricsSnippet: { type: Type.STRING },
            fullLyrics: { type: Type.STRING },
            isAlternative: { type: Type.BOOLEAN },
          },
          required: ["slot", "hymnName", "fullLyrics"],
        },
      },
    },
  });

  try {
    return JSON.parse(response.text || "[]");
  } catch (e) {
    console.error("Failed to parse Gemini response", e);
    return [];
  }
}

export async function parseLiturgySheet(base64Pdf: string): Promise<LiturgyData> {
  if (!process.env.GEMINI_API_KEY) {
    throw new Error("GEMINI_API_KEY is not configured.");
  }

  const prompt = `
    Extract the following liturgical elements from this Sunday Liturgy sheet:
    1. Liturgy Sheet Title: The main title of the sheet (e.g., 4th Sunday of Lent). DO NOT add extra info like "- Year A" or "- Year B" if it's not the primary title.
    2. First Reading: Title (e.g., A reading from the first Book of Samuel), Reference (MUST include book name, e.g., 1 Samuel 16:1,5-7,10-13), and the full text.
    3. Second Reading: Title, Reference (MUST include book name), and the full text.
    4. Gospel: Title, Reference (MUST include book name), and the full text.
    5. Responsorial Psalm: The Response text and each Verse text separately (up to 5 verses).
    6. Gospel Antiphon: The text of the Gospel Antiphon.
    7. Prayer of the Faithful Response: The response text for the prayers.

    Return the data in the following JSON format:
    {
      "title": "...",
      "firstReading": { "title": "...", "reference": "...", "text": "..." },
      "secondReading": { "title": "...", "reference": "...", "text": "..." },
      "gospel": { "title": "...", "reference": "...", "text": "..." },
      "psalm": { "response": "...", "verses": ["verse 1...", "verse 2...", ...] },
      "gospelAntiphon": "...",
      "prayerOfTheFaithfulResponse": "...",
      "prayerTitle": "...",
      "prayerText": "..."
    }
  `;

  const response = await ai.models.generateContent({
    model: "gemini-3-flash-preview",
    contents: [
      {
        parts: [
          { text: prompt },
          {
            inlineData: {
              mimeType: "application/pdf",
              data: base64Pdf,
            },
          },
        ],
      },
    ],
    config: {
      responseMimeType: "application/json",
    },
  });

  try {
    return JSON.parse(response.text || "{}");
  } catch (e) {
    console.error("Failed to parse Liturgy response", e);
    throw new Error("Failed to extract liturgy data.");
  }
}
