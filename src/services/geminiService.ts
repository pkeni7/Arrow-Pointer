import { GoogleGenAI, Type } from "@google/genai";

const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

export interface BoundingBox {
  ymin: number;
  xmin: number;
  ymax: number;
  xmax: number;
}

export interface DetectionResult {
  label: number;
  description: string;
  sap_code: string;
  name: string;
  arrow_box: BoundingBox;
  target_box: BoundingBox;
}

export const detectArrowsAndTargets = async (base64Image: string, mimeType: string): Promise<DetectionResult[]> => {
  const prompt = `
    Detect all arrows in the image. For each arrow:
    1. Identify the direction it is pointing.
    2. Find the nearest object or region it is pointing to.
    3. Provide a bounding box for the arrow itself.
    4. Provide a bounding box for the object or region it is pointing to.
    5. Provide a short description of the pointed object.
    6. CRITICAL: Look for any numeric labels, callouts, or serial numbers (e.g., "1", "2", "(3)", "Item 4") written directly in the image next to the arrow or the pointed object. Use this as the "label".
    7. If no numeric labels are found in the image, assign a serial number starting from 1, sorted spatially from top-to-bottom and then left-to-right.
    8. Look for a "Name" and an "SAP Code" (often found in a table or nearby text) associated with the pointed object or its detected label. If not found, use "Unknown".

    Return the results as a JSON array of objects.
    Bounding boxes should be in [ymin, xmin, ymax, xmax] format, normalized from 0 to 1000.
  `;

  const response = await ai.models.generateContent({
    model: "gemini-3-flash-preview",
    contents: [
      {
        parts: [
          { text: prompt },
          {
            inlineData: {
              data: base64Image.split(',')[1],
              mimeType: mimeType,
            },
          },
        ],
      },
    ],
    config: {
      responseMimeType: "application/json",
      responseSchema: {
        type: Type.ARRAY,
        items: {
          type: Type.OBJECT,
          properties: {
            label: { type: Type.INTEGER },
            description: { type: Type.STRING },
            sap_code: { type: Type.STRING, description: "The SAP Code found in the table or nearby text" },
            name: { type: Type.STRING, description: "The Name found in the table or nearby text" },
            arrow_box: {
              type: Type.OBJECT,
              properties: {
                ymin: { type: Type.NUMBER },
                xmin: { type: Type.NUMBER },
                ymax: { type: Type.NUMBER },
                xmax: { type: Type.NUMBER },
              },
              required: ["ymin", "xmin", "ymax", "xmax"],
            },
            target_box: {
              type: Type.OBJECT,
              properties: {
                ymin: { type: Type.NUMBER },
                xmin: { type: Type.NUMBER },
                ymax: { type: Type.NUMBER },
                xmax: { type: Type.NUMBER },
              },
              required: ["ymin", "xmin", "ymax", "xmax"],
            },
          },
          required: ["label", "description", "sap_code", "name", "arrow_box", "target_box"],
        },
      },
    },
  });

  try {
    const results = JSON.parse(response.text || "[]");
    return results;
  } catch (e) {
    console.error("Failed to parse Gemini response", e);
    return [];
  }
};
