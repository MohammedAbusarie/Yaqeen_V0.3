/**
 * OCR processing module for extracting student IDs from attendance sheet images.
 * Uses Tesseract.js for OCR and OpenCV.js for table detection.
 */

/**
 * @typedef {Object} OcrResult
 * @property {string} id - Extracted ID (may be uncertain)
 * @property {number} confidence - Confidence score (0-100)
 * @property {string} rawText - Raw OCR text
 * @property {string} imageName - Source image filename
 */

/**
 * @typedef {Object} OcrProcessingResult
 * @property {OcrResult[]} confident - IDs with high confidence (>= 80)
 * @property {OcrResult[]} uncertain - IDs needing review (< 80)
 * @property {string[]} errors - Processing errors
 */

let openCvReady = false;

/**
 * Callback for when OpenCV.js is loaded
 */
window.onOpenCvReady = function() {
  openCvReady = true;
  console.log("OpenCV.js loaded successfully");
};

/**
 * Preprocess image for better OCR accuracy
 * @param {HTMLCanvasElement} canvas - Canvas with image
 * @returns {HTMLCanvasElement} - Preprocessed canvas
 */
function preprocessImage(canvas) {
  const ctx = canvas.getContext("2d");
  const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
  const data = imageData.data;

  // Convert to grayscale and apply threshold
  for (let i = 0; i < data.length; i += 4) {
    const gray = Math.round(0.299 * data[i] + 0.587 * data[i + 1] + 0.114 * data[i + 2]);
    const threshold = gray > 128 ? 255 : 0;
    data[i] = threshold;     // R
    data[i + 1] = threshold; // G
    data[i + 2] = threshold; // B
    // data[i + 3] stays as alpha
  }

  ctx.putImageData(imageData, 0, 0);
  return canvas;
}

/**
 * Detect table structure using OpenCV (if available) or fallback to simple grid detection
 * @param {HTMLCanvasElement} canvas - Canvas with image
 * @returns {Array<{x: number, y: number, width: number, height: number}>} - Array of cell bounding boxes
 */
function detectTableCellsSimple(canvas) {
  // Simple fallback: assume ID column is on the left
  const cellWidth = Math.floor(canvas.width * 0.25); // Assume ID column is ~25% of width
  const cellHeight = Math.floor(canvas.height / 30); // Assume ~30 rows
  const cells = [];
  
  for (let y = Math.floor(canvas.height * 0.1); y < canvas.height * 0.9; y += cellHeight) {
    cells.push({
      x: Math.floor(canvas.width * 0.05),
      y: y,
      width: cellWidth,
      height: cellHeight
    });
  }
  return cells;
}

function detectTableCells(canvas) {
  // If OpenCV is not ready, use a simple fallback: assume ID column is on the left
  // This is a simplified approach - in production, you'd want proper table detection
  if (!openCvReady || typeof cv === 'undefined') {
    return detectTableCellsSimple(canvas);
  }

  try {
    // Use OpenCV for proper table detection
    const src = cv.imread(canvas);
    const gray = new cv.Mat();
    cv.cvtColor(src, gray, cv.COLOR_RGBA2GRAY);

    // Apply threshold
    const binary = new cv.Mat();
    cv.threshold(gray, binary, 127, 255, cv.THRESH_BINARY_INV);

    // Detect horizontal and vertical lines
    const horizontal = new cv.Mat();
    const vertical = new cv.Mat();
    const horizontalKernel = cv.getStructuringElement(cv.MORPH_RECT, new cv.Size(40, 1));
    const verticalKernel = cv.getStructuringElement(cv.MORPH_RECT, new cv.Size(1, 40));
    
    cv.morphologyEx(binary, horizontal, cv.MORPH_OPEN, horizontalKernel);
    cv.morphologyEx(binary, vertical, cv.MORPH_OPEN, verticalKernel);

    // Find intersections (simplified - in production, use HoughLines)
    // For now, return simplified cell regions
    const cells = [];
    const cellWidth = Math.floor(canvas.width * 0.25);
    const cellHeight = Math.floor(canvas.height / 30);
    
    for (let y = Math.floor(canvas.height * 0.1); y < canvas.height * 0.9; y += cellHeight) {
      cells.push({
        x: Math.floor(canvas.width * 0.05),
        y: y,
        width: cellWidth,
        height: cellHeight
      });
    }

    src.delete();
    gray.delete();
    binary.delete();
    horizontal.delete();
    vertical.delete();
    horizontalKernel.delete();
    verticalKernel.delete();

    return cells;
  } catch (error) {
    console.warn("OpenCV table detection failed, using fallback:", error);
    // Fallback to simple grid
    return detectTableCellsSimple(canvas);
  }
}

/**
 * Crop a region from canvas
 * @param {HTMLCanvasElement} sourceCanvas - Source canvas
 * @param {number} x - X coordinate
 * @param {number} y - Y coordinate
 * @param {number} width - Width
 * @param {number} height - Height
 * @returns {HTMLCanvasElement} - Cropped canvas
 */
function cropCanvas(sourceCanvas, x, y, width, height) {
  const cropped = document.createElement("canvas");
  cropped.width = width;
  cropped.height = height;
  const ctx = cropped.getContext("2d");
  ctx.drawImage(sourceCanvas, x, y, width, height, 0, 0, width, height);
  return cropped;
}

/**
 * Extract student IDs from an image using OCR
 * @param {File} imageFile - Image file to process
 * @param {(progress: number, message: string) => void} onProgress - Progress callback
 * @returns {Promise<OcrProcessingResult>}
 */
export async function processImageWithOcr(imageFile, onProgress) {
  if (!window.Tesseract) {
    throw new Error("Tesseract.js is not loaded. Please check the CDN script.");
  }

  const results = {
    confident: [],
    uncertain: [],
    errors: []
  };

  try {
    onProgress(0, `Loading image: ${imageFile.name}`);

    // Load image to canvas
    const img = new Image();
    const imgUrl = URL.createObjectURL(imageFile);
    
    await new Promise((resolve, reject) => {
      img.onload = resolve;
      img.onerror = reject;
      img.src = imgUrl;
    });

    const canvas = document.createElement("canvas");
    canvas.width = img.width;
    canvas.height = img.height;
    const ctx = canvas.getContext("2d");
    ctx.drawImage(img, 0, 0);
    URL.revokeObjectURL(imgUrl);

    onProgress(20, "Preprocessing image...");
    const preprocessed = preprocessImage(canvas);

    onProgress(40, "Detecting table structure...");
    const cells = detectTableCells(preprocessed);

    onProgress(50, `Processing ${cells.length} cells with OCR...`);

    // Initialize Tesseract worker
    const worker = await window.Tesseract.createWorker('eng');
    await worker.setParameters({
      tessedit_char_whitelist: '0123456789',
    });

    let processed = 0;
    for (const cell of cells) {
      const cellCanvas = cropCanvas(preprocessed, cell.x, cell.y, cell.width, cell.height);
      
      try {
        const { data } = await worker.recognize(cellCanvas);
        const text = data.text.trim();
        
        // Extract 9-digit IDs using regex
        const idMatches = text.match(/\d{9}/g);
        
        if (idMatches && idMatches.length > 0) {
          // Use the first 9-digit match
          const id = idMatches[0];
          const confidence = data.confidence || 0;

          const result = {
            id: id,
            confidence: confidence,
            rawText: text,
            imageName: imageFile.name
          };

          if (confidence >= 80) {
            results.confident.push(result);
          } else {
            results.uncertain.push(result);
          }
        } else if (text.match(/\d{6,}/)) {
          // Found some digits but not exactly 9 - mark as uncertain
          const partialId = text.match(/\d+/)?.[0] || '';
          if (partialId.length >= 6) {
            results.uncertain.push({
              id: partialId.padEnd(9, '?'),
              confidence: 50,
              rawText: text,
              imageName: imageFile.name
            });
          }
        }
      } catch (error) {
        console.warn(`Error processing cell at (${cell.x}, ${cell.y}):`, error);
        results.errors.push(`Cell processing error: ${error.message}`);
      }

      processed++;
      const progress = 50 + (processed / cells.length) * 40;
      onProgress(progress, `Processed ${processed}/${cells.length} cells...`);
    }

    await worker.terminate();
    onProgress(100, "OCR processing complete!");

  } catch (error) {
    results.errors.push(`Image processing error: ${error.message}`);
    console.error("OCR processing error:", error);
  }

  return results;
}

/**
 * Process multiple images
 * @param {File[]} imageFiles - Array of image files
 * @param {(progress: number, message: string) => void} onProgress - Progress callback
 * @returns {Promise<OcrProcessingResult>}
 */
export async function processMultipleImages(imageFiles, onProgress) {
  const allResults = {
    confident: [],
    uncertain: [],
    errors: []
  };

  for (let i = 0; i < imageFiles.length; i++) {
    const file = imageFiles[i];
    const fileProgress = (fileProgressValue, message) => {
      const overallProgress = (i / imageFiles.length) * 100 + (fileProgressValue / imageFiles.length);
      onProgress(overallProgress, `[${i + 1}/${imageFiles.length}] ${message}`);
    };

    try {
      const result = await processImageWithOcr(file, fileProgress);
      allResults.confident.push(...result.confident);
      allResults.uncertain.push(...result.uncertain);
      allResults.errors.push(...result.errors);
    } catch (error) {
      allResults.errors.push(`Failed to process ${file.name}: ${error.message}`);
    }
  }

  // Remove duplicate IDs (keep the one with highest confidence)
  const idMap = new Map();
  
  [...allResults.confident, ...allResults.uncertain].forEach(result => {
    const existing = idMap.get(result.id);
    if (!existing || result.confidence > existing.confidence) {
      idMap.set(result.id, result);
    }
  });

  // Re-categorize based on confidence
  allResults.confident = [];
  allResults.uncertain = [];
  
  idMap.forEach((result) => {
    if (result.confidence >= 80) {
      allResults.confident.push(result);
    } else {
      allResults.uncertain.push(result);
    }
  });

  return allResults;
}

/**
 * Generate text file content from OCR results
 * @param {OcrResult[]} approvedIds - Approved IDs (confident + manually approved uncertain)
 * @returns {string} - Text file content
 */
export function generateTextFile(approvedIds) {
  // Extract just the IDs, one per line
  return approvedIds.map(result => result.id).join('\n');
}

