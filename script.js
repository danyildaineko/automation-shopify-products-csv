// ============================================
// CONFIGURATION
// ============================================
const CONFIG = {
  STATIC_VALUES: {
    VENDOR: "AlpadaDoors",
    PRODUCT_CATEGORY: "Hardware > Building Materials > Doors > Home Doors",
    GOOGLE_PRODUCT_CATEGORY: 4634,
    WEIGHT_UNIT: "lb",
    CUSTOM_SIZE_LABEL: "Custom Size" 
  },
  
  VARIANT_DEFAULTS: {
    INVENTORY_POLICY: "continue",
    FULFILLMENT: "manual",
    REQUIRES_SHIPPING: true,
    TAXABLE: true,
    INVENTORY_BASE: -1,
    INVENTORY_VARIANT: 0
  },

  METAFIELDS: {
    MATERIAL: "Steel Glass",
    COLOR: "Black" 
  },
  
  PRICE_MARKUP: 1.20, // +20% for compare price
  WEIGHT_GRAMS_TO_LBS: 453.592
};

const COLUMNS = {
  HANDLE: 1,
  TITLE: 2,
  BODY_HTML: 3,
  VENDOR: 4,
  PRODUCT_CATEGORY: 5,
  TYPE: 6,
  TAGS: 7,
  PUBLISHED: 8,
  OPTION1_NAME: 9,
  OPTION1_VALUE: 10,
  VARIANT_SKU: 18,
  VARIANT_GRAMS: 19,
  VARIANT_INVENTORY_TRACKER: 20,
  VARIANT_INVENTORY_QTY: 21,
  VARIANT_INVENTORY_POLICY: 22,
  VARIANT_FULFILLMENT: 23,
  VARIANT_PRICE: 24,
  VARIANT_COMPARE_PRICE: 25,
  VARIANT_REQUIRES_SHIPPING: 26,
  VARIANT_TAXABLE: 27,
  IMAGE_SRC: 29,
  IMAGE_POSITION: 30,
  IMAGE_ALT: 31,
  GIFT_CARD: 32,
  GOOGLE_PRODUCT_CATEGORY: 35,
  COLLECTION: 46,
  DOOR_STYLE: 47,
  DOOR_TYPE: 48,
  NUMBER_OF_LITES: 49,
  COLOR: 50,
  DOOR_GLASS_FINISH: 51,
  DOOR_MATERIAL: 52,
  DOOR_SURFACE_FINISH: 53,
  STYLE: 54,
  VARIANT_IMAGE: 59,
  VARIANT_WEIGHT_UNIT: 60,
  STATUS: 63
};

const SHOPIFY_HEADERS = [
  "Handle", "Title", "Body (HTML)", "Vendor", "Product Category", "Type", "Tags", "Published",
  "Option1 Name", "Option1 Value", "Option1 Linked To", "Option2 Name", "Option2 Value",
  "Option2 Linked To", "Option3 Name", "Option3 Value", "Option3 Linked To", "Variant SKU",
  "Variant Grams", "Variant Inventory Tracker", "Variant Inventory Qty", "Variant Inventory Policy",
  "Variant Fulfillment Service", "Variant Price", "Variant Compare At Price", "Variant Requires Shipping",
  "Variant Taxable", "Variant Barcode", "Image Src", "Image Position", "Image Alt Text", "Gift Card",
  "SEO Title", "SEO Description", "Google Shopping / Google Product Category", "Google Shopping / Gender",
  "Google Shopping / Age Group", "Google Shopping / MPN", "Google Shopping / Condition",
  "Google Shopping / Custom Product", "Google Shopping / Custom Label 0", "Google Shopping / Custom Label 1",
  "Google Shopping / Custom Label 2", "Google Shopping / Custom Label 3", "Google Shopping / Custom Label 4",
  "Collection (product.metafields.custom.collection)", "Door Style (product.metafields.custom.door_style)",
  "Door Type (product.metafields.custom.door_type)", "Number of Lites (product.metafields.custom.number_of_lite)",
  "Color (product.metafields.shopify.color-pattern)", "Door glass finish (product.metafields.shopify.door-glass-finish)",
  "Door material (product.metafields.shopify.door-material)", "Door surface finish (product.metafields.shopify.door-surface-finish)",
  "Style (product.metafields.shopify.style)", "Complementary products (product.metafields.shopify--discovery--product_recommendation.complementary_products)",
  "Related products (product.metafields.shopify--discovery--product_recommendation.related_products)",
  "Related products settings (product.metafields.shopify--discovery--product_recommendation.related_products_display)",
  "Search product boosts (product.metafields.shopify--discovery--product_search_boost.queries)",
  "Variant Image", "Variant Weight Unit", "Variant Tax Code", "Cost per item", "Status"
];

// ============================================
// SUITE MANAGEMENT
// ============================================
function readSuites() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Suites');
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();
  
  return data
    .filter(row => row[1] === true)
    .map(row => ({
      id: row[0],
      name: row[2],
      destSheet: row[3],
      sourceFileId: row[4],
      sourceSheet: row[5],
      startRow: row[6],
      endRow: row[7],
      handleSuffix: row[8],
      sectionSize: row[9],
      skuSheet: row[10] || null
    }));
}

function runSuite() {
  const start = new Date();
  const suites = readSuites();
  
  if (suites.length === 0) {
    Logger.log('⚠️ No enabled suites found');
    return;
  }
  
  Logger.log(`🚀 Processing ${suites.length} suite(s)`);
  
  suites.forEach(suite => {
    try {
      Logger.log(`\n▶ ${suite.name}`);
      copyBriefs(suite);
      Logger.log(`✅ ${suite.name} completed`);
    } catch (e) {
      Logger.log(`❌ ${suite.name} failed: ${e.message}`);
    }
  });
  
  Logger.log(`\n✅ All completed in ${(new Date() - start) / 1000}s`);
}

// ============================================
// SIZE & PRICE & WEIGHTS
// ============================================
function readSizesFromSource(suite) {
  try {
    const sourceSS = SpreadsheetApp.openById(suite.sourceFileId);
    const sizesSheet = sourceSS.getSheetByName('Variants');
    
    if (!sizesSheet) {
      Logger.log('⚠️ Variants sheet not found');
      return null;
    }
    
    const lastRow = sizesSheet.getLastRow();
    if (lastRow < 3) return null;

     // Read from B3 onwards (columns B-H)
    const allData = sizesSheet.getRange(3, 2, lastRow - 2, 7).getValues();

    return {
      columns: {
        size_exterior: 0,           // Column B
        price_exterior: 1,          // Column C
        weights_exterior: 2,        // Column D
        size_exterior_double: 4,    // Column F
        price_exterior_double: 5,   // Column G
        weights_exterior_double: 6  // Column H
      },
      data: allData
    };
  } catch (e) {
      Logger.log(`⚠️ Error reading sizes: ${e.message}`);
      return null;
    }
  }

// ============================================
// COPY SKU FROM OLD TABS
// ============================================
function readSKUsFromSource(suite) {
  Logger.log(`📦 SKU sheet config: "${suite.skuSheet}"`);

  if (!suite.skuSheet) {
    Logger.log('📦 No SKU sheet configured - skipping');
    return {};
  }

  try {
    // Read from active spreadsheet (same file), not source file
    const activeSS = SpreadsheetApp.getActiveSpreadsheet();
    const skuSheet = activeSS.getSheetByName(suite.skuSheet);
    
    if (!skuSheet) {
      Logger.log(`⚠️ SKU sheet "${suite.skuSheet}" not found`);
      return {};
    }
    
    const lastRow = skuSheet.getLastRow();
    Logger.log(`📦 SKU sheet has ${lastRow} rows`);
    
    if (lastRow < 2) return {};
    
    const data = skuSheet.getRange(2, 1, lastRow - 1, 18).getValues();

    // Debug: check first row
    if (data.length > 0) {
      Logger.log(`📦 DEBUG First row column J (index 9): "${data[0][9]}"`);
      Logger.log(`📦 DEBUG First row column R (index 17): "${data[0][17]}"`);
      Logger.log(`📦 DEBUG First row has ${data[0].length} columns`);
    }

    const skuMap = {};
    let mapped = 0;

    data.forEach(row => {
      const size = row[9];  // Column J
      const sku = row[17];  // Column R

      if (size && sku) {
        skuMap[size.toString().trim()] = sku.toString().trim();
        mapped++;
        if (mapped <= 3) {
          Logger.log(`  Sample: ${size} -> ${sku}`);
        }
      }
    });
    
    Logger.log(`📦 Loaded ${Object.keys(skuMap).length} SKU mappings from "${suite.skuSheet}"`);
    return skuMap;
    
  } catch (e) {
    Logger.log(`⚠️ Error reading SKUs: ${e.message}`);
    return {};
  }
}

// ============================================
// IMAGE HANDLING
// ============================================
function parseImageUrl(url) {
  if (!url || typeof url !== 'string' || !url.startsWith('http')) {
    return null;
  }
  
  const parts = url.split('/');
  const filename = parts[parts.length - 1];
  const nameWithoutExt = filename.replace(/\.(webp|jpg|jpeg|png|gif)$/i, '');
  
  let type = '';
  let handle = nameWithoutExt;
  
  if (nameWithoutExt.endsWith('-open')) {
    type = 'open';
    handle = nameWithoutExt.replace(/-open$/, '');
  } else if (nameWithoutExt.endsWith('-closed')) {
    type = 'closed';
    handle = nameWithoutExt.replace(/-closed$/, '');
  }
  
  return { url: url.trim(), handle: handle.toLowerCase(), type };
}

function readImageLinks(sourceFileId) {
  try {
    const sourceSS = SpreadsheetApp.openById(sourceFileId);
    const imagesSheet = sourceSS.getSheetByName('Images Links');
    
    if (!imagesSheet) {
      Logger.log('⚠️ Images Links sheet not found');
      return [];
    }
    
    const lastRow = imagesSheet.getLastRow();
    if (lastRow < 2) return [];
    
    const urls = imagesSheet.getRange(2, 1, lastRow - 1, 1).getValues()
      .map(row => row[0])
      .filter(url => url && url.toString().trim());
    
    return urls
      .map(url => parseImageUrl(url))
      .filter(img => img !== null);
    
  } catch (e) {
    Logger.log(`⚠️ Error reading images: ${e.message}`);
    return [];
  }
}

function buildImageHash(imageLinks) {
  const hash = {};
  
  imageLinks.forEach(img => {
    if (!hash[img.handle]) {
      hash[img.handle] = { open: null, closed: null, others: [] };
    }
    
    if (img.type === 'open') {
      hash[img.handle].open = img.url;
    } else if (img.type === 'closed') {
      hash[img.handle].closed = img.url;
    } else {
      hash[img.handle].others.push(img.url);
    }
  });
  
  return hash;
}

function getProductImages(imageHash, handle) {
  // Try exact match first (O(1))
  if (imageHash[handle]) {
    return sortImages(imageHash[handle]);
  }
  
  // Fuzzy match (only runs if no exact match)
  const matches = { open: null, closed: null, others: [] };
  
  for (const imgHandle in imageHash) {
    if (handle.includes(imgHandle) || imgHandle.includes(handle)) {
      const imgs = imageHash[imgHandle];
      if (imgs.open && !matches.open) matches.open = imgs.open;
      if (imgs.closed && !matches.closed) matches.closed = imgs.closed;
      matches.others.push(...imgs.others);
    }
  }
  
  return sortImages(matches);
}

function sortImages(matches) {
  const images = [];
  if (matches.open) images.push(matches.open);
  if (matches.closed) images.push(matches.closed);
  images.push(...matches.others);
  return images;
}

// ============================================
// HANDLE GENERATION
// ============================================
function generateHandle(title, collection, suffix) {
  const cleanTitle = title.toString().toLowerCase()
    .replace(/[^a-z0-9\s]/g, '')
    .replace(/\s+/g, '-');
  
  const cleanCollection = collection.toString().toLowerCase()
    .replace(/[^a-z0-9\s]/g, '')
    .replace(/\s+/g, '-');
  
  return (cleanTitle + '-' + cleanCollection + suffix)
    .replace(/--+/g, '-')
    .replace(/^-|-$/g, '');
}

// Wrapping description from brief in tag <p>
function wrapDescription(text) {
  if (!text) return '';
  
  const str = text.toString().trim();
  
  // Check if already has <p> tag
  if (str.includes('<p>') || str.includes('<p ')) {
    return str;
  }
  
  // Wrap in <p> tags
  return `<p>${str}</p>`;
}

// ============================================
// PRODUCT DATA MAPPING
// ============================================
function fillProductData(outputRow, sourceRow, handle) {
  outputRow[COLUMNS.HANDLE - 1] = handle;
  outputRow[COLUMNS.TITLE - 1] = sourceRow[0];
  outputRow[COLUMNS.BODY_HTML - 1] = wrapDescription(sourceRow[1]);
  outputRow[COLUMNS.TYPE - 1] = sourceRow[2];
  outputRow[COLUMNS.TAGS - 1] = sourceRow[7];
  outputRow[COLUMNS.COLLECTION - 1] = sourceRow[3];
  outputRow[COLUMNS.DOOR_TYPE - 1] = sourceRow[4];
  outputRow[COLUMNS.DOOR_STYLE - 1] = sourceRow[5];
  outputRow[COLUMNS.COLOR - 1] = CONFIG.METAFIELDS.COLOR;
  outputRow[COLUMNS.DOOR_MATERIAL - 1] = CONFIG.METAFIELDS.MATERIAL;
  outputRow[COLUMNS.VENDOR - 1] = CONFIG.STATIC_VALUES.VENDOR;
  outputRow[COLUMNS.PRODUCT_CATEGORY - 1] = CONFIG.STATIC_VALUES.PRODUCT_CATEGORY;
  outputRow[COLUMNS.PUBLISHED - 1] = true;
  outputRow[COLUMNS.OPTION1_NAME - 1] = "Size";
  outputRow[COLUMNS.GIFT_CARD - 1] = false;
  outputRow[COLUMNS.GOOGLE_PRODUCT_CATEGORY - 1] = CONFIG.STATIC_VALUES.GOOGLE_PRODUCT_CATEGORY;
  outputRow[COLUMNS.VARIANT_INVENTORY_QTY - 1] = CONFIG.VARIANT_DEFAULTS.INVENTORY_BASE;
  outputRow[COLUMNS.STATUS - 1] = "active";
}

// ============================================
// VARIANT DATA FILLING
// ============================================
function fillVariantData(outputRow, isFirst, isLast, sizeData) {
  outputRow[COLUMNS.VARIANT_INVENTORY_POLICY - 1] = CONFIG.VARIANT_DEFAULTS.INVENTORY_POLICY;
  outputRow[COLUMNS.VARIANT_FULFILLMENT - 1] = CONFIG.VARIANT_DEFAULTS.FULFILLMENT;
  outputRow[COLUMNS.VARIANT_REQUIRES_SHIPPING - 1] = CONFIG.VARIANT_DEFAULTS.REQUIRES_SHIPPING;
  outputRow[COLUMNS.VARIANT_TAXABLE - 1] = CONFIG.VARIANT_DEFAULTS.TAXABLE;
  outputRow[COLUMNS.VARIANT_WEIGHT_UNIT - 1] = CONFIG.STATIC_VALUES.WEIGHT_UNIT;

  if (isFirst) {
    outputRow[COLUMNS.VARIANT_INVENTORY_QTY - 1] = CONFIG.VARIANT_DEFAULTS.INVENTORY_BASE;
  } else {
    outputRow[COLUMNS.VARIANT_INVENTORY_QTY - 1] = CONFIG.VARIANT_DEFAULTS.INVENTORY_VARIANT;
  }
  
  if (isLast && sizeData?.isCustomSize) {
    outputRow[COLUMNS.OPTION1_VALUE - 1] = CONFIG.STATIC_VALUES.CUSTOM_SIZE_LABEL;
    if (sizeData.price) {
      const price = sizeData.price;
      outputRow[COLUMNS.VARIANT_PRICE - 1] = price;
      outputRow[COLUMNS.VARIANT_COMPARE_PRICE - 1] = Math.round(price * CONFIG.PRICE_MARKUP * 100) / 100;
    }
    if (sizeData.weight) {
      outputRow[COLUMNS.VARIANT_GRAMS - 1] = Math.round(sizeData.weight * CONFIG.WEIGHT_GRAMS_TO_LBS);
    }
  } else {
    if (sizeData?.size) {
      outputRow[COLUMNS.OPTION1_VALUE - 1] = sizeData.size;
    }
    if (sizeData?.price) {
      const price = sizeData.price;
      outputRow[COLUMNS.VARIANT_PRICE - 1] = price;
      outputRow[COLUMNS.VARIANT_COMPARE_PRICE - 1] = Math.round(price * CONFIG.PRICE_MARKUP * 100) / 100;
    }
    if (sizeData?.weight) {
      outputRow[COLUMNS.VARIANT_GRAMS - 1] = Math.round(sizeData.weight * CONFIG.WEIGHT_GRAMS_TO_LBS);
    }
  }
}

// ============================================
// MAIN COPY FUNCTION
// ============================================
function copyBriefs(suite) {
  const destSS = SpreadsheetApp.getActiveSpreadsheet();
  let dest = destSS.getSheetByName(suite.destSheet);
  
  if (!dest) {
    dest = destSS.insertSheet(suite.destSheet);
  }
  
  dest.getRange(1, 1, dest.getMaxRows(), 63).clearContent();
  dest.getRange(1, 1, 1, SHOPIFY_HEADERS.length).setValues([SHOPIFY_HEADERS]);
  
  const sourceSS = SpreadsheetApp.openById(suite.sourceFileId);
  const source = sourceSS.getSheetByName(suite.sourceSheet);
  const sizesConfig = readSizesFromSource(suite);
  const imageLinks = readImageLinks(suite.sourceFileId);
  const skuMap = readSKUsFromSource(suite);
  
  const numProducts = suite.endRow - suite.startRow + 1;
  const sourceData = source.getRange(suite.startRow, 2, numProducts, 8).getValues();
  
  const totalRows = numProducts * suite.sectionSize;
  const outputData = Array(totalRows).fill(null).map(() => Array(63).fill(''));
  
  Logger.log(`📷 Parsed ${imageLinks.length} image URLs`);
  
  // Build image hash once
  const imageHash = buildImageHash(imageLinks);
  
  sourceData.forEach((row, i) => {
    const baseRow = i * suite.sectionSize;
    const handle = generateHandle(row[0], row[3], suite.handleSuffix);
    const productImages = getProductImages(imageHash, handle);
    
    fillProductData(outputData[baseRow], row, handle);
    
    const isDouble = row[4].toString().toLowerCase().includes('double');
    const colPrefix = isDouble ? '_double' : '';
    
    for (let j = 0; j < suite.sectionSize; j++) {
      const currentRow = baseRow + j;
      const isLast = (j === suite.sectionSize - 1);
      const isFirst = (j === 0);
      
      outputData[currentRow][COLUMNS.HANDLE - 1] = handle;
      
      let sizeData = null;

      if (sizesConfig) {
        if (isLast) {
          const lastIdx = sizesConfig.data.length - 1;
          if (lastIdx >= 0) {
            sizeData = { 
              isCustomSize: true,
              price: sizesConfig.data[lastIdx][sizesConfig.columns[`price_exterior${colPrefix}`]],
              weight: sizesConfig.data[lastIdx][sizesConfig.columns[`weights_exterior${colPrefix}`]]
            };
          }
        } else if (j < sizesConfig.data.length) {
          const dataRow = sizesConfig.data[j];
          sizeData = {
            size: dataRow[sizesConfig.columns[`size_exterior${colPrefix}`]],
            price: dataRow[sizesConfig.columns[`price_exterior${colPrefix}`]],
            weight: dataRow[sizesConfig.columns[`weights_exterior${colPrefix}`]]
          };
        }
      }
      
      fillVariantData(outputData[currentRow], isFirst, isLast, sizeData);
      
      // SKU mapping
      if (sizeData?.size && skuMap[sizeData.size]) {
        outputData[currentRow][COLUMNS.VARIANT_SKU - 1] = skuMap[sizeData.size];
      }

      // Variant images mapping
      if (productImages.length > 0) {
        outputData[currentRow][COLUMNS.VARIANT_IMAGE - 1] = productImages[0];
      }
    }
    
    productImages.forEach((imageUrl, imgIdx) => {
      const imgRow = baseRow + imgIdx;
      if (imgRow < baseRow + suite.sectionSize) {
        outputData[imgRow][COLUMNS.IMAGE_SRC - 1] = imageUrl;
        outputData[imgRow][COLUMNS.IMAGE_POSITION - 1] = imgIdx + 1;
        outputData[imgRow][COLUMNS.IMAGE_ALT - 1] = row[0];
      }
    });
  });
  
  dest.getRange(2, 1, totalRows, 63).setValues(outputData);
  Logger.log(`✓ ${numProducts} products × ${suite.sectionSize} variants = ${totalRows} rows`);
}