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
    COLOR: "matte-black-ral-7021; grey-blue-ral-5008; sepia-brown-ral-8014; olive-gray-ral-7002; pebble-gray-ral-7032; window-gray-ral-7040; pure-white-ral-9010",
    DOOR_GLASS_FINISH: "ribbed-veritcal; ribbed-horizontal; grey-glass; clear; frost; opaque",
    DOOR_MATERIAL: "steel; glass",
    DOOR_SUITABLE_LOCATION: "front",
    DOOR_SURFACE_FINISH: ""
  },

  REPEATABLE_IMAGES: {
    EXTERIOR_ARCHED: [
      'https://cdn.shopify.com/s/files/1/0633/0558/0744/files/arched-interior-french-double-doors-3.webp?v=1762696444',
      'https://cdn.shopify.com/s/files/1/0633/0558/0744/files/arched-interior-french-double-doors-4.webp?v=1762696445'
    ]
  },

  CONSTANT_SIZES: {
    SUITE_2: {
      SINGLE: ['24"W x 80"H Door', '28"W x 80"H Door', '30"W x 80"H Door', '32"W x 80"H Door', '32"W x 81"H Door', '32"W x 96"H Door', '34"W x 96"H Door', '36"W x 80"H Door', '36"W x 96"H Door', '38"W x 96"H Door', '40"W x 96"H Door', '42"W x 96"H Door'],
      DOUBLE: ['48"W x 80"H Doors', '56"W x 80"H Doors', '60"W x 80"H Doors', '64"W x 80"H Doors', '60"W x 96"H Doors', '64"W x 96"H Doors', '68"W x 96"H Doors', '72"W x 80"H Doors', '72"W x 96"H Doors', '76"W x 96"H Doors', '80"W x 96"H Doors', '84"W x 96"H Doors']
    }
  },

  APPROVED_SIZES: {
    SINGLE: ['Custom Size', '32"W x 81"H', '32"W x 84"H', '32"W x 96"H', '36"W x 80"H', '36"W x 81"H', '36"W x 84"H', '36"W x 96"H', '42"W x 102"H'],
    DOUBLE: ['Custom Size', '60"W x 96"H', '61"W x 81"H', '61"W x 96"H', '64"W x 80"H', '64"W x 96"H', '72"W x 80"H', '72"W x 84"H', '72"W x 96"H']
  },

  PRICE_MARKUP: 1.20, // +20% for markup price
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
  DOOR_APPLICATION: 51,
  DOOR_GLASS_FINISH: 52,
  DOOR_MATERIAL: 53,
  DOOR_SUITABLE_LOCATION: 54,
  DOOR_SURFACE_FINISH: 55,
  STYLE: 56,
  SHOW_VARIANT: 57,
  VARIANT_IMAGE: 62,
  VARIANT_WEIGHT_UNIT: 63,
  STATUS: 66
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
  "Color (product.metafields.shopify.color-pattern)", "Door application (product.metafields.shopify.door-application)",
  "Door glass finish (product.metafields.shopify.door-glass-finish)", "Door material (product.metafields.shopify.door-material)",
  "Door suitable location (product.metafields.shopify.door-suitable-location)", "Door surface finish (product.metafields.shopify.door-surface-finish)",
  "Style (product.metafields.shopify.style)", "Show Variant (variant.metafields.custom.show_variant)",
  "Complementary products (product.metafields.shopify--discovery--product_recommendation.complementary_products)",
  "Related products (product.metafields.shopify--discovery--product_recommendation.related_products)",
  "Related products settings (product.metafields.shopify--discovery--product_recommendation.related_products_display)",
  "Search product boosts (product.metafields.shopify--discovery--product_search_boost.queries)",
  "Variant Image", "Variant Weight Unit", "Variant Tax Code", "Cost per item", "Status"
];

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
// SUITE MANAGEMENT
// ============================================
function readSuites() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Suites');
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 12).getValues();

  return data
    .filter(row => row[1] === true)
    .map(row => {
      // Debug: log what's in column K
      Logger.log(`🔍 Suite "${row[2]}" - Column K (row[10]): "${row[10]}" (type: ${typeof row[10]}, length: ${row[10]?.length})`);

      return {
        id: row[0],
        name: row[2],
        destSheet: row[3],
        sourceFileId: row[4],
        sourceSheet: row[5],
        startRow: row[6],
        endRow: row[7],
        handleSuffix: row[8],
        sectionSize: row[9],
        skuSheet: row[10] || null,
        published: row[11] === true
      };
    });
}

// ============================================
// SIZE & PRICE & WEIGHTS
// ============================================
function readSizesFromSource(suite) {
  // Check for constant sizes first
  if (suite.id === 5 && CONFIG.CONSTANT_SIZES.SUITE_2) {
    return {
      isConstant: true,
      constantSizes: CONFIG.CONSTANT_SIZES.SUITE_2
    };
  }
  
  // Original logic for reading from sheet
  try {
    const sourceSS = SpreadsheetApp.openById(suite.sourceFileId);
    const sizesSheet = sourceSS.getSheetByName('Variants');
    
    if (!sizesSheet) {
      Logger.log('⚠️ Variants sheet not found');
      return null;
    }
    
    const lastRow = sizesSheet.getLastRow();
    if (lastRow < 3) return null;

    const allData = sizesSheet.getRange(3, 2, lastRow - 2, 7).getValues();

    return {
      isConstant: false,
      columns: {
        size_exterior: 0,
        price_exterior: 1,
        weights_exterior: 2,
        size_exterior_double: 4,
        price_exterior_double: 5,
        weights_exterior_double: 6
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
  if (!suite.skuSheet) return suite.id === 2 ? [] : {};

  try {
    const activeSS = SpreadsheetApp.getActiveSpreadsheet();
    const skuSheet = activeSS.getSheetByName(suite.skuSheet);
    if (!skuSheet) return suite.id === 2 ? [] : {};

    const lastRow = skuSheet.getLastRow();
    if (lastRow < 2) return suite.id === 2 ? [] : {};

    const data = skuSheet.getRange(2, 1, lastRow - 1, 18).getValues();

    if (suite.id === 2) {
      const SOURCE_SECTION_SIZE = 100; // Original size in source
      const skusByProduct = [];
      for (let i = 0; i < data.length; i += SOURCE_SECTION_SIZE) {
        const productSkus = data.slice(i, i + SOURCE_SECTION_SIZE)
          .map(row => row[17])
          .filter(sku => sku)
          .slice(0, suite.sectionSize);
        skusByProduct.push(productSkus);
      }
      return skusByProduct;
    }

    const skuMap = {};
    data.forEach(row => {
      const handle = row[0];
      const sku = row[17];
      if (handle && sku) {
        let cleanHandle = handle.toString().trim().replace(/-(exterior|test|ext|int|interior)$/i, '');
        if (!skuMap[cleanHandle]) skuMap[cleanHandle] = [];
        if (skuMap[cleanHandle].length < suite.sectionSize) {
          skuMap[cleanHandle].push(sku.toString().trim());
        }
      }
    });
    return skuMap;
  } catch (e) {
    return suite.id === 2 ? [] : {};
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
  
  return { url: url.trim(), handle: normalizeHandle(handle), type };
}

function readImageLinks(sourceFileId, suite) {
  try {
    const sourceSS = SpreadsheetApp.openById(sourceFileId);
    const imagesSheet = sourceSS.getSheetByName('Images Links');
    Logger.log(`🖼️ Reading images from file: https://docs.google.com/spreadsheets/d/${sourceFileId} (sheet: ${imagesSheet ? imagesSheet.getName() : 'N/A'})`);
    
    if (!imagesSheet) {
      Logger.log('⚠️ Images Links sheet not found');
      return [];
    }
    
    const lastRow = imagesSheet.getLastRow();
    if (lastRow < 2) return [];
    
    let columnIndex = 1;
    
    if (suite?.id === 2) {
      columnIndex = 2;
    }
    
    const numRows = lastRow - 1;
    const dataRange = imagesSheet.getRange(2, columnIndex, numRows, 1);
    const rawValues = dataRange.getValues();
    const richValues = dataRange.getRichTextValues();
    
    if (suite?.id === 2) {
      const sampleSize = Math.min(numRows, 5);
      if (sampleSize > 0) {
        const sampleRange = imagesSheet.getRange(2, columnIndex, sampleSize, 1);
        const sampleValues = sampleRange.getValues().map(row => row[0]);
        const sampleDisplay = sampleRange.getDisplayValues().map(row => row[0]);
        const sampleFormulas = sampleRange.getFormulas().map(row => row[0]);
        const sampleRich = sampleRange.getRichTextValues().map(row => {
          const rich = row[0];
          return rich ? { text: rich.getText(), url: rich.getLinkUrl() } : null;
        });
      }
    }
    
    let urls = rawValues.map((row, idx) => {
      const value = row[0];
      if (value && value.toString().trim()) {
        return value.toString().trim();
      }
      
      const rich = richValues[idx] && richValues[idx][0];
      if (rich) {
        const linkUrl = rich.getLinkUrl();
        if (linkUrl) {
          return linkUrl.trim();
        }
        const text = rich.getText();
        if (text && text.trim()) {
          return text.trim();
        }
      }
      return null;
    }).filter(url => url && url.toString().trim());
    
    if (suite?.id === 2 && urls.length === 0 && columnIndex !== 1) {
      Logger.log('🖼️ Suite 2 fallback to column 1 (Side Lites column empty)');
      columnIndex = 1;
      const fallbackRange = imagesSheet.getRange(2, columnIndex, lastRow - 1, 1);
      const fallbackRaw = fallbackRange.getValues();
      const fallbackRich = fallbackRange.getRichTextValues();
      urls = fallbackRaw.map((row, idx) => {
        const value = row[0];
        if (value && value.toString().trim()) {
          return value.toString().trim();
        }
        const rich = fallbackRich[idx] && fallbackRich[idx][0];
        if (rich) {
          const linkUrl = rich.getLinkUrl();
          if (linkUrl) return linkUrl.trim();
          const text = rich.getText();
          if (text && text.trim()) return text.trim();
        }
        return null;
      }).filter(url => url && url.toString().trim());
    }
    
    const uniqueUrls = Array.from(new Set(urls.map(url => url.toString().trim())));
    
    Logger.log(`🖼️ Collected ${urls.length} raw / ${uniqueUrls.length} unique image URL(s) from column ${columnIndex}`);
    
    const parsed = uniqueUrls
      .map(url => parseImageUrl(url))
      .filter(img => img !== null);
    
    return parsed;
    
  } catch (e) {
    Logger.log(`⚠️ Error reading images: ${e.message}`);
    return [];
  }
}

function buildImageHash(imageLinks) {
  const hash = {};
  
  imageLinks.forEach(img => {
    const normalizedHandle = normalizeHandle(img.handle);
    
    if (!hash[normalizedHandle]) {
      hash[normalizedHandle] = { open: null, closed: null, others: [] };
    }
    
    if (img.type === 'open') {
      hash[normalizedHandle].open = img.url;
    } else if (img.type === 'closed') {
      hash[normalizedHandle].closed = img.url;
    } else {
      hash[normalizedHandle].others.push(img.url);
    }
  });
  
  return hash;
}

function getProductImages(imageHash, handle) {
  const normalizedHandle = normalizeHandle(handle);
  const productTokens = tokenizeHandle(normalizedHandle);
  // Try exact match first (O(1))
  if (imageHash[normalizedHandle]) {
    return sortImages(imageHash[normalizedHandle]);
  }
  
  // Fuzzy match (only runs if no exact match)
  const matches = { open: null, closed: null, others: [] };
  
  for (const imgHandle in imageHash) {
    if (normalizedHandle.includes(imgHandle) || imgHandle.includes(normalizedHandle)) {
      const imgs = imageHash[imgHandle];
      if (imgs.open && !matches.open) matches.open = imgs.open;
      if (imgs.closed && !matches.closed) matches.closed = imgs.closed;
      matches.others.push(...imgs.others);
    }
  }
  
  if (!matches.open && !matches.closed && matches.others.length === 0 && productTokens.length > 0) {
    for (const imgHandle in imageHash) {
      const imageTokens = tokenizeHandle(imgHandle);
      if (imageTokens.every(token => productTokens.includes(token))) {
        const imgs = imageHash[imgHandle];
        if (imgs.open && !matches.open) matches.open = imgs.open;
        if (imgs.closed && !matches.closed) matches.closed = imgs.closed;
        matches.others.push(...imgs.others);
      }
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

function normalizeHandle(handle) {
  if (!handle) return '';
  
  return handle
    .toString()
    .toLowerCase()
    .replace(/side[\s_-]?lite(s)?/g, (match, plural) => plural ? 'side-lites' : 'side-lite')
    .replace(/--+/g, '-')
    .replace(/^-|-$/g, '');
}

function tokenizeHandle(handle) {
  return normalizeHandle(handle)
    .split('-')
    .filter(Boolean);
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

// Extract Number of Lites from title
// Equivalent to Excel formula: =IFERROR(LEFT(B2, SEARCH("Lite", B2) + 3), "")
function extractNumberOfLites(title) {
  if (!title || typeof title !== 'string') {
    return '';
  }

  const liteIndex = title.indexOf('Lite');
  if (liteIndex === -1) {
    return '';
  }

  // Excel SEARCH returns 1-indexed position, +3 extends past start of "Lite"
  // In JS (0-indexed): substring(0, liteIndex + 4) captures text up to "Lite" + first 3 chars
  // This captures patterns like "3 Lite", "Full Lite", etc.
  return title.substring(0, liteIndex + 4);
}

// ============================================
// PRODUCT DATA MAPPING
// ============================================
function fillProductData(outputRow, sourceRow, handle, published) {
  const type = sourceRow[2].toString().toLowerCase();
  let doorApplication = "internal";
  if (type.includes('exterior')) {
    doorApplication = "external";
  }
  
  outputRow[COLUMNS.HANDLE - 1] = handle;
  outputRow[COLUMNS.TITLE - 1] = sourceRow[0];
  outputRow[COLUMNS.BODY_HTML - 1] = wrapDescription(sourceRow[1]);
  outputRow[COLUMNS.VENDOR - 1] = CONFIG.STATIC_VALUES.VENDOR;
  outputRow[COLUMNS.PRODUCT_CATEGORY - 1] = CONFIG.STATIC_VALUES.PRODUCT_CATEGORY;
  outputRow[COLUMNS.TYPE - 1] = sourceRow[2];
  outputRow[COLUMNS.TAGS - 1] = sourceRow[7];
  outputRow[COLUMNS.PUBLISHED - 1] = published;
  outputRow[COLUMNS.OPTION1_NAME - 1] = "Size";
  outputRow[COLUMNS.VARIANT_INVENTORY_QTY - 1] = CONFIG.VARIANT_DEFAULTS.INVENTORY_BASE;
  outputRow[COLUMNS.GIFT_CARD - 1] = false;
  outputRow[COLUMNS.GOOGLE_PRODUCT_CATEGORY - 1] = CONFIG.STATIC_VALUES.GOOGLE_PRODUCT_CATEGORY;
  outputRow[COLUMNS.COLLECTION - 1] = sourceRow[3];
  outputRow[COLUMNS.DOOR_TYPE - 1] = sourceRow[4];
  outputRow[COLUMNS.DOOR_STYLE - 1] = sourceRow[5];
  outputRow[COLUMNS.NUMBER_OF_LITES - 1] = extractNumberOfLites(sourceRow[0]);
  outputRow[COLUMNS.COLOR - 1] = CONFIG.METAFIELDS.COLOR;
  outputRow[COLUMNS.DOOR_APPLICATION - 1] = doorApplication;
  outputRow[COLUMNS.DOOR_GLASS_FINISH - 1] = CONFIG.METAFIELDS.DOOR_GLASS_FINISH;
  outputRow[COLUMNS.DOOR_MATERIAL - 1] = CONFIG.METAFIELDS.DOOR_MATERIAL;
  outputRow[COLUMNS.DOOR_SUITABLE_LOCATION - 1] = CONFIG.METAFIELDS.DOOR_SUITABLE_LOCATION;
  outputRow[COLUMNS.DOOR_SURFACE_FINISH - 1] = CONFIG.METAFIELDS.DOOR_SURFACE_FINISH;
  outputRow[COLUMNS.STYLE - 1] = formatStyleMetafield(sourceRow[45] || sourceRow[5]); 
  outputRow[COLUMNS.STATUS - 1] = published ? "active" : "draft";

}

function formatStyleMetafield(value) {
  if (!value) return '';
  return value.toString()
    .replace(/([a-z])([A-Z])/g, '$1; $2')
    .split(/\s*[;,]\s*|\n/)
    .map(s => s.trim().toLowerCase().replace(/\s+/g, '-'))
    .filter(s => s)
    .join('; ');
}

function calculateSuite2Price(sizeString) {
  const match = sizeString.match(/(\d+)"W x (\d+)"H/);
  if (!match) return null;
  const width = parseInt(match[1]);
  const height = parseInt(match[2]);
  return ((width * height) * 0.34 + 1000) * 1.5;
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
  
  if (isFirst && sizeData?.isCustomSize) {
    outputRow[COLUMNS.OPTION1_VALUE - 1] = CONFIG.STATIC_VALUES.CUSTOM_SIZE_LABEL;
    if (sizeData.price) outputRow[COLUMNS.VARIANT_PRICE - 1] = sizeData.price;
    if (sizeData.weight) outputRow[COLUMNS.VARIANT_GRAMS - 1] = Math.round(sizeData.weight * CONFIG.WEIGHT_GRAMS_TO_LBS);
  } else {
    if (sizeData?.size) outputRow[COLUMNS.OPTION1_VALUE - 1] = sizeData.size;
    if (sizeData?.price) {
      outputRow[COLUMNS.VARIANT_PRICE - 1] = sizeData.price;
      outputRow[COLUMNS.VARIANT_COMPARE_PRICE - 1] = Math.round(sizeData.price * CONFIG.PRICE_MARKUP * 100) / 100;
    }
    if (sizeData?.weight) outputRow[COLUMNS.VARIANT_GRAMS - 1] = Math.round(sizeData.weight * CONFIG.WEIGHT_GRAMS_TO_LBS);
  }
}


function copyBriefs(suite) {
  const destSS = SpreadsheetApp.getActiveSpreadsheet();
  let dest = destSS.getSheetByName(suite.destSheet);

  if (!dest) dest = destSS.insertSheet(suite.destSheet);

  dest.getRange(1, 1, dest.getMaxRows(), 66).clearContent();

  // Headers - now all suites use standard headers
  dest.getRange(1, 1, 1, SHOPIFY_HEADERS.length).setValues([SHOPIFY_HEADERS]);

  const sourceSS = SpreadsheetApp.openById(suite.sourceFileId);
  const source = sourceSS.getSheetByName(suite.sourceSheet);
  const sizesConfig = readSizesFromSource(suite);
  const imageLinks = readImageLinks(suite.sourceFileId, suite);
  const skuMap = readSKUsFromSource(suite);

  const numProducts = suite.endRow - suite.startRow + 1;
  const sourceData = source.getRange(suite.startRow, 2, numProducts, 46).getValues();
  const totalRows = numProducts * suite.sectionSize;
  const numCols = 66;
  const outputData = Array(totalRows).fill(null).map(() => Array(numCols).fill(''));

  Logger.log(`📷 Parsed ${imageLinks.length} image URLs`);

  const imageHash = suite.id === 2 ? null : buildImageHash(imageLinks);

  sourceData.forEach((row, i) => {
    const baseRow = i * suite.sectionSize;
    const handle = generateHandle(row[0], row[3], suite.handleSuffix);
    const handleWithoutSuffix = generateHandle(row[0], row[3], '');

    // Image logic
    let productImages = [];
    if (suite.id === 2) {
      if (i < imageLinks.length && imageLinks[i]?.url) {
        productImages = [imageLinks[i].url];
      }
    } else {
      productImages = getProductImages(imageHash, handle);
      if (suite.id !== 1 && (row[2].toString().toLowerCase().includes('exterior') ||
          row[5].toString().toLowerCase().includes('arched'))) {
        productImages.push(...CONFIG.REPEATABLE_IMAGES.EXTERIOR_ARCHED);
      }
    }

    // Fill product data for all suites
    fillProductData(outputData[baseRow], row, handle, suite.published);

    // Fill variants
    for (let j = 0; j < suite.sectionSize; j++) {
      const currentRow = baseRow + j;
      const isLast = (j === suite.sectionSize - 1);
      const isFirst = (j === 0);

      outputData[currentRow][0] = handle;

      let sizeData = null;
      if (sizesConfig) {
        if (sizesConfig.isConstant) {
          const isDouble = row[0].toString().toLowerCase().includes('double');
          const sizeList = isDouble ? sizesConfig.constantSizes.DOUBLE : sizesConfig.constantSizes.SINGLE;
          if (j < sizeList.length) {
            sizeData = { size: sizeList[j], price: calculateSuite2Price(sizeList[j]) };
          }
        } else {
          const isDouble = row[4].toString().toLowerCase().includes('double');
          const colPrefix = isDouble ? '_double' : '';
          if (isFirst) {
            const lastIdx = sizesConfig.data.length - 1;
            if (lastIdx >= 0) {
              sizeData = {
                isCustomSize: true,
                price: sizesConfig.data[lastIdx][sizesConfig.columns[`price_exterior${colPrefix}`]],
                weight: sizesConfig.data[lastIdx][sizesConfig.columns[`weights_exterior${colPrefix}`]]
              };
            }
          } else if (j - 1 < sizesConfig.data.length) {
            // normal size (offset by 1 since custom is first)
            const dataRow = sizesConfig.data[j - 1];
            sizeData = {
              size: dataRow[sizesConfig.columns[`size_exterior${colPrefix}`]],
              price: dataRow[sizesConfig.columns[`price_exterior${colPrefix}`]],
              weight: dataRow[sizesConfig.columns[`weights_exterior${colPrefix}`]]
            };
          }
        }
      }
      fillVariantData(outputData[currentRow], isFirst, isLast, sizeData);

      // SKU mapping
      if (suite.id === 2 && Array.isArray(skuMap)) {
        if (skuMap[i]?.[j]) outputData[currentRow][COLUMNS.VARIANT_SKU - 1] = skuMap[i][j];
      } else if (skuMap[handleWithoutSuffix]?.[j]) {
        outputData[currentRow][COLUMNS.VARIANT_SKU - 1] = skuMap[handleWithoutSuffix][j];
      }

      if (productImages.length > 0) {
        outputData[currentRow][COLUMNS.VARIANT_IMAGE - 1] = productImages[0];
      }

      // Show variant metafield logic - skip first variant (baseline)
      if (!isFirst) {
        const size = outputData[currentRow][COLUMNS.OPTION1_VALUE - 1];
        if (size) {
          const isDouble = row[0].toString().toLowerCase().includes('double') ||
                          row[4].toString().toLowerCase().includes('double');
          const approvedList = isDouble ? CONFIG.APPROVED_SIZES.DOUBLE : CONFIG.APPROVED_SIZES.SINGLE;
          const isApproved = approvedList.includes(size);
          outputData[currentRow][COLUMNS.SHOW_VARIANT - 1] = isApproved ? 'true' : 'false';
        }
      }
    }

    // Images
    productImages.forEach((imageUrl, imgIdx) => {
      const imgRow = baseRow + imgIdx;
      if (imgRow < baseRow + suite.sectionSize) {
        outputData[imgRow][COLUMNS.IMAGE_SRC - 1] = imageUrl;
        outputData[imgRow][COLUMNS.IMAGE_POSITION - 1] = imgIdx + 1;
        outputData[imgRow][COLUMNS.IMAGE_ALT - 1] = row[0];
      }
    });
  });

  dest.getRange(2, 1, totalRows, numCols).setValues(outputData);
  Logger.log(`✓ ${numProducts} products × ${suite.sectionSize} variants = ${totalRows} rows`);
}
