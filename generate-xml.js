const { google } = require('googleapis');
const { create } = require('xmlbuilder2');
const fs = require('fs');

// Configuration
const SHEET_ID = process.env.SHEET_ID;
const SHEET_NAMES = ['Товари Ncase', 'Вручну додані', 'Товари Kiborg', 'Товари Viktailor'];
const FILTER_ENABLED = false;

// Column indices (0-based)
const COL_ID = 0;
const COL_STOCK_QTY = 1;
const COL_NAME = 2;
const COL_NAME_UA = 3;
const COL_PRICE = 4;
const COL_CATEGORY_ID = 5;
const COL_PICTURES = 6;
const COL_VENDOR = 7;
const COL_DESCRIPTION = 8;
const COL_DESCRIPTION_UA = 9;
const COL_PARAM = 10;

// Authenticate
async function authenticate() {
  console.log('Starting authentication...');
  const credentials = process.env.GOOGLE_SERVICE_ACCOUNT_KEY
    ? JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_KEY)
    : null;

  if (!credentials || typeof credentials !== 'object') {
    console.error('Invalid or missing GOOGLE_SERVICE_ACCOUNT_KEY in environment');
    throw new Error('Invalid or missing GOOGLE_SERVICE_ACCOUNT_KEY in environment');
  }
  console.log('Credentials parsed successfully');

  const auth = new google.auth.GoogleAuth({
    credentials: credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
  });
  const sheets = google.sheets({ version: 'v4', auth: await auth.getClient() });
  console.log('Authentication completed');
  return sheets;
}

// Fetch data
async function getSheetData(sheets, sheetName) {
  console.log(`Fetching data from sheet: ${sheetName}`);
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: `${sheetName}!A:L`,
  });
  const data = response.data.values || [];
  console.log(`Fetched ${data.length} rows from ${sheetName}`);
  return data;
}

// Main function
async function generateXML() {
  console.log('Starting XML generation process...');
  const sheets = await authenticate();
  let allItems = [];

  // Fetch items
  for (const name of SHEET_NAMES) {
    const data = await getSheetData(sheets, name);
    if (data.length <= 1) continue;
    const items = data.slice(1);
    let filteredItems = items;
    if (FILTER_ENABLED) {
      filteredItems = items.filter(row =>
        row[COL_STOCK_QTY] === true || row[COL_STOCK_QTY] === 'TRUE' || row[COL_STOCK_QTY] === 'true'
      );
      console.log(`Filtered ${filteredItems.length} items from ${name}`);
    } else {
      console.log(`No filtering, using ${items.length} items from ${name}`);
    }
    allItems = allItems.concat(filteredItems);
  }
  console.log(`Total items collected: ${allItems.length}`);

  // Fetch categories
  const categoriesData = await getSheetData(sheets, 'Зв\'язування категорій');
  const bindingCategories = categoriesData.slice(1);
  console.log(`Fetched ${bindingCategories.length} category bindings`);

  // Build XML
  const date = new Date().toISOString();
  console.log(`Generating XML with date: ${date}`);
  const doc = create({ version: '1.0', encoding: 'UTF-8' })
    .ele('yml_catalog', { date });

  const shop = doc.ele('shop');

  // Currencies
  const currencies = shop.ele('currencies');
  currencies.ele('currency', { id: 'UAH', rate: '1' });

  // Categories
  const categories = shop.ele('categories');
  let categoryCount = 0;
  for (const bindingCategory of bindingCategories) {
    const id = bindingCategory[0];
    const name = bindingCategory[1];
    if (id && name) {
      categories.ele('category', { id, rz_id: id }).txt(name);
      categoryCount++;
    }
  }
  console.log(`Added ${categoryCount} categories`);

  // Offers
  const offers = shop.ele('offers');
  let offerCount = 0;
  for (const row of allItems) {
    const id = row[COL_ID];
    const available = row[COL_STOCK_QTY] ? 'true' : 'false';
    const stock_quantity = available === 'true' ? 30 : 0;
    const name = row[COL_NAME] || '';
    const name_ua = row[COL_NAME_UA] || '';
    const price = row[COL_PRICE] || '';
    const category_id = row[COL_CATEGORY_ID] || '';
    const image = row[COL_PICTURES] || '';
    const vendor = row[COL_VENDOR] || '';
    const description = row[COL_DESCRIPTION] || '';
    const description_ua = row[COL_DESCRIPTION_UA] || '';
    const params = row[COL_PARAM] || '';

    const offer = offers.ele('offer', { id, available });

    offer.ele('name').txt(name);
    offer.ele('name_ua').txt(name_ua);
    offer.ele('price').txt(price);
    offer.ele('currencyId').txt('UAH');
    offer.ele('categoryId').txt(category_id);

    if (image) {
      const pics = image.includes(',') ? image.split(',') : [image];
      for (const pic of pics) {
        if (pic.trim()) offer.ele('picture').txt(pic.trim());
      }
    }

    if (vendor) offer.ele('vendor').txt(vendor);

    offer.ele('stock_quantity').txt(stock_quantity.toString());

    offer.ele('description').dat(description.trim());
    offer.ele('description_ua').dat(description_ua.trim());

    const paramValues = params.split('\n');
    for (const param of paramValues) {
      const [nameParam, value] = param.split(' - ');
      if (nameParam && value) {
        offer.ele('param', { name: nameParam.trim() }).txt(value.trim());
      }
    }
    offerCount++;
  }
  console.log(`Added ${offerCount} offers`);

  // Output XML
  const xmlString = doc.end({ prettyPrint: true });
  fs.writeFileSync('output.xml', xmlString);
  console.log('XML generated: output.xml');
  console.log(`XML size: ${Buffer.byteLength(xmlString)} bytes`);
}

generateXML().catch(error => {
  console.error('Error during XML generation:', error.message);
});
