const { google } = require('googleapis');
const { create } = require('xmlbuilder2');
const fs = require('fs');

// Configuration
const SHEET_ID = process.env.SHEET_ID; // Replace or use env var
const KEY_FILE = process.env.KEY_JSON; // Path to your service account JSON key
const SHEET_NAMES = ['Товари Ncase', 'Вручну додані', 'Товари Kiborg', 'Товари Viktailor'];
const FILTER_ENABLED = false; // true to filter by stock_quantity === true/'TRUE'/'true'

// Column indices (0-based, matching your original)
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

// Authenticate with service account
async function authenticate() {
  const auth = new google.auth.GoogleAuth({
    keyFile: JSON.parse(process.env.KEY_JSON || '{}'),
    scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
  });
  return google.sheets({ version: 'v4', auth: await auth.getClient() });
}

// Fetch data from a sheet
async function getSheetData(sheets, sheetName) {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: `${sheetName}!A:L`, // Adjust range if needed
  });
  return response.data.values || [];
}

// Main function
async function generateXML() {
  const sheets = await authenticate();
  let allItems = [];

  // Fetch items from product sheets
  for (const name of SHEET_NAMES) {
    const data = await getSheetData(sheets, name);
    if (data.length <= 1) continue;
    const items = data.slice(1);
    let filteredItems = items;
    if (FILTER_ENABLED) {
      filteredItems = items.filter(row => 
        row[COL_STOCK_QTY] === true || row[COL_STOCK_QTY] === 'TRUE' || row[COL_STOCK_QTY] === 'true'
      );
    }
    allItems = allItems.concat(filteredItems);
  }

  // Fetch categories from "Зв'язування категорій"
  const categoriesData = await getSheetData(sheets, 'Зв\'язування категорій');
  const bindingCategories = categoriesData.slice(1);

  // Build XML
  const date = new Date().toISOString();
  const doc = create({ version: '1.0', encoding: 'UTF-8' })
    .ele('yml_catalog', { date });

  const shop = doc.ele('shop');

  // Currencies
  const currencies = shop.ele('currencies');
  currencies.ele('currency', { id: 'UAH', rate: '1' });

  // Categories
  const categories = shop.ele('categories');
  for (const bindingCategory of bindingCategories) {
    const id = bindingCategory[0];
    const name = bindingCategory[1];
    if (id && name) {
      categories.ele('category', { id, rz_id: id }).txt(name);
    }
  }

  // Offers
  const offers = shop.ele('offers');
  for (const row of allItems) {
    const id = row[COL_ID];
    const available = row[COL_STOCK_QTY] ? 'true' : 'false';
    const stock_quantity = available === 'true' ? 30 : 0;
    const name = row[COL_NAME] || '';
    const name_ua = row[COL_NAME_UA] || '';
    const price = row[COL_PRICE] || '';
    const currency = 'UAH';
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
    offer.ele('currencyId').txt(currency);
    offer.ele('categoryId').txt(category_id);

    if (image) {
      const pics = image.includes(',') ? image.split(',') : [image];
      for (const pic of pics) {
        if (pic.trim()) offer.ele('picture').txt(pic.trim());
      }
    }

    if (vendor) offer.ele('vendor').txt(vendor);

    offer.ele('stock_quantity').txt(stock_quantity.toString());

    // Use CDATA for descriptions
    offer.ele('description').dat(description.trim());
    offer.ele('description_ua').dat(description_ua.trim());

    const paramValues = params.split('\n');
    for (const param of paramValues) {
      const [nameParam, value] = param.split(' - ');
      if (nameParam && value) {
        offer.ele('param', { name: nameParam.trim() }).txt(value.trim());
      }
    }
  }

  // Output XML to file
  const xmlString = doc.end({ prettyPrint: true });
  fs.writeFileSync('output.xml', xmlString);
  console.log('XML generated: output.xml');
}

generateXML().catch(console.error);
