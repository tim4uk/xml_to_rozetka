const { google } = require('googleapis');
const { create } = require('xmlbuilder2');
const fs = require('fs');

// Configuration
const SHEET_ID = process.env.SHEET_ID;
// const SHEET_NAMES = ['Товари Ncase', 'Вручну додані', 'Товари Kiborg', 'Товари Viktailor'];
const SHEET_NAMES = ['Вручну додані'];
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

// --- Sanitizers --------------------------------------------------------------
function sanitizeForXmlText(text) {
  if (!text) return '';
  return String(text)
    .replace(/&reg;?/gi, '®')
    .replace(/&copy;?/gi, '©')
    .replace(/&trade;?/gi, '™')
    .replace(/&nbsp;?/gi, ' ')
    // екранізуємо ЛИШЕ «сирі» амперсанди, не чіпаючи валідні XML-ентіті
    .replace(/&(?!amp;|lt;|gt;|quot;|apos;)/gi, '&amp;');
}

function sanitizeForCdata(text) {
  // те саме + безпечне розбиття ']]>' усередині CDATA
  const t = sanitizeForXmlText(text);
  return t.replace(/]]>/g, ']]]]><![CDATA[>');
}
// ---------------------------------------------------------------------------

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
    const filteredItems = FILTER_ENABLED
      ? items.filter(row =>
          row[COL_STOCK_QTY] === true || row[COL_STOCK_QTY] === 'TRUE' || row[COL_STOCK_QTY] === 'true'
        )
      : items;
    allItems = allItems.concat(filteredItems);
  }
  console.log(`Total items collected: ${allItems.length}`);

  // Fetch categories
  const categoriesData = await getSheetData(sheets, "Зв'язування категорій");
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
      categories.ele('category', { id, rz_id: id }).txt(sanitizeForXmlText(name));
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

    offer.ele('name').txt(sanitizeForXmlText(name));
    offer.ele('name_ua').txt(sanitizeForXmlText(name_ua));
    offer.ele('price').txt(sanitizeForXmlText(price));
    offer.ele('currencyId').txt('UAH');
    offer.ele('categoryId').txt(sanitizeForXmlText(category_id));

    if (image) {
      const pics = image.includes(',') ? image.split(',') : [image];
      for (const pic of pics) {
        const p = pic.trim();
        if (p) offer.ele('picture').txt(sanitizeForXmlText(p));
      }
    }

    if (vendor) offer.ele('vendor').txt(sanitizeForXmlText(vendor));

    offer.ele('stock_quantity').txt(String(stock_quantity));

    // CDATA з додатковою санітизацією
    offer.ele('description').dat(sanitizeForCdata(description.trim()));
    offer.ele('description_ua').dat(sanitizeForCdata(description_ua.trim()));

    const paramValues = params.split('\n');
    for (const param of paramValues) {
      const [nameParam, value] = param.split(' - ');
      if (nameParam && value) {
        offer
          .ele('param', { name: sanitizeForXmlText(nameParam.trim()) })
          .txt(sanitizeForXmlText(value.trim()));
      }
    }
    offerCount++;
  }
  console.log(`Added ${offerCount} offers`);

  // Output XML + страхувальна пост-обробка
  let xmlString = doc.end({ prettyPrint: true });

  // Якщо раптом прослизнули інші HTML-ентіті (типу &laquo;), перетворюємо на &amp;...
  const suspicious = xmlString.match(/&(?!amp;|lt;|gt;|quot;|apos;)[a-zA-Z]+;?/g);
  if (suspicious) {
    console.warn('Found suspicious entities, auto-escaping:', [...new Set(suspicious)].slice(0, 10));
    xmlString = xmlString.replace(/&(?!amp;|lt;|gt;|quot;|apos;)([a-zA-Z]+);?/g, '&amp;$1;');
  }

  fs.writeFileSync('output.xml', xmlString, 'utf8');
  console.log('XML generated: output.xml');
  console.log(`XML size: ${Buffer.byteLength(xmlString)} bytes`);
}

generateXML().catch(error => {
  console.error('Error during XML generation:', error);
  process.exit(1);
});
