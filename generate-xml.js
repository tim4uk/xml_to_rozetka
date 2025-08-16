// generate-xml.js
import fs from 'fs';
import { google } from 'googleapis';
import { create } from 'xmlbuilder2';

// Авторизація
const auth = new google.auth.GoogleAuth({
  credentials: JSON.parse(process.env.GCP_SERVICE_ACCOUNT_KEY),
  scopes: ["https://www.googleapis.com/auth/spreadsheets.readonly"],
});

const sheets = google.sheets({ version: "v4", auth });

// ID таблиці
const SPREADSHEET_ID = "1y0dydLLy-l44qoVKVBaBE_Z7si2b1M55hKvyegiY21Y";
const SHEET_NAMES = ["Товари Ncase", "Вручну додані", "Товари Kiborg", "Товари Viktailor"];
const CATEGORY_SHEET = "Зв'язування категорій";

async function fetchSheet(name) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: name,
  });
  return res.data.values || [];
}

async function main() {
  let allItems = [];
  for (const name of SHEET_NAMES) {
    const data = await fetchSheet(name);
    if (data.length <= 1) continue;
    allItems = allItems.concat(data.slice(1)); // без заголовків
  }

  const categoriesData = await fetchSheet(CATEGORY_SHEET);

  // XML root
  const root = create({ version: "1.0", encoding: "UTF-8" })
    .ele("yml_catalog", { date: new Date().toISOString() })
      .ele("shop");

  // Валюта
  root.ele("currencies")
    .ele("currency", { id: "UAH", rate: "1" }).up().up();

  // Категорії
  const categories = root.ele("categories");
  categoriesData.slice(1).forEach(row => {
    if (row[0] && row[1]) {
      categories.ele("category", { id: row[0], rz_id: row[0] }).txt(row[1]).up();
    }
  });

  // Товари
  const offers = root.ele("offers");
  allItems.forEach(row => {
    const [
      id, stockQty, name, nameUa, price, categoryId,
      pictures, vendor, description, descriptionUa, params
    ] = row;

    const available = stockQty ? "true" : "false";
    const stock_quantity = available === "true" ? 30 : 0;

    const offer = offers.ele("offer", { id, available });
    offer.ele("name").txt(name || "").up();
    offer.ele("name_ua").txt(nameUa || "").up();
    offer.ele("price").txt(price || "").up();
    offer.ele("currencyId").txt("UAH").up();
    offer.ele("categoryId").txt(categoryId || "").up();

    if (pictures) {
      (pictures.includes(",") ? pictures.split(",") : [pictures]).forEach(pic => {
        if (pic) offer.ele("picture").txt(pic.trim()).up();
      });
    }

    if (vendor) {
      offer.ele("vendor").txt(vendor).up();
    }

    offer.ele("stock_quantity").txt(stock_quantity).up();

    // CDATA
    offer.ele("description").dat(description || "").up();
    offer.ele("description_ua").dat(descriptionUa || "").up();

    if (params) {
      params.split("\n").forEach(param => {
        const [pName, val] = param.split(" - ");
        if (pName && val) {
          offer.ele("param", { name: pName.trim() }).txt(val.trim()).up();
        }
      });
    }
  });

  const xml = root.end({ prettyPrint: true });
  fs.writeFileSync("feed.xml", xml, "utf-8");
}

main().catch(err => {
  console.error(err);
  process.exit(1);
});
