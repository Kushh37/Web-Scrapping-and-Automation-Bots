const puppeteer = require('puppeteer');
const fs = require('fs');
const csv = require('csv-parser');
const xlsx = require('xlsx');

// Function to read the CSV file and return an array
const readCsvFile = async (filePath) => {
  return new Promise((resolve, reject) => {
    const addresses = [];
    fs.createReadStream(filePath)
      .pipe(csv())
      .on('data', (row) => {
        addresses.push(row.address);
      })
      .on('end', () => {
        resolve(addresses);
      })
      .on('error', reject);
  });
};

(async () => {
  const addresses = await readCsvFile('Z:/Geowarehouse/address.csv');
  const browser = await puppeteer.launch({headless: false});
  const page = await browser.newPage();

  const wb = xlsx.utils.book_new();
  const wsName = "Data";
  const wsData = [["Owner Name", "Last Sale", "Sale Date", "GeoWarehouse Address"]];

  await page.goto('https://iam.itsorealestate.ca/idp/login');
  await page.type('#clareity', 'kwkw1367');
  await page.type('#security', 'tysonandzeus');
  await page.click('#loginbtn');
  await page.waitForNavigation({waitUntil: 'networkidle0'});
  await page.goto('https://matrix.itsorealestate.ca/Matrix/Default.aspx');
  await page.waitForSelector('#ctl02_m_pnlTopMenu > div.TopMenuInner > table > tbody > tr > td.TopMenuLeft > a > img');

  for (const address of addresses) {
    console.log(address);
    const page = await browser.newPage();

    await page.goto('https://matrix.itsorealestate.ca/Matrix/special/thirdpartyformpost.aspx?n=GeoWarehouse');
    await page.waitForSelector('#searchTextBox');
    
    await page.type('#searchTextBox', address+' Kitchener');
    page.keyboard.press('Enter');

    await page.waitForSelector('#main-search-results-pane > div.search-result-data > div.ng-isolate-scope > div');
    
    //property report click
    await page.click('#main-search-results-pane > div.search-result-data > div.ng-isolate-scope > div > div.right > p:nth-child(4) > a');
    console.log('Property report clicked');

    await page.waitForSelector('#prbasicsection > div.property-report-basic-section.middle');
    console.log('2Property report clicked');
  
    const htmlContent = await page.content();

    //Dont remove this htmlcontent log
    console.log(htmlContent);
    const data = await page.evaluate(() => {
      const headers = Array.from(document.querySelectorAll('.header, .header2'));
      let ownerName = '';
      let lastSale = '';
      let saleDate = '';
      let geoWarehouseAddress = '';
  
      headers.forEach(header => {
        if (header.textContent.trim() === 'Owner Name') {
          const detailsElement = header.nextElementSibling;
          if (detailsElement && detailsElement.classList.contains('ng-binding')) {
            ownerName = detailsElement.textContent.trim();
          }
        }
  
        if (header.textContent.trim() === 'Last Sale') {
          const detailsElement = header.nextElementSibling;
          if (detailsElement && detailsElement.classList.contains('ng-binding')) {
            lastSale = detailsElement.textContent.trim();
          }
          const moreDetailsElement = detailsElement.nextElementSibling;
          if (moreDetailsElement && moreDetailsElement.classList.contains('ng-binding')) {
            saleDate = moreDetailsElement.textContent.trim();
          }
        }
  
        if (header.textContent.trim() === 'GeoWarehouse Address') {
          const detailsElement = header.nextElementSibling;
          if (detailsElement && detailsElement.classList.contains('ng-binding')) {
            geoWarehouseAddress = detailsElement.textContent.trim();
          }
        }
      });
  
      
      return { ownerName, lastSale, saleDate, geoWarehouseAddress };
    });
    console.log(`Owner Name: ${data.ownerName}\nLast Sale: ${data.lastSale}\nSale Date: ${data.saleDate}\nGeoWarehouse Address: ${data.geoWarehouseAddress}\n`);
    wsData.push([data.ownerName, data.lastSale, data.saleDate, data.geoWarehouseAddress]);
    await page.close();

  }

  const ws = xlsx.utils.aoa_to_sheet(wsData);
  xlsx.utils.book_append_sheet(wb, ws, wsName);
  xlsx.writeFile(wb, 'output.xlsx');

  await browser.close();
})();
