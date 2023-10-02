const puppeteer = require('puppeteer');
const fs = require('fs');
const ExcelJS = require('exceljs');
const https = require('https');

const colmns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V"];

async function urlToBase64(url) {
    return new Promise((resolve, reject) => {
        https.get(url, { responseType: 'arraybuffer' }, (response) => {
            const chunks = [];

            response.on('data', (chunk) => {
                chunks.push(chunk);
            });

            response.on('end', () => {
                const buffer = Buffer.concat(chunks); // binary data

                const base64Data = buffer.toString('base64'); // convert buffer to Base64 string
                resolve(base64Data);
            });
        }).on('error', (error) => {
            reject(error);
        });
    });
}

(async () => {
    const browser = await puppeteer.launch({ headless: false, defaultViewport: null });
    const page = await browser.newPage();

    let all_link = [];

    let page_number = 1;
    while (true) {
        await page.goto('https://www.examplepage.com/search?page=' + page_number); //This study was conducted to capture data from an e-commerce website. Please do not run it without permission from the owners of the websites.

        const links = await page.evaluate(() => {
            return Array.from(document.querySelectorAll(".product-item")).map(product => product.children[0].children[0].children[0].children[0].children[1].children[0].href);
        });

        if (links.length == 0) break;
        //if (page_number == 4) break;

        all_link = [...all_link, ...links];

        page_number++;
    }

    console.log(all_link.length);

    try {
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Sheet 1');

        async function add_row(product_details, row_index) { // Add row to excel file
            for (let i = 0; i < 3; i++) {
                if (product_details.images[i] == undefined) break;
                const base64Image = "data:image/png;base64," + await urlToBase64(product_details.images[i]);

                let column = colmns[i];
                const cell = worksheet.getCell(column + row_index);

                const imageId = workbook.addImage({
                    base64: base64Image,
                    extension: 'png',
                });

                const A_colm = worksheet.getColumn(column);
                A_colm.width = 20;

                const zero_row = worksheet.getRow(row_index);
                zero_row.height = 85;

                worksheet.addImage(imageId, {
                    tl: { col: cell.col - 1, row: cell.row - 1 },
                    br: { col: cell.col, row: cell.row }
                });
            }

            worksheet.getCell('D' + row_index).value = product_details.name;
            worksheet.getCell('E' + row_index).value = product_details.brand;
            worksheet.getCell('F' + row_index).value = product_details.description;
        }

        for (let i = 0; i < all_link.length; i++) { // Get product details from all links and add to excel file
            try {
                await page.goto(all_link[i]);

                await page.evaluate(() => { // Remove absolute whatsapp icon
                    var images = document.querySelectorAll("img");
                    images[images.length - 4].remove();
                });

                const product_details = await page.evaluate(() => { // Get product details
                    return {
                        name: document.querySelectorAll("h2")[0].innerText,
                        brand: document.querySelectorAll("h2")[0].parentNode.children[0].innerText,
                        images: [...new Set(Array.from(document.querySelectorAll(".slick-list.draggable")[0].children[0].children).map(a => a.children[0].children[0].src))],
                        description: document.querySelectorAll(".tab-pane")[0].innerText
                    }
                });

                console.log(product_details);

                if (product_details == undefined) continue;

                await add_row(product_details, i + 2);

            } catch (error) {
                console.error('Err:', error);
            }
        }

        workbook.xlsx.writeFile('output.xlsx') // Write excel file
            .then(() => {
                console.log("The file was saved!");
            })
            .catch((err) => {
                console.error('Err:', err);
            });

    } catch (error) {
        console.error('Err:', error);
    }

    //fs.writeFileSync('links.csv', all_link.map(row => `${row}`).join("\n"), 'utf8');

    await browser.close();
})();