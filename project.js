//  node .\project.js --name='iphone 11' 


let minimist = require('minimist');
let axios = require('axios');
let jsdom = require("jsdom");
let excel = require("excel4node");


let args = minimist(process.argv);
// console.log(args.name);

let stringArray = args.name.split(" ");
let amazonString = ""
for (let i = 0; i < stringArray.length; i++) {
    amazonString += stringArray[i] + "+";
}
let flipkartString = ""
for (let i = 0; i < stringArray.length; i++) {
    flipkartString += stringArray[i] + "%20";
}

let amazonLink = "https://www.amazon.in/s?k=" + amazonString;
let flipkartLink = "https://www.flipkart.com/search?q=" + flipkartString + "&otracker=search&otracker1=search&marketplace=FLIPKART&as-show=on&as=off"

// console.log(amazonLink);
// console.log(flipkartLink);

let wb = new excel.Workbook();

let amazonDownloadPromise = axios.get(amazonLink);
amazonDownloadPromise.then(function (response) {
    let html = response.data;
    let dom = new jsdom.JSDOM(html);
    let document = dom.window.document;


    let amazonItems = [];
    let amazonItemsBlocks = document.querySelectorAll("div.s-include-content-margin.s-latency-cf-section.s-border-bottom");
    for (let i = 0; i < amazonItemsBlocks.length; i++) {
        let amazonItem = {};
        let itemName = amazonItemsBlocks[i].querySelector("span.a-size-medium.a-color-base.a-text-normal");
        amazonItem.Name = itemName.textContent;

        let itemPrice = amazonItemsBlocks[i].querySelector("span.a-price-whole")
        amazonItem.Price = "₹" + itemPrice.textContent;

        let itemLinkTag = amazonItemsBlocks[i].querySelector("a");
        let itemLink = "https://www.amazon.in" + itemLinkTag.href
        amazonItem.Link = itemLink;

        amazonItems.push(amazonItem);
    }


    let sheet = wb.addWorksheet("Amazon Products");
    sheet.column(1).setWidth(30);
    sheet.column(2).setWidth(20);
    sheet.row(1).setHeight(30);

    sheet.cell(1, 1).string("Product Name").style({
        font: {
            bold: true
        },
        alignment: {
            horizontal: 'center',
        },
    });
    sheet.cell(1, 2).string("Product Price").style({
        font: {
            bold: true
        },
        alignment: {
            horizontal: 'center',
        },
    });
    sheet.cell(1, 3).string("Link").style({
        font: {
            bold: true
        },
        alignment: {
            horizontal: 'center',
        },
    });
    for (let j = 0; j < amazonItems.length; j++) {
        sheet.cell(3 + j, 1).string(amazonItems[j].Name).style({
            alignment: {
                wrapText: true
            }
        });
        sheet.cell(3 + j, 2).string(amazonItems[j].Price).style({
            alignment: {
                horizontal: 'center'
            }
        });
        sheet.cell(3 + j, 3).link(amazonItems[j].Link, "Click here").style({
            alignment: {
                wrapText: true
            }
        });
    }
    wb.write("Products.xls");


}).catch(function (err) {
    console.log(err.message);
})

let flipkartDownloadPromise = axios.get(flipkartLink);
flipkartDownloadPromise.then(function (response) {
    let html = response.data;
    let dom = new jsdom.JSDOM(html);
    let document = dom.window.document;

    let flipkartItems = [];
    let flipkartItemsBlocks = document.querySelectorAll("div._2kHMtA");
    for (let i = 0; i < flipkartItemsBlocks.length; i++) {
        let flipkartItem = {};
        let itemName = flipkartItemsBlocks[i].querySelector("div._4rR01T");
        flipkartItem.Name = itemName.textContent;

        let itemPrice = flipkartItemsBlocks[i].querySelector("div._30jeq3._1_WHN1")
        flipkartItem.Price = itemPrice.textContent;

        let itemLinkTag = flipkartItemsBlocks[i].querySelector("a");
        let itemLink = "https://www.flipkart.com" + itemLinkTag.href
        flipkartItem.Link = itemLink;

        flipkartItems.push(flipkartItem);
    }

    let sheet = wb.addWorksheet("Flipkart Products");
    sheet.column(1).setWidth(30);
    sheet.column(2).setWidth(20);
    sheet.row(1).setHeight(30);

    sheet.cell(1, 1).string("Product Name").style({
        font: {
            bold: true
        },
        alignment: {
            horizontal: 'center',
        },
    });
    sheet.cell(1, 2).string("Product Price").style({
        font: {
            bold: true
        },
        alignment: {
            horizontal: 'center',
        },
    });
    sheet.cell(1, 3).string("Link").style({
        font: {
            bold: true
        },
        alignment: {
            horizontal: 'center',
        },
    });
    for (let j = 0; j < flipkartItems.length; j++) {
        sheet.cell(3 + j, 1).string(flipkartItems[j].Name).style({
            alignment: {
                wrapText: true
            }
        });
        sheet.cell(3 + j, 2).string(flipkartItems[j].Price).style({
            alignment: {
                horizontal: 'center'
            }
        });
        sheet.cell(3 + j, 3).link(flipkartItems[j].Link, "Click here").style({
            alignment: {
                wrapText: true
            }
        });
    }
    wb.write("Products.xls");


}).catch(function (err) {
    console.log(err.message);
})