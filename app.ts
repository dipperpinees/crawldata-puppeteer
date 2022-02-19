const puppeteer = require('puppeteer');
const excel = require('excel4node');

const crawlData = async () => {
    const browser = await puppeteer.launch({ headless: false });
    const page = await browser.newPage();   

    await page.goto('https://www.thegioididong.com/laptop-asus');

    const pageCrawl = await page.evaluate(() => {
        let elems:any = document.querySelectorAll('.listproduct > li');
        elems = [...elems];
        let data = elems.map((item:any) => ({
            "Tên": item.querySelector('h3').textContent.trim(),
            "Giá gốc": item.querySelector('.price-old')?.textContent,
            "Giá hiện tại": item.querySelector('.price')?.textContent,
            "Ảnh sản phẩm": item.querySelector('.item-img img')?.getAttribute("src") || item.querySelector('.item-img img')?.getAttribute("data-src"),
            "Màn hình": item.querySelector('.utility p:nth-child(1) span:last-child')?.textContent,
            "CPU": item.querySelector('.utility p:nth-child(2) span:last-child')?.textContent,
            "CARD": item.querySelector('.utility p:nth-child(3) span:last-child')?.textContent,
            "PIN": item.querySelector('.utility p:nth-child(4) span:last-child')?.textContent,
        }));
        return data;
    })

    console.log(pageCrawl)

    await writeExcel("data", pageCrawl, ["Tên", "Giá gốc", "Giá hiện tại", "Ảnh sản phẩm", "Tỷ lệ khuyến mãi", "Màn hình", "CPU", "CARD", "PIN"]);

    await browser.close(); 
}

const writeExcel = (name:string, data:any, columnList:string[]) => {
    let workbook = new excel.Workbook();
  
    let worksheet = workbook.addWorksheet('Sheet 1');
  
    let style = workbook.createStyle({
      font: {
        size: 12
      },
    });
    worksheet.cell(1,1).string("STT").style(style);
    columnList.forEach((column:string, index:number) => {
        worksheet.cell(1,index + 2).string(column).style(style);
    })

    data.forEach((item:any, indexData:number) => {
        worksheet.cell(indexData + 2, 1).number(indexData + 1).style(style);
        columnList.forEach((column:string, indexColumn:number) => {
            if(item[column]) {
                worksheet.cell(indexData + 2,indexColumn + 2).string(item[column]).style(style);
            }
        })
    })
  
    workbook.write(`./xlsx/${name}.xlsx`);
}

crawlData();