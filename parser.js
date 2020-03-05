require("chromedriver"); //导入chrome浏览器 driver
const { Builder, By, Key } = require("selenium-webdriver");
const chrome = require("selenium-webdriver/chrome");
const Excel = require("exceljs");
const fs = require("fs");
const Options = chrome.Options;
const chromePath = require("chromedriver").path;

async function parser(webUrl, driver) {
  await driver.get("https://tool.chinaz.com/speedcom/" + webUrl);
  await waitJquery(driver);
  for (let m = 0; m < 0; m++) {
    await driver
      .findElement(By.className("allRetry fl mt3 ml10  col-red"))
      .click();
    console.log("Refresh");
    await waitJquery(driver);
  }

  let items = await driver.findElements(
    By.className("row listw compare clearfix")
  );
  let all = [];
  let header = [
    "监测点",
    "对比网址",
    "解析IP",
    "HTTP状态",
    "总耗时",
    "解析时间",
    "连接时间",
    "下载时间",
    "文件大小",
    "下载速度",
    "操作",
    "赞助商"
  ];
  all.push(header);

  for (let i = 0; i < items.length; i++) {
    let dGroup = await items[i].findElements(
      By.className("row bor-l1s clearfix")
    );
    for (let j = 0; j < dGroup.length; j++) {
      let colArr = [];

      let aa = await items[i]
        .findElement(By.className("subw"))
        .findElement(By.xpath(".//span"))
        .getText();
      colArr.push(aa);

      let ddGroup = await dGroup[j].findElements(By.tagName("dd"));
      for (let k = 0; k < ddGroup.length; k++) {
        let bb;
        bb = await exedd(ddGroup[k]);
        colArr.push(bb);
      }
      let cc = await items[i]
        .findElement(By.className("subwr bor-l1s"))
        .findElement(By.tagName("a"))
        .getText();
      colArr.push(cc);

      all.push(colArr);
    }
  }
  console.log("Parser " + webUrl);
  return all;
}

async function waitJquery(driver) {
  while (true) {
    var validation1;
    var validation2;

    driver.executeScript("return jQuery.active").then(result => {
      validation1 = result;
    });
    driver.executeScript("return document.readyState").then(result => {
      validation2 = result;
    });

    if (validation1 == "0" && validation2 == "complete") {
      console.log("Load Jquery");
      return;
    }
    await delay(1000);
  }
}

function delay(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

async function exedd(dd) {
  try {
    return await dd.findElement(By.tagName("a")).getText();
  } catch (error) {
    return await dd.findElement(By.tagName("p")).getText();
  }
}

async function exportExcel(all, webUrl, dateStr) {
  var fileName = "parser-" + webUrl + "-" + dateStr + ".xlsx";
  var workbook = new Excel.stream.xlsx.WorkbookWriter({
    filename: "./" + dateStr + "/" + fileName
  });
  var worksheet = workbook.addWorksheet("Sheet");

  for (let i in all) {
    worksheet.addRow(all[i]).commit();
  }
  workbook.commit();
  console.log("Finished Exporting " + fileName);
}

async function readFilePromise() {
  return new Promise((resolve, reject) => {
    fs.readFile("./input.txt", "utf-8", async (err, input) => {
      if (err) return reject(err);
      resolve(input);
    });
  });
}

async function mkdirPromise(dateStr) {
  return new Promise((resolve, reject) => {
    fs.mkdir(
      "./" + dateStr,
      {
        recursive: true
      },
      async err => {
        if (err) return reject(err);
        console.log("Mkdir ./" + dateStr);
        resolve("Mkdir ./" + dateStr);
      }
    );
  });
}

async function start() {
  console.log("Project Start");
  var date = new Date();
  var dateStr =
    date.getFullYear() +
    ("00" + (date.getMonth() + 1)).slice(-2) +
    ("00" + date.getDate()).slice(-2) +
    "-" +
    ("00" + date.getHours()).slice(-2) +
    ("00" + date.getMinutes()).slice(-2) +
    ("00" + date.getSeconds()).slice(-2);

  let input = await readFilePromise();
  let localArr = input.split("\r\n")[0].split(",");
  let remoteArr = input.split("\r\n")[1].split(",");

  await mkdirPromise(dateStr);

  for (let i in localArr) {
    for (let j in remoteArr) {
      let driver;
      try {
        driver = await new Builder()
          .forBrowser("chrome")
          .setChromeOptions(
            new Options()
              .excludeSwitches(["ignore-certificate-errors", chromePath])
              .addArguments("disable-infobars")
              .addArguments("--headless")
          )
          .build();
        let webUrl = localArr[i] + "-" + remoteArr[j];
        let all = await parser(webUrl, driver);
        await exportExcel(all, webUrl, dateStr);
      } catch (err) {
        console.error(err);
      } finally {
        driver.close();
      }
    }
  }
  console.log("Done Parser");
}

start();
