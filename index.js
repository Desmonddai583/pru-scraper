'use strict';

const puppeteer = require('puppeteer');
const Excel = require('exceljs');
const R = require('ramda');
const moment = require('moment');
const fs = require('fs');
const yargs = require('yargs');

const generatePaymentReport = async () => {
  const currentYear = new Date().getFullYear();

  const width = 1024;
  const height = 1600;
  const browser = await puppeteer.launch({
    headless: true,
    timeout: 0,
    'defaultViewport' : { 'width' : width, 'height' : height }
  });

  const getNewPageWhenLoaded =  async () => {
    return new Promise(x =>
      browser.on('targetcreated', async target => {
        if (target.type() === 'page') {
          const newPage = await target.page();
          const newPagePromise = new Promise(y =>
            newPage.once('domcontentloaded', () => y(newPage))
          );
          const isPageLoaded = await newPage.evaluate(
            () => document.readyState
          );
          return isPageLoaded.match('complete|interactive')
            ? x(newPage)
            : x(newPagePromise);
        }
      })
    );
  };

  const page = (await browser.pages())[0];

  await page.goto('https://salesforce.prudential.com.hk/sap/login');
  await page.waitForSelector('input[name=username], input[name=password]');

  await page.type('input[name=username]', username);
  await page.type('input[name=password]', password);
  await page.click('#submit');

  await page.waitForSelector('.apmenu-menu .apmenu-submenu > li:nth-child(1) > div');
  await page.hover('.apmenu-menu > li:nth-child(1)');
  await page.click('.apmenu-menu .apmenu-submenu > li:nth-child(1) > div');

  const aesPagePromise = getNewPageWhenLoaded();
  const aesPage = await aesPagePromise; 

  await aesPage.waitForSelector(".search_key.non_own_and_downline_case.non_overview_report");
  await aesPage.click('.search_key.non_own_and_downline_case.non_overview_report input');
  await aesPage.click('#normal_search_button_td input');

  let clientPolicys = [];
  await aesPage.waitForSelector("#result_list");
  const pageLength = (await aesPage.$$('#result_list_nav_page a')).length;
  for (let i = 0; i <= pageLength; i++) {
    if (i != 0) {
      await aesPage.click(`#result_list_nav_page a:nth-child(${i})`);
    }
    await aesPage.waitForSelector("#result_list");
    const clientLength = (await aesPage.$$('#result_list tr')).length;
    for (let j = 3; j < clientLength - 2; j++) {
      await aesPage.click(`#result_list tr:nth-child(${j}) a`);
      const clientPagePromise = getNewPageWhenLoaded();
      const clientPage = await clientPagePromise; 
      await clientPage.waitForSelector('#aes_search_container');
      const isExpired = await clientPage.evaluate(() => {
        const element = document.querySelector('#aes_search_container');
        return element.innerText.includes("由於此客戶所有保單或投保申請已失效超過最少兩年半，未能查閱客戶詳細資料"); 
      });
      if (isExpired) {
        clientPage.close();
        continue;
      }
      await clientPage.waitForSelector('[tab=policy_list]');
      const clientName = await clientPage.evaluate(() => {
        return document.querySelector('.result_details:nth-of-type(2) tr:nth-child(3) td:nth-child(4)').innerText.trim();
      });
      await clientPage.click('[tab=policy_list]');
      await clientPage.waitForSelector("#result_list");
      const noRecord = await clientPage.evaluate(() => {
        const element = document.querySelector('#result_list');
        return element.innerText.includes("找不到紀錄"); 
      });
      if (noRecord) {
        clientPage.close();
        continue;
      }
      const policyLength = (await clientPage.$$('#result_list tr')).length;
      for (let k = 3; k <= policyLength; k++) {
        const policyText = await clientPage.evaluate((index) => {
          return document.querySelector(`#result_list tr:nth-child(${index})`).innerText;
        }, k);
        console.log(policyText);
        if (policyText.includes("CANCEL FROM INCEPTION") || policyText.includes("VOID FROM INCEPTION")) {
          continue;
        }
        await clientPage.click(`#result_list tr:nth-child(${k}) td:nth-child(1) a`);
        const policyDetailPagePromise = getNewPageWhenLoaded();
        const policyDetailPage = await policyDetailPagePromise; 
        await policyDetailPage.waitForSelector('#details_table');
        const policyDetail = await policyDetailPage.evaluate((vars) => {
          const policyHolderSelector = document.querySelector('#blockBasicInfo .result_details:nth-child(1) tr:nth-child(3) td:nth-child(2)');
          let policyHolder = policyHolderSelector.innerText.replace(/\s+/g, " ").trim().split(/\(/)[0].trim().split(/\s/);
          policyHolder = policyHolder[policyHolder.length - 1];
          if (policyHolder !== vars.clientName) {
            return null;
          }
          const policyDueDateSelector = document.querySelector('#blockBasicInfo .result_details:nth-of-type(3) tr:nth-child(3) td:nth-child(4)')
          const policyDueDate = policyDueDateSelector.innerText.trim();
          if (parseInt(policyDueDate.split("/")[2]) != vars.currentYear) {
            return null;
          }
          const premiumSelector = document.querySelector('#blockBasicInfo .result_details:nth-of-type(3)>tbody>tr:nth-child(5)>td:nth-child(2)');
          const premium = parseFloat(premiumSelector.innerText.trim().replace(/,/g, ''));
          if (premium === 0) {
            return null;
          }
          const policyNumSelector = document.querySelector('#blockPolicyDetail .result_details:nth-child(2) tr:nth-child(1) td:nth-child(2)');
          const paymentNumSelector = document.querySelector('#blockBasicInfo .result_details:nth-of-type(3) tr:nth-child(2) td:nth-child(4)');
          const policyNameSelector = document.querySelector('#blockPolicyDetail .result_details:nth-child(2) tr:nth-child(1) td:nth-child(4)');
          const policyInsuredPersonSelector = document.querySelector('#blockBasicInfo .result_details:nth-child(1) tr:nth-child(3) td:nth-child(4)');
          let policyInsuredPerson = policyInsuredPersonSelector.innerText.replace(/\s+/g, " ").trim().split(/\(/)[0].trim().split(/\s/);
          policyInsuredPerson = policyInsuredPerson[policyInsuredPerson.length - 1];
          const policyIssueDateSelector = document.querySelector('#blockBasicInfo .result_details:nth-child(2) tr:nth-child(5) td:nth-child(4)');
          const pdaSelector = document.querySelector('#blockBasicInfo .result_details:nth-of-type(3)>tbody>tr:nth-child(6) td:nth-child(2)');
          const pda = parseFloat(pdaSelector.innerText.trim().replace(/,/g, ''));
          const beneficiarySelector = document.querySelector('#blockBasicInfo .result_details:nth-child(10n) tr:nth-child(3) td:nth-child(3)');
          const phoneSelector = document.querySelector('#blockBasicInfo .result_details:nth-child(1) tr:nth-child(4) td:nth-child(2)');
          const addressSelector = document.querySelector('#blockBasicInfo .result_details:nth-child(1) tr:nth-child(5) td:nth-child(2)');
          return {
            policyNum: policyNumSelector.innerText.trim(),
            paymentNum: paymentNumSelector.innerText.trim(),
            policyName: policyNameSelector.innerText.trim().replace(/\s/g, ''),
            policyHolder,
            policyInsuredPerson,
            policyIssueDate: policyIssueDateSelector.innerText.trim(),
            policyDueDate,
            premium,
            pda,
            beneficiary: beneficiarySelector ? beneficiarySelector.innerText.trim() : '',
            phone: phoneSelector.innerText.trim(),
            address: addressSelector.innerText.trim().replace("\n", "\r\n"),
          };
        }, {clientName, currentYear});
        if (policyDetail) {
          clientPolicys.push(policyDetail);
        }
        policyDetailPage.close();
      }
      clientPage.close();
    }
  }

  clientPolicys = R.map((p) => {
    p['premiumUSD'] = p.premium - p.pda;
    p['premiumHKD'] = Math.ceil(p['premiumUSD'] * 7.8 * 100) / 100 ;
    return p;
  }, clientPolicys);
  clientPolicys = R.groupBy(R.prop('policyHolder'), clientPolicys);
  R.forEachObjIndexed((policys, policyHolder) => {
    clientPolicys[policyHolder] = R.sort((p1, p2) => { 
      return moment(p1, "DD/MM/YYYY") - moment(p2, "DD/MM/YYYY"); 
    }, policys);
  }, clientPolicys);
  
  const borderStyle = {
    top: {style:'thin', color: {argb:'000000'}},
    left: {style:'thin', color: {argb:'000000'}},
    bottom: {style:'thin', color: {argb:'000000'}},
    right: {style:'thin', color: {argb:'000000'}}
  };
  const cellStyle =  {
    font: {size: 15, bold: true}, 
  };
  fs.mkdirSync('payment');
  const months = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"];
  months.forEach(month => fs.mkdirSync(`payment/${month}月`));
  await R.forEachObjIndexed(async (policys, policyHolder) => {
    const workbook = new Excel.Workbook();
    const worksheet = workbook.addWorksheet("缴费信息");
    worksheet.columns = [
      {header: '保单号码', key: 'policyNum', width: 20, style: cellStyle},
      {header: '缴费编号', key: 'paymentNum', width: 20, style: cellStyle}, 
      {header: '保单计划名称', key: 'policyName', width: 40, style: cellStyle},
      {header: '保单持有人', key: 'policyHolder', width: 15, style: cellStyle},
      {header: '受保人', key: 'policyInsuredPerson', width: 15, style: cellStyle},
      {header: '首期保费日', key: 'policyIssueDate', width: 15, style: cellStyle},
      {header: '保费到期日', key: 'policyDueDate', width: 15, style: cellStyle},
      {header: '保费(USD)', key: 'premium', width: 15, style: cellStyle},
      {header: "pda\r\n储蓄账户余额", key: 'pda', width: 20, style: cellStyle},
      {header: "应缴保费\r\nUSD", key: 'premiumUSD', width: 15, style: cellStyle},
      {header: "应缴保费\r\nHKD", key: 'premiumHKD', width: 15, style: cellStyle},
      {header: '受益人', key: 'beneficiary', width: 30, style: cellStyle},
      {header: '联系方式', key: 'phone', width: 25, style: cellStyle},
      {header: '通讯地址', key: 'address', width: 40, style: cellStyle},
    ];
    const header = worksheet.getRow(1)
    header.fill = {
      type: 'pattern',
      pattern:'solid',
      fgColor:{ argb:'F2A741' },
    }
    header.height = 35;
    worksheet.getRow(1).border = borderStyle;
    header.eachCell((cell, _) => {
      cell.alignment = {vertical: 'middle', horizontal: 'center', wrapText: true};
      cell.alignment.wrapText = true;
    });

    R.forEach(p => {
      const policyRow = worksheet.addRow(p);
      policyRow.border = borderStyle;
      policyRow.eachCell((cell, _) => {
        cell.alignment = {vertical: 'middle', horizontal: 'center', wrapText: true};
      });
    }, policys);

    const sumRow = worksheet.addRow({
      pda: '合计',
      premiumUSD: R.sum(R.map((p) => p.premiumUSD, policys)),
      premiumHKD: R.sum(R.map((p) => p.premiumHKD, policys)),
    })
    sumRow.border = borderStyle;
    sumRow.fill = {
      type: 'pattern',
      pattern:'solid',
      fgColor:{ argb:'F2A741' },
    }
    sumRow.eachCell((cell, _) => {
      cell.alignment = {vertical: 'middle', horizontal: 'center', wrapText: true};
    });
    
    const month = parseInt(policys[0].policyDueDate.split("/")[1]);
    await workbook.xlsx.writeFile(`payment/${month}月/${policyHolder}.xlsx`);
  }, clientPolicys);

  browser.close();
};

const argv = yargs
  .command('pru', 'fetch and generate client data from pru')
  .option('report', {
    alias: 'r',
    description: 'pru report type',
    type: 'string',
  })
  .option('username', {
    alias: 'u',
    description: 'pru agent username',
    type: 'string',
  })
  .option('password', {
    alias: 'p',
    description: 'pru agent password',
    type: 'string',
  })
  .help()
  .alias('help', 'h')
  .argv;

if (!argv.report) {
  console.log("please input report type");
  process.exit(1);
}

if (!argv.username || !argv.password) {
  console.log("please input username & password");
  process.exit(1);
}

const username = argv.username;
const password = argv.password;

switch (argv.report) {
  case 'payment':
    generatePaymentReport();
    break;
  default:
    console.log(`Sorry, the report type ${argv.report} does not exist.`);
    process.exit(1);
}