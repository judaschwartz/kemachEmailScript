const formId = '' // ID of the form
const spreadsheetId = '' // ID of the result spreadsheetId
const year = '23'
const closeDate = 'August 1st'
const paymentDate = 'August 15, 2023'
const pickupDate = 'SUNDAY, September 10th'
const yt = 'Sukkos'//'Pesach'
const sheetName = `Orders ${yt[0]}${year}` // name of the result sheet
const hasCoupons = yt[0] === 'P' ? 2 : 0
const coupons = [['S4', 400], ['S2', 200], ['MR', 250], ['KK', 300], ['LS', 200], ['KO', 300], ['RE', 400]]

function triggerOnSubmit(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
  const orderRow = sheet.getActiveRange().getRow()
  primaryOrderProcessor(sheet, orderRow, true)
}

function runManually() {
  const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName)
  const numColumns = sheet.getLastColumn()
  const numRows = sheet.getLastRow() + 1
  for (let row = 5; row < numRows; row++) {
    if (sheet.getRange(row, numColumns-1).isBlank() && !sheet.getRange(row, 2).isBlank()) {
      console.log('Drafting Email for ', sheet.getRange(row, 2).getValue())
      primaryOrderProcessor(sheet, row, false)
    }
  }
}

function primaryOrderProcessor(sheet, orderRow, send = true) {
  const numRows = sheet.getLastRow()
  const numColumns = sheet.getLastColumn()
  const priceHeaderData = sheet.getRange(1, 1, 3, numColumns).getValues()
  const dateStamp = sheet.getRange(orderRow, 1).getValue()
  const semail = sheet.getRange(orderRow, 2).getValue()
  const orderNum = semail.endsWith('@matanbsayser.org') ? semail.replace('@matanbsayser.org', '') : 96+orderRow
  sheet.getRange(orderRow, numColumns-(2+hasCoupons)).setValue(orderNum)
  const totalItemsCellName = sheet.getRange(orderRow, numColumns-(3+hasCoupons)).getA1Notation()
  let totalFormula = hasCoupons ? "=0" : `=IF(${totalItemsCellName}>0,15,0)`
  let firstClm, colmnl, cpnCol, stayCol = ''
  for (let j=2;j<numColumns-(3+hasCoupons);j++) {
    if (priceHeaderData[2][j]) {
      colmnl = sheet.getRange(2, j+1, 2, 1).getA1Notation().match(/([A-Z]+)/)[0]
      totalFormula += "+("+colmnl+orderRow+"*"+colmnl+"$3)"
      firstClm = firstClm ? firstClm : colmnl
    }
    if (hasCoupons && priceHeaderData[1][j] === 'Coupon code') {
      stayCol = sheet.getRange(1, j+1, 1, 1).getA1Notation().match(/([A-Z]+)/)[0]
    }
    if (hasCoupons && priceHeaderData[1][j] === 'Staying Home') {
      cpnCol = sheet.getRange(1, j+1, 1, 1).getA1Notation().match(/([A-Z]+)/)[0]
    }
  }
  sheet.getRange(orderRow, numColumns-(3+hasCoupons)).setValue(`=SUM(${firstClm}${orderRow}:${colmnl}${orderRow})`)
  sheet.getRange(orderRow, numColumns-(1+hasCoupons)).setValue(totalFormula)
  if (hasCoupons) {
    const totalCellName = sheet.getRange(orderRow, numColumns-3).getA1Notation()
    let couponFormula= `=-1*IF("Yes"=${stayCol}${orderRow}, MIN(CEILING.MATH(${totalCellName}/2), 0`
    coupons.map((coupon) => {
      couponFormula += `+(IFERROR(SEARCH("${coupon[0]}", ${cpnCol}${orderRow})*${coupon[1]}, 0))`
    })
    couponFormula += '), 0)'
    sheet.getRange(orderRow, numColumns-2).setValue(couponFormula)
    sheet.getRange(orderRow, numColumns-1).setValue(`=IF(${totalCellName}>0,15,0)+${totalCellName}+${sheet.getRange(orderRow, numColumns-2).getA1Notation()}`)
  }
  const lasturl = getEditLink(dateStamp, semail)
  if (lasturl){
    sheet.getRange(orderRow, numColumns).setValue(lasturl)
  }
  const dataRange = sheet.getRange(1, 1, numRows, numColumns)
  const data = dataRange.getValues()
  const { subject, message, ordersForEmail } = composeEmail(data, orderRow, numColumns)
  const draft = GmailApp.createDraft(semail, subject, message, {htmlBody: message})
  if (send && ordersForEmail < 2) {draft.send();}
}

function checkMyQuota() {
  const remaining = MailApp.getRemainingDailyQuota()
  console.log("You can still send " + remaining + " emails today.")
  return remaining
}

function composeEmail (data, orderRow, numColumns) {
  const header_row = data[1]
  const price_row = data[2]
  const row = data[orderRow-1]
  const ordersForEmail = data.filter(r => r[1] === row[1] && r[numColumns - (4+hasCoupons)]).length
  let subject = ordersForEmail > 1 ? '**POSSIBLE DUPLICATE**' : ''
  subject += `${yt} 20${year} Kemach ${row[1].endsWith('@matanbsayser.org') ? 'PAPER ' : ''}Order ${yt[0]}${year}-${row[numColumns - (3+hasCoupons)]}`
  let message = ordersForEmail > 1 ? '<div style="color:red;background: yellow;padding: 5px 20px;"><h2>WARINING THERE IS ALREADY ANOTHER ORDER FROM THE SAME EMAIL ADDRESS</h2><p style="font-size: 15px;">If you intended to make both orders you can ignore this message, otherwise please edit one of your orders and set all the quantities to zero to cancel the duplicate.</p><h2>IF YOU LEAVE BOTH ORDERS AS IS YOU WILL BE RESPONSIBLE TO PAY FOR BOTH OF THEM</h2></div>' : ''
  message +=`<div style="font-size: calc(0.5vw + 10px); max-width: 800px; margin: auto;"><div style="border: solid 3px #000; padding: 10px 15px; font-size: 20px;"><b>ID:</b> ${yt[0]}${year}-${row[numColumns - (3+hasCoupons)]}<b> NAME:</b> ${row[2].toUpperCase()}</div><img src="http://levavrohom.org/Capture.PNG" /><p>Your ${yt} 20${year} Kemach order was Successfully Submitted.<br><span style="background: #e1e100;">*IMPORTANT*</span> Please double check that the information below is correct.</p><h1 style="text-align: center;">Order # ${yt[0]}${year}-${row[numColumns - (3+hasCoupons)]} Summary</h1>`
  message += "<table style='width: 100%; font-size:calc(0.6vw + 9px); border-spacing: 0;'><tr><th>Item</th><th>Price</th><th><b>Qty</b></th><th><b>Total</b></th></tr>"
  let style ='background:#ddd;'
  for (let j=2;j<numColumns-(2+hasCoupons);j++) {
    if (row[j]!="" && row[j]!="0" && j!=8 && j!=9 && j!=10) {
      message += `<tr style='padding:5px; ${style}'><td style='max-width: 350px'>${header_row[j]}</td>`
      message += `<td style='padding:0 15px 0 5px;'>${price_row[j] ? "$"+price_row[j] : ''}</td>`
      message += `<td><b>${ ((price_row[j]) ? row[j] : "")}</b></td>`
      message += `<td>${price_row[j] ? '<b>$'+(row[j]*price_row[j])+'</b>' : row[j].toString().substr(0,18)}</td></tr>`
      style = style ? '' : 'background:#ddd;'
    }
  }
 
  if (row[numColumns-2]) {
    message += `<tr style='padding:5px; ${style}'><td style='max-width: 350px'>Processing Fee</td><td></td><td></td><td style='padding:0 15px 0 5px;'><b>$15</b></td></tr>`
    if (hasCoupons && row[numColumns-3] < 0) {
      message += `<tr style='padding:5px; ${style = style ? '' : 'background:#ddd;'}'><td style='max-width: 350px'>Coupon</td><td></td><td style='text-align: right; padding:0 15px 0 5px;'></td><td><b>${row[numColumns-3]}</b></td></tr>`
    }
    message += `<tr style='background:yellow; padding:5px; font-size:200%;'><td style='max-width: 350px'>Order Total</td><td></td><td></td><td style='padding:0 15px 0 5px;'><b>$${row[numColumns-2]}</b></td></tr></table>`
    message += `<p>This order can be changed by clicking on the button below until <b>${closeDate}</b>, please DO NOT create a new order.</p>`
    message += `<a style="background:darkblue; border-radius: 10px; padding: 15px 0; font-size:120%; display: block; width: 80%; margin: 20px 10%; text-align:center; color: #fff; border: 1px solid #000" href="${row[numColumns-1]}">EDIT MY ORDER</a>`
    message += `<p>PAYMENT MUST BE RECEIVED BY <b>${paymentDate}</b> Payment options include cash, check, Zelle or credit card.</p>`
    message += '<p>If you pay with <b>ZELLE</b> there are no additional fees<br>Zelle payments can be sent to <b>kemach@matanbsayser.org</b><br>Please include your order number in the memo to ensure we can credit you for your payment</p>'
    message += 'We also accept credit card payments for your Kemach order! There is an additional 3% charge to cover the credit card processing fee if you choose to pay with credit card.'
    const ccFee = Math.floor(row[numColumns-2]*0.03)
    message += `<table style='width: 100%; font-size:calc(0.6vw + 9px); border-spacing: 0;'><tr style='background:#ddd;'><td>Three percent additional charge <b>ONLY</b> IF PAYING WITH CREDIT CARD</td><td><b>$${ccFee}</b></td></tr>`
    message += `<tr style='padding: 5px;'><td style='background: #ffff8e; width: calc(100% - 105px)'>Order Total For Credit Card payment ONLY</td><td style='width 105px; text-align: center'><b>$${row[numColumns-2] + ccFee}</b></td></tr></table>`
    message += `<a style="background:red; padding:15px 0; font-size:150%; border-radius: 10px; display: block; width: 60%; margin: 20px 20%; text-align:center; color: #fff; border: 1px solid #000" href="matanbsayser.org/Kemach?amnt=${row[numColumns-2] + ccFee}&orderid=${row[numColumns-(3+hasCoupons)]}">PAY ONLINE</a>`
    message += `<p>If you are paying with cash or check please make sure you submit your payment to 1928 Janette Avenue Cleveland, Ohio 44118 before ${paymentDate}. When you drop off the payment PLEASE make sure it is in a sealed envelope with your name and address clearly written on it.</p>`
    message += `DISTRIBUTION IS SLATED FOR <span style='background:yellow;'>${pickupDate}.</span> at Hillcrest Foods 2735 East 40th, Cleveland, Ohio 44115<br>Important: Please enter through back road.<br>`
    message += 'In order for the entire project to function successfully, we need manpower on that day to assist us with the distribution process. Please let us know if you are able to volunteer by either calling 440-7KEMACH (753-6224) or email kemachcleveland@gmail.com'
    message += '<p>If you have any questions please call the Kemach office 440-7KEMACH (753-6224) and leave a message.</p></div>'
  } else {
    message += "<tr style='background:yellow; padding:5px; font-size:200%;'><td style='max-width: 350px'>Order Canceled</td><td></td><td></td><td style='padding:0 15px 0 5px;'><b>$0</b></td></tr></table>"
    message += `<a style="background:darkblue; border-radius: 10px; padding: 15px 0; font-size:120%; display: block; width: 80%; margin: 20px 10%; text-align:center; color: #fff; border: 1px solid #000" href="${row[numColumns-1]}">EDIT MY ORDER</a>`
  }
  return { subject, message, ordersForEmail }
}

function getEditLink(timeStamp='', semail = '') {
  let form
  try {
    form = FormApp.openById(formId)
  } catch {
    console.log('form unavailble')
    Utilities.sleep(15*1000)
    console.log('retry form')
    form = FormApp.openById(formId)
  }
  let lasturl = ''
  const responses = form.getResponses()
  for (const response of responses) {
    if (timeStamp) {
      rts = response.getTimestamp()
      if (timeStamp == rts.toString()) {
        lasturl = response.getEditResponseUrl()
        console.log({ rts, lasturl })
      }
    } else {
      const femail = response.getRespondentEmail()
      if (femail == semail) {
        lasturl = response.getEditResponseUrl()
        console.log(lasturl)
      }
    }
  }
  return lasturl
}
