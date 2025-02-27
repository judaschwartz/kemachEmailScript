const year = '25'
const yt = 'Pesach' // 'Sukkos'
const closeDate = 'March 4th'
const paymentDate = 'March 24th'
const pickupDate = 'SUNDAY, April 6th'
const processingFee = 15
const priceSplitter = '- $'
const skipRowsInEmail = [0,5,7,12]
const coupons = [['S4', 400], ['S2', 200], ['MR', 200], ['KK', 300], ['KO', 300], ['RE', 400], ['LS', 200]]
const couponCodeCol = 'Coupon code'
const stayingHomeCol = 'Staying Home'
const orderSheet = `Orders ${yt[0]}${year}`
const paymentsSheet = `Payments ${yt[0]}${year}`
const hasCoupons = yt.startsWith('P') ? 2 : 0

function getRemainingDailyEmails() {
  const remaining = MailApp.getRemainingDailyQuota()
  console.log("You can still send " + remaining + " emails today.")
  return remaining
}

function triggerOnSubmit(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
  const orderRow = sheet.getActiveRange().getRow()
  primaryOrderProcessor(sheet, orderRow, true)
}

function runManually() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(orderSheet)
  const numColumns = sheet.getLastColumn()
  const numRows = sheet.getLastRow() + 1
  for (let row = 5; row < numRows; row++) {
    if (sheet.getRange(row, numColumns-(2+hasCoupons)).isBlank() && !sheet.getRange(row, 2).isBlank()) {
      console.log('Drafting Email for ', sheet.getRange(row, 2).getValue())
      primaryOrderProcessor(sheet, row, false)
    }
  }
}

function intilizeProdNamePriceRows() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(orderSheet)
  const cols = sheet.getLastColumn()
  const data = sheet.getRange(1, 1, 1, cols).getValues()
  const pRows = hasCoupons ? ['subtotal', 'Coupon discount'] : []
  sheet.getRange(2, 1, 2, cols + 4 + hasCoupons).setValues([
    [...data[0].map((v, i) => !skipRowsInEmail.includes(i) ? v.split(priceSplitter)[0]?.trim() : ''), '# of Items ordered', 'Order ID', ...pRows, 'total', 'edit link'],
    [...data[0].map((v, i) => !skipRowsInEmail.includes(i) ? v.split(priceSplitter)[1]?.trim() : ''),,,,,]
  ])
}

function primaryOrderProcessor(sheet, orderRow, send) {
  const numRows = sheet.getLastRow()
  const numColumns = sheet.getLastColumn()
  const semail = setOrderValues(sheet, orderRow, numColumns)
  let { subject, message } = composeEmail(sheet, orderRow, numRows, numColumns)
  const draft = GmailApp.createDraft(semail, subject, message, {htmlBody: message})
  send && draft.send()
}

function setOrderValues(sheet, orderRow, numColumns) {
  const priceHeaderData = sheet.getRange(1, 1, 3, numColumns).getValues()
  const semail = sheet.getRange(orderRow, 2).getValue()
  const dateStamp = sheet.getRange(orderRow, 1).getValue()
  const orderNum = semail.endsWith('@matanbsayser.org') ? semail.replace('@matanbsayser.org', '') : 96 + orderRow
  const lasturl = getEditLink(sheet.getFormUrl(), dateStamp, semail)
  const totalItemsCellName = sheet.getRange(orderRow, numColumns-(3+hasCoupons)).getA1Notation()
  let totalFormula = hasCoupons ? '=0' : `=IF(${totalItemsCellName}>0,${processingFee}`
  let firstProdCol, lastProdCol, cpnCol, stayCol
  for (let j=2; j < numColumns-1; j++) {
    if (priceHeaderData[2][j]) {
      lastProdCol = j + 1
      firstProdCol = firstProdCol || lastProdCol
    }
    if (hasCoupons && priceHeaderData[1][j] === stayingHomeCol) {
      stayCol = sheet.getRange(1, j+1, 1, 1).getA1Notation().match(/([A-Z]+)/)[0]
    }
    if (hasCoupons && priceHeaderData[1][j] === couponCodeCol) {
      cpnCol = sheet.getRange(1, j+1, 1, 1).getA1Notation().match(/([A-Z]+)/)[0]
    }
  }
  firstProdCol = sheet.getRange(2, firstProdCol, 2, 1).getA1Notation().match(/([A-Z]+)/)[0]
  lastProdCol = sheet.getRange(2, lastProdCol, 2, 1).getA1Notation().match(/([A-Z]+)/)[0]
  totalFormula += `+SUMPRODUCT(${firstProdCol}${orderRow}:${lastProdCol}${orderRow}, ${firstProdCol}$3:${lastProdCol}$3)${hasCoupons ? '' : ',"")'}`  
  const countFormula = `=SUM(${firstProdCol}${orderRow}:${lastProdCol}${orderRow})`
  const endInfo = [countFormula, orderNum, totalFormula]
  if (hasCoupons) {
    const totalCellName = sheet.getRange(orderRow, numColumns-3).getA1Notation()
    let couponFormula= `=-1*IF("Yes"=${stayCol}${orderRow}, MIN(CEILING.MATH(${totalCellName}/2), 0`
    coupons.forEach((coupon) => {
      couponFormula += `+(IFERROR(SEARCH("${coupon[0]}", ${cpnCol}${orderRow})*${coupon[1]}, 0))`
    })
    couponFormula += '), 0)'
    const gtf = `=IF(${totalItemsCellName}>0,${processingFee}+${totalCellName}+${sheet.getRange(orderRow, numColumns-2).getA1Notation()},"")`
    endInfo.push(couponFormula, gtf)
  }
  sheet.getRange(orderRow, numColumns - (3 + hasCoupons), 1, 4 + hasCoupons).setValues([[...endInfo, lasturl]])
  return semail
}

function composeEmail (sheet, orderRow, numRows, numColumns) {
  const data = sheet.getRange(1, 1, numRows, numColumns).getValues()
  const header_row = data[1]
  const price_row = data[2]
  const row = data[orderRow-1]
  const orderId = `${yt[0]}${year}-${row[numColumns - (3+hasCoupons)]}`
  const ordersForEmail = data.filter(r => r[1] === row[1] && r[numColumns - 2]).length
  let subject = ordersForEmail > 1 ? '**POSSIBLE DUPLICATE**' : ''
  subject += `${yt} 20${year} Kemach ${row[1].endsWith('@matanbsayser.org') ? 'PAPER ' : ''}Order ${orderId}`
  let message = ordersForEmail > 1 ? '<div style="color:red;background: yellow;padding: 5px 20px;"><h2>WARINING THERE IS ALREADY ANOTHER ORDER FROM THE SAME EMAIL ADDRESS</h2><p style="font-size: 15px;">If you intended to make both orders you can ignore this message, otherwise please edit one of your orders and set all the quantities to zero to cancel the duplicate.</p><h2>IF YOU LEAVE BOTH ORDERS AS IS YOU WILL BE RESPONSIBLE TO PAY FOR BOTH OF THEM</h2></div>' : ''
  message +=`<div style="font-size: calc(0.5vw + 10px); max-width: 800px; margin: auto;"><div style="border: solid 3px #000; padding: 10px 15px; font-size: 20px;"><b>ID:</b> ${orderId}<b> NAME:</b> ${row[2].toUpperCase()}</div><img src="http://levavrohom.org/Capture.PNG" /><p>Your ${yt} 20${year} Kemach order was Successfully Submitted.<br><span style="background: #e1e100;">*IMPORTANT*</span> Please double check that the information below is correct.</p><h1 style="text-align: center;">Order # ${orderId} Summary</h1>`
  message += "<table style='width: 100%; font-size:calc(0.6vw + 9px); border-spacing: 0;'>"
  let style = 'background:#ddd;'
  message += `<tr style='${style}'><th>Item</th><th>Price</th><th>Qty</th><th>Total</th></tr>`
  for (let j=2;j<numColumns-(4+hasCoupons);j++) {
    if (row[j]!="" && parseInt(row[j]) !== 0 && !skipRowsInEmail.includes(j)) {
      style = style ? '' : 'background:#ddd;'
      message += `<tr style='padding:5px; ${style}'><td style='max-width: 350px'>${header_row[j]}</td>`
      message += `<td style='padding:0 15px 0 5px;'>${price_row[j] ? "$"+price_row[j] : ''}</td>`
      message += `<td><b>${ ((price_row[j]) ? row[j] : "")}</b></td>`
      message += `<td>${price_row[j] ? '<b>$'+(row[j]*price_row[j])+'</b>' : row[j].toString().substr(0,18)}</td></tr>`
    }
  }
  if (row[numColumns-2]) {
    message += `<tr style='padding:5px; ${style ? '' : 'background:#ddd;'}'><td style='max-width: 350px'>Processing Fee</td><td></td><td></td><td style='padding:0 15px 0 5px;'><b>$${processingFee}</b></td></tr>`
    if (hasCoupons && row[numColumns-3] < 0) {
      message += `<tr style='padding:5px; ${style}'><td style='max-width: 350px'>Coupon</td><td></td><td style='text-align: right; padding:0 15px 0 5px;'></td><td><b>${row[numColumns-3]}</b></td></tr>`
    }
    message += `<tr style='background:yellow; padding:5px; font-size:200%;'><td style='max-width: 350px'>Order Total</td><td></td><td></td><td style='padding:0 15px 0 5px;'><b>$${row[numColumns-2]}</b></td></tr>`
    message += `<tr style='padding:5px; background:#ddd;'><td style='max-width: 350px'>${header_row[numColumns-(4+hasCoupons)]}</td><td></td><td><b>${row[numColumns-(4+hasCoupons)]}</b></td><td></td></tr>`
    message += `</table>`
    message += `<p style="margin-bottom: 0;">This order can be changed by clicking on the button below until <b>${closeDate}</b>.<br>Please DO NOT create a second order.</p><small>To cancel this order edit the order and change the quantity of all items ordered to zero</small>`
    message += `<a style="background:darkblue; border-radius: 10px; padding: 15px 0; font-size:120%; display: block; width: 80%; margin: 20px 10%; text-align:center; color: #fff; border: 1px solid #000" href="${row[numColumns-1]}">EDIT MY ORDER</a>`
    message += `<p>PAYMENT MUST BE RECEIVED BY <b>${paymentDate}</b> Payment options include cash, check, Zelle or credit card.</p>`
    message += '<p>If you pay with <b>ZELLE</b> there are no additional fees<br>Zelle payments can be sent to <b>kemach@matanbsayser.org</b><br>Please include your order number in the memo to ensure we can credit you for your payment</p>'
    message += 'We also accept credit card payments for your Kemach order! There is an additional 3% charge to cover the credit card processing fee if you choose to pay with credit card.'
    const ccFee = Math.floor(row[numColumns-2]*0.03)
    message += `<table style='width: 100%; font-size:calc(0.6vw + 9px); border-spacing: 0;'><tr style='background:#ddd;'><td>Three percent additional charge <b>ONLY</b> IF PAYING WITH CREDIT CARD</td><td><b>$${ccFee}</b></td></tr>`
    message += `<tr style='padding: 5px;'><td>Order Total For Credit Card payment ONLY</td><td style='width 105px; text-align: center'><b>$${row[numColumns-2] + ccFee}</b></td></tr></table>`
    message += `<a style="background:red; padding:15px 0; font-size:150%; border-radius: 10px; display: block; width: 60%; margin: 20px 20%; text-align:center; color: #fff; border: 1px solid #000" href="matanbsayser.org/Kemach?amnt=${row[numColumns-2] + ccFee}&orderid=${row[numColumns-(3+hasCoupons)]}">PAY ONLINE</a>`
    message += `<p>If you are paying with cash or check please make sure you submit your payment to 1928 Janette Avenue Cleveland, Ohio 44118 before ${paymentDate}. When you drop off the payment PLEASE make sure it is in a sealed envelope with your name and address clearly written on it.</p>`
    message += `DISTRIBUTION IS SLATED FOR <span style='background:yellow;'>${pickupDate}.</span> at Hillcrest Foods 2735 East 40th, Cleveland, Ohio 44115<br>Important: Please enter through the back road.<br>`
    message += 'In order for the entire project to function successfully, we need manpower on that day to assist us with the distribution process. Please let us know if you are able to volunteer by either calling 440-7KEMACH (753-6224) or email kemachcleveland@gmail.com'
    message += '<p>If you have any questions please call the Kemach office 440-7KEMACH (753-6224) and leave a message.</p></div>'
  } else {
    const validOrder = data.find(r => r[1] === row[1] && r[numColumns - 2])
    message += "<tr style='background:yellow; padding:5px; font-size:200%;'><td style='max-width: 350px'>Order Canceled</td><td></td><td></td><td style='padding:0 15px 0 5px;'><b>$0</b></td></tr></table>"
    if (!validOrder) {
      message += `<a style="background:darkblue; border-radius: 10px; padding: 15px 0; font-size:120%; display: block; width: 80%; margin: 20px 10%; text-align:center; color: #fff; border: 1px solid #000" href="${row[numColumns-1]}">RECREATE ORDER</a>`
    } else {
      message += '<p><span style="background: #e1e100;">*IMPORTANT*</span> This order was cancelled, however our records show that there is another order for this email address that order can be edited via the button below</p>'
      message += `<a style="background:darkblue; border-radius: 10px; padding: 15px 0; font-size:120%; display: block; width: 80%; margin: 20px 10%; text-align:center; color: #fff; border: 1px solid #000" href="${validOrder[numColumns-1]}">EDIT ORDER ${yt[0]}${year}-${validOrder[numColumns - (3+hasCoupons)]}</a>`
    }
  }
  return { subject, message }
}

function getEditLink(formUrl, timeStamp='', semail = '') {
  let form
  try {
    form = FormApp.openByUrl(formUrl)
  } catch {
    console.log('form unavailble')
    Utilities.sleep(60*1000)
    console.log('retry form')
    form = FormApp.openByUrl(formUrl)
  }
  let lasturl = ''
  const responses = form.getResponses()
  for (const response of responses) {
    if (timeStamp) {
      rts = response.getTimestamp()
      if (timeStamp == rts.toString()) {
        lasturl = response.getEditResponseUrl()
      }
    } else {
      const femail = response.getRespondentEmail()
      if (femail == semail) {
        lasturl = response.getEditResponseUrl()
      }
    }
  }
  return lasturl
}

function processOrderData(startRow = 4) {
  const timeStamp = (new Date()).toISOString().split('.')[0].replace(/-|:/g,"_")
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = ss.getSheetByName(orderSheet)
  const cols = sheet.getLastColumn()
  const numRows = sheet.getLastRow()
  const data = sheet.getRange(1, 1, numRows, cols).getValues()
  const recorded = []
  const phoneCols = [10, 11, 12]
  let result = ''
  for (let row of data.slice(startRow)) {
    if (row[cols - (4 + hasCoupons)] > 0) {
      let phone = ''
      for (let i = 0; i < cols; i++) {
        let val = row[1] ? String(row[i]).toUpperCase().trim() : ''
        if (i === 2) {rowStr = val.substring(0, 20).padEnd(20)}
        else if (i === 4 || i === 6) {rowStr += val.substring(0, 10).padEnd(10)}
        else if (phoneCols.includes(i)) {
          phone = !val ? phone : val
          if (i === phoneCols.at(-1)) {rowStr += phone.padStart(10, '9')}
        }
        else if ([stayingHomeCol, 'Volunteer'].includes(data[1][i])) {rowStr += val === 'YES' ? 'Y' : 'N'}
        else if (data[1][i] === 'Order ID') {rowStr += val.substring(0, 3).padStart(3, '0')}
        else if (data[1][i] === 'total') {rowStr += val.substring(0, 4).padStart(4, '0')}
        else if (data[2][i]) {rowStr += val.substring(0, 2).padStart(2, '0')}
        else if (hasCoupons) {
          if (data[1][i] === couponCodeCol) {rowStr += val.substring(0, 2).padStart(2, 'X')}
          // else if (data[1][i] === 'Coupon discount') {rowStr += val.substr(-3).padStart(3, '0')}
          else if (data[1][i] === 'Produce package') {rowStr += val === 'YES' ? 'Y' : 'N'}
        }
      }
      if (recorded.includes(row[1])) {
        console.log(`possible duplicate order for ${row[1]}`)
      }
      recorded.push(row[1])
      result += rowStr + "\n"
    }
  }
  const folder = DriveApp.getFileById(ss.getId()).getParents().next()
  console.log(`creating fixed length text file orderResults_${timeStamp}.txt with headers: last name (20) first (10) wife (10) phone (10) staying (1) volenteer (1) products (2 each) num itms orders (3) order ID (3) total (4)`)
  folder.createFile(`orderResults_${timeStamp}.txt`, result)
}

function processPaymentData(startRow = 2) {
  const timeStamp = (new Date()).toISOString().split('.')[0].replace(/-|:/g,"_")
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = ss.getSheetByName(paymentsSheet)
  const cols = sheet.getLastColumn()
  const numRows = sheet.getLastRow()
  const data = sheet.getRange(startRow, 1, numRows, cols).getValues()
  let result = ''
  for (let row of data) {
    let adj = '0000'
    if (row[3] > 0) { // } && !['total', 'count'].some(w => row[0].indexOf(w) + 1)) {
      let rowStr = row[0] ? String(row[0]).substring(0, 3).padStart(3, '0') : 'XXX'
      rowStr += row[13].substring(0, 20).toUpperCase().padEnd(20)
      rowStr += row[14].substring(0, 10).toUpperCase().padEnd(10)
      rowStr += row[15].substring(0, 10).toUpperCase().padEnd(10)
      for (let i = 5; i < cols; i++) {
        if (i < 9 || i > 16) {rowStr += row[i] ? String(parseInt(row[i])).substring(0, 4).padStart(4, '0') : '0000'}
        if (i === 9) {adj = parseInt(row[i]) < -1 ? String(Math.abs(parseInt(row[i]))).substring(0, 4).padStart(4, '0') : '0000'}
      }
      result += rowStr + adj + "\n"
    }
  }
  const folder = DriveApp.getFileById(ss.getId()).getParents().next()
  console.log(`creating fixed length text file paymentResults_${timeStamp}.txt with headers: order ID (3) last name (20) first (10) wife (10) phone (10) payments 8 cols CC,	Zelle,	Check,	Cash. twice (4 each)`)
  folder.createFile(`paymentResults_${timeStamp}.txt`, result)
}

function processOrderAndPaymentData(startRow = 4) {
  const timeStamp = (new Date()).toISOString().split('.')[0].replace(/-|:/g,"_")
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = ss.getSheetByName(orderSheet)
  const cols = sheet.getLastColumn()
  const numRows = sheet.getLastRow()
  const data = sheet.getRange(1, 1, numRows, cols).getValues()
  const payments = ss.getSheetByName(paymentsSheet)
  const paymentCols = payments.getLastColumn()
  const paymentRows = payments.getLastRow()
  const paymentData = payments.getRange(1, 1, paymentRows, paymentCols).getValues()
  const recorded = []
  let result = ''
  for (let row of data.slice(startRow)) {
    if (row[cols - (4 + hasCoupons)] > 0) {
      let phone = ''
      for (let i = 0; i < cols; i++) {
        let val = row[1] ? String(row[i]).toUpperCase().trim() : ''
        if (i === 2) {rowStr = val.substring(0, 20).padEnd(20)}
        else if (i === 4 || i === 6) {rowStr += val.substring(0, 10).padEnd(10)}
        else if (i > 10 && i < 14) {
          phone = phone ? phone : val
          if (i === 13) {rowStr += phone.padStart(10, '9')}
        }
        else if ([data[1][i], data[1][i-1]].includes(stayingHomeCol)) {rowStr += val[0]}
        else if ([data[1][i], data[1][i+1]].includes('Order ID')) {rowStr += val.substring(0, 3).padStart(3, '0')}
        else if (data[1][i] === 'total') {rowStr += val.substring(0, 4).padStart(4, '0')}
        else if (data[2][i]) {rowStr += val.substring(0, 2).padStart(2, '0')}
        else if (hasCoupons) {
          if (data[1][i] === couponCodeCol) {rowStr += val.substring(0, 2).padStart(2, 'X')}
          else if (i === cols - 4) {rowStr += val[0]}
        }
      }
      const paymentRow = paymentData.find(p => p[0] === row[cols - (3 + hasCoupons)])
      if (paymentRow) {
        for (let p = 5; p < paymentCols; p++) {
          if (p < 9 || p > 16) {rowStr += paymentRow[p] ? String(parseInt(paymentRow[p])).substring(0, 4).padStart(4, '0') : '0000'}
        }
      } else {
        rowStr += '0'.repeat(32)
        console.log(`no payment record found for order #${row[cols - (3 + hasCoupons)]}`)
      }
      if (recorded.includes(row[1])) {
        console.log(`possible duplicate order for ${row[1]}`)
      }
      recorded.push(row[1])
      result += rowStr + "\n"
    }
  }
  const folder = DriveApp.getFileById(ss.getId()).getParents().next()
  console.log(`creating fixed length text file results_${timeStamp}.txt with headers: last name (20) first (10) wife (10) phone (10) staying (1) volunteer (1) products (2 each) num items ordered (3) order ID (3) total (4) payments 8 cols CC,	Zelle,	Check,	Cash. twice (4 each)`)
  folder.createFile(`results_${timeStamp}.txt`, result)
}
