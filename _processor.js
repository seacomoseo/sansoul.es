/* global
  SpreadsheetApp
  Utilities
  Logger
  MailApp
  DriveApp
  ContentService
*/

const now = Utilities.formatDate(new Date(), 'Europe/Madrid', 'yyyy-MM-dd-HH-mm-ss')
const ss = SpreadsheetApp.getActiveSpreadsheet()
let domain = 'try'
let sheet = ss.getActiveSheet()
let sheetName = sheet.getSheetName()

// Anti-spam: silent success response (bots think it worked)
function spamResponse () {
  return response(200, 'success', 'Data processed successfully')
}

// Anti-spam: check all layers
function isSpamSubmission (params) {
  // 1. Honeypot: if _gotcha has content, it's a bot
  if (params._gotcha) return true

  // 2. JS token validation (only if present — backward compat)
  if (params._token) {
    try {
      const decoded = Utilities.newBlob(Utilities.base64Decode(params._token)).getDataAsString()
      const parts = decoded.split(':')
      if (parts.length !== 2) return true
      const tokenMinute = parseInt(parts[1])
      const nowMinute = Math.floor(new Date().getTime() / 60000)
      const diff = nowMinute - tokenMinute
      if (diff > 30 || diff < -2) return true // expired or future
    } catch (_) {
      return true // malformed token
    }
  }

  // 3. Content spam detection
  if (isSpamContent(params)) return true

  // 4. Email format validation
  const email = params.Email || params.email || params.Mail || params.mail
  if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return true

  return false
}

// Anti-spam: content analysis
function isSpamContent (params) {
  const skipKeys = ['_headers', '_subject', '_id', '_domain', '_sheetname', '_gotcha', '_token', 'CC', 'URL', 'Timestamp']
  const allValues = Object.keys(params)
    .filter(k => !skipKeys.includes(k))
    .map(k => params[k] || '')
    .join(' ')

  if (!allValues.trim()) return false

  // Too many URLs
  const urlCount = (allValues.match(/https?:\/\//gi) || []).length
  if (urlCount > 2) return true

  // Cyrillic or Chinese characters (uncommon in Spanish forms)
  if (/[\u0400-\u04FF\u4E00-\u9FFF]/.test(allValues)) return true

  // Common spam keywords
  const spamWords = ['viagra', 'cialis', 'casino', 'poker', 'lottery', 'winner', 'click here', 'buy now', 'free money', 'seo service', 'web traffic', 'backlink']
  const lower = allValues.toLowerCase()
  if (spamWords.some(w => lower.includes(w))) return true

  return false
}

// Anti-spam: rate limiting per form (uses PropertiesService)
function isRateLimited (sheetName, maxPerHour) {
  maxPerHour = maxPerHour || 20
  try {
    const props = PropertiesService.getScriptProperties()
    const hour = new Date().getHours()
    const key = 'rate_' + sheetName + '_' + hour
    const count = parseInt(props.getProperty(key) || '0')
    if (count >= maxPerHour) return true
    props.setProperty(key, String(count + 1))
    // Clean previous hour key
    const prevHour = (hour + 23) % 24
    props.deleteProperty('rate_' + sheetName + '_' + prevHour)
  } catch (_) { /* PropertiesService may fail in some contexts */ }
  return false
}

// Main processor function
// eslint-disable-next-line
function processor (e) {
  try {
    const params = e.parameter
    const subject = params._subject || 'Form Submission'
    const formId = params._id
    const headers = JSON.parse(params._headers || '[]')
    domain = params._domain

    if (!domain || !formId) {
      throw new Error('Missing domain or form ID')
    }

    // Anti-spam checks
    if (isSpamSubmission(params)) return spamResponse()

    // Determine sheet name (needed for rate limiting)
    sheetName = params._sheetname || `${domain}#${formId}`
    if (isRateLimited(sheetName)) return spamResponse()

    sheet = getOrCreateSheet(sheetName) || sheet

    logToSheet(sheet, JSON.stringify(e.parameters))

    // Update sheet headers
    const configHeaders = headers.length ? headers : Object.keys(params).map(key => ({ name: key }))
    const sheetHeaders = updateSheetHeaders(sheet, configHeaders)

    // Append values to sheet
    const values = appendDataToSheet(sheet, e, sheetHeaders, domain, formId)

    // Send emails
    sendEmails(domain, subject, values, sheet, e.parameter, sheetHeaders)

    return response(200, 'success', 'Data processed successfully')
  } catch (error) {
    const logSheet = logToSheet(sheet, `Error processing form data: ${error}\nRemaining email quota: ${MailApp.getRemainingDailyQuota() - 1}`)

    MailApp.sendEmail({
      bcc: 'lorensansol@gmail.com',
      subject: `Error in form ${sheetName}`,
      htmlBody:
        `<p>There was an error processing the form <a href="${ss.getUrl()}#gid=${sheet.getSheetId()}">${sheetName}</a>:</p>` +
        `<p><strong><code>${error}</code></strong></p>` +
        `<p><a href="${ss.getUrl()}#gid=${logSheet.getSheetId()}">View logs</a></p>`,
      name: domain
    })

    Logger.log(`Error processing form data: ${error}`)

    return response(400, 'error', error.toString())
  }
}

function getOrCreateSheet (sheetName) {
  if (!sheetName) throw new Error('Sheet name is empty')
  let sheet = ss.getSheetByName(sheetName)
  if (!sheet) sheet = ss.insertSheet(sheetName)
  return sheet
}

function updateSheetHeaders (sheet, configHeaders) {
  const lastColumn = sheet.getLastColumn()
  let existingHeaders = []
  if (lastColumn > 0) {
    existingHeaders = sheet.getRange(1, 1, 1, lastColumn).getValues()[0]
  }

  // Filter out the headers that are not already in the sheet
  const missingHeaders = configHeaders
    .filter(field => !existingHeaders.includes(field.name))
    .map(field => field.name)

  // If there are any missing headers, add them after the last column
  if (missingHeaders.length > 0) {
    sheet.getRange(1, lastColumn + 1, 1, missingHeaders.length)
      .setValues([missingHeaders])
  }

  const sheetHeaders = [...existingHeaders, ...missingHeaders]
    .map(name => {
      const type = configHeaders.find(item => item.name === name)?.type
      return { name, type }
    })
  return sheetHeaders
}

function appendDataToSheet (sheet, e, headers, domain, formId) {
  const cellsArray = []
  const cells = headers.map(({ name, type }) => {
    const params = e.parameters[name] || []
    const valuesArray = []
    const values = params.map(param => {
      let value, valueArray
      if (param && typeof param === 'string' && type === 'file') {
        if (param === 'null') {
          valueArray = 'null'
          value = valueArray
        } else {
          valueArray = processBase64File(param, domain, formId, sheet)
          value = valueArray.thumbnail || valueArray.view
        }
      } else if (param && typeof param === 'string' && param.startsWith('+')) {
        value = `'${param}`
      } else {
        value = param || ''
      }
      valuesArray.push(valueArray || value)
      return value
    })
    cellsArray.push(valuesArray.length ? valuesArray : [''])
    return values.join(', ')
  })
  sheet.appendRow([...cells])
  return cellsArray
}

function processBase64File (param, domain, formId, sheet) {
  try {
    logToSheet(sheet, param)
    const folder = getOrCreateFolder(domain, formId)
    const match = param.match(/([^:;,|])+/g)
    const mimeType = match[1]
    const base64 = match[3]
    const fileName = match[4] || now
    if (!mimeType) throw new Error('MIME type not found in base64 string')
    const decodedBytes = Utilities.base64Decode(base64)
    const blob = Utilities.newBlob(decodedBytes, mimeType, fileName)
    const file = folder.createFile(blob)
    const id = file.getId()
    const view = `https://drive.google.com/file/d/${id}/view`
    const mimeTypeRegex = /^(image\/|video\/|application\/(pdf|vnd\.google-apps|vnd\.openxmlformats-officedocument|(vnd\.)?msword|vnd\.ms-excel|vnd\.ms-powerpoint|vnd\.oasis\.opendocument))/
    const thumbnail = mimeTypeRegex.test(mimeType) ? `https://lh3.googleusercontent.com/d/${id}` : ''
    logToSheet(sheet, `=HYPERLINK("${view}"; "${fileName}")`)
    return { mimeType, thumbnail, view, fileName }
  } catch (error) {
    logToSheet(sheet, `Error processing base64 file: ${error.message}`)
    return ''
  }
}

function getOrCreateFolder (domain, formId) {
  const parentFolder = DriveApp.getFileById(ss.getId()).getParents().next()
  const mainFolderName = ss.getName()
  let mainFolder
  const mainFolderIterator = parentFolder.getFoldersByName(mainFolderName)
  if (mainFolderIterator.hasNext()) {
    mainFolder = mainFolderIterator.next()
  } else {
    mainFolder = parentFolder.createFolder(mainFolderName)
  }
  const folderName = `${domain}/${formId}`
  let folder
  const folderIterator = mainFolder.getFoldersByName(folderName)
  if (folderIterator.hasNext()) {
    folder = folderIterator.next()
  } else {
    folder = mainFolder.createFolder(folderName)
  }
  return folder
}

function logToSheet (sheet, message) {
  let logSheet = ss.getSheetByName('logs')
  if (!logSheet) {
    logSheet = ss.insertSheet('logs')
    logSheet.appendRow(['Timestamp', 'Form', 'Message'])
  }
  const date = Utilities.formatDate(new Date(), 'Europe/Madrid', 'yyyy-MM-dd HH:mm:ss')
  const sheetInfo = `=HYPERLINK("#gid=${sheet.getSheetId()}"; "${sheet.getSheetName()}")`
  logSheet.appendRow([date, sheetInfo, message])
  return logSheet
}

// Send emails
function sendEmails (domain, subject, values, sheet, params, headers) {
  const bcc = Array.isArray(params.CC)
    ? params.CC.join(',')
    : (typeof params.CC === 'string' ? params.CC : '')
  if (!bcc) return
  const replyTo = [params.Email ?? [], params.email ?? [], params.Mail ?? [], params.mail ?? []].flat().filter(arr => arr).join(',')
  const htmlBody =
    '<table><tbody style="vertical-align:top">' +
    headers.map(({ name, type }, i) => {
      if (!name || name === 'CC') return ''
      const valuesArray = values[i] || []
      return valuesArray.map((value, j) => {
        if (type === 'file' && typeof value === 'object') {
          const file = value.thumbnail ? `<img src="${value.thumbnail}" width="320">` : value.fileName
          value = `<a href="${value.view}">${file}</a>`
        } else if (typeof value === 'string' && value.includes('\n')) {
          value = `<pre style="font-family:inherit;margin:0">  ${value.replace(/\n/gm, '\n  ')}</pre>`
        }
        const index = valuesArray.length > 1 ? ` ${j + 1}` : ''
        const row =
          '<tr>' +
            '<td>' +
              '<span style="opacity:0;font-size:.01px">- **</span>' +
              `<strong>${name}${index}:</strong>` +
              '<span style="opacity:0">**</span>' +
            '</td>' +
            `<td>${value}</td>` +
          '</tr>'
        return row
      }).join('')
    }).join('') +
    '</tbody></table>'
  MailApp.sendEmail({
    bcc,
    replyTo,
    subject,
    htmlBody,
    name: domain
  })
  logToSheet(sheet, `Remaining email quota: ${MailApp.getRemainingDailyQuota()}`)
}

function response (statusCode, result, message) {
  return ContentService
    .createTextOutput(JSON.stringify({ result, message }))
    .setMimeType(ContentService.MimeType.JSON)
    // .setHeader('Access-Control-Allow-Origin', '*')
    // .setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
    // .setHeader('Access-Control-Allow-Headers', 'Content-Type')
    // .setResponseCode(statusCode)
}
