const fs = require('fs')
const path = require('path')
const xlsx = require('node-xlsx')

const outFile = 'template.xlsx' // 模版文件名字
const loopPath = '/Users/admin/Desktop/leilei' // 处理的文件集合
const toPath = './template'

let needColumn = 11
let fromNameRow = 1

// 从单人的传的内容解析excel表格
async function parseFromExcel(file, fPath) {
    const filePath = path.join(fPath, file)
    const workSheetsFromBuffer = xlsx.parse(fs.readFileSync(filePath));
    // 去最后一个sheet
    const work = workSheetsFromBuffer.length <= 5 ? (workSheetsFromBuffer[2] ? workSheetsFromBuffer[2].data : workSheetsFromBuffer[0].data) : workSheetsFromBuffer[workSheetsFromBuffer.length-1].data
    if (work[fromNameRow] === undefined) {
        return null
    }
    let result = []
    let name = ''
    for (let i = 0; i < 15; i++) {
        if (i === fromNameRow) {
            name = work[i][0]
        }
        if (i<4) {
            continue;
        }
        result.push(work[i][needColumn])
    }
    return {
        name: name.split(/：|:/)[1] && name.split(/：|:/)[1].trim() || name.split(/：|:/)[0].trim(),
        data: result
     }
}

let nameRow = 3
let nameLine = 4
async function getSumNameAndRowMap(file) {
    // 获取模版的地址
    const filePath = path.join(__dirname, toPath, file)        
    const workSheetsFromBuffer = xlsx.parse(fs.readFileSync(filePath));
    let sheetIndex=0
    let nameLineAndSheetMap = {}
    for (let sheet of workSheetsFromBuffer) {
        // 名字与列数的map
        let rowNameValue = {}
        let sheetData = sheet.data
        for (let i = 0; i < 15; i++) {
            // 名字在第4行
            if (i === nameRow) {
                // 名字从第五列开始
                let nameRowContent = sheetData[i]
                for (let j = 0; j < nameRowContent.length; j++) {
                    if (j< nameLine) {
                        continue
                    } else {
                        let userName = nameRowContent[j] && nameRowContent[j].trim() || ''
                        nameLineAndSheetMap[userName] = { sheetIndex, column: j, row: nameRow,  }
                        // rowNameValue[nameRowContent[j]] = j
                    }
                }
            } else {
                if (i > nameRow) {
                    break
                }
                continue
            }
        }
        sheetIndex ++
    }
    return nameLineAndSheetMap   
}

async function getSumContent(file) {
    const filePath = path.join(__dirname, toPath, file)
    return xlsx.parse(fs.readFileSync(filePath));
}

async function main(fPath) {
    let totalMap = await getSumNameAndRowMap(outFile)
    let content = await getSumContent(outFile)

    const files = await fs.readdirSync(fPath);
    for (let file of files) {
        if (path.extname(file) !== '.xlsx' && path.extname(file) !== '.xls') {
            continue
        } 
        let parseResult = await parseFromExcel(file, fPath)
        if (!parseResult) { continue }
        const { name, data } = parseResult
        // 根据名称取出对应映射中的信息
        if (!totalMap[name]) {
            continue
        }
        let { sheetIndex, column, row } = totalMap[name]
        let index = row
        for (let d of data) {
            index++
            let targetRow = content[sheetIndex].data[index]
            content[sheetIndex].data[index][column] = d
        }
    }
    let buffer = await xlsx.build(content)
    let fileName = path.basename(fPath)
    fs.writeFileSync(`./result/${fileName}.xlsx`, buffer, {
        'flag': 'w'
    })

    console.log(`${fileName} ok!`)
}

async function loopGenMonth() {
    const files = await fs.readdirSync(loopPath);
    for (let file of files) {
        if (path.extname(file) !== '' || file === '.DS_Store') {
            continue
        }
        const fPath = path.join(loopPath, file)
        await main(fPath).catch(console.log)
    }
    
}

loopGenMonth().catch(console.log)